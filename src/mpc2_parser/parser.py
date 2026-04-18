"""ASC and HAD file parsers + filename metadata extractor.

ASC format (IPS PgU Touch potentiostat):
    4 whitespace-separated columns, scientific notation, dot decimal:
    Time/s   Potential/V   I/A   S/[A/m²]

HAD format:
    Key : Value lines with metadata (Probenflaeche, Anzahl Werte RP, ...)
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

import numpy as np


# --------------------------------------------------------------------------- #
# Dataclasses                                                                 #
# --------------------------------------------------------------------------- #

@dataclass
class ASCData:
    """Raw curve data from an .ASC file."""
    time_s: np.ndarray           # seconds
    potential_v: np.ndarray      # volts
    current_a: np.ndarray        # amperes
    current_density_am2: np.ndarray  # A/m²

    @property
    def n_points(self) -> int:
        return len(self.time_s)

    @property
    def potential_mv(self) -> np.ndarray:
        return self.potential_v * 1000.0

    @property
    def current_density_macm2(self) -> np.ndarray:
        """J in mA/cm² = (A/m²) / 10."""
        return self.current_density_am2 / 10.0


@dataclass
class HADMetadata:
    """Metadata from an .HAD file."""
    erstellungsdatum: str | None = None
    erstellungszeit: str | None = None
    sachbearbeiter: str | None = None
    anzahl_werte_rp: int | None = None        # rest-potential / warmup points
    anzahl_werte: int | None = None            # total points
    probenflaeche_mm2: float | None = None     # sample area
    i_bereich: str | None = None               # current range label
    s_bereich: str | None = None               # current density range label
    kanalzahl: int | None = None
    kommentar: str | None = None
    raw: dict[str, str] = field(default_factory=dict)


@dataclass
class FilenameMetadata:
    """Metadata extracted from the filename (best-effort)."""
    messung_id: str | None = None
    material: str | None = None
    probenbez: str | None = None
    temperature_c: float | None = None        # bath temperature (Prüflösung)
    probe_temperature_c: float | None = None  # sample temperature if present
    activation_mv: float | None = None
    activation_s: float | None = None
    notes: str | None = None
    raw_stem: str = ""


# --------------------------------------------------------------------------- #
# ASC parser                                                                  #
# --------------------------------------------------------------------------- #

def parse_asc(path: Path | str) -> ASCData:
    """Parse an IPS PgU Touch .ASC file.

    Data is whitespace-separated with scientific-notation floats using '.' as
    the decimal separator (already Python-native, no locale fixup needed).
    """
    path = Path(path)
    arr = np.loadtxt(path, dtype=float)
    if arr.ndim == 1:
        arr = arr.reshape(1, -1)
    if arr.shape[1] < 4:
        raise ValueError(f"{path.name}: expected 4+ columns, got {arr.shape[1]}")
    return ASCData(
        time_s=arr[:, 0],
        potential_v=arr[:, 1],
        current_a=arr[:, 2],
        current_density_am2=arr[:, 3],
    )


# --------------------------------------------------------------------------- #
# HAD parser                                                                  #
# --------------------------------------------------------------------------- #

_HAD_INT_KEYS = {"anzahl_werte_rp", "anzahl_werte", "kanalzahl"}
_HAD_FLOAT_KEYS = {"probenflaeche_mm2"}

# German umlauts may be mangled in the HAD (Probenfl?che). Use fuzzy matching.
_HAD_KEY_MAP = {
    "erstellungsdatum": "erstellungsdatum",
    "erstellungszeit": "erstellungszeit",
    "sachbearbeiter": "sachbearbeiter",
    "anzahl werte rp": "anzahl_werte_rp",
    "anzahl werte": "anzahl_werte",
    # The HAD encodes German umlauts in a way that often mangles to "?" -
    # after normalization "Probenfläche/mm²" becomes "probenflche/mm" or
    # similar. We just check that both "probenfl" and "mm" appear.
    "probenfl": "probenflaeche_mm2",
    "i-bereich": "i_bereich",
    "s-bereich": "s_bereich",
    "kanalzahl": "kanalzahl",
    "kommentar": "kommentar",
}


def _normalize_had_key(key: str) -> str:
    """Strip unit suffixes and mangled bytes, lowercase, collapse spaces."""
    k = key.strip().lower()
    # Replace any non-ASCII (e.g. mangled ² or ä) with empty string
    k = re.sub(r"[^a-z0-9/\- ]", "", k)
    k = re.sub(r"\s+", " ", k).strip()
    return k


def parse_had(path: Path | str) -> HADMetadata:
    """Parse an IPS .HAD metadata file."""
    path = Path(path)
    raw: dict[str, str] = {}
    # HAD files are often latin-1 encoded due to German umlauts
    text = path.read_text(encoding="latin-1")
    for line in text.splitlines():
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        key = _normalize_had_key(key)
        val = val.strip()
        if not key or val in ("(null)", ""):
            continue
        raw[key] = val

    md = HADMetadata(raw=raw)
    for raw_key, field_name in _HAD_KEY_MAP.items():
        for k, v in raw.items():
            if k.startswith(raw_key):
                if field_name in _HAD_INT_KEYS:
                    try:
                        setattr(md, field_name, int(v))
                    except ValueError:
                        pass
                elif field_name in _HAD_FLOAT_KEYS:
                    try:
                        # German decimal comma -> dot
                        setattr(md, field_name, float(v.replace(",", ".")))
                    except ValueError:
                        pass
                else:
                    setattr(md, field_name, v)
                break
    return md


# --------------------------------------------------------------------------- #
# Filename parser                                                             #
# --------------------------------------------------------------------------- #

_TEMP_RE = re.compile(r"(\d{1,3}(?:[.,]\d+)?)\s*°?\s*C", re.IGNORECASE)
_ACTIV_RE = re.compile(r"-(\d{2,4})\s*mV?[^\d]*(\d{1,3})\s*s", re.IGNORECASE)


def parse_filename(filename: str) -> FilenameMetadata:
    """Extract what metadata we can from the filename.

    Filenames follow patterns like:
        0048_CR_S32906_W6-1_on weld_40°C_Probe 33,2°C_7mm_elektrochem aktiviert -450mV_45s.ASC
        0399_3D-Druck III_K2-1_146-120-1_30C_400mV.ASC
    """
    stem = Path(filename).stem
    md = FilenameMetadata(raw_stem=stem)

    parts = stem.split("_")
    if parts:
        # measurement id: leading numeric token, maybe with trailing x/-suffix
        m = re.match(r"^(\d{3,5}[a-z]?(?:-\d+)?)", parts[0])
        if m:
            md.messung_id = m.group(1)

    # Find material code. Patterns:
    #   S32906, S31803 (UNS designations)
    #   1.4404, 1.4466 (EN werkstoffnummer)
    #   Also capture project/batch names like "3D-Druck III", "CR" etc. as fallback
    for p in parts:
        if re.match(r"^[A-Z]\d{4,6}$", p) or re.match(r"^\d\.\d{4}$", p):
            md.material = p
            break
    # Fallback: look for project identifier like "3D-Druck III"
    if not md.material:
        for p in parts:
            if re.match(r"^(3D-Druck|CR)\b", p, re.IGNORECASE):
                md.material = p.strip()
                break

    # Probenbezeichnung (K2-1, W6-1, P30, ...)
    for p in parts:
        if re.match(r"^[KWP]\d+-?\d*$", p):
            md.probenbez = p
            break

    # Temperatures - first °C is bath, optional second "Probe X°C" is sample
    temps = _TEMP_RE.findall(stem)
    if temps:
        try:
            md.temperature_c = float(temps[0].replace(",", "."))
        except ValueError:
            pass
    # Explicit "Probe XX°C" or "Probe XX,X°C"
    probe = re.search(r"Probe\s+(\d{1,3}[.,]?\d*)\s*°?\s*C", stem, re.IGNORECASE)
    if probe:
        try:
            md.probe_temperature_c = float(probe.group(1).replace(",", "."))
        except ValueError:
            pass

    # Activation pulse: -450mV_45s etc.
    act = _ACTIV_RE.search(stem)
    if act:
        try:
            md.activation_mv = -float(act.group(1))
            md.activation_s = float(act.group(2))
        except ValueError:
            pass

    # Trailing comment after '!' or 'Fehlmessung' marker
    if "Fehlmessung" in stem or "fehlmessung" in stem:
        md.notes = "Fehlmessung"

    return md


# --------------------------------------------------------------------------- #
# Serialization helpers                                                       #
# --------------------------------------------------------------------------- #

def to_serializable(obj: Any) -> Any:
    """Convert dataclasses / numpy arrays to JSON-friendly types."""
    if isinstance(obj, np.ndarray):
        return obj.tolist()
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        return float(obj)
    if hasattr(obj, "__dataclass_fields__"):
        return {k: to_serializable(v) for k, v in asdict(obj).items()}
    if isinstance(obj, dict):
        return {k: to_serializable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [to_serializable(v) for v in obj]
    return obj
