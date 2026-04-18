"""High-level orchestration: ASC + HAD + filename → Measurement object."""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from pathlib import Path

from .parser import (
    ASCData, HADMetadata, FilenameMetadata,
    parse_asc, parse_had, parse_filename, to_serializable,
)
from .analysis import DLEPRResult, analyze_dlepr, SplitMethod


@dataclass
class Measurement:
    """Complete measurement: raw data + metadata + analysis results."""
    source_file: str
    asc: ASCData
    had: HADMetadata
    filename_meta: FilenameMetadata
    analysis: DLEPRResult

    def to_summary_dict(self) -> dict:
        """Flat dict for tabular output (Messübersicht row)."""
        return {
            "Messung": self.filename_meta.messung_id,
            "Material": self.filename_meta.material,
            "Materialzustand": None,
            "Probenbez": self.filename_meta.probenbez,
            "Messfläche [mm²]": self.had.probenflaeche_mm2,
            "Temperatur [°C]": self.filename_meta.temperature_c,
            "ProbeTemperatur [°C]": self.filename_meta.probe_temperature_c,
            "Ruhepotential [mV]": round(self.analysis.ruhepotential_mv, 2)
                if self.analysis.ruhepotential_mv == self.analysis.ruhepotential_mv
                else None,  # NaN check
            "Ja [mA/cm²]": round(self.analysis.ja_ma_cm2, 5),
            "Jr [mA/cm²]": round(self.analysis.jr_ma_cm2, 5),
            "Jr/Ja": round(self.analysis.jr_ja, 6),
            "Qa [As]": round(self.analysis.qa_as, 8),
            "Qr [As]": round(self.analysis.qr_as, 8),
            "Qr/Qa": round(self.analysis.qr_qa, 6),
            "Prüfer": self.had.sachbearbeiter,
            "Datum": self.had.erstellungsdatum,
            "Uhrzeit": self.had.erstellungszeit,
            "Split-Index": self.analysis.split_index,
            "Split-Methode": self.analysis.split_method,
            "Bemerkung": self.filename_meta.notes,
            "Datei": Path(self.source_file).name,
        }

    def to_json_dict(self) -> dict:
        """Full serializable dict including raw curve (for JSON export)."""
        return to_serializable({
            "source_file": self.source_file,
            "asc": self.asc,
            "had": self.had,
            "filename_meta": self.filename_meta,
            "analysis": self.analysis,
        })


def process_measurement(
    asc_path: Path | str,
    had_path: Path | str | None = None,
    split_method: SplitMethod = "vertex",
    split_override: int | None = None,
) -> Measurement:
    """Parse ASC + its HAD sibling + filename and run DL-EPR analysis.

    If `had_path` is None, looks for a file with the same stem but .HAD
    extension next to the ASC file.
    """
    asc_path = Path(asc_path)
    if had_path is None:
        candidate = asc_path.with_suffix(".HAD")
        if not candidate.exists():
            candidate = asc_path.with_suffix(".had")
        had_path = candidate if candidate.exists() else None

    asc = parse_asc(asc_path)
    had = parse_had(had_path) if had_path and Path(had_path).exists() else HADMetadata()
    fm = parse_filename(asc_path.name)
    result = analyze_dlepr(
        asc, had,
        split_method=split_method,
        split_override=split_override,
    )
    return Measurement(
        source_file=str(asc_path),
        asc=asc,
        had=had,
        filename_meta=fm,
        analysis=result,
    )
