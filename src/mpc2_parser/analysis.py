"""DL-EPR analysis: split detection, peak detection, charge integration.

The DL-EPR test consists of:
  1. A brief rest-potential period ("Anzahl Werte RP") where the cell equilibrates
     at a cathodic protection potential. Current is typically ~0.
  2. A forward (anodic) sweep: potential ramps UP through the activation region.
     Maximum current density here = Ja (activation peak).
  3. A reverse (cathodic) sweep: potential ramps DOWN. Maximum current density
     here = Jr (reactivation peak), which indicates the Degree of Sensitization.

The "split point" is the index where sweep direction reverses.
"""
from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Literal

import numpy as np

from .parser import ASCData, HADMetadata


SplitMethod = Literal["vertex", "midpoint", "manual"]


# --------------------------------------------------------------------------- #
# Split-point detection                                                       #
# --------------------------------------------------------------------------- #

def detect_split_vertex(
    asc: ASCData,
    rp_skip: int = 0,
    smooth_window: int = 11,
) -> tuple[int, dict]:
    """Method A — Physics-based: find the actual potential reversal vertex.

    Pros:
      - Physically correct: this IS the definition of the DL-EPR "double loop".
      - Works across arbitrary sweep rates, durations, and data lengths.
      - Independent of Anzahl Werte RP — robust even if HAD is missing.
    Cons:
      - Sensitive to noise in the potential signal (mitigated by smoothing).
      - Can shift a few rows if vertex isn't a clean peak.

    Strategy: smooth the potential signal, take the argmax. The smoothing
    window handles measurement noise; the argmax is the sweep-direction vertex.

    Returns (split_index, diagnostic_dict).
    """
    pot = asc.potential_v[rp_skip:]
    n = len(pot)
    if n < 20:
        raise ValueError("Too few points to detect vertex")

    # Simple moving-average smoothing for noise robustness
    w = max(3, min(smooth_window, n // 10))
    if w % 2 == 0:
        w += 1  # odd window for symmetric centering
    kernel = np.ones(w) / w
    pot_smooth = np.convolve(pot, kernel, mode="same")

    vertex_local = int(np.argmax(pot_smooth))

    # If the maximum is at position 0, the potential never rose — no reversal detected.
    # This typically means the file is an activation-only sweep, not a full DL-EPR double loop.
    if vertex_local == 0:
        raise ValueError(
            "Kein Umkehrpunkt erkannt: Die Potenzialkurve steigt in dieser Datei nie an. "
            "Handelt es sich um eine vollständige DL-EPR-Doppelschleife (Aktivierungs- "
            "UND Reaktivierungs-Sweep)? Reine Aktivierungs-Dateien (z.B. '…aktivieren…') "
            "können nicht ausgewertet werden."
        )

    vertex = rp_skip + vertex_local

    diagnostics = {
        "method": "vertex",
        "rp_skip": rp_skip,
        "smooth_window": w,
        "vertex_potential_v": float(asc.potential_v[vertex]),
        "forward_duration_s": float(asc.time_s[vertex] - asc.time_s[rp_skip]),
        "reverse_duration_s": float(asc.time_s[-1] - asc.time_s[vertex]),
    }
    return vertex, diagnostics


def detect_split_midpoint(
    asc: ASCData,
    rp_skip: int = 0,
) -> tuple[int, dict]:
    """Method B — Heuristic: split at the midpoint after skipping RP.

    Pros:
      - Simple, deterministic, explainable in one sentence.
      - Fast, no signal processing needed.
      - Matches Manuel's historical pattern (split ≈ 1060-1088 for ~2000-row files).
    Cons:
      - Assumes forward and reverse sweeps have equal duration (usually true
        for DL-EPR but not guaranteed).
      - Doesn't use the actual curve — purely statistical.

    Returns (split_index, diagnostic_dict).
    """
    n = asc.n_points
    sweep_points = n - rp_skip
    split = rp_skip + sweep_points // 2
    diagnostics = {
        "method": "midpoint",
        "rp_skip": rp_skip,
        "total_points": n,
        "sweep_points": sweep_points,
    }
    return split, diagnostics


# --------------------------------------------------------------------------- #
# DL-EPR analysis                                                             #
# --------------------------------------------------------------------------- #

@dataclass
class DLEPRResult:
    """Computed DL-EPR values for a single measurement."""
    # Split metadata
    split_index: int
    split_method: str
    rp_skip: int
    split_diagnostics: dict

    # Peaks (current densities in mA/cm²)
    ja_ma_cm2: float         # activation peak
    jr_ma_cm2: float         # reactivation peak
    ja_index: int
    jr_index: int

    # Charges (As = Coulombs)
    qa_as: float             # sum of delta_t * I over activation segment
    qr_as: float             # same over reactivation segment

    # Ratios (dimensionless)
    jr_ja: float
    qr_qa: float

    # Rest potential (mV) — taken from beginning of forward sweep
    ruhepotential_mv: float

    # Row indices in Manuel's Excel convention (1-based, +1 for header row)
    # So row 2 in Excel == index 0 in numpy
    excel_jr_range: tuple[int, int]  # MAX(F{start}:F{end})
    excel_qa_range: tuple[int, int]  # SUM(I{start}:I{end})
    excel_qr_range: tuple[int, int]  # SUM(I{start}:I{end})

    def to_dict(self) -> dict:
        return asdict(self)


def find_reverse_endpoint(
    asc: ASCData,
    split_index: int,
    potential_tolerance_mv: float = 5.0,
) -> int:
    """DL-EPR convention: reverse sweep ends when potential returns to start.

    After the vertex, the reverse sweep ramps the potential back down. Once it
    crosses below the starting potential (rest potential), we're in the post-
    sweep recovery / over-cathodic region which should NOT be integrated into
    Qr. This matches Manuel's manual cutoff choices in his Auswertung sheets.

    Strategy: find the first index i > split where potential is within
    `potential_tolerance_mv` of (or below) the starting potential.
    """
    start_potential_v = asc.potential_v[0]
    tol_v = potential_tolerance_mv / 1000.0
    threshold = start_potential_v + tol_v  # cross-below threshold

    # Walk forward from split; find first crossing below threshold
    for i in range(split_index + 1, asc.n_points):
        if asc.potential_v[i] <= threshold:
            return i
    return asc.n_points - 1  # if we never cross, use the end


def analyze_dlepr(
    asc: ASCData,
    had: HADMetadata | None = None,
    split_method: SplitMethod = "vertex",
    split_override: int | None = None,
    ruhepotential_row: int = 87,
    truncate_reverse: bool = True,
) -> DLEPRResult:
    """Compute Ja, Jr, Qa, Qr, ratios from the raw curve.

    Parameters
    ----------
    asc : parsed ASC data
    had : parsed HAD metadata (kept for compatibility; not used for Qa skipping
          to match Manuel's convention of integrating from the very first row).
    split_method : "vertex" (physics-based) or "midpoint" (heuristic)
    split_override : manual split index (overrides split_method if given)
    ruhepotential_row : 1-based row index in Excel where Manuel reads
                        Ruhepotential (default 87 matches his template).
    truncate_reverse : if True, end Qr integration when potential returns to
                       the starting value (matches Manuel's convention).
    """
    # We keep rp_skip=0 to match Manuel's convention: Qa is integrated from the
    # very first data point. During the actual RP period current is ~0, so this
    # adds nothing to Qa and avoids inconsistency with Manuel's Excel formulas.
    rp_skip = 0

    if split_override is not None:
        split = split_override
        diag = {"method": "manual", "override": True}
    elif split_method == "vertex":
        split, diag = detect_split_vertex(asc, rp_skip=rp_skip)
    elif split_method == "midpoint":
        split, diag = detect_split_midpoint(asc, rp_skip=rp_skip)
    else:
        raise ValueError(f"Unknown split_method: {split_method}")

    j_macm2 = asc.current_density_macm2
    n = asc.n_points

    if split <= 1 or split >= n - 1:
        raise ValueError(f"Invalid split index {split} (n={n})")

    # Determine reverse-sweep endpoint
    if truncate_reverse:
        rev_end = find_reverse_endpoint(asc, split)
        diag["reverse_endpoint"] = rev_end
        diag["reverse_endpoint_reason"] = "potential_returned_to_start"
    else:
        rev_end = n
        diag["reverse_endpoint"] = rev_end
        diag["reverse_endpoint_reason"] = "end_of_file"

    fwd_slice = slice(0, split)
    rev_slice = slice(split, rev_end)

    ja_idx = int(np.argmax(j_macm2[fwd_slice]))
    ja = float(j_macm2[ja_idx])

    jr_idx_local = int(np.argmax(j_macm2[rev_slice]))
    jr_idx = split + jr_idx_local
    jr = float(j_macm2[jr_idx])

    # Qa, Qr — Manuel's convention: delta_Q[i] = delta_t[i] * I[i]
    delta_t = np.diff(asc.time_s, append=asc.time_s[-1])
    delta_q = delta_t * asc.current_a  # coulombs per step

    qa = float(np.sum(delta_q[fwd_slice]))
    qr = float(np.sum(delta_q[rev_slice]))

    jr_ja = jr / ja if ja != 0 else float("nan")
    qr_qa = qr / qa if qa != 0 else float("nan")

    rp_array_idx = ruhepotential_row - 2
    if 0 <= rp_array_idx < n:
        ruhe_mv = float(asc.potential_mv[rp_array_idx])
    else:
        ruhe_mv = float("nan")

    # Excel ranges (1-based, row 1 = header)
    excel_jr_range = (split + 2, rev_end + 1)
    excel_qa_range = (2, split + 1)
    excel_qr_range = (split + 2, rev_end + 1)

    return DLEPRResult(
        split_index=split,
        split_method=diag.get("method", split_method),
        rp_skip=rp_skip,
        split_diagnostics=diag,
        ja_ma_cm2=ja,
        jr_ma_cm2=jr,
        ja_index=ja_idx,
        jr_index=jr_idx,
        qa_as=qa,
        qr_as=qr,
        jr_ja=jr_ja,
        qr_qa=qr_qa,
        ruhepotential_mv=ruhe_mv,
        excel_jr_range=excel_jr_range,
        excel_qa_range=excel_qa_range,
        excel_qr_range=excel_qr_range,
    )
