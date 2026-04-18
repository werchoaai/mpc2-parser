"""Datenintegritäts-Checks für DL-EPR-Messungen.

Berechnet einen Integritäts-Score (0-100) basierend auf mehreren physikalisch
motivierten Kriterien. Jeder Check gibt Status + Begründung zurück, sodass der
Benutzer sofort sieht, was geprüft wurde und wo eine Auffälligkeit ist.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

import numpy as np

from .parser import ASCData
from .analysis import DLEPRResult


Status = Literal["ok", "warn", "fail"]


@dataclass
class Check:
    name: str
    status: Status
    score: float  # 0..1
    weight: float
    detail: str
    value: str = ""


@dataclass
class IntegrityReport:
    score: int  # 0..100
    grade: str  # A, B, C, D
    checks: list[Check] = field(default_factory=list)

    def weighted_score(self) -> float:
        total_w = sum(c.weight for c in self.checks)
        if total_w == 0:
            return 0.0
        return 100.0 * sum(c.score * c.weight for c in self.checks) / total_w

    @property
    def n_ok(self) -> int:
        return sum(1 for c in self.checks if c.status == "ok")

    @property
    def n_warn(self) -> int:
        return sum(1 for c in self.checks if c.status == "warn")

    @property
    def n_fail(self) -> int:
        return sum(1 for c in self.checks if c.status == "fail")


def evaluate_integrity(asc: ASCData, result: DLEPRResult) -> IntegrityReport:
    """Compute data quality score and per-check breakdown."""
    checks: list[Check] = []

    # ── Check 1: Datenpunktzahl ────────────────────────────────────────────
    n = asc.n_points
    if n >= 1000:
        checks.append(Check("Anzahl Datenpunkte", "ok", 1.0, 1.0,
                            f"{n} Punkte — ausreichend für robuste Peak-Erkennung.",
                            value=f"{n}"))
    elif n >= 500:
        checks.append(Check("Anzahl Datenpunkte", "warn", 0.6, 1.0,
                            f"Nur {n} Punkte — Peak-Erkennung funktioniert, aber Qa/Qr-Integration "
                            "kann leicht ungenauer werden.",
                            value=f"{n}"))
    else:
        checks.append(Check("Anzahl Datenpunkte", "fail", 0.2, 1.0,
                            f"Nur {n} Punkte — zu wenig für eine verlässliche DL-EPR-Auswertung.",
                            value=f"{n}"))

    # ── Check 2: Rückkehr zum Startpotenzial ───────────────────────────────
    start_mv = asc.potential_v[0] * 1000.0
    end_mv = asc.potential_v[-1] * 1000.0
    delta_mv = abs(end_mv - start_mv)
    if delta_mv <= 15:
        checks.append(Check("Rückkehr zum Startpotenzial", "ok", 1.0, 1.5,
                            f"Endpotenzial weicht nur {delta_mv:.1f} mV vom Startwert ab — "
                            "saubere DL-EPR-Schleife.",
                            value=f"Δ {delta_mv:.1f} mV"))
    elif delta_mv <= 40:
        checks.append(Check("Rückkehr zum Startpotenzial", "warn", 0.6, 1.5,
                            f"Endpotenzial weicht {delta_mv:.1f} mV vom Startwert ab — "
                            "Rückwärtsrampe endet nicht exakt beim Ausgangspotenzial.",
                            value=f"Δ {delta_mv:.1f} mV"))
    else:
        checks.append(Check("Rückkehr zum Startpotenzial", "fail", 0.3, 1.5,
                            f"Endpotenzial weicht {delta_mv:.1f} mV vom Startwert ab — "
                            "die Rückwärtsrampe ist unvollständig oder überzogen.",
                            value=f"Δ {delta_mv:.1f} mV"))

    # ── Check 3: Symmetrie forward/reverse Dauer ──────────────────────────
    fwd_duration = asc.time_s[result.split_index] - asc.time_s[0]
    rev_duration = asc.time_s[-1] - asc.time_s[result.split_index]
    if fwd_duration > 0:
        ratio = rev_duration / fwd_duration
        diff_pct = abs(ratio - 1.0) * 100.0
    else:
        ratio = 0.0
        diff_pct = 100.0

    if diff_pct <= 20:
        checks.append(Check("Sweep-Symmetrie", "ok", 1.0, 1.0,
                            f"Vorwärts- und Rückwärtsrampe dauern fast gleich lang "
                            f"(Verhältnis {ratio:.2f}) — Split-Punkt mittig platziert.",
                            value=f"{ratio:.2f}"))
    elif diff_pct <= 40:
        checks.append(Check("Sweep-Symmetrie", "warn", 0.6, 1.0,
                            f"Rampen-Verhältnis {ratio:.2f} — leicht asymmetrisch. Split-Punkt "
                            "möglicherweise nicht optimal.",
                            value=f"{ratio:.2f}"))
    else:
        checks.append(Check("Sweep-Symmetrie", "fail", 0.3, 1.0,
                            f"Rampen-Verhältnis {ratio:.2f} — stark asymmetrisch. Split-Punkt "
                            "sollte manuell überprüft werden.",
                            value=f"{ratio:.2f}"))

    # ── Check 4: Ja Peak klar über Grundlinie ──────────────────────────────
    j = asc.current_density_macm2
    # Baseline = median of first 5% of data
    baseline = float(np.median(np.abs(j[:max(10, n // 20)])))
    ja_ratio = result.ja_ma_cm2 / max(baseline, 1e-6)

    if ja_ratio >= 50:
        checks.append(Check("Ja-Peak Signal/Rausch", "ok", 1.0, 1.5,
                            f"Aktivierungspeak Ja = {result.ja_ma_cm2:.2f} mA/cm² liegt "
                            f"deutlich über der Grundlinie ({ja_ratio:.0f}× Signal/Rausch).",
                            value=f"{ja_ratio:.0f}×"))
    elif ja_ratio >= 10:
        checks.append(Check("Ja-Peak Signal/Rausch", "warn", 0.7, 1.5,
                            f"Ja-Peak ist erkennbar aber mit geringerem Abstand zur Grundlinie "
                            f"({ja_ratio:.0f}× Signal/Rausch).",
                            value=f"{ja_ratio:.0f}×"))
    else:
        checks.append(Check("Ja-Peak Signal/Rausch", "fail", 0.3, 1.5,
                            f"Ja-Peak nur {ja_ratio:.1f}× über Grundlinie — sehr schwach oder "
                            "die Messung ist nicht ausreichend aktiviert.",
                            value=f"{ja_ratio:.1f}×"))

    # ── Check 5: Jr Peak detektierbar ──────────────────────────────────────
    jr_ratio = result.jr_ma_cm2 / max(baseline, 1e-6)
    if jr_ratio >= 5:
        checks.append(Check("Jr-Peak Signal/Rausch", "ok", 1.0, 1.5,
                            f"Reaktivierungspeak Jr = {result.jr_ma_cm2:.3f} mA/cm² — "
                            f"klar detektierbar ({jr_ratio:.1f}× Signal/Rausch).",
                            value=f"{jr_ratio:.1f}×"))
    elif jr_ratio >= 2:
        checks.append(Check("Jr-Peak Signal/Rausch", "warn", 0.7, 1.5,
                            f"Jr-Peak liegt nur {jr_ratio:.1f}× über Grundlinie — "
                            "geringe Sensibilisierung oder ungenaue Peak-Detektion möglich.",
                            value=f"{jr_ratio:.1f}×"))
    else:
        checks.append(Check("Jr-Peak Signal/Rausch", "fail", 0.2, 1.5,
                            f"Jr-Peak nur {jr_ratio:.1f}× über Grundlinie — praktisch "
                            "nicht unterscheidbar vom Rauschen.",
                            value=f"{jr_ratio:.1f}×"))

    # ── Check 6: Reverse-Endpunkt gut getroffen ────────────────────────────
    rev_end = result.split_diagnostics.get("reverse_endpoint", n - 1)
    tail_fraction = (n - rev_end) / n
    if 0.02 <= tail_fraction <= 0.2:
        checks.append(Check("Qr-Integrations-Endpunkt", "ok", 1.0, 1.0,
                            f"Qr-Integration endet {tail_fraction*100:.0f}% vor dem Dateiende — "
                            "DL-EPR-konformer Cutoff beim Startpotenzial erkannt.",
                            value=f"{tail_fraction*100:.1f}%"))
    elif tail_fraction < 0.02:
        checks.append(Check("Qr-Integrations-Endpunkt", "warn", 0.6, 1.0,
                            "Potenzial kehrt nicht eindeutig zum Startwert zurück — Qr-Cutoff "
                            "wird am Dateiende gesetzt, könnte den Tail leicht überzählen.",
                            value=f"{tail_fraction*100:.1f}%"))
    else:
        checks.append(Check("Qr-Integrations-Endpunkt", "warn", 0.6, 1.0,
                            f"Sehr langer Tail nach Reaktivierung ({tail_fraction*100:.0f}% der Daten) — "
                            "bitte visuell prüfen ob die Sweep-Parameter korrekt waren.",
                            value=f"{tail_fraction*100:.1f}%"))

    # ── Check 7: Ir/Ia im plausiblen DOS-Bereich ──────────────────────────
    dos = result.jr_ja
    if 0 < dos < 0.6:
        checks.append(Check("DOS-Wert plausibel", "ok", 1.0, 0.8,
                            f"Jr/Ja = {dos:.4f} liegt im erwarteten Bereich für DL-EPR "
                            "(0 = kein Sensibilisierungsgrad, 0.6 = extrem).",
                            value=f"{dos:.4f}"))
    elif 0 < dos < 1.0:
        checks.append(Check("DOS-Wert plausibel", "warn", 0.7, 0.8,
                            f"Jr/Ja = {dos:.4f} — ungewöhnlich hoher DOS-Wert. "
                            "Bitte Probenvorbereitung oder Elektrolytzusammensetzung prüfen.",
                            value=f"{dos:.4f}"))
    else:
        checks.append(Check("DOS-Wert plausibel", "fail", 0.3, 0.8,
                            f"Jr/Ja = {dos:.4f} außerhalb des physikalisch plausiblen Bereichs — "
                            "Auswertung überprüfen.",
                            value=f"{dos:.4f}"))

    report = IntegrityReport(score=0, grade="F", checks=checks)
    score_f = report.weighted_score()
    report.score = int(round(score_f))
    report.grade = (
        "A" if report.score >= 90 else
        "B" if report.score >= 75 else
        "C" if report.score >= 60 else
        "D" if report.score >= 40 else "F"
    )
    return report
