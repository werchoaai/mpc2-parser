"""Output variants for MPC2 parser results."""
from .variant1_auswertung import write_auswertung_workbook
from .variant2_messuebersicht import append_to_messuebersicht
from .variant4_combined import write_combined_workbook

__all__ = [
    "write_auswertung_workbook",
    "append_to_messuebersicht",
    "write_combined_workbook",
]
