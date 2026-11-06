"""BC Roma roster: initials, jersey number and surname."""

import csv
from dataclasses import dataclass
from pathlib import Path

ROSTER_PATH = Path(__file__).parent.parent / "templates" / "roster.csv"


@dataclass(frozen=True)
class Player:
    """A roster player.

    Attributes:
        initials: Initials used in the general plan (e.g. ``"NM"``).
        number: Jersey number, or None when not assigned.
        surname: Player surname.
    """

    initials: str
    number: int | None
    surname: str

    @property
    def header(self) -> str:
        """Block title used in the individual schedule, e.g. ``"#1 MANNION"``."""
        label = self.surname.upper()
        return f"#{self.number} {label}" if self.number is not None else label


def load_roster(path: Path = ROSTER_PATH) -> list[Player]:
    """Load the roster in schedule order.

    Args:
        path: CSV with columns ``initials``, ``number`` and ``surname``.

    Returns:
        Players in the order their blocks are placed in the template.
    """
    with path.open(newline="", encoding="utf-8") as handle:
        return [
            Player(
                row["initials"].strip().upper(),
                int(row["number"]) if row["number"].strip() else None,
                row["surname"].strip(),
            )
            for row in csv.DictReader(handle)
        ]
