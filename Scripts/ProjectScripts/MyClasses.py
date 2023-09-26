from dataclasses import dataclass

@dataclass
class MatchItem:
    cfo140_row_index: int
    rpnd_row_index: int
    old_summ: float
    new_summ: float
    key: str