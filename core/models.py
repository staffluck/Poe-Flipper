from dataclasses import dataclass
from typing import Dict

@dataclass
class Item:
    category: str
    group: str
    name: str
    explicits: Dict[str, tuple]
    implicits: Dict[str, tuple]
    mean: float
    depends_on_links: bool