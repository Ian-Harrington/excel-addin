from dataclasses import dataclass
from decimal import Decimal


@dataclass
class HeaderCellMapping:
    """Using common format for defaults"""
    row: int = 7
    identifier: int = 3
    description: int = 6
    price: int = 8


@dataclass
class LineItem:
    identifier: str | None
    description: str | None
    price: Decimal | None

    def __str__(self) -> str:
        desc_str = str(self.description or "?")
        abbreviated_desc = desc_str[:min(len(desc_str), 6)]
        return f"[{self.identifier}]: ${self.price or " ? "} - {abbreviated_desc}"

    def is_valid(self) -> bool:
        return (
            self.identifier is not None
            and self.description is not None
            and self.price is not None
        )

    def as_tuple(self) -> tuple[str, str, str, Decimal]:
        if not self.is_valid():
            raise ValueError(f"Cannot convert invalid LineItem to tuple: {self}")
        uom = "FULL" if self.price < 10 else "WT"
        return (self.identifier, self.description, uom, self.price)


@dataclass
class SheetFormat:
    name: str
    header_mapping: HeaderCellMapping


@dataclass
class ParsedSheet:
    name: str
    items: list[LineItem]