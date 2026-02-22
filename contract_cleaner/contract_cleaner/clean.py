import logging

from contract_cleaner.models import LineItem


log = logging.getLogger(__name__)


def clean_items(items: list[LineItem]) -> list[LineItem]:
    """Removes duplicate items or ones with missing values"""
    item_dict: dict[str, LineItem] = {}
    for item in items:
        if (
            item.identifier is None
            or item.description is None
            or item.price is None
        ):
            log.info(f"Dropping {item} - missing value(s)")
            continue
        if item.identifier in item_dict.keys():
            # keep higher priced duplicate
            if item.price > item_dict[item.identifier].price:
                log.info(f"Dropping {item_dict[item.identifier]} - duplicate identifier with lower price than {item}")
                item_dict[item.identifier] = item
            continue
        item_dict[item.identifier] = item
    return list(item_dict.values())

