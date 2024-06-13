import logging
from dataclasses import Field, fields, is_dataclass

from openpyxl import Workbook
from openpyxl.cell import Cell

log = logging.getLogger(__name__)

# - find the cell/range, based on the field name, or perhaps some alias
# - cell/range comment could also provide input
# - openpyxl does some conversion based upon the formatting
# - perform conversion based upon the defined name contents first
# - convertion based on cell
# - perform conversion based upon the field metadata, and then the default


def _resolve_defined_name(field: Field, wb: Workbook, defn):
    log.debug(f"Resolving {field.name} from {defn.attr_text}")
    for dest in defn.destinations:
        # this resolves to the cell or range - data_type is n/s/d
        cell: Cell = wb[dest[0]][dest[1]]
        log.debug(f"\t= {cell.value!r}: {cell.data_type}")
        return cell.value

    # if no matching field, but the type is structured then use that

    # field.type is another dataclass then we can recurse on the dotted names

    raise ValueError(f"Could not find any destinations for {defn}")


def read_excel(cls, wb: Workbook):
    if not is_dataclass(cls):
        raise TypeError(f"{cls!s} is not a dataclass")

    kwargs = {
        field.name: _resolve_defined_name(field, wb, wb.defined_names[field.name])
        for field in fields(cls)
    }
    return cls(**kwargs)
