import logging
from dataclasses import Field, fields, is_dataclass
from typing import Type, get_args, get_origin

from openpyxl import Workbook
from openpyxl.cell import Cell

log = logging.getLogger(__name__)

# problem is parallel recursive descent of the type spec and the excel structure

# openpyxl: Cell, Tuple[Cell], List[Tuple[Cell]]
# xlwings:
# xlrd:


# - find the cell/range, based on the field name, or perhaps some alias
# - cell/range comment could also provide input
# - openpyxl does some conversion based upon the formatting
# - perform conversion based upon the defined name contents first
# - convertion based on cell
# - perform conversion based upon the field metadata, and then the default


def _ndims(type: type):
    if issubclass(get_origin(type) or type, (tuple, list)):
        return 1 + sum(_ndims(arg) for arg in get_args(type))
    return 0


# each type has an origin and args
# we only accept list/tuple with two levels of nesting and arg
# arg can be any atomic type with a conversion, or datetime, or date.
# do the base case type -> cell first.
# the do the recursive bit - need to capture 1d lists first.

# if we see a (Cell, ) tuple that can convert to an atomic.
# atomic == not a list or tuple

# how do dicts fit in - they have to parse the cell structure
# dataclasses can also be parsed from just a range by matching the
# cell names, or maybe aliases.


def _resolve_defined_name(field: Field, type_: type, wb: Workbook, defn):
    log.debug(f"Resolving {field.name} from {defn.attr_text}")

    # usetyping.get_origin to figure out if it's a list
    for dest in defn.destinations:
        # this resolves to the cell or range - data_type is n/s/d
        contents = wb[dest[0]][dest[1]]

        log.debug(f"{field.name} -> {defn.attr_text} : {contents.data_type}")

        if isinstance(contents, Cell):
            return contents.value

        if get_origin(field.type) is list:
            if get_origin(get_args(field.type)[0]) not in (list, tuple):
                return [cell.value for row in contents for cell in row]

            return [tuple(cell.value for cell in row) for row in contents]

        raise ValueError("unknown value type {type(rows)}")

    # if no matching field, but the type is structured then use that

    # field.type is another dataclass then we can recurse on the dotted names

    raise ValueError(f"Could not find any destinations for {defn}")


def read_excel(cls, wb: Workbook):
    if not is_dataclass(cls):
        raise TypeError(f"{cls!s} is not a dataclass")

    kwargs = {
        field.name: _resolve_defined_name(
            field, field.type, wb, wb.defined_names[field.name]
        )
        for field in fields(cls)
    }
    return cls(**kwargs)
