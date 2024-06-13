import logging
from dataclasses import dataclass
from datetime import datetime
from typing import List, Tuple

import pytest
from dataclasses_excel import read_excel
from openpyxl import Workbook, load_workbook

log = logging.getLogger(__name__)


@pytest.fixture()
def wb(shared_datadir):
    path = shared_datadir / "test_named_ranges.xlsx"
    log.info(f"Loading {path}")
    wb = load_workbook(path)
    yield wb
    wb.close()


def test_dataclass(wb: Workbook):
    @dataclass
    class MyClass:
        integer: int
        string: str
        datetime: datetime
        integer_list: List[int]
        integer_matrix: List[Tuple[int]]

    obj = read_excel(MyClass, wb)

    assert obj == MyClass(
        integer=123,
        string="string",
        datetime=datetime(2024, 1, 1),
        integer_list=[1, 2, 3],
        integer_matrix=[(1, 2), (3, 4)],
    )


# def test_defined_names(wb: Workbook):
#     defn: DefinedName
#     for key, defn in wb.defined_names.items():
#         print(vars(defn))
#         for dest in defn.destinations:
#             print(dest)
#             # this will get the range of cells values...
#             print(wb[dest[0]][dest[1]])
