import logging
from dataclasses import dataclass
from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook

from dataclasses_excel import read_excel

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
        foo: int
        bar: str
        baf: datetime

    obj = read_excel(MyClass, wb)

    assert obj == MyClass(foo=123, bar="string", baf=datetime(2024, 1, 1))


# def test_defined_names(wb: Workbook):
#     defn: DefinedName
#     for key, defn in wb.defined_names.items():
#         print(vars(defn))
#         for dest in defn.destinations:
#             print(dest)
#             # this will get the range of cells values...
#             print(wb[dest[0]][dest[1]])
