# SPDX-FileCopyrightText: 2024-present U.N. Owen <void@some.where>
#
# SPDX-License-Identifier: MIT
import sys

if __name__ == "__main__":
    from dataclasses_excel.cli import dataclasses_excel

    sys.exit(dataclasses_excel())
