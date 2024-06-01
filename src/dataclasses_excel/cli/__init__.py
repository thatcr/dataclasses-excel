# SPDX-FileCopyrightText: 2024-present U.N. Owen <void@some.where>
#
# SPDX-License-Identifier: MIT
import click

from dataclasses_excel.__about__ import __version__


@click.group(context_settings={"help_option_names": ["-h", "--help"]}, invoke_without_command=True)
@click.version_option(version=__version__, prog_name="dataclasses-excel")
def dataclasses_excel():
    click.echo("Hello world!")
