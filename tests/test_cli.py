from click.testing import CliRunner

from dataclasses_excel.cli import dataclasses_excel


def test_hello_world():
    runner = CliRunner()
    result = runner.invoke(dataclasses_excel, [])
    assert result.exit_code == 0
    assert result.output == "Hello world!\n"
