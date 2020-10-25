import os
import sys
import pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parents[1]))
import excel_to_html
import pytest
os.chdir(os.path.dirname(os.path.abspath(__file__)))


@pytest.mark.parametrize(
    "file_path,kwargs,output",
    [
        ("test.xlsx", {}, "output1.html"),
        ("test.xlsx", {'min_row': 1}, "output2.html"),
        ("test.xlsx", {'min_row': 1, 'max_row': 3}, "output3.html"),
        ("test.xlsx", {'min_row': 1, 'max_row': 3, 'min_col': 1}, "output4.html"),
        ("test2.xlsx", {}, "output5.html"),
        ("test2.xlsx", {'min_row': 2}, "output6.html"),

    ],
)
def test(file_path, kwargs, output):
    body = excel_to_html.main(file_path, **kwargs)
    with open(output, 'r') as f:
        official_body = f.read()
    assert body == official_body
