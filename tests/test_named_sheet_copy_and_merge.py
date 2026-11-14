"""
Named sheet copy, extension and merge

A draft is made by copying an inbound sheet that has column names. Columns
are then appended under new names, rows are appended under new names, and
the draft is merged into a summary workbook.  Throughout, the column and
row names must follow the data, and the header must be exported once.
"""

import pyexcel as pe
from pyexcel._compact import OrderedDict

from .base import clean_up_files

INBOUND = [
    ["SKU", "Name", "Qty"],
    ["A1", "apple", 3],
    ["B2", "banana", 5],
]

HEADER_ONCE = "SKU,Name,Qty,Price,Currency\r\n"


def _inbound_sheet():
    return pe.Sheet(
        [list(row) for row in INBOUND],
        "inbound",
        name_columns_by_row=0,
    )


def test_clone_keeps_the_sheet_and_column_names():
    sheet = _inbound_sheet()
    draft = sheet.clone()
    assert draft.name == "inbound"
    assert draft.colnames == ["SKU", "Name", "Qty"]
    assert draft.rownames == []
    assert draft.to_array() == INBOUND
    assert draft.column["SKU"] == ["A1", "B2"]


def test_deepcopy_keeps_the_sheet_and_column_names():
    import copy

    sheet = _inbound_sheet()
    draft = copy.deepcopy(sheet)
    assert draft.name == "inbound"
    assert draft.colnames == ["SKU", "Name", "Qty"]
    assert draft.to_array() == INBOUND


def test_clone_keeps_the_row_names():
    sheet = _inbound_sheet()
    sheet.rownames = ["r1", "r2"]
    draft = sheet.clone()
    assert draft.rownames == ["r1", "r2"]
    assert draft.colnames == ["SKU", "Name", "Qty"]
    assert draft.to_array() == [
        ["", "SKU", "Name", "Qty"],
        ["r1", "A1", "apple", 3],
        ["r2", "B2", "banana", 5],
    ]


def test_clone_of_a_sheet_that_has_only_a_header():
    sheet = pe.Sheet(
        [["SKU", "Name"]],
        "header only",
        name_columns_by_row=0,
    )
    assert sheet.number_of_rows() == 0
    draft = sheet.clone()
    assert draft.name == "header only"
    assert draft.colnames == ["SKU", "Name"]
    assert draft.to_array() == [["SKU", "Name"]]


def test_clone_of_an_unnamed_sheet_is_unchanged():
    sheet = pe.Sheet([[1, 2], [3, 4]])
    draft = sheet.clone()
    assert draft.name == sheet.name
    assert draft.colnames == []
    assert draft.rownames == []
    assert draft.to_array() == [[1, 2], [3, 4]]


def test_clone_does_not_share_data_with_the_original():
    sheet = _inbound_sheet()
    draft = sheet.clone()
    draft[0, 0] = "changed"
    assert sheet[0, 0] == "A1"
    assert draft[0, 0] == "changed"


def test_extend_columns_by_name_keeps_the_header():
    draft = _inbound_sheet().clone()
    draft.column += OrderedDict(
        [("Price", [1.2, 2.3]), ("Currency", ["CNY", "CNY"])],
    )
    assert draft.colnames == ["SKU", "Name", "Qty", "Price", "Currency"]
    assert draft.column["Price"] == [1.2, 2.3]
    assert draft.to_array() == [
        ["SKU", "Name", "Qty", "Price", "Currency"],
        ["A1", "apple", 3, 1.2, "CNY"],
        ["B2", "banana", 5, 2.3, "CNY"],
    ]


def test_extend_columns_of_a_single_data_row_keeps_the_header():
    draft = pe.Sheet(
        [["SKU", "Qty"], ["A1", 5]],
        "draft",
        name_columns_by_row=0,
    )
    draft.extend_columns(OrderedDict([("Note", ["n1"])]))
    assert draft.colnames == ["SKU", "Qty", "Note"]
    assert draft.to_array() == [["SKU", "Qty", "Note"], ["A1", 5, "n1"]]


def test_append_the_columns_of_a_named_sheet():
    draft = _inbound_sheet()
    extra = pe.Sheet(
        [["Price", "Currency"], [1.2, "CNY"], [2.3, "CNY"]],
        "extra",
        name_columns_by_row=0,
    )
    draft.column += extra
    assert draft.colnames == ["SKU", "Name", "Qty", "Price", "Currency"]
    assert draft.column["Currency"] == ["CNY", "CNY"]
    assert draft.to_array() == [
        ["SKU", "Name", "Qty", "Price", "Currency"],
        ["A1", "apple", 3, 1.2, "CNY"],
        ["B2", "banana", 5, 2.3, "CNY"],
    ]


def test_extend_rows_by_name_keeps_the_rows_and_the_columns():
    sheet = _inbound_sheet()
    sheet.rownames = ["r1", "r2"]
    sheet.extend_columns(OrderedDict([("Price", [1.2, 2.3])]))
    sheet.extend_rows(
        OrderedDict(
            [
                ("r3", ["C3", "cherry", 7, 3.5]),
                ("r4", ["D4", "durian", 9, 4.5]),
            ],
        ),
    )
    assert sheet.rownames == ["r1", "r2", "r3", "r4"]
    assert sheet.colnames == ["SKU", "Name", "Qty", "Price"]
    assert sheet.named_row_at("r3") == ["C3", "cherry", 7, 3.5]
    assert sheet.to_array() == [
        ["", "SKU", "Name", "Qty", "Price"],
        ["r1", "A1", "apple", 3, 1.2],
        ["r2", "B2", "banana", 5, 2.3],
        ["r3", "C3", "cherry", 7, 3.5],
        ["r4", "D4", "durian", 9, 4.5],
    ]


def test_rows_added_under_names_get_the_names_not_the_old_rows():
    sheet = _inbound_sheet()
    sheet.extend_rows(OrderedDict([("C3", ["C3", "cherry", 7])]))
    assert sheet.rownames == ["", "", "C3"]
    assert sheet.named_row_at("C3") == ["C3", "cherry", 7]
    assert sheet.colnames == ["SKU", "Name", "Qty"]


def test_draft_export_has_the_header_exactly_once():
    draft = _inbound_sheet().clone()
    draft.column += OrderedDict(
        [("Price", [1.2, 2.3]), ("Currency", ["CNY", "CNY"])],
    )
    assert draft.csv == (
        HEADER_ONCE + "A1,apple,3,1.2,CNY\r\n" + "B2,banana,5,2.3,CNY\r\n"
    )


def test_a_sheet_added_to_a_book_keeps_its_names():
    book = pe.Book()
    book += _inbound_sheet()
    stored = book["inbound"]
    assert stored.colnames == ["SKU", "Name", "Qty"]
    assert stored.to_array() == INBOUND
    assert book.to_dict()["inbound"] == INBOUND


def test_a_book_merged_into_a_book_keeps_its_names():
    first = pe.Book({"inbound": _inbound_sheet()})
    second = pe.Book({"other": pe.Sheet([[1, 2]], "other")})
    merged = first + second
    assert merged.sheet_names() == ["memory_inbound", "memory_other"]
    assert merged["memory_inbound"].colnames == ["SKU", "Name", "Qty"]
    assert merged["memory_inbound"].to_array() == INBOUND


def test_sheet_plus_sheet_keeps_the_column_names():
    book = _inbound_sheet() + pe.Sheet([["Note"], ["n1"], ["n2"]], "extra")
    assert book["inbound"].colnames == ["SKU", "Name", "Qty"]
    assert book["inbound"].to_array() == INBOUND


def test_summary_workbook_export_matches_the_fixed_data():
    source = _inbound_sheet()
    draft = source.clone()
    draft.column += OrderedDict(
        [("Price", [1.2, 2.3]), ("Currency", ["CNY", "CNY"])],
    )
    draft.row += OrderedDict([("R4", ["C3", "cherry", 7, 3.5, "CNY"])])

    book = pe.Book()
    book += source
    book += draft
    book.save_as("summary_named_sheets.xlsx")

    reopened = pe.get_book(file_name="summary_named_sheets.xlsx")
    names = reopened.sheet_names()
    assert len(names) == 2
    assert names[0] == "inbound"
    assert names[1].startswith("inbound")
    assert reopened[names[0]].to_array() == INBOUND
    assert reopened[names[1]].to_array() == [
        ["", "SKU", "Name", "Qty", "Price", "Currency"],
        ["", "A1", "apple", 3, 1.2, "CNY"],
        ["", "B2", "banana", 5, 2.3, "CNY"],
        ["R4", "C3", "cherry", 7, 3.5, "CNY"],
    ]
    clean_up_files(["summary_named_sheets.xlsx"])


def test_unnamed_sheet_keeps_its_previous_results():
    sheet = pe.Sheet([[1, 2], [3, 4]])
    sheet.extend_rows([[5, 6]])
    assert sheet.colnames == []
    assert sheet.rownames == []
    assert sheet.to_array() == [[1, 2], [3, 4], [5, 6]]

    sheet = pe.Sheet([[1, 2], [3, 4]])
    sheet.extend_columns([[5, 6], [7, 8]])
    assert sheet.colnames == []
    assert sheet.to_array() == [[1, 2, 5, 7], [3, 4, 6, 8]]

    book = pe.Book()
    book += pe.Sheet([[9, 9]], "unnamed")
    assert book["unnamed"].array == [[9, 9]]
    assert book["unnamed"].colnames == []
