from typing import List, Any

class _Cell:
    text: str

class _Row:
    cells: List[_Cell]

class _Table:
    rows: List[_Row]

class Document:
    def __init__(self, path: str = ...) -> None: ...
    paragraphs: List[Any]
    tables: List[_Table]


def DocumentFactory(path: str) -> Document: ...
