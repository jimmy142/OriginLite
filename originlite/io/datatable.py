# originlite/io/datatable.py
from dataclasses import dataclass
from typing import List
import csv
import numpy as np


def _excel_col_name(n: int) -> str:
    """
    0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA, 27 -> AB, ...
    """
    s = ""
    n = int(n)
    while True:
        n, r = divmod(n, 26)
        s = chr(ord('A') + r) + s
        if n == 0:
            return s
        n -= 1  # Excel-style carry


@dataclass
class DataTable:
    headers: List[str]
    data: np.ndarray  # shape (n_rows, n_cols)

    @staticmethod
    def from_csv(path: str) -> "DataTable":
        with open(path, "r", newline="") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except Exception:
                dialect = csv.excel
            reader = csv.reader(f, dialect)
            rows = list(reader)
        if not rows:
            raise ValueError("Empty file")

        def is_number(s: str) -> bool:
            try:
                float(s)
                return True
            except Exception:
                return False

        maybe_header = rows[0]
        header_ratio = sum(is_number(x) for x in maybe_header) / max(len(maybe_header), 1)

        # Decide if row 0 is a header (mostly non-numeric) — but regardless,
        # we will REPLACE headers with A, B, C... for easier expressions.
        if header_ratio < 0.6:
            numeric_rows = rows[1:]
        else:
            numeric_rows = rows

        arr = []
        for r in numeric_rows:
            try:
                arr.append([float(x) for x in r])
            except Exception:
                arr.append([np.nan for _ in r])

        A = np.array(arr, dtype=float)
        if A.size == 0:
            raise ValueError("No numeric rows found")

        # Trim to the shortest row length
        min_cols = min(len(r) for r in numeric_rows)
        A = A[:, :min_cols]

        # Drop rows with NaNs introduced by bad lines
        mask = ~np.isnan(A).any(axis=1)
        A = A[mask]
        if A.size == 0:
            raise ValueError("All rows contained non-numeric values after cleaning")

        # Force Excel-like headers: A, B, C, ..., AA, AB, ...
        headers = [_excel_col_name(i) for i in range(A.shape[1])]

        return DataTable(headers=headers, data=A)

    # Column utilities
    def add_column(self, name: str, values: np.ndarray) -> None:
        values = np.asarray(values, dtype=float).reshape(-1)
        if values.shape[0] != self.data.shape[0]:
            raise ValueError("Length of new column does not match row count")
        self.data = np.column_stack([self.data, values])
        self.headers.append(name)

    def delete_column(self, index: int) -> None:
        if index < 0 or index >= self.data.shape[1]:
            return
        self.data = np.delete(self.data, index, axis=1)
        del self.headers[index]

    def rename_column(self, index: int, new_name: str) -> None:
        if 0 <= index < len(self.headers):
            self.headers[index] = new_name
