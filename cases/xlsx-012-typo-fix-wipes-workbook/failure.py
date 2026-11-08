"""Reproduces the failure: a pandas round trip (read_excel -> replace -> to_excel)."""

import pandas as pd


def solve(workdir, meta):
    path = workdir / "customers.xlsx"
    df = pd.read_excel(path, sheet_name="Customers")
    df["City"] = df["City"].replace("Bejing", "Beijing")
    df.to_excel(path, sheet_name="Customers", index=False)
