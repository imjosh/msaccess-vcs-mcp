"""Isolated read-only DAO discovery. Never open the installed file in Access."""

from __future__ import annotations

import json
import sys


def read_version(path: str):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    engine = database = None
    try:
        engine = win32com.client.Dispatch("DAO.DBEngine.120")
        database = engine.OpenDatabase(path, False, True)  # shared, read-only
        return database.Properties("AppVersion").Value
    finally:
        if database is not None:
            database.Close()
        database = engine = None
        pythoncom.CoUninitialize()


def main() -> None:
    try:
        reply = {"version": read_version(sys.argv[1])}
    except Exception as exc:
        reply = {"version": None, "error": str(exc)}
    print(json.dumps(reply))


if __name__ == "__main__":
    main()
