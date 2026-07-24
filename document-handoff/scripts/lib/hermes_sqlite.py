#!/usr/bin/env python3
"""Dependency-free SQLite bridge for the document-handoff Hermes adapter."""

import base64
import json
import sqlite3
import sys
from pathlib import Path


def quote_identifier(value):
    return '"' + str(value).replace('"', '""') + '"'


def json_value(value):
    if isinstance(value, (bytes, bytearray, memoryview)):
        return {"$binary_base64": base64.b64encode(bytes(value)).decode("ascii")}
    return value


def row_dict(row):
    return {key: json_value(row[key]) for key in row.keys()}


def connect(database):
    path = Path(database)
    if not path.is_file():
        raise ValueError("database does not exist")
    native = str(path)
    if native.startswith("\\\\"):
        # SQLite URI authorities reject valid Windows UNC paths. Use the native
        # path plus query_only for that case.
        connection = sqlite3.connect(native)
    else:
        connection = sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA query_only=ON")
    return connection


def table_names(connection):
    return {
        row[0]
        for row in connection.execute(
            "SELECT name FROM sqlite_master "
            "WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        )
    }


def table_columns(connection, table):
    return [
        str(row[1])
        for row in connection.execute(
            "PRAGMA table_info({})".format(quote_identifier(table))
        )
    ]


def first_column(columns, names):
    return next((name for name in names if name in columns), None)


def discover(connection, request):
    session_table = request.get("session_table") or "sessions"
    if session_table not in table_names(connection):
        return {"rows": [], "id_column": None, "cwd_column": None}
    columns = table_columns(connection, session_table)
    id_column = request.get("id_column") or first_column(
        columns, ["id", "session_id", "sessionId"]
    )
    cwd_column = request.get("cwd_column") or first_column(
        columns, ["cwd", "project_path", "workspace_path", "workspace"]
    )
    if not id_column or id_column not in columns:
        return {"rows": [], "id_column": None, "cwd_column": cwd_column}
    session_ids = [str(value) for value in request.get("session_ids", [])]
    sql = "SELECT * FROM {}".format(quote_identifier(session_table))
    parameters = []
    if session_ids:
        sql += " WHERE {} IN ({})".format(
            quote_identifier(id_column), ",".join("?" for _ in session_ids)
        )
        parameters = session_ids
    rows = [row_dict(row) for row in connection.execute(sql, parameters)]
    return {"rows": rows, "id_column": id_column, "cwd_column": cwd_column}


def export_session(connection, request):
    session_id = str(request["session_id"])
    session_table = request.get("session_table") or "sessions"
    session_id_column = request.get("id_column") or "id"
    rows = []
    connection.execute("BEGIN")
    try:
        for table in sorted(table_names(connection)):
            columns = table_columns(connection, table)
            key = (
                session_id_column
                if table == session_table and session_id_column in columns
                else first_column(columns, ["session_id", "sessionId", "session"])
            )
            if not key:
                continue
            sql = "SELECT * FROM {} WHERE {} = ?".format(
                quote_identifier(table), quote_identifier(key)
            )
            for row in connection.execute(sql, (session_id,)):
                rows.append({"table": table, "row": row_dict(row)})
    finally:
        connection.rollback()
    return {"rows": rows}


def main():
    try:
        request = json.load(sys.stdin)
        with connect(request["database"]) as connection:
            if request.get("action") == "discover":
                result = discover(connection, request)
            elif request.get("action") == "export":
                result = export_session(connection, request)
            else:
                raise ValueError("unknown action")
        json.dump(result, sys.stdout, separators=(",", ":"))
        sys.stdout.write("\n")
        return 0
    except Exception:
        print("Hermes SQLite bridge failed", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
