import sqlite3
import pymysql
import os

def export_sqlite_to_mysql(sqlite_file, mysql_file):
    # Connect to SQLite
    sqlite_conn = sqlite3.connect(sqlite_file)
    sqlite_cursor = sqlite_conn.cursor()

    # Get list of tables
    sqlite_cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = sqlite_cursor.fetchall()

    with open(mysql_file, "w", encoding="utf-8") as f:
        f.write("-- SQLite to MySQL export\n\n")
        f.write("SET FOREIGN_KEY_CHECKS=0;\n\n")

        for (table_name,) in tables:
            # Skip SQLite internal tables
            if table_name.startswith("sqlite_"):
                continue

            # Get schema
            sqlite_cursor.execute(f"PRAGMA table_info({table_name})")
            columns = sqlite_cursor.fetchall()

            col_defs = []
            for col in columns:
                col_name = col[1]
                col_type = col[2].upper()
                # Map SQLite types to MySQL types
                if "INT" in col_type:
                    mysql_type = "INT"
                elif "CHAR" in col_type or "CLOB" in col_type or "TEXT" in col_type:
                    mysql_type = "TEXT"
                elif "BLOB" in col_type:
                    mysql_type = "BLOB"
                elif "REAL" in col_type or "FLOA" in col_type or "DOUB" in col_type:
                    mysql_type = "DOUBLE"
                else:
                    mysql_type = "TEXT"

                col_defs.append(f"`{col_name}` {mysql_type}")

            # Write CREATE TABLE
            f.write(f"DROP TABLE IF EXISTS `{table_name}`;\n")
            f.write(f"CREATE TABLE `{table_name}` ({', '.join(col_defs)});\n\n")

            # Dump data
            sqlite_cursor.execute(f"SELECT * FROM {table_name}")
            rows = sqlite_cursor.fetchall()

            for row in rows:
                values = []
                for val in row:
                    if val is None:
                        values.append("NULL")
                    elif isinstance(val, (int, float)):
                        values.append(str(val))
                    else:
                        values.append("'" + str(val).replace("'", "''") + "'")
                f.write(f"INSERT INTO `{table_name}` VALUES ({', '.join(values)});\n")
            f.write("\n")

        f.write("SET FOREIGN_KEY_CHECKS=1;\n")

    sqlite_conn.close()
    print(f"✅ Export complete! File saved at {mysql_file}")


if __name__ == "__main__":
    sqlite_file = "attendance.db"   # <-- change to your sqlite file
    mysql_file = "attendance_mysql.sql" # <-- output MySQL dump
    export_sqlite_to_mysql(sqlite_file, mysql_file)
