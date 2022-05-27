import mysql.connector
import formulas
import json
from .cell import *


NUM_TYPE = 0
STR_TYPE = 1


def export_to_mysql(file_path, database_name, host, username, password):
    print("parsing excel file...")
    # xl_model = formulas.ExcelModel().loads(file_path).finish()
    # xl_dict = xl_model.to_dict()
    xl_dict = None
    with open(file_path, 'r', encoding='utf-8') as f:
        xl_dict = json.load(f)

    db = mysql.connector.connect(
        host=host,
        user=username,
        password=password
    )
    cursor = db.cursor()

    cursor.execute("CREATE DATABASE " + database_name)

    db = mysql.connector.connect(
        host=host,
        user=username,
        password=password,
        database=database_name
    )
    cursor = db.cursor()

    cursor.execute("""
    CREATE TABLE Cell (
        Sheet VARCHAR(255) NOT NULL,
        Col VARCHAR(255) NOT NULL, 
        RowVal INT NOT NULL,
        ValType INT NOT NULL,
        NumVal DOUBLE,
        StrVal VARCHAR(511),
        CONSTRAINT PK_Cell PRIMARY KEY (Sheet, Col, RowVal)
    )
    """)
    cursor.execute("""
    CREATE TABLE Constant (
        Name VARCHAR(255) NOT NULL,
        ValType INT NOT NULL,
        NumVal DOUBLE,
        StrVal VARCHAR(511),
        CONSTRAINT PK_Constant PRIMARY KEY (Name)
    )
    """)

    print("inserting into database...")
    for k in xl_dict:
        cell = Cell.make_from_key(k)
        val = xl_dict[k]
        if isinstance(cell, Constant):
            if isinstance(val, float):
                cursor.execute(f"INSERT INTO Constant (Name, ValType, NumVal) VALUES (\"{k}\", {NUM_TYPE}, {val})")
            else:
                val = str(val).replace("\'", "\'\'")
                cursor.execute(f"INSERT INTO Constant (name, ValType, StrVal) VALUES (\"{k}\", {STR_TYPE}, '{val}')")
        else:
            if isinstance(val, float):
                cursor.execute(f"INSERT INTO Cell (Sheet, Col, RowVal, ValType, NumVal) VALUES (\"{cell.sheet}\", \"{cell.col}\", {cell.row}, {NUM_TYPE}, {val})")
            else:
                val = str(val).replace("\'", "\'\'")
                cursor.execute(f"INSERT INTO Cell (Sheet, Col, RowVal, ValType, StrVal) VALUES (\"{cell.sheet}\", \"{cell.col}\", {cell.row}, {STR_TYPE}, '{val}')")

    db.commit()


class MySqlDatabase:
    def __init__(self, host, username, password, database):
        self.db = mysql.connector.connect(
            host=host,
            user=username,
            password=password,
            database=database
        )


    def get(self, cell):
        cursor = self.db.cursor()
        
        if isinstance(cell, Constant):
            cursor.execute(f"SELECT ValType, NumVal, StrVal FROM Constant WHERE name = \"{cell.form_key()}\"")
        else:
            cursor.execute(f"SELECT ValType, NumVal, StrVal FROM Cell WHERE Sheet = \"{cell.sheet}\" AND Col = \"{cell.col}\" AND RowVal = {cell.row}")

        result = cursor.fetchall()
        
        final_result = None
        if len(result) > 0:
            if result[0][0] == NUM_TYPE:
                final_result = result[0][1]
            else:
                final_result = result[0][2]

        if final_result is None:
            return ""
        return final_result


    def find_text(self, target):
        cursor = self.db.cursor()

        val = str(target).replace("\'", "\'\'")
        cursor.execute(f"SELECT Sheet, Col, RowVal FROM Cell WHERE StrVal = '{val}'")
        results = cursor.fetchall()

        return list(map(lambda r: Cell(r[0], r[1], r[2]), results))


    def find_num(self, target):
        cursor = self.db.cursor()

        cursor.execute(f"SELECT Sheet, Col, RowVal FROM Cell WHERE NumVal = {target}")
        results = cursor.fetchall()

        return list(map(lambda r: Cell(r[0], r[1], r[2]), results))


    def get_final_cell_in_col(self, cell):
        cursor = self.db.cursor()

        cursor.execute(f"SELECT RowVal FROM Cell WHERE Sheet = \"{cell.sheet}\" AND Col = \"{cell.col}\" ORDER BY RowVal DESC")
        results = cursor.fetchall()

        if len(results) > 0:
            return Cell(cell.sheet, cell.col, results[0][0])
        return cell
