# `excelformula`

Database integration of Excel formulae. Supports importing excel files, exporting as JSON or SQL databases, calculating individual cells and calculating columns. Built on top of my personal fork of [`formulas`](https://github.com/vinci1it2000/formulas) which is accessible [here](https://github.com/benrobinson16/formulas).

## Installation

```shell
pip3 install git+https://github.com/benrobinson16/excelformula.git@main
```

## Using with `JsonDatabase`

To use this package with a simple JSON file acting as a key/value database, an excel file first has to be exported to json form.

```python
import excelformula

excelformula.export_to_json("excel.xlsx", "data.json")
```

Now, the file can be used to calculate specific rows/columns:

```python
from excelformula import *

db = JsonDatabase("data.json")
cache = Cache()
s = FormulaSolver(db, cache)

cell1 = Cell("'[excel.xlsx]OVERALL SCORES'", "AU", 2)
cell2 = Cell("'[excel.xlsx]OVERALL SCORES'", "AU", 184)

cell_result = s.calculate_cell(cell1)
col_result = s.calculate_column(cell1, cell2)
```

Labels can be applied to the column result like so:

```python
# first column in sheet
label_cell1 = Cell("'[excel.xlsx]OVERALL SCORES'", "A", 2)
label_cell2 = Cell("'[excel.xlsx]OVERALL SCORES'", "A", 184)
labels = s.calculate_column(label_cell1, label_cell2)

# now the column result with labels
new_col_result = s.calculate_column(cell1, cell2, labels=labels)
```

## Using a MySQL database

The framework can also use a MySQL database with the following tables:

```SQL
CREATE TABLE Cell (
    Sheet VARCHAR(255) NOT NULL,
    Col VARCHAR(255) NOT NULL, 
    RowVal INT NOT NULL,
    ValType INT NOT NULL,
    NumVal DOUBLE,
    StrVal VARCHAR(511),
    CONSTRAINT PK_Cell PRIMARY KEY (Sheet, Col, RowVal)
);

CREATE TABLE Constant (
    Name VARCHAR(255) NOT NULL,
    ValType INT NOT NULL,
    NumVal DOUBLE,
    StrVal VARCHAR(511),
    CONSTRAINT PK_Constant PRIMARY KEY (Name)
);
```

- Here, the `Constant` entity is for names constants stored in the excel file which do not have a row/column.

- `ValType` indicated whether a number or string value should be expected. (`NumVal` vs `StrVal`)

The database can be used like the JSON database as so:

```python
from excelformula import *

db = MySqlDatabase("<host>", "<username>", "<password>", "<database name>")
cache = Cache()
s = FormulaSolver(db, cache)

cell1 = Cell("'[excel.xlsx]OVERALL SCORES'", "AU", 2)
cell2 = Cell("'[excel.xlsx]OVERALL SCORES'", "AU", 184)

cell_result = s.calculate_cell(cell1)
col_result = s.calculate_column(cell1, cell2)
```

However, because it is a relational database, you can query the cells to find the locations of the cells you want to calculate. For example, to calculate the cell in the column called "2017 Rank Score" and the row "T US EQUITY" you could do the following:

```python
from excelformula import *

db = MySqlDatabase("<host>", "<username>", "<password>", "<database name>")
cache = Cache()
s = FormulaSolver(db, cache)

header_cell = db.find_text("2017 Rank Score")[0]

# need to filter down row cells because "T US EQUITY" appears on many sheets
row_cells = db.find_text("T US EQUITY")
row_cell = list(filter(lambda x: x.sheet == header_cell.sheet, row_cells))[0]

target_cell = Cell(header_cell.sheet, header_cell.col, target_cell.row)

result = s.calculate_cell(target_cell)
```

The querying can also be used to find the final row in a column in order to calculate the entire column. For example, to calculate the entire "2017 Rank Score" column:

```python
from excelformula import *

db = MySqlDatabase("<host>", "<username>", "<password>", "<database name>")
cache = Cache()
s = FormulaSolver(db, cache)

header_cell = db.find_text("2017 Rank Score")[0]
start_cell = header_cell.increment_row()
end_cell = db.get_final_cell_in_col(header_cell)

col_result = s.calculate_column(cell1, cell2)
```

> **Note that there is currently an issue with `formulas` that causes the output dictionary to not include some column headers. All the data is intact, but some column headers are skipped randomly. I have tried to address this issue, but currently to no avail.**

## Using your own databse

By mocking the interface of `JsonDatabase` or `SqlDatabase`, you can implement the package for any database:

```python
from excelformula import *

class MyCustomDatabase:
    def __init__(self):
        # initialise your database
        pass
        
    def get(self, cell):
        # return your cell's value
        return ""
        
db = MyCustomDatabase()
cache = Cache()
s = FormulaSolver(db, cache)
```

The same can be done with a custom cache to control the memory usage precisely.

## Excel Functions

The majority of common Excel functions do work with the `formulas` package that this project uses. However, the following are notable exceptions:

- `INDIRECT()` does not work, because it would involve rewriting the system for evaluating dependencies to include querying the database. Formula evalutaion is currently provided by `formulas` not `excelformula` where the database is connected.
- `FILTER()` again does not work because it would involve significantly changing the underlying formula interpreter to support it.

> Note that columns DO and DP of the Processed Data sheet do not work. There appears to be a problem with the actual formulae here, as they appear blank in Excel and raise a `FormulaError` when parsed. To fix this, just remove the contents of these cells.

Most other functions do work. I have added several functions in a fork of `formulas` which is accessible [here](https://github.com/benrobinson16/formulas).

## Internal Implementation

Internally, the `FormulaSolver` class is responsible for evaluating a cell. It does so recursively, evaluating the dependencies of the cell in turn before running its own compiled function. As mentioned, `formulas` provides the functionality of actually running a formula, and also of identifying the necessary dependencies.

Both ranges and individual cell dependencies need to be dealt with, and `FormulaSolver` exapands ranges of cells into a list of cell locations that can then be evaluated in turn.

Loading the excel file in is currently a performance bottleneck. This is managed by the `formulas` package, and despite investigating the issue, I have not been able to improve it.

Loading is *much* faster (a couple of seconds instead of minutes) when the data is loaded from a JSON file or database. Hence, it is prefereable to ditch the excel file format as soon as possible, and converting to json/database.

To improve evalutaion performance of my dependency evaluation system, memoization is used with a cache dictionary. This ensure that the cell is not calculated more than once as future calls hit the cache instead of requiring re-computation. This has a huge impact particularly when calculating an entire column, each cell of which relies on the entirety of another column.
