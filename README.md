# `excelformula`

Database integration of Excel formulae. Supports importing excel files, exporting as JSON or SQL databases, calculating individual cells and calculating columns.

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

## Using your own databse

By mocking the interface of `JsonDatabase` or `SqlDatabase`, you can implement the package for any database:

```python
from excelformula import *

class MyCustomDatabase:
    def __init__(self):
        # initialis your database
        pass
        
    def get(self, cell):
        # return your cell's value
        return ""
        
db = MyCustomDatabase()
cache = Cache()
s = FormulaSolver(db, cache)
```

The same can be done with a custom cache to control the memory usage precisely.
