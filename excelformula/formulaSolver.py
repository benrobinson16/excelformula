from .cell import Cell, Constant
import formulas
import numpy


class FormulaSolver:
    def __init__(self, db, cache):
        self.db = db
        self.cache = cache


    def depndencies(self, range):
        components = range.split(':')
        if len(components) == 1:
            return Cell.make_from_key(components[0])
        else:
            edgeCells = Cell.make_from_range(range)
            first = edgeCells[0]
            last = edgeCells[1]

            if first.row == last.row:
                output = [first]
                while not output[-1].equals(last):
                    output.append(output[-1].increment_col())
                return output
            else:
                output = [first]
                while not output[-1].equals(last):
                    output.append(output[-1].increment_row())
                return output


    def calculate_cell(self, cell):
        val = self.db.get(cell)
        if isinstance(val, str):
            if len(val) > 0 and val[0] == '=':
                if self.cache.get(cell) is not None:
                    return self.cache.get(cell)
                func = formulas.Parser().ast(val)[1].compile()
                inputs = list(func.inputs)
                input_values = {}
                for input in inputs:
                    cells = self.depndencies(input)
                    if isinstance(cells, list):
                        y = []
                        for c in cells:
                            v = self.calculateCell(c)
                            if isinstance(v, numpy.ndarray):
                                y.append(v.item())
                            else:
                                y.append(v)
                        input_values[input] = numpy.array(y, dtype=object)
                    else:
                        input_values[input] = self.calculate_cell(cells)
                        if isinstance(input_values[input], list):
                            input_values[input] = numpy.asarray(input_values[input])
                result = func(**input_values)
                self.cache.set(cell, result)
                return result
            else:
                return val
        else:
            return val


    def calculate_column(self, start, end, labels=None):
        output = []
        cell = start

        while not cell.equals(end):
            v = self.calculate_cell(cell)
            if isinstance(v, numpy.ndarray):
                output.append(v.item())
            else:
                output.append(v)
            cell = cell.increment_row()

        if labels is not None:
            output = list(zip(labels, output))

        return output