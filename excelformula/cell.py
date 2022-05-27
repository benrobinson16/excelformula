import formulas.tokens.operand


class Constant:
    def __init__(self, key):
        self.key = key


    def formKey(self):
        return self.key


class Cell:
    def __init__(self, sheet, col, row):
        self.sheet = sheet
        self.col = col
        self.row = int(row)


    @staticmethod
    def make_from_key(key):
        components = key.split('!')
        sheet = components[0]

        if sheet == "'[excel.xlsx]'":
            # named constants don't have a sheet
            return Constant(key)

        col = ""
        row = ""
        for c in components[1]:
            if c.isdigit():
                row += c
            else:
                col += c

        return Cell(sheet, col, row)


    @staticmethod
    def make_from_range(range):
        components = range.split('!')
        sheet = components[0]
        cells = []

        for part in components[1].split(':'):
            col = ""
            row = ""
            for c in part:
                if c.isdigit():
                    row += c
                else:
                    col += c

            cells.append(Cell(sheet, col, row))

        return cells
        

    def increment_row(self):
        return Cell(self.sheet, self.col, self.row + 1)


    def increment_col(self):
        return Cell(self.sheet, self.nextCol(), self.row)


    def next_col(self):
        val = formulas.tokens.operand._col2index(self.col)
        val += 1
        return formulas.tokens.operand._index2col(val)

    
    def set_col(self, newCol):
        return Cell(self.sheet, newCol, self.row)


    def set_row(self, newRow):
        return Cell(self.sheet, self.col, newRow)
        

    def form_key(self):
        return f"{self.sheet}!{self.col}{self.row}"


    def equals(self, other):
        return self.sheet == other.sheet and self.col == other.col and self.row == other.row