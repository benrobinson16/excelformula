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
    def makeFromKey(key):
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
    def makeFromRange(range):
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
        

    def incrementRow(self):
        return Cell(self.sheet, self.col, self.row + 1)


    def incrementCol(self):
        return Cell(self.sheet, self.nextCol(), self.row)


    def nextCol(self):
        val = formulas.tokens.operand._col2index(self.col)
        val += 1
        return formulas.tokens.operand._index2col(val)

    
    def setCol(self, newCol):
        return Cell(self.sheet, newCol, self.row)


    def setRow(self, newRow):
        return Cell(self.sheet, self.col, newRow)
        

    def formKey(self):
        return f"{self.sheet}!{self.col}{self.row}"


    def equals(self, other):
        return self.sheet == other.sheet and self.col == other.col and self.row == other.row