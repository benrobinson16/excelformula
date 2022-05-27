class Cache:
    def __init__(self):
        self.dict = {}


    def get(self, cell):
        if cell.formKey() in self.dict:
            return self.dict[cell.formKey()]
        else:
            return None


    def set(self, cell, val):
        self.dict[cell.formKey()] = val