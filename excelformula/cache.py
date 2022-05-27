class Cache:
    def __init__(self):
        self.dict = {}


    def get(self, cell):
        if cell.form_key() in self.dict:
            return self.dict[cell.form_key()]
        else:
            return None


    def set(self, cell, val):
        self.dict[cell.form_key()] = val