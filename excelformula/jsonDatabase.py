import json
import formulas


def export_to_json(input_path, output_file):
    xl_model = formulas.ExcelModel().loads(input_path).finish()
    xl_dict = xl_model.to_dict()

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(xl_dict, f, ensure_ascii=False, indent=4)


class JsonDatabase:
    def __init__(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            self.xl_dict = json.load(f)


    def get(self, cell):
        if cell.formKey() in self.xl_dict:
            return self.xl_dict[cell.form_key()]
        else:
            return ""