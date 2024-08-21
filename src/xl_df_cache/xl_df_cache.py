import pandas as pd
from openpyxl import Workbook
import re
import warnings


def clean_excel_sheet_name(name):
    # Remove any characters that are not alphanumeric, underscore, or hyphen
    cleaned_name = re.sub(r"[^\w\-]", "_", name)

    # Ensure that the name starts with a letter
    if not cleaned_name[0].isalpha():
        cleaned_name = "S" + cleaned_name

    # Limit the name to a maximum of 31 characters
    cleaned_name = cleaned_name[:31]

    return cleaned_name


def disambiguate_key(key, keys):
    counter = 1
    while key in keys:
        modified_key = f"{key[:27]}_{counter}"
        if len(modified_key) <= 31:
            key = modified_key
        else:
            key = f"{key[:31 - len(str(counter)) - 1]}_{counter}"
        counter += 1
    return key


class XlDfCache(dict):
    def __setitem__(self, key, value):
        key = clean_excel_sheet_name(key)

        if key in self:
            key = disambiguate_key(key, self.keys())

        super().__setitem__(key.strip(), value)

    def write(self, filename):
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            for sheet, df in self.items():
                df.to_excel(writer, sheet_name=sheet)
