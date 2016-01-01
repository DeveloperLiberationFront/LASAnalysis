from __future__ import unicode_literals

import openpyxl as pxl
import re
import csv
import os

import formula_list as fl
import config

formulas = fl.get_formula_list()

# Pattern for formulas.
formula_pattern = re.compile('[A-Z][A-Z\.0-9]*\(')

# Dictionary to store formula used per file.
file_formula_usage = {}

# Load the setting to be used.
setting = config.Q45


def get_formulas_from_file(filename):
    try:
        # Formulas in this file.
        file_formulas = []

        wb = pxl.load_workbook(filename)
        # Get the list of sheets in the workbook.
        sheet_names = wb.get_sheet_names()
        for sheet_name in sheet_names:
            sheet = wb.get_sheet_by_name(sheet_name)

            # Iterate over the cells.
            for row in sheet.iter_rows():
                for cell in row:
                    cell_string = unicode(cell.value)
                    formulas_used = extract_formula_from_text(cell_string)

                    for formula in formulas_used:
                        if not (formula in file_formulas) and (formula in formulas):
                            file_formulas.append(formula)
    except Exception:
        print "Error in file: " + filename
    finally:
        # Add the formula to the dictionary storage with file as key.
        file_formula_usage[filename.replace(setting['prefix'], "")] = ', '.join(map(str, file_formulas))

    return


def extract_formula_from_text(cell_string):
    formulas_used = []
    if cell_string.startswith('='):
        # Find all matches iteratively.
        for match in re.finditer(formula_pattern, cell_string):
            if not is_in_quotes(match.start(), cell_string):
                # Get rid of the trailing '('.
                formula = match.group()[0:-1]
                formulas_used.append(formula)

    return formulas_used


def is_in_quotes(start, string):
    is_in_single_quote = False
    is_in_double_quote = False

    for index in xrange(0, start+1):
        if string[index] == '"':
            is_in_double_quote = not is_in_double_quote
        elif string[index] == '\'':
            is_in_single_quote = not is_in_single_quote

    return is_in_single_quote or is_in_double_quote


def write_output():
    with open(setting['output_file'], 'wb') as output_file:
        csv_writer = csv.writer(output_file, dialect='excel')

        # Write headers.
        csv_writer.writerow(["Excel File", "Formulas Used"])

        for key, value in file_formula_usage.iteritems():
            csv_writer.writerow([key, value])

    return


def get_formulas_from_dir():
    for input_file in os.listdir(setting['dir']):
        if input_file.endswith(".xlsx") or input_file.endswith(".xlsm") or input_file.endswith(".xls"):
            file_path = os.path.join(setting['dir'], input_file)
            get_formulas_from_file(file_path)

    write_output()
    return


def main():
    get_formulas_from_dir()
    return 0


if __name__ == '__main__':
    main()
