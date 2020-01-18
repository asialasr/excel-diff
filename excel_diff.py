__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

import xlrd
import csv
import difflib
import os # mkdir, path.exists
import shutil # rmtree
import glob
import ntpath
import argparse
import logger
import CsvDiffToSheet as cdts
import SheetDiffToXlsx as sdtx

TEMP_FOLDER='temp'
OUTPUT_FOLDER='output'

def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

def sheet_to_csv(sheet, out_path):
    with open(out_path, 'w') as temp_csv:
        wr = csv.writer(temp_csv, quoting=csv.QUOTE_ALL)

        for row_num in range(sheet.nrows):
            wr.writerow(sheet.row_values(row_num))
        
        temp_csv.close()
        return True
    return False

def csv_diff(csvl, csvr, out_path):
    with open(csvl, 'r') as csvl_file:
        with open(csvr, 'r') as csvr_file:
            d = difflib.Differ()
            diff = difflib.ndiff(csvl_file.read().splitlines(1), csvr_file.read().splitlines(1))
    with open(out_path, 'w') as diff_file:
        for line in diff:
            diff_file.writelines(line)
        diff_file.close()
        return True
    return False

def remove_temp_directories():
    if os.path.exists(TEMP_FOLDER):
        shutil.rmtree(TEMP_FOLDER)

def setup_temp_directories():
    remove_temp_directories()

    os.mkdir(TEMP_FOLDER)
    os.mkdir(TEMP_FOLDER + '/lhs')
    os.mkdir(TEMP_FOLDER + '/rhs')
    os.mkdir(TEMP_FOLDER + '/diff_sheets')
    os.mkdir(TEMP_FOLDER + '/csv_diff')

def setup_output_directory():
    if not os.path.exists(OUTPUT_FOLDER):
        os.mkdir(OUTPUT_FOLDER)

def process_xlsx(lhs_path, rhs_path):
    setup_temp_directories()
    logger.initialize_directory_structure()
    setup_output_directory()

    left_temp_path = TEMP_FOLDER + '/lhs'
    right_temp_path = TEMP_FOLDER + '/rhs'
    with xlrd.open_workbook(lhs_path) as xlsx_file:
        for sheet_num in range(xlsx_file.nsheets):
            # TODO(sasiala): match sheets up so that diff is complete
            sheet_path = left_temp_path + '/sheet_' + str(sheet_num) + '.csv'
            if not sheet_to_csv(xlsx_file.sheet_by_index(sheet_num), sheet_path):
                print("Sheet to csv failed")
                return False

    with xlrd.open_workbook(rhs_path) as xlsx_file:
        for sheet_num in range(xlsx_file.nsheets):
            # TODO(sasiala): match sheets up so that diff is complete
            sheet_path = right_temp_path + '/sheet_' + str(sheet_num) + '.csv'
            if not sheet_to_csv(xlsx_file.sheet_by_index(sheet_num), sheet_path):
                print("Sheet to csv failed (right)")
                return False
    
    temp_lhs_sheets = glob.glob(TEMP_FOLDER + '/lhs/sheet_*.csv')
    lhs_filenames = [path_leaf(i) for i in temp_lhs_sheets]
    temp_rhs_sheets = glob.glob(TEMP_FOLDER + '/rhs/sheet_*.csv')
    rhs_filenames = [path_leaf(i) for i in temp_rhs_sheets]

    if not csv_diff(TEMP_FOLDER + '/lhs/sheet_0.csv', TEMP_FOLDER + '/rhs/sheet_0.csv', TEMP_FOLDER + '/csv_diff/sheet_0.diff'):
        print("Csv diff failed")
        return False
    if not cdts.diff_to_sheet(TEMP_FOLDER + '/csv_diff/sheet_0.diff', TEMP_FOLDER + '/diff_sheets/sheet_0.csv'):
        print("Diff to sheet failed")
        return False
    if not sdtx.csv_to_sheet(TEMP_FOLDER + '/diff_sheets/sheet_0.csv', OUTPUT_FOLDER + '/final_out.xlsx'):
        print("CSV to Sheet failed")
        return False

    # TODO(sasiala): remove temp directories, also remove at other returns
    # remove_temp_directories()
    return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('lhspath')
    parser.add_argument('rhspath')
    args = parser.parse_args()
    
    process_xlsx(args.lhspath, args.rhspath)

if __name__ == "__main__":
    # execute only if run as a script
    main()