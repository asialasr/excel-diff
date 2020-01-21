__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

import xlrd
import csv
import difflib
import os # mkdir, path.exists
import shutil # rmtree
import ntpath
import argparse
import logger
import re
import xlsxwriter
import CsvDiffToSheet as cdts
import SheetDiffToXlsx as sdtx

TEMP_FOLDER='temp'
OUTPUT_FOLDER='output'

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

def generate_csvs_for_xlsx(xlsx_path, temp_path):
    with xlrd.open_workbook(xlsx_path) as xlsx_file:
        with open(f'{temp_path}/sheet_names.txt', 'w') as sheet_name_file:
            sheet_name_file.write('\n'.join(xlsx_file.sheet_names()))
            for sheet_num in range(xlsx_file.nsheets):
                # TODO(sasiala): match sheets up so that diff is complete
                sheet = xlsx_file.sheet_by_index(sheet_num)
                sheet_path = f'{temp_path}/{sheet.name}.csv'
                if not sheet_to_csv(xlsx_file.sheet_by_index(sheet_num), sheet_path):
                    print(f'Sheet to csv failed on {sheet_path}')
                    return False
            return True
        return False
    return False

def get_unified_sheets(lhs_sheet_file, lhs_temp_path, rhs_sheet_file, rhs_temp_path):
    unified = []
    with open(lhs_sheet_file, 'r') as lhs_sheets:
        with open(rhs_sheet_file, 'r') as rhs_sheets:
            with open('temp_out.txt', 'w') as out_file:
                d = difflib.Differ()
                diff_gen = difflib.ndiff(lhs_sheets.read().splitlines(1), rhs_sheets.read().splitlines(1))
                diff = [i.rstrip() for i in diff_gen]
                
    skip_next = False
    for i in range(len(diff)):
        if skip_next:
            skip_next = False
            continue

        if re.match('^  .*$', diff[i]):
            unified.append(['b', re.sub('^  ', '', diff[i])])
        elif re.match('^\+ .*$', diff[i]):
            unified.append(['n', re.sub('^\+ ', '', diff[i])])
        elif re.match('^- .*$', diff[i]):
            this_line_sub = re.sub('^- ', '', diff[i])
            if i < len(diff) - 1 and re.match('^\+ .*$', diff[i+1]):
                # r stands for rename (potential)
                next_line_sub = re.sub('^\+ ', '', diff[i + 1])

                with open(f'{lhs_temp_path}/{this_line_sub}.csv', 'r') as lhs_csv:
                    with open(f'{rhs_temp_path}/{next_line_sub}.csv', 'r') as rhs_csv:
                        seq = difflib.SequenceMatcher(None, lhs_csv.read(), rhs_csv.read())
                
                if seq.quick_ratio() > .5:
                    unified.append(['r', f'{this_line_sub},{next_line_sub}'])
                    # TODO(sasiala): reconsider output format for r; are commas allowed in sheet names?
                    skip_next = True
                else:
                    unified.append(['n', this_line_sub])
            else:
                unified.append(['d', this_line_sub])
        else:
            print('Unexpected format in unified sheets')

    return unified

def process_xlsx(lhs_path, rhs_path):
    setup_temp_directories()
    logger.initialize_directory_structure()
    logger.set_log_level(logger.LogLevel.DEBUG)
    setup_output_directory()

    left_temp_path = TEMP_FOLDER + '/lhs'
    right_temp_path = TEMP_FOLDER + '/rhs'
    generate_csvs_for_xlsx(lhs_path, left_temp_path)
    generate_csvs_for_xlsx(rhs_path, right_temp_path)
    
    # TODO(sasiala): use workbook.sheet_names() for lhs and rhs to see new/deleted sheets
    
    lhs_sheet_names = []
    with open(f'{left_temp_path}/sheet_names.txt', 'r') as sheet_name_file:
        lhs_sheet_names = sheet_name_file.read().split('\n')
    
    rhs_sheet_names = []
    with open(f'{right_temp_path}/sheet_names.txt', 'r') as sheet_name_file:
        rhs_sheet_names = sheet_name_file.read().split('\n')

    get_unified_sheets(f'{left_temp_path}/sheet_names.txt', left_temp_path, f'{right_temp_path}/sheet_names.txt', right_temp_path)

    # TODO(sasiala): this doesn't account for missing sheets in one book
    xlsx_path = f'{OUTPUT_FOLDER}/final_out.xlsx'
    workbook = xlsxwriter.Workbook(xlsx_path)
    for i in range(len(lhs_sheet_names)):
        if not csv_diff(f'{TEMP_FOLDER}/lhs/{lhs_sheet_names[i]}.csv', f'{TEMP_FOLDER}/rhs/{rhs_sheet_names[i]}.csv', f'{TEMP_FOLDER}/csv_diff/{lhs_sheet_names[i]}.diff'):
            print("Csv diff failed")
            return False
        if not cdts.diff_to_sheet(f'{TEMP_FOLDER}/csv_diff/{lhs_sheet_names[i]}.diff', f'{TEMP_FOLDER}/diff_sheets/{lhs_sheet_names[i]}.csv'):
            print("Diff to sheet failed")
            return False
        if not sdtx.csv_to_sheet(workbook, f'{TEMP_FOLDER}/diff_sheets/{lhs_sheet_names[i]}.csv', lhs_sheet_names[i]):
            print("CSV to Sheet failed")
            return False
    workbook.close()
    
    # TODO(sasiala): change to merging sheets into new workbook, instead of outputting separate workbooks

    # TODO(sasiala): remove temp directories, also remove at other returns
    # remove_temp_directories()
    return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('lhspath')
    parser.add_argument('rhspath')
    args = parser.parse_args()
    
    # TODO(sasiala): add logging command line option
    process_xlsx(args.lhspath, args.rhspath)

if __name__ == "__main__":
    # execute only if run as a script
    main()