__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

import xlrd
import csv
import difflib
import os # mkdir, path.exists
import shutil # rmtree
import argparse
import logger
import re
import xlsxwriter
import CsvDiffToSheet as cdts
import SheetDiffToXlsx as sdtx

# TODO(sasiala): raise exceptions at errors

TEMP_FOLDER='temp'
LHS_TEMP_FOLDER=TEMP_FOLDER+'/lhs'
RHS_TEMP_FOLDER=TEMP_FOLDER+'/rhs'
SHEET_DIFF_TEMP_FOLDER=TEMP_FOLDER+'/diff_sheets'
CSV_DIFF_TEMP_FOLDER=TEMP_FOLDER+'/csv_diff'
OUTPUT_FOLDER='output'
GUI_OUTPUT_FOLDER=f'{OUTPUT_FOLDER}/gui_out'

save_temp=False

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
    os.mkdir(LHS_TEMP_FOLDER)
    os.mkdir(RHS_TEMP_FOLDER)
    os.mkdir(SHEET_DIFF_TEMP_FOLDER)
    os.mkdir(CSV_DIFF_TEMP_FOLDER)

def remove_output_directories():
    if os.path.exists(OUTPUT_FOLDER):
        shutil.rmtree(OUTPUT_FOLDER)

def setup_output_directory(gui_output):
    remove_output_directories()

    if not os.path.exists(OUTPUT_FOLDER):
        os.mkdir(OUTPUT_FOLDER)
    if gui_output and not os.path.exists(GUI_OUTPUT_FOLDER):
        os.mkdir(GUI_OUTPUT_FOLDER)

def generate_csvs_for_xlsx(xlsx_path, temp_path):
    with xlrd.open_workbook(xlsx_path) as xlsx_file:
        with open(f'{temp_path}/sheet_names.txt', 'w') as sheet_name_file:
            sheet_name_file.write('\n'.join(xlsx_file.sheet_names()))
            for sheet_num in range(xlsx_file.nsheets):
                sheet = xlsx_file.sheet_by_index(sheet_num)
                sheet_path = f'{temp_path}/{sheet.name}.csv'
                if not sheet_to_csv(xlsx_file.sheet_by_index(sheet_num), sheet_path):
                    print(f'Sheet to csv failed on {sheet_path}')
                    return False
            return True
        return False
    return False

def split_intermediate_renamed_sheet(intermediate_sheet_name):
    return intermediate_sheet_name.split(']')

def format_intermediate_renamed_sheet(lhs_name, rhs_name):
    return f'{lhs_name}]{rhs_name}'

EXISTING_SHEET='b'
RENAMED_SHEET='r'
NEW_SHEET='n'
DELETED_SHEET='d'

def get_unified_sheets(lhs_sheet_file, lhs_temp_path, rhs_sheet_file, rhs_temp_path):
    unified = []
    with open(lhs_sheet_file, 'r') as lhs_sheets:
        with open(rhs_sheet_file, 'r') as rhs_sheets:
            d = difflib.Differ()
            diff_gen = difflib.ndiff(lhs_sheets.read().splitlines(1), rhs_sheets.read().splitlines(1))
            diff = [i.rstrip() for i in diff_gen]
                
    skip_next = False
    for i in range(len(diff)):
        if skip_next:
            skip_next = False
            continue

        if re.match('^  .*$', diff[i]):
            unified.append([EXISTING_SHEET, re.sub('^  ', '', diff[i])])
        elif re.match('^\+ .*$', diff[i]):
            unified.append([NEW_SHEET, re.sub('^\+ ', '', diff[i])])
        elif re.match('^- .*$', diff[i]):
            this_line_sub = re.sub('^- ', '', diff[i])
            if i < len(diff) - 1 and re.match('^\+ .*$', diff[i+1]):
                # r stands for rename (potential)
                next_line_sub = re.sub('^\+ ', '', diff[i + 1])

                with open(f'{lhs_temp_path}/{this_line_sub}.csv', 'r') as lhs_csv:
                    with open(f'{rhs_temp_path}/{next_line_sub}.csv', 'r') as rhs_csv:
                        seq = difflib.SequenceMatcher(None, lhs_csv.read(), rhs_csv.read())
                
                if seq.quick_ratio() > .5:
                    unified.append([RENAMED_SHEET, format_intermediate_renamed_sheet(this_line_sub, next_line_sub)])
                    skip_next = True
                else:
                    unified.append([DELETED_SHEET, this_line_sub])
            else:
                unified.append([DELETED_SHEET, this_line_sub])
        else:
            print('Unexpected format in unified sheets')

    return unified

def format_unchanged_sheetname(sheet_name):
    return f'_UNCH_{sheet_name}'

def format_renamed_sheetname(lhs_name, rhs_name):
    # prints length of original sheet name to indicate where split occurs
    return f'_RNM{len(lhs_name)}_{lhs_name}_{rhs_name}'

def check_modified_sheet(workbook, sheet_pair, output_sheet_path, sheet_name):
    if sheet_pair[0]==EXISTING_SHEET or sheet_pair[0]==RENAMED_SHEET:
        left_csv_path=''
        right_csv_path=''
        if sheet_pair[0]==EXISTING_SHEET:
            left_csv_path=f'{LHS_TEMP_FOLDER}/{sheet_pair[1]}.csv'
            right_csv_path=f'{RHS_TEMP_FOLDER}/{sheet_pair[1]}.csv'
            sheet_name.append(format_unchanged_sheetname(sheet_pair[1]))
        else:
            temp_sheet_names=split_intermediate_renamed_sheet(sheet_pair[1])
            left_csv_path=f'{LHS_TEMP_FOLDER}/{temp_sheet_names[0]}.csv'
            right_csv_path=f'{RHS_TEMP_FOLDER}/{temp_sheet_names[1]}.csv'
            sheet_name.append(format_renamed_sheetname(temp_sheet_names[0], temp_sheet_names[1]))

        output_sheet_path.append(f'{SHEET_DIFF_TEMP_FOLDER}/{sheet_name[0]}.csv')

        if not csv_diff(left_csv_path, right_csv_path, f'{CSV_DIFF_TEMP_FOLDER}/{sheet_name[0]}.diff'):
            print("Csv diff failed")
            return False
        if not cdts.diff_to_sheet(f'{CSV_DIFF_TEMP_FOLDER}/{sheet_name[0]}.diff', output_sheet_path[0]):
            print("Diff to sheet failed")
            return False
        return True
    return False

def format_new_sheetname(sheet_name):
    return f'_NEW_{sheet_name}'

def check_new_sheet(workbook, sheet_pair, output_sheet_path, sheet_name):
    if sheet_pair[0]==NEW_SHEET:
        sheet_name.append(format_new_sheetname(sheet_pair[1]))
        output_sheet_path.append(f'{RHS_TEMP_FOLDER}/{sheet_pair[1]}_proc.csv')

        # TODO(sasiala): a) will this work correctly? b) may need to fix original csv gen & thus fix everything else
        with open(f'{RHS_TEMP_FOLDER}/{sheet_pair[1]}.csv', 'r') as unprocessed_csv:
            with open(output_sheet_path[0], 'w') as processed_csv:
                processed_csv.write(re.sub('\n\n','\n',unprocessed_csv.read()))
                return True
    return False

def format_deleted_sheetname(sheet_name):
    return f'_DEL_{sheet_name}'

def check_deleted_sheet(workbook, sheet_pair, output_sheet_path, sheet_name):
    if sheet_pair[0]==DELETED_SHEET:
        sheet_name.append(format_deleted_sheetname(sheet_pair[1]))
        output_sheet_path.append(f'{LHS_TEMP_FOLDER}/{sheet_pair[1]}_proc.csv')

        # TODO(sasiala): a) will this work correctly? b) may need to fix original csv gen & thus fix everything else
        with open(f'{LHS_TEMP_FOLDER}/{sheet_pair[1]}.csv', 'r') as unprocessed_csv:
            with open(output_sheet_path[0], 'w') as processed_csv:
                processed_csv.write(re.sub('\n\n','\n',unprocessed_csv.read()))
                return True
    return False

def process_sheet(workbook, sheet_pair):
    # TODO(sasiala): I abhor using lists here, but it made the rest of this code "nicer" (read: easier).
    #                Refactor to make these strings
    output_sheet_path=[]
    sheet_name=[]

    if check_modified_sheet(workbook, sheet_pair, output_sheet_path, sheet_name):
        logger.log(f'excel_diff.log', 'Modified Sheet (mod or renamed)', logger.LogLevel.DEBUG)
    elif check_new_sheet(workbook, sheet_pair, output_sheet_path, sheet_name):
        logger.log('excel_diff.log', 'New Sheet', logger.LogLevel.DEBUG)
    elif check_deleted_sheet(workbook, sheet_pair, output_sheet_path, sheet_name):
        logger.log('excel_diff.log', 'Deleted Sheet', logger.LogLevel.DEBUG)
    else:
        logger.log('excel_diff.log', 'Error: unrecognized sheet modification', logger.LogLevel.ERROR)
        return False

    if not sdtx.csv_to_sheet(workbook, output_sheet_path[0], sheet_name[0]):
        logger.log('excel_diff.log', 'CSV to Sheet failed', logger.LogLevel.ERROR)
        return False
    return True

def generate_gui_sheet_names(unified_sheets):
    with open(f'{GUI_OUTPUT_FOLDER}/sheets.txt', 'w') as fout:
        for i in unified_sheets:
            if i[0]==EXISTING_SHEET:
                fout.write(f'{format_unchanged_sheetname(i[1])}.csv\n')
            elif i[0]==RENAMED_SHEET:
                sheet_names = split_intermediate_renamed_sheet(i[1])
                fout.write(f'{format_renamed_sheetname(sheet_names[0], sheet_names[1])}.csv\n')
            elif i[0]==NEW_SHEET:
                fout.write(f'{format_new_sheetname(i[1])}.csv\n')
            elif i[0]==DELETED_SHEET:
                fout.write(f'{format_deleted_sheetname(i[1])}.csv\n')
            else:
                logger.log('excel_diff.log', 'Error: unrecognized sheet pair type', logger.LogLevel.ERROR)
                return False
        return True
    logger.log('excel_diff.log', 'Error: unable to open sheet_names.txt for gui out', logger.LogLevel.ERROR)
    return False

def get_gui_destination_filename(sheet_pair):
    if sheet_pair[0]==EXISTING_SHEET:
        return f'{format_unchanged_sheetname(sheet_pair[1])}.csv'
    elif sheet_pair[0]==RENAMED_SHEET:
        sheet_names = split_intermediate_renamed_sheet(sheet_pair[1])
        return f'{format_renamed_sheetname(sheet_names[0], sheet_names[1])}.csv'
    elif sheet_pair[0]==NEW_SHEET:
        return f'{format_new_sheetname(sheet_pair[1])}.csv'
    elif sheet_pair[0]==DELETED_SHEET:
        return f'{format_deleted_sheetname(sheet_pair[1])}.csv'
    logger.log('excel_diff.log', 'Error: unrecognized sheet pair type', logger.LogLevel.ERROR)
    return f''

def find_sheet(sheet_pair):
    ret = f''
    if sheet_pair[0]==EXISTING_SHEET:
        ret = f'{SHEET_DIFF_TEMP_FOLDER}/'
    elif sheet_pair[0]==RENAMED_SHEET:
        ret = f'{SHEET_DIFF_TEMP_FOLDER}/'
    elif sheet_pair[0]==NEW_SHEET:
        return f'{RHS_TEMP_FOLDER}/{sheet_pair[1]}.csv'
    elif sheet_pair[0]==DELETED_SHEET:
        return f'{LHS_TEMP_FOLDER}/{sheet_pair[1]}.csv'
    else:
        logger.log('excel_diff.log', 'Error: unrecognized sheet pair type', logger.LogLevel.ERROR)
        return f''
    
    return ret + get_gui_destination_filename(sheet_pair)

def copy_sheets(unified_sheets, source_dir, destination_dir):
    for sheet in unified_sheets:
        sheet_path = find_sheet(sheet)
        if not sheet_path == '':
            shutil.copy(sheet_path, f'{destination_dir}/{get_gui_destination_filename(sheet)}')
        else:
            logger.log('excel_diff.log', 'Error: file not found in copy_sheets', logger.LogLevel.ERROR)
            return False
    return True

def generate_gui_output(unified_sheets):
    generate_gui_sheet_names(unified_sheets)
    copy_sheets(unified_sheets, SHEET_DIFF_TEMP_FOLDER, GUI_OUTPUT_FOLDER)


def process_xlsx(lhs_path, rhs_path, gui_output):
    setup_temp_directories()
    logger.initialize_directory_structure()
    setup_output_directory(gui_output)

    generate_csvs_for_xlsx(lhs_path, LHS_TEMP_FOLDER)
    generate_csvs_for_xlsx(rhs_path, RHS_TEMP_FOLDER)
    
    lhs_sheet_names = []
    with open(f'{LHS_TEMP_FOLDER}/sheet_names.txt', 'r') as sheet_name_file:
        lhs_sheet_names = sheet_name_file.read().split('\n')
    
    rhs_sheet_names = []
    with open(f'{RHS_TEMP_FOLDER}/sheet_names.txt', 'r') as sheet_name_file:
        rhs_sheet_names = sheet_name_file.read().split('\n')

    unified_sheets = get_unified_sheets(f'{LHS_TEMP_FOLDER}/sheet_names.txt', LHS_TEMP_FOLDER, f'{RHS_TEMP_FOLDER}/sheet_names.txt', RHS_TEMP_FOLDER)

    xlsx_path = f'{OUTPUT_FOLDER}/final_out.xlsx'
    workbook = xlsxwriter.Workbook(xlsx_path)
    for sheet_pair in unified_sheets:
        process_sheet(workbook, sheet_pair)
    workbook.close()

    if gui_output:
        generate_gui_output(unified_sheets)

    # TODO(sasiala): add logging

    if not save_temp:
        remove_temp_directories()

    return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('lhspath')
    parser.add_argument('rhspath')
    # TODO(sasiala): rethink argument naming
    parser.add_argument('-v', '--verbose', action='count', default=0, help='Increase Verbosity')
    parser.add_argument('-st', '--save-temp', action='store_true', help='Save all temporary files')
    parser.add_argument('-g', '--gui-out', action='store_true', help='Output files for GUI instead of xlsx')
    args = parser.parse_args()

    try:
        loglevel = logger.LogLevel(args.verbose)
    except:
        loglevel = logger.LogLevel.ERROR
    logger.set_log_level(loglevel)

    global save_temp
    save_temp=args.save_temp

    process_xlsx(args.lhspath, args.rhspath, args.gui_out)

if __name__ == "__main__":
    # execute only if run as a script
    main()