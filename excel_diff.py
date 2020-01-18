__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

import xlrd
import csv
import difflib
import re
import xlsxwriter
import os # mkdir, path.exists
import shutil # rmtree
import glob
import ntpath
import argparse
import logger

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

def check_change_sub(line, out_file):
    # TODO(sasiala): change to similar format to new/deleted lines
    is_change_sub = re.match('^\- .*\n\?.*\n\+ .*$', line)
    
    if is_change_sub:
        split_lines = re.split('\n\?.*\n\+ ', line)
        first_line = re.sub('^- ', '', split_lines[0]).split(',')
        second_line = split_lines[1].split(',')
        iter_size = len(first_line)
        if (len(first_line) > len(second_line)):
            iter_size = len(second_line)
        
        # TODO(sasiala): deal with lines of diff sizes (skipping rest of output, currently)
        new_list = []
        for col in range(iter_size):
            if not first_line[col] == second_line[col]:
                new_list.append('||'.join([first_line[col], second_line[col]]))
            else:
                new_list.append(first_line[col])
        out_file.write('Change/Sub,')
        out_file.write(','.join(new_list))
        out_file.write('\n')
        return True
    return False

def check_change_add(line, out_file):
    # TODO(sasiala): change to similar format to new/deleted lines
    is_change_add = re.match('^- .*\n\+ .*\n\? .*$', line)

    if is_change_add:
        split_lines = re.split('\n\?.*$', line)
        temp_lines = split_lines[0].split('\n')
        first_line = re.sub('^- ', '', temp_lines[0]).split(',')
        second_line = re.sub('^\+ ', '', temp_lines[1]).split(',')
        iter_size = len(first_line)
        if (len(first_line) > len(second_line)):
            iter_size = len(second_line)

        # TODO(sasiala): deal with lines of diff sizes (skipping rest of output, currently)
        new_list = []
        for col in range(iter_size):
            if not first_line[col] == second_line[col]:
                new_list.append('||'.join([first_line[col], second_line[col]]))
            else:
                new_list.append(first_line[col])
        out_file.write('Change/Add,')
        out_file.write(','.join(new_list))
        out_file.write('\n')
        return True
    return False

def check_change_add_and_sub(line, out_file):
    is_add_and_sub = re.match('^- .*\n\? .*\n\+ .*\n\? .*$', line)
    
    if is_add_and_sub:
        line = line + '\n'
        lines = re.split('\n\? .*[\n|$]', line)
        # TODO(sasiala): split on "," or similar, instead of , (need to think about strings w/ comma)
        first_line = re.sub('^- ', '', lines[0]).split(',')
        second_line = re.sub('^\+ ', '', lines[1]).split(',')
        iter_size = len(first_line)
        if (len(first_line) > len(second_line)):
            iter_size = len(second_line)

        # TODO(sasiala): deal with lines of diff sizes (skipping rest of output, currently)
        new_list = []
        for col in range(iter_size):
            if not first_line[col] == second_line[col]:
                new_list.append('||'.join([first_line[col], second_line[col]]))
            else:
                new_list.append(first_line[col])
        out_file.write('Change/Add/Sub,')
        out_file.write(','.join(new_list))
        out_file.write('\n')
        return True
    return False


def check_new_line(line, out_file):
    is_new_line = re.match('\+ .*$', line)

    if is_new_line:
        split_lines = line.split('\n')
        for i in split_lines:
            if not re.match('^\+ $', i):
                out_file.write('New Line,')
                out_file.write(re.sub('^\+ ', '', i))
                out_file.write('\n')
        return True
    return False

def check_deleted_line(line, out_file):
    is_deleted_line = re.match('^- .*$', line)

    if is_deleted_line:
        split_lines = line.split('\n')
        for i in split_lines:
            if not re.match('^- $', i):
                out_file.write('Deleted Line,')
                out_file.write(re.sub('^- ', '', i))
                out_file.write('\n')
        return True
    return False

def check_compound(line, out_file):
    line = line.split('\n')
    left_over = line
    if (len(line) >= 4):
        # TODO(sasiala): am I sure these can't be in the middle of the string?
        if (check_change_add_and_sub('\n'.join(line[0:4]), out_file)):
            left_over = line[4:]
            logger.log('diff_to_sheet.log', 'Compound:Change/Add/Sub')
        elif (check_change_add('\n'.join(line[0:3]), out_file)):
            left_over = line[3:]
            logger.log('diff_to_sheet.log', 'Compound:Change/Add')
        elif (check_change_sub('\n'.join(line[0:3]), out_file)):
            left_over = line[3:]
            logger.log('diff_to_sheet.log', 'Compound:Change/Sub')
        
        for i in left_over:
            if (check_new_line(i, out_file)):
                logger.log('diff_to_sheet.log', 'Compound:New Line')
            elif (check_deleted_line(i, out_file)):
                logger.log('diff_to_sheet.log', 'Compound:Deleted Line')
            elif re.match('- $', i) or re.match('^\+ $', i):
                logger.log('diff_to_sheet.log', 'Compound:Skipped empty +/- line')
            else:
                # unexpected format in diff
                logger.log('diff_to_sheet.log', 'Compound:Curious (unexpected diff format)...')
                logger.log('diff_to_sheet.log', i)
                logger.log('diff_to_sheet.log', '/Curious')
                # TODO(sasiala): return False
        return True
    return False

def diff_to_sheet(csv_diff_path, out_path):
    with open(out_path, 'w') as out_file:
        with open(csv_diff_path, 'r') as csv_diff:
            lines = csv_diff.read().split('\n  \n')
            for line in lines:
                change_sub_split = re.split('\n\?.*\n\+ ', line) # TODO

                if check_change_sub(line, out_file):
                    logger.log('diff_to_sheet.log', 'Change/Sub')
                elif check_change_add(line, out_file):
                    logger.log('diff_to_sheet.log', 'Change/Add')
                elif check_change_add_and_sub(line, out_file):
                    logger.log('diff_to_sheet.log', 'Change/Add/Sub')
                elif check_compound(line, out_file):
                    logger.log('diff_to_sheet.log', 'Compound')
                elif len(change_sub_split) == 1:
                    if len(line.split('  ')) == 2:
                        logger.log('diff_to_sheet.log', 'No Change')
                        out_file.write('No Change,')
                        out_file.write(line.split('  ')[1])
                        out_file.write('\n')
                else:
                    # unexpected format in diff
                    logger.log('diff_to_sheet.log', 'Curious (unexpected diff format)...')
                    logger.log('diff_to_sheet.log', line)
                    logger.log('diff_to_sheet.log', '/Curious')
                    # TODO(sasiala): return False
            return True
        return False
    return False

def csv_to_xlsx(csv_path, xlsx_path):
    # TODO(sasiala): convert to only add sheets in this function
    workbook = xlsxwriter.Workbook(xlsx_path)
    worksheet = workbook.add_worksheet()
    change_add_format = workbook.add_format({'bold':True, 'bg_color':'green'})
    change_sub_format = workbook.add_format({'bold':True, 'bg_color':'red'})
    change_add_sub_format = workbook.add_format({'bold':True, 'bg_color':'pink'})
    no_change_format = workbook.add_format()
    with open(csv_path, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        for r, row in enumerate(csv_reader):
            temp = list(enumerate(row))
            line_format = change_add_format
            if not len(temp) == 0:
                if (temp[0][1] == 'Change/Sub'):
                    logger.log('csv_to_xlsx.log', 'Change/Sub')
                    line_format = change_sub_format
                elif (temp[0][1] == 'Change/Add'):
                    logger.log('csv_to_xlsx.log', 'Change/Add')
                elif (temp[0][1] == 'Change/Add/Sub'):
                    logger.log('csv_to_xlsx.log', 'Change/Add/Sub')
                    line_format = change_add_sub_format
                elif (temp[0][1] == 'New Line'):
                    logger.log('csv_to_xlsx.log', 'New Line')
                elif (temp[0][1] == 'Deleted Line'):
                    logger.log('csv_to_xlsx.log', 'Deleted Line')
                    line_format = change_sub_format
                elif (temp[0][1] == 'No Change'):
                    logger.log('csv_to_xlsx.log', 'No Change')
                    line_format = no_change_format
                else:
                    logger.log('csv_to_xlsx.log', 'Curious...')
            
            for c, col in enumerate(row):
                worksheet.write(r, c, col, line_format)
                
        workbook.close()
        return True
    workbook.close()
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
    if not diff_to_sheet(TEMP_FOLDER + '/csv_diff/sheet_0.diff', TEMP_FOLDER + '/diff_sheets/sheet_0.csv'):
        print("Diff to sheet failed")
        return False
    if not csv_to_xlsx(TEMP_FOLDER + '/diff_sheets/sheet_0.csv', OUTPUT_FOLDER + '/final_out.xlsx'):
        print("CSV to XLSX failed")
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