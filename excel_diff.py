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
    split_lines = re.split('\n\?.*\n\+ ', line)
    
    if len(split_lines) == 2:
        first_line = split_lines[0].split('- ')[1].split(',')
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
    split_lines = re.split('\n\?.*$', line)

    if len(split_lines) == 2:
        temp_lines = split_lines[0].split('\n')
        first_line = temp_lines[0].split('- ')[1].split(',')
        second_line = re.split('\+ ', temp_lines[1])[1].split(',')
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

def check_new_line(line, out_file):
    split_lines = re.split('\+ ', line)

    if len(split_lines) == 4:
        out_file.write('New Line,')
        out_file.write(split_lines[2])
        out_file.write('\n')
        return True
    return False

def check_deleted_line(line, out_file):
    split_lines = re.split('- ', line)

    if len(split_lines) == 4:
        out_file.write('Deleted Line,')
        out_file.write(split_lines[2])
        out_file.write('\n')
        return True
    return False

def diff_to_sheet(csv_diff_path, out_path):
    with open(out_path, 'w') as out_file:
        with open(csv_diff_path, 'r') as csv_diff:
            lines = csv_diff.read().split('\n  \n')
            for line in lines:
                change_sub_split = re.split('\n\?.*\n\+ ', line) # TODO

                if check_change_sub(line, out_file):
                    print('Change/Sub')
                elif check_change_add(line, out_file):
                    print('Change/Add')
                elif check_new_line(line, out_file):
                    print('New Line')
                elif check_deleted_line(line, out_file):
                    print('Deleted Line')
                elif len(change_sub_split) == 1:
                    if len(line.split('  ')) == 2:
                        print('No Change')
                        out_file.write('No Change,')
                        out_file.write(line.split('  ')[1])
                        out_file.write('\n')
                else:
                    # unexpected format in diff
                    csv_diff.close()
                    out_file.close()
                    print(len(temp))
                    print(temp)
                    print(line)
                    return False
            csv_diff.close()
            out_file.close()
            return True
        out_file.close()
        return False
    return False

def csv_to_xlsx(csv_path, xlsx_path):
    # TODO(sasiala): convert to only add sheets in this function
    workbook = xlsxwriter.Workbook(xlsx_path)
    worksheet = workbook.add_worksheet()
    change_add_format = workbook.add_format({'bold':True, 'bg_color':'green'})
    change_sub_format = workbook.add_format({'bold':True, 'bg_color':'red'})
    no_change_format = workbook.add_format()
    with open(csv_path, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        for r, row in enumerate(csv_reader):
            temp = list(enumerate(row))
            line_format = change_add_format
            if not len(temp) == 0:
                if (temp[0][1] == 'Change/Sub'):
                    print('Change/Sub')
                    line_format = change_sub_format
                elif (temp[0][1] == 'Change/Add'):
                    print('Change/Add')
                elif (temp[0][1] == 'New Line'):
                    print('New Line')
                elif (temp[0][1] == 'Deleted Line'):
                    print('Deleted Line')
                    line_format = change_sub_format
                elif (temp[0][1] == 'No Change'):
                    print('No Change')
                    line_format = no_change_format
                else:
                    print('Curious...')
            
            for c, col in enumerate(row):
                worksheet.write(r, c, col, line_format)
                
        workbook.close()
        return True
        workbook.close()
    return False

def remove_temp_directories():
    if os.path.exists('temp'):
        shutil.rmtree('temp')

def setup_temp_directories():
    remove_temp_directories()

    os.mkdir('temp')
    os.mkdir('temp/lhs')
    os.mkdir('temp/rhs')
    os.mkdir('temp/diff_sheets')
    os.mkdir('temp/csv_diff')

def setup_output_directory():
    if not os.path.exists('output'):
        os.mkdir('output')

def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

def process_xlsx(lhs_path, rhs_path):
    setup_temp_directories()
    setup_output_directory()

    left_temp_path = 'temp/lhs'
    right_temp_path = 'temp/rhs'
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
    
    # TODO(sasiala): automate diffs with glob
    temp_lhs_sheets = glob.glob('temp/lhs/sheet_*.csv')
    lhs_filenames = [path_leaf(i) for i in temp_lhs_sheets]
    temp_rhs_sheets = glob.glob('temp/rhs/sheet_*.csv')
    rhs_filenames = [path_leaf(i) for i in temp_rhs_sheets]
    print(lhs_filenames)
    print(rhs_filenames)

    if not csv_diff('temp/lhs/sheet_0.csv', 'temp/rhs/sheet_0.csv', 'temp/csv_diff/sheet_0.diff'):
        print("Csv diff failed")
        return False
    if not diff_to_sheet('temp/csv_diff/sheet_0.diff', 'temp/diff_sheets/sheet_0.csv'):
        print("Diff to sheet failed")
        return False
    if not csv_to_xlsx('temp/diff_sheets/sheet_0.csv', 'output/final_out.xlsx'):
        print("CSV to XLSX failed")
        return False

    # TODO(sasiala): remove temp directories, also remove at other returns
    # remove_temp_directories()
    return True

def main():
    process_xlsx("tests\\test_xlsx_l.xlsx", 'tests/test_xlsx_2.xlsx')

if __name__ == "__main__":
    # execute only if run as a script
    main()