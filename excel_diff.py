__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

import xlrd
import csv
import difflib
import re

CSV_PATH="temp.csv"
DIFF_OUT_PATH="temp.diff"

def sheet_to_csv(sheet):
    with open(CSV_PATH, 'w') as temp_csv:
        wr = csv.writer(temp_csv, quoting=csv.QUOTE_ALL)

        for row_num in range(sheet.nrows):
            wr.writerow(sheet.row_values(row_num))
        
        temp_csv.close()
        return True
    return False

def csv_diff(csvl, csvr):
    with open(csvl, 'r') as csvl_file:
        with open(csvr, 'r') as csvr_file:
            d = difflib.Differ()
            diff = difflib.ndiff(csvl_file.read().splitlines(1), csvr_file.read().splitlines(1))
    with open(DIFF_OUT_PATH, 'w') as diff_file:
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
        out_file.write('Change/add,')
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

def diff_to_sheet(out_path, csv_diff_path):
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
                    

def process_xlsx(path):
    with xlrd.open_workbook(path) as xlsx_file:
        for sheet_num in range(xlsx_file.nsheets):
            if not sheet_to_csv(xlsx_file.sheet_by_index(sheet_num)):
                print("Sheet to csv failed")
                return False
            if not csv_diff('temp.csv', 'tempr.csv'):
                print("Csv diff failed")
                return False
            if not diff_to_sheet('temp_diff_out.csv', 'temp.diff'):
                print("Diff to sheet failed")
                return False
        return True
    return False

def main():
    process_xlsx("C:\\Users\\Sean\\Documents\\Moving\\Packed list.xlsx")

if __name__ == "__main__":
    # execute only if run as a script
    main()