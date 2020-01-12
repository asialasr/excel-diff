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

def diff_to_sheet(out_path, csv_diff_path):
    with open(out_path, 'w') as out_file:
        with open(csv_diff_path, 'r') as csv_diff:
            lines = csv_diff.read().split('\n  \n')
            for line in lines:
                change_sub_split = re.split('\n\?.*\n\+ ', line)
                is_change_sub = len(change_sub_split) == 2
                change_add_split = re.split('\n\?.*$', line)
                is_change_add = len(change_add_split) == 2
                remove_split = re.split('- ', line)
                is_remove = len(remove_split) == 4
                add_split = re.split('\+ ', line)
                is_add = len(add_split) == 4

                if is_change_sub:
                    temp_first_line = change_sub_split[0].split('- ')[1]
                    first_line = temp_first_line.split(',')
                    second_line = change_sub_split[1].split(',')
                    iter_size = len(first_line)
                    if (len(first_line) > len(second_line)):
                        iter_size = len(second_line)
                    
                    # TODO(sasiala): deal with lines of diff sizes (skipping rest of output, currently)
                    new_list = []
                    for col in range(iter_size):
                        new_list.append('||'.join([first_line[col], second_line[col]]))
                    out_file.write(','.join(new_list))
                    out_file.write('\n')
                elif is_change_add:
                    temp_lines = change_add_split[0].split('\n')
                    first_line = temp_lines[0].split('- ')[1].split(',')
                    second_line = re.split('\+ ', temp_lines[1])[1].split(',')
                    iter_size = len(first_line)
                    if (len(first_line) > len(second_line)):
                        iter_size = len(second_line)

                    # TODO(sasiala): deal with lines of diff sizes (skipping rest of output, currently)
                    new_list = []
                    for col in range(iter_size):
                        new_list.append('||'.join([first_line[col], second_line[col]]))
                    out_file.write(','.join(new_list))
                    out_file.write('\n')
                elif is_add:
                    out_file.write(add_split[2])
                    out_file.write('\n')
                elif is_remove:
                    out_file.write(remove_split[2])
                    out_file.write('\n')
                elif len(change_sub_split) == 1:
                    if len(line.split('  ')) == 2:
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