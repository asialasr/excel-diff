import xlrd
import csv
import difflib

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

def process_xlsx(path):
    with xlrd.open_workbook(path) as xlsx_file:
        for sheet_num in range(xlsx_file.nsheets):
            if not sheet_to_csv(xlsx_file.sheet_by_index(sheet_num)):
                return False
            
        return True
    return False

def main():
    process_xlsx("C:\\Users\\Sean\\Documents\\Moving\\Packed list.xlsx")
    csv_diff("temp.csv", "tempr.csv")

if __name__ == "__main__":
    # execute only if run as a script
    main()