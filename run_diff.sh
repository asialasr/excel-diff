export PATH="$PATH:C:\Users\Sean\AppData\Local\Programs\Python\Python36"

lhs_path=tests/test_xlsx_l.xlsx
rhs_path=tests/test_xlsx_2.xlsx

python excel_diff.py $lhs_path $rhs_path -st -vvvvvv --gui-out