__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

# TODO(sasiala): again, reconsider naming

import csv
import xlsxwriter
import logger

def csv_to_sheet(workbook, csv_path):
    # TODO(sasiala): convert to only add sheets in this function
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
                    logger.log('csv_to_xlsx.log', 'Change/Sub', logger.LogLevel.DEBUG)
                    line_format = change_sub_format
                elif (temp[0][1] == 'Change/Add'):
                    logger.log('csv_to_xlsx.log', 'Change/Add', logger.LogLevel.DEBUG)
                elif (temp[0][1] == 'Change/Add/Sub'):
                    logger.log('csv_to_xlsx.log', 'Change/Add/Sub', logger.LogLevel.DEBUG)
                    line_format = change_add_sub_format
                elif (temp[0][1] == 'New Line'):
                    logger.log('csv_to_xlsx.log', 'New Line', logger.LogLevel.DEBUG)
                elif (temp[0][1] == 'Deleted Line'):
                    logger.log('csv_to_xlsx.log', 'Deleted Line', logger.LogLevel.DEBUG)
                    line_format = change_sub_format
                elif (temp[0][1] == 'No Change'):
                    logger.log('csv_to_xlsx.log', 'No Change', logger.LogLevel.DEBUG)
                    line_format = no_change_format
                else:
                    logger.log('csv_to_xlsx.log', 'Curious...', logger.LogLevel.ERROR)
            
            for c, col in enumerate(row):
                worksheet.write(r, c, col, line_format)
                
        return True
    return False
