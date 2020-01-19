__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

# TODO(sasiala): may need to update naming to match general python conventions
import logger
import re

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

# TODO(sasiala): new line is broken
def check_new_line(line, out_file):
    is_new_line = re.match('^\+ .+$', line)

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
    is_deleted_line = re.match('^- .+$', line)

    if is_deleted_line:
        split_lines = line.split('\n')
        for i in split_lines:
            if not re.match('^- $', i):
                out_file.write('Deleted Line,')
                out_file.write(re.sub('^- ', '', i))
                out_file.write('\n')
        return True
    return False

def check_compound(line_in, out_file):
    line = line_in.split('\n')
    left_over = line
    logger.log('diff_to_sheet.log', f'Line: {line_in}', logger.LogLevel.DEBUG)
    # TODO(sasiala): am I sure these can't be in the middle of the string?
    # TODO(sasiala): there is a case for - .*\n+ .*$ that needs to be handled.  
    #                It is similar to check_change_sub and can be seen in sheet 0 of test.
    if (len(line) >= 4 and check_change_add_and_sub('\n'.join(line[0:4]), out_file)):
        left_over = line[4:]
        logger.log('diff_to_sheet.log', 'Compound:Change/Add/Sub', logger.LogLevel.DEBUG)
    elif (len(line) >= 3 and check_change_add('\n'.join(line[0:3]), out_file)):
        left_over = line[3:]
        logger.log('diff_to_sheet.log', 'Compound:Change/Add', logger.LogLevel.DEBUG)
    elif (len(line) >= 3 and check_change_sub('\n'.join(line[0:3]), out_file)):
        left_over = line[3:]
        logger.log('diff_to_sheet.log', 'Compound:Change/Sub', logger.LogLevel.DEBUG)
        
    for i in left_over:
        # TODO(sasiala): should the check for an empty line be here? 
        #                Or should that be fixed when generating the file used as input here?
        if re.match('^- $', i) or re.match('^\+ $', i) or re.match('^  $', i) or re.match('^$', i):
            logger.log('diff_to_sheet.log', 'Compound:Skipped empty +/- line', logger.LogLevel.DEBUG)
        elif (check_new_line(i, out_file)):
            logger.log('diff_to_sheet.log', 'Compound:New Line', logger.LogLevel.DEBUG)
        elif (check_deleted_line(i, out_file)):
            logger.log('diff_to_sheet.log', 'Compound:Deleted Line', logger.LogLevel.DEBUG)
        elif re.match('^  .+$', i):
            logger.log('diff_to_sheet.log', 'Compound:No Change', logger.LogLevel.DEBUG)
            out_file.write('No Change,')
            out_file.write(re.sub('^  ', '', i))
            out_file.write('\n')
        else:
            # unexpected format in diff
            logger.log('diff_to_sheet.log', 'Compound:Curious (unexpected diff format)...', logger.LogLevel.ERROR)
            logger.log('diff_to_sheet.log', i, logger.LogLevel.ERROR)
            logger.log('diff_to_sheet.log', '/Curious', logger.LogLevel.ERROR)
            # TODO(sasiala): return False
    return True

def diff_to_sheet(csv_diff_path, out_path):
    logger.log('diff_to_sheet.log', f'Creating sheet csv for {csv_diff_path}', logger.LogLevel.DEBUG)
    with open(out_path, 'w') as out_file:
        with open(csv_diff_path, 'r') as csv_diff:
            lines = csv_diff.read().split('\n  \n')
            for line in lines:
                if not check_compound(line, out_file):
                    # unexpected format in diff
                    logger.log('diff_to_sheet.log', 'Curious (unexpected diff format)...', logger.LogLevel.ERROR)
                    logger.log('diff_to_sheet.log', line, logger.LogLevel.ERROR)
                    logger.log('diff_to_sheet.log', '/Curious', logger.LogLevel.ERROR)
                    # TODO(sasiala): return False
            logger.log('diff_to_sheet.log', f'Successfully completed sheet csv for {csv_diff_path}', logger.LogLevel.DEBUG)
            return True
        logger.log('diff_to_sheet.log', f'Failed to create sheet csv for {csv_diff_path}.  Error: could not open csv_diff', logger.LogLevel.ERROR)
        return False
    logger.log('diff_to_sheet.log', f'Failed to create sheet csv for {csv_diff_path}.  Error: could not open out_path', logger.LogLevel.ERROR)
    return False
