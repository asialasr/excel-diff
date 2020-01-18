__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

import os
import ntpath
import shutil

log_dir_path='log'

# TODO(sasiala): path_leaf defined in multiple places.  Reconsider.
def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

def set_log_dir_path(new_log_dir):
    global log_dir_path
    log_dir_path=new_log_dir

def remove_directory_structure():
    if os.path.exists(log_dir_path):
        shutil.rmtree(log_dir_path)

def initialize_directory_structure():
    remove_directory_structure()
    os.mkdir(log_dir_path)

def log(log_path, message):
    COMBINED_LOG_FILE='combined.log'
    with open(log_dir_path + '\\' + COMBINED_LOG_FILE, 'a+') as combined_log:
        with open(log_dir_path + '\\' + log_path, 'a+') as log_file:
            log_file.write(message + '\n')
            combined_log.write(path_leaf(log_path) + ': ' + message + '\n')
