__author__ = "Sean Asiala"
__copyright__ = "Copyright (C) 2020 Sean Asiala"

import os
import ntpath
import shutil
from enum import Enum

# TODO(sasiala): re-implement logging using the logging module
class LogLevel(Enum):
        NONE=0
        FATAL=NONE+1
        ERROR=FATAL+1
        WARN=ERROR+1
        INFO=WARN+1
        DEBUG=INFO+1
        TRACE=DEBUG+1
        ALL=TRACE+1

log_dir_path='log'
log_level_thresh=LogLevel.FATAL

def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

def set_log_dir_path(new_log_dir):
    global log_dir_path
    log_dir_path=new_log_dir

def set_log_level(new_log_level):
    global log_level_thresh
    log_level_thresh=new_log_level

def remove_directory_structure():
    if os.path.exists(log_dir_path):
        shutil.rmtree(log_dir_path)

def initialize_directory_structure():
    remove_directory_structure()
    os.mkdir(log_dir_path)

def log(log_path, message, log_level):
    if log_level.value > log_level_thresh.value:
        return
    
    COMBINED_LOG_FILE='combined.log'
    with open(log_dir_path + '\\' + COMBINED_LOG_FILE, 'a+') as combined_log:
        with open(log_dir_path + '\\' + log_path, 'a+') as log_file:
            log_file.write(message + '\n')
            combined_log.write(path_leaf(log_path) + ': ' + message + '\n')
