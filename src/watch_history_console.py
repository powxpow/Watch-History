'''Command line version of watch-history. Let the user select their Google Takeout 
    file that contains their YouTube Watch History. The file can be either a zip, 
    a JSON file, or a HTML file'''
#core
import argparse
import logging as log
import sys
from pathlib import Path, PurePath
#classes
# pylint: disable=no-name-in-module, import-error
from classes.whrun import WatchHistoryRun
from classes.whdata import WatchHistoryDataHandler as whdh
from classes.whexcel import ExcelBuilder as excel

@staticmethod
def get_parameters():
    '''Attempt to get the Google Takeout fille and the Output directory. 
       If the user has provided command line arguments, use that.
       Otherwise ask the user for them.'''
    args = get_args()
    source_prompt = "Enter Google Takeout file path"
    src_default = "tests/data/good-sample-j.json"
    source = args.source_file or get_from_user(source_prompt, src_default)
    output_dir_prompt = "Enter Output directory"
    out_dir = args.output_dir or get_from_user(output_dir_prompt, "~/Downloads")
    return source, out_dir

@staticmethod
def get_args():
    '''Attempt to get the Google Takeout file and output directory from
       command line arguments'''
    desc = "Process Google Takeout file and output spreadsheet to directory."
    parser = argparse.ArgumentParser(description=desc)
    parser.add_argument("source_file", nargs="?", help="Google Takeout file")
    parser.add_argument("output_dir", nargs="?", help="Output directory")
    args = parser.parse_args()
    return args

@staticmethod
def get_from_user(prompt, default=None):
    '''Generic requesting info from the user.'''
    if default:
        prompt = f"{prompt} (default: {default}): "
    else:
        prompt = f"{prompt}: "
    user_input = input(prompt)
    return user_input or default

def main():
    '''main'''
    log_handler = log.StreamHandler(sys.stdout)
    log_fmt = '%(asctime)s %(levelname)s\t%(message)s'
    log.basicConfig(level=log.INFO, format=log_fmt, handlers=[log_handler])
    source, out_dir = get_parameters()
    watch_history = WatchHistoryRun(None, whdh(), spreadsheet=excel())
    src = watch_history.get_source_path(source)
    if src is not None:
        xlsx_file = src.name.replace(src.suffix, '.xlsx')
        dest_file = PurePath(Path(out_dir), xlsx_file)
        watch_history.run(src, dest_file)
    else:
        log.info("Nothing to do.")

if __name__ == '__main__':
    main()
