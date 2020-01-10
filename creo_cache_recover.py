import linecache
import re
from pathlib import Path


def unique_name(file_name):
    counter = 0
    file_name_pattern = file_name + '.{:d}'
    while True:
        counter += 1
        if not Path(file_name_pattern.format(counter)).exists():
            return file_name_pattern.format(counter)


entries = Path('.')
for entry in sorted(entries.iterdir(),
                    key=lambda f: f.stat().st_mtime,
                    reverse=False):
    with open(entry, mode='rb') as file:
        content = file.readlines()
        check_string = content[0].decode('ISO-8859-1')
        if not check_string.startswith("#UGC"):
            continue
        name_line = content[9].decode('ISO-8859-1')
        name = name_line[11:-2].rstrip()
        entry.rename(entry.with_name(unique_name(name)))
