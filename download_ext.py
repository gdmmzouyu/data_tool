import duckdb
import os
import shutil

# Auto download matching version of excel extension
conn = duckdb.connect()
conn.execute('INSTALL excel')

# Locate extension cache and copy to project extensions directory
ext_cache = os.path.join(os.path.expanduser('~'), '.duckdb', 'extensions')
os.makedirs('extensions', exist_ok=True)

found = False
for root, dirs, files in os.walk(ext_cache):
    if 'excel.duckdb_extension' in files:
        src = os.path.join(root, 'excel.duckdb_extension')
        shutil.copy(src, 'extensions/excel.duckdb_extension')
        print(f'Extension ready: {src}')
        found = True
        break

if not found:
    print('Extension file not found')
    exit(1)