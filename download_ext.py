import duckdb
import os
import shutil

# 自动下载对应版本的 excel 扩展
conn = duckdb.connect()
conn.execute('INSTALL excel')

# 找到扩展缓存并复制到项目 extensions 目录
ext_cache = os.path.join(os.path.expanduser('~'), '.duckdb', 'extensions')
os.makedirs('extensions', exist_ok=True)

found = False
for root, dirs, files in os.walk(ext_cache):
    if 'excel.duckdb_extension' in files:
        src = os.path.join(root, 'excel.duckdb_extension')
        shutil.copy(src, 'extensions/excel.duckdb_extension')
        print(f'扩展已准备好: {src}')
        found = True
        break

if not found:
    print('未找到扩展文件')
    exit(1)