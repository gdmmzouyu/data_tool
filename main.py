# 【第一行，绝对不能动！Python强制要求】
from __future__ import annotations

# 【第二部分：系统库导入】
import sys
import os

# 实时模块总开关：True=启用，False=完全移除（界面/代码/依赖全不加载）
ENABLE_REALTIME_MODULE = True

# 【第三部分：资源路径函数（读取master.key必备）】
def get_res_path(rel_path: str):
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        # 修正：用脚本自身所在目录，而非工作目录，调试更稳定
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel_path)

# 密钥路径（固定写法）
KEY_PATH = get_res_path("master.key")
import warnings
# 抑制requests依赖版本不匹配的环境警告
warnings.filterwarnings("ignore", message="urllib3.*doesn't match a supported version")

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import chardet
import duckdb
import uuid
import datetime
import math
import re
from typing import Optional, Callable
import threading
import tempfile
import atexit
import requests
import json
import html
import concurrent.futures
import time
import hashlib
from excel_tools import ExcelFileTools, DuckExcelEngine
from wps_excel_business import WpsExcelBusiness
from multi_sheet_business import MultiSheetBusiness
from realtime_module import create_realtime_tab

from cryptography.fernet import Fernet
import os

# ========== duckdb加载扩展不联网配置 ==========
# 打包专用离线初始化（自动适配开发/EXE环境）
# -------------------
def init_duckdb_offline():
    try:
        conn = duckdb.connect()
        
        # 【锁死】彻底关闭所有联网行为，打包后也绝对不会联网
        conn.execute("SET autoinstall_extensions = false")
        conn.execute("SET autoload_known_extensions = false")
        
        # 自动获取扩展目录路径（开发环境用当前目录，EXE环境用临时解压目录）
        if getattr(sys, 'frozen', False):
            # 打包成EXE后运行，扩展在临时解压目录
            ext_dir = os.path.join(sys._MEIPASS, "extensions")
        else:
            # 开发环境运行，扩展在项目目录
            ext_dir = os.path.join(os.path.dirname(__file__), "extensions")
        
        # 加载所有扩展（按你实际有的来写）
        conn.execute(f"LOAD '{os.path.join(ext_dir, 'excel.duckdb_extension')}'")
        # conn.execute(f"LOAD '{os.path.join(ext_dir, 'parquet.duckdb_extension')}'")
        # conn.execute(f"LOAD '{os.path.join(ext_dir, 'json.duckdb_extension')}'")
        
        print("✅ DuckDB离线初始化成功")
        return conn
    except Exception as e:
        print(f"❌ 初始化失败: {str(e)}")
        input("按回车退出...")
        sys.exit(1)


# ========== 全局临时文件管理器 ==========
_temp_files_lock = threading.Lock()
_temp_files_set = set()

def register_temp_file(file_path: str):
    """注册临时文件，程序退出时自动清理"""
    with _temp_files_lock:
        _temp_files_set.add(os.path.abspath(file_path))

def cleanup_all_temp_files():
    """清理所有注册的临时文件，异常静默跳过"""
    with _temp_files_lock:
        if not _temp_files_set:
            return
        for path in list(_temp_files_set):
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
        _temp_files_set.clear()

# 注册程序退出兜底清理（极端情况如崩溃外的正常退出均会执行）
atexit.register(cleanup_all_temp_files)


class AppConfig:
    APP_TITLE = "茂名移动数据处理工具V3.5 作者：邹宇"
    WINDOW_SIZE = "1100x1140"
    MIN_SIZE = (1100, 1140)
    # 字体字号配置，字体名由 AppTools.get_font 跨平台自动适配
    FONT_SIZE_MAIN = 9
    FONT_SIZE_SMALL = 8
    FONT_SIZE_BOLD = 9
    FONT_SIZE_BASE = 11
    LOG_MATCH_LIMIT = 25
    LOG_MATCH_LIMIT_REAL = 25  # 实时日志单独定义
    LOG_PROCESS_LIMIT = 12
    PREVIEW_MAX_ROWS = 100
    DEFAULT_SEPARATOR = "&&"
    ENCODING_FALLBACK = "utf-8"
    SPECIAL_CHARS = {
        "换行\\n": "\n", "制表\\t": "\t", "回车\\r": "\r",
        "空格": " ", "逗号": ",", "竖线|": "|",
        "单引'": "'", "双引\"": '"', "和号&": "&"
    }

    # ========== 全局标准配色（全模块唯一颜色源） ==========
    COLOR_INFO = "#0872DC"       # 处理中
    COLOR_SUCCESS = "#1B5E20"    # 已完成
    COLOR_WARNING = "#E39117"    # 警告
    COLOR_ERROR = "#F24D4D"      # 错误
    COLOR_PENDING = "#838181"      # 待处理

class FileTools:
    
    
    @staticmethod
    def get_file_size_mb(file_path: str) -> float:
        try:
            return os.path.getsize(file_path) / 1024 / 1024
        except FileNotFoundError:
            return 0
        except PermissionError:
            return 0
        except OSError:
            return 0

    @staticmethod
    def detect_file_encoding(file_path: str) -> str:
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(200000)
            result = chardet.detect(raw_data)
            enc = result.get('encoding') or AppConfig.ENCODING_FALLBACK
            if enc.lower() in ('gb2312', 'gbk', 'gb18030'):
                return 'gb18030'
            if enc.lower() in ('utf-8', 'utf8'):
                return 'utf-8-sig'
            return enc
        except FileNotFoundError:
            raise RuntimeError("文件不存在，无法检测编码")
        except PermissionError:
            raise RuntimeError("无文件读取权限，请检查文件权限")
        except OSError:
            pass

        # 备用编码逐次探测
        for enc in ['utf-8-sig', 'gb18030', 'latin-1']:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    f.read(1024)
                return enc
            except UnicodeDecodeError:
                continue
            except PermissionError:
                raise RuntimeError("无文件读取权限，请检查文件权限")
            except OSError:
                continue
        return AppConfig.ENCODING_FALLBACK

    @staticmethod
    def is_file_available(file_path: str) -> bool:
        return os.path.exists(file_path) and os.path.getsize(file_path) > 0

    @staticmethod
    def read_filter_blank_lines(file_path: str, encoding: str) -> tuple[list[str], bool]:
        valid_lines = []
        try:
            with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                for line in f:
                    if line.strip():
                        valid_lines.append(line)
            return valid_lines, len(valid_lines) > 0
        except FileNotFoundError:
            raise RuntimeError("文件不存在，读取失败")
        except PermissionError:
            raise RuntimeError("无文件读取权限")
        except IsADirectoryError:
            raise RuntimeError("目标路径是文件夹，不是文件")
        except UnicodeDecodeError:
            raise RuntimeError(f"使用 {encoding} 编码解码失败，请更换编码")
        except OSError as e:
            raise RuntimeError(f"文件读取失败：{str(e)}")

    @staticmethod
    def check_file_valid(file_path: str) -> tuple[bool, str]:
        if not os.path.exists(file_path):
            return False, "文件不存在"
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            return False, "文件为空（字节大小为0）"
        try:
            enc = FileTools.detect_file_encoding(file_path)
            _, has_valid = FileTools.read_filter_blank_lines(file_path, enc)
            if not has_valid:
                return False, "文件仅包含换行/空格等空白内容，无有效数据"
            return True, "文件校验通过"
        except RuntimeError as e:
            return False, str(e)

    @staticmethod
    def is_excel_file(file_path: str) -> bool:
        return ExcelFileTools.is_excel_file(file_path)

    @staticmethod
    def check_excel_valid(file_path: str) -> tuple[bool, str]:
        return ExcelFileTools.check_excel_valid(file_path)  

    @staticmethod
    def get_save_file_path(tag: str) -> str:
        time_str = datetime.datetime.now().strftime('%Y%m%d_%H时%M分')
        app_dir = FileTools.get_app_root_dir()  # 改为 FileTools
        return os.path.join(app_dir, f"{tag}_{time_str}.txt")

    @staticmethod
    def get_file_full_path(file_path: str) -> str:
        return os.path.abspath(file_path)

    @staticmethod
    def atomic_rename(tmp_path: str, final_path: str):
        """原子替换文件：同分区下为系统级原子操作，杜绝中间状态数据丢失
        要么替换成功（新文件覆盖旧文件），要么完全失败（旧文件保留）"""
        try:
            if not os.path.exists(tmp_path):
                raise RuntimeError("临时文件不存在，写入失败")
            # os.replace 跨平台原子覆盖：自动替换目标文件，无中间删除步骤
            # 同目录/同分区下为原子操作，崩溃/断电不会出现旧文件已删、新文件未生成的状态
            os.replace(tmp_path, final_path)
        except FileNotFoundError:
            raise RuntimeError("临时文件不存在，写入失败")
        except PermissionError:
            raise RuntimeError("目标文件被占用或无写入权限，请关闭占用该文件的程序")
        except OSError as e:
            raise RuntimeError(f"文件写入失败: {str(e)}")

    @staticmethod
    def safe_remove(file_path: str):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except Exception:
            pass

    @staticmethod
    def count_file_lines_fast(file_path: str) -> int:
        """二进制分块快速统计文件行数，自动修正末尾无换行的情况，结果与实际行数一致"""
        count = 0
        buffer_size = 8 * 1024 * 1024  # 每次读取8MB
        try:
            file_size = os.path.getsize(file_path)
            # 空文件直接返回0
            if file_size == 0:
                return 0

            with open(file_path, 'rb') as f:
                # 分块统计换行符数量
                while True:
                    chunk = f.read(buffer_size)
                    if not chunk:
                        break
                    count += chunk.count(b'\n')

                # 校验文件末尾：最后一个字节不是换行符，则实际行数+1
                f.seek(-1, os.SEEK_END)
                last_byte = f.read(1)
                if last_byte != b'\n':
                    count += 1

            return count
        except Exception:
            return 0

    
    @staticmethod
    def prepare_file_load(file_path: str) -> dict:
        """统一预处理文件加载：校验有效性、检测编码、判断是否大文件
        返回结构化结果，供各业务Tab自行处理UI与数据加载"""
        full_path = FileTools.get_file_full_path(file_path)
        is_valid, check_msg = FileTools.check_file_valid(file_path)
        
        if not is_valid:
            return {
                "valid": False,
                "message": check_msg,
                "full_path": full_path
            }

        enc = FileTools.detect_file_encoding(file_path)
        size_mb = FileTools.get_file_size_mb(file_path)

        return {
            "valid": True,
            "message": "文件校验通过",
            "full_path": full_path,
            "encoding": enc,
            "size_mb": size_mb,
        }
    
    
    @staticmethod
    def get_app_root_dir() -> str:
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    


class DuckRelation:
    """轻量 Relation 包装类，持有所属连接，解决原生对象无法挂属性+递归绑定问题"""
    # 全局查询计数器，保证别名唯一
    _query_id = 0

    def __init__(self, conn: duckdb.DuckDBPyConnection, rel: duckdb.DuckDBPyRelation):
        self.conn = conn
        self._rel = rel

    def query(self, sql: str):
        """执行SQL，__src__ 代表当前关系，自动生成唯一别名避免递归
        SQL 中固定使用 __src__ 指代当前数据集
        """
        DuckRelation._query_id += 1
        alias = f"__src_{DuckRelation._query_id}__"
        safe_sql = sql.replace("__src__", alias)
        return DuckRelation(self.conn, self._rel.query(alias, safe_sql))

    # ========== 代理原生常用方法 ==========
    def select(self, expr: str):
        return DuckRelation(self.conn, self._rel.select(expr))

    def filter(self, expr: str):
        return DuckRelation(self.conn, self._rel.filter(expr))

    def limit(self, n: int, offset: int = 0):
        return DuckRelation(self.conn, self._rel.limit(n, offset))

    def distinct(self):
        return DuckRelation(self.conn, self._rel.distinct())

    def count(self, expr: str = "*"):
        return self._rel.count(expr)

    def fetchone(self):
        return self._rel.fetchone()

    @property
    def description(self):
        return self._rel.description

    @property
    def native_rel(self):
        return self._rel
    
    # ========== Excel 读取封装 ==========
    @staticmethod
    def read_excel_multi_column(file: str, sheet_name: str = None, header: bool = True) -> DuckRelation:
        conn = DuckTools._create_connection()
        native_rel = DuckExcelEngine.read_multi_column(conn, file, sheet_name, header)
        return DuckRelation(conn, native_rel)

    @staticmethod
    def read_excel_single_column(file: str, sheet_name: str = None, header: bool = True, col_index: int = 0) -> DuckRelation:
        conn = DuckTools._create_connection()
        native_rel = DuckExcelEngine.read_single_column(conn, file, sheet_name, header, col_index)
        return DuckRelation(conn, native_rel)

    @staticmethod
    def peek_excel_first_row(file: str, sheet_name: str = None, header: bool = True) -> list:
        return DuckExcelEngine.peek_first_row(file, sheet_name, header)

    @staticmethod
    def excel_to_clean_view(conn, file: str, sheet_name: str, header: bool, view_name: str, dedup: bool = True):
        return DuckExcelEngine.read_clean_view(conn, file, sheet_name, header, view_name, dedup)


class DuckTools:
    """纯 duckdb 实现的全量数据处理工具，替代原 DataTools"""

    @staticmethod  
    def _create_connection() -> duckdb.DuckDBPyConnection:
        temp_dir = tempfile.gettempdir()
        os.makedirs(temp_dir, exist_ok=True)
        
        conn = duckdb.connect(
            database=':memory:',
            config={
                'temp_directory': temp_dir,
                'memory_limit': '4GB',
                # 唯一全版本通用的线程控制参数，强制单线程
                'threads': 1
            }
        )
        return conn

    @staticmethod
    def peek_first_row(file: str, sep: str, encoding: str) -> list:
        """
        轻量快速读取第一行，仅用于列预览/列数统计
        不做清洗、不去重、不扫全表，大文件下毫秒级返回
        """
        conn = DuckTools._create_connection()
        try:
            utf8_file = DuckTools._ensure_utf8_file(file, encoding)
            rel = conn.read_csv(
                utf8_file,
                sep=sep,
                header=False,
                null_padding=True,
                all_varchar=True,
                sample_size=100  # 只采样前100行推断格式，极致提速
            )
            rows = rel.limit(1).fetchall()
            if not rows:
                return []
            return [str(v) if v is not None else '' for v in rows[0]]
        except Exception:
            return []
        finally:
            conn.close()

    # ========== 内部工具：编码转码 ==========
    @staticmethod
    def _ensure_utf8_file(file_path: str, encoding: str) -> str:
        enc_lower = encoding.lower()
        if enc_lower in ("utf-8", "utf8"):
            return file_path
        tmp_path = file_path + ".utf8.tmp"
        with open(file_path, "r", encoding=encoding, errors="replace") as fr, \
             open(tmp_path, "w", encoding="utf-8") as fw:
            while True:
                chunk = fr.read(8 * 1024 * 1024)
                if not chunk:
                    break
                fw.write(chunk)
        # 注册为临时文件，退出时自动清理
        register_temp_file(tmp_path)
        return tmp_path

    # ========== 内部工具：生成匹配键 ==========
    @staticmethod
    def _build_key_sql(cols: list, key_cols: list = None, ignore_cols: list = None) -> str:
        if ignore_cols is not None:
            use_cols = [c for i, c in enumerate(cols) if i not in ignore_cols]
        else:
            use_cols = [cols[i] for i in key_cols]
        if not use_cols:
            return "''"
        parts = [f'TRIM(COALESCE("{c}", \'\'))' for c in use_cols]
        return " || CHR(1) || ".join(parts)

    # ========== 内部工具：单连接内读取清洗去重 ==========
    @staticmethod
    def _read_clean_view(conn, file: str, sep: str, encoding: str, view_name: str, dedup: bool = True):
        utf8_file = DuckTools._ensure_utf8_file(file, encoding)
        src_alias = f"{view_name}_src"
        rel = conn.read_csv(
            utf8_file,
            sep=sep,
            header=False,
            null_padding=True,
            all_varchar=True
        )
        cols = [c[0] for c in rel.description]

        # 基础清洗：原生TRIM去首尾空格，性能远优于正则
        trim_parts = [
            f"COALESCE(TRIM(\"{c}\"), '') AS \"{c}\""
            for c in cols
        ]
        rel = rel.select(", ".join(trim_parts))

        # 过滤全空行
        filter_parts = [f"\"{c}\" != ''" for c in cols]
        rel = rel.filter(" OR ".join(filter_parts))

        col_str = ", ".join([f'"{c}"' for c in cols])
        if dedup:
            # 整行全字段去重（默认开启，用于普通匹配/并集场景）
            sql_distinct = f"""
                SELECT {col_str}
                FROM (
                    SELECT *, ROW_NUMBER() OVER () AS __orig_row__
                    FROM {src_alias}
                )
                QUALIFY ROW_NUMBER() OVER (PARTITION BY {col_str} ORDER BY __orig_row__) = 1
                ORDER BY __orig_row__
            """
            rel_clean = rel.query(src_alias, sql_distinct)
        else:
            # 不做整行去重，仅增加行号列（用于对比场景，后续按key去重）
            sql = f"""
                SELECT {col_str}, ROW_NUMBER() OVER () AS __orig_row__
                FROM {src_alias}
            """
            rel_clean = rel.query(src_alias, sql)
        
        rel_clean.create_view(view_name)
        return cols

    # ========== 单列读取 ==========
    @staticmethod
    def read_single_column(file: str, encoding: str) -> DuckRelation:
        conn = DuckTools._create_connection()
        utf8_file = DuckTools._ensure_utf8_file(file, encoding)
        # 完全不指定列名，让DuckDB自动生成column0
        rel = conn.read_csv(
            utf8_file,
            sep="\0",
            header=False,
            null_padding=True,
            all_varchar=True
        )
        # 用位置索引0获取第一列，然后重命名为col0，绝对不会出错
        rel = rel.select("COALESCE(TRIM(column0), '') AS col0")
        # 过滤空行
        rel = rel.filter("col0 != ''")
        # 去重保序
        sql = """
            SELECT col0
            FROM (
                SELECT *, ROW_NUMBER() OVER () AS __rid__
                FROM __src__
            )
            QUALIFY ROW_NUMBER() OVER (PARTITION BY col0 ORDER BY __rid__) = 1
            ORDER BY __rid__
        """
        return DuckRelation(conn, rel).query(sql)

    # ========== 多列读取 ==========
    @staticmethod
    def read_multi_column(file: str, sep: str, encoding: str) -> DuckRelation:
        conn = DuckTools._create_connection()
        utf8_file = DuckTools._ensure_utf8_file(file, encoding)
        rel = conn.read_csv(
            utf8_file,
            sep=sep,
            header=False,
            null_padding=True,
            all_varchar=True
        )
        cols = [c[0] for c in rel.description]

        trim_parts = [
            f"COALESCE(TRIM(\"{c}\"), '') AS \"{c}\""
            for c in cols
        ]
        rel = rel.select(", ".join(trim_parts))

        filter_parts = [f"\"{c}\" != ''" for c in cols]
        rel = rel.filter(" OR ".join(filter_parts))

        col_str = ", ".join([f'"{c}"' for c in cols])
        sql = f"""
            SELECT {col_str}
            FROM (
                SELECT *, ROW_NUMBER() OVER () AS __rid__
                FROM __src__
            )
            QUALIFY ROW_NUMBER() OVER (PARTITION BY {col_str} ORDER BY __rid__) = 1
            ORDER BY __rid__
        """
        return DuckRelation(conn, rel).query(sql)

    # ========== 全局去重（保序）==========
    @staticmethod
    def distinct(rel: DuckRelation) -> DuckRelation:
        cols = [c[0] for c in rel.description]
        col_str = ", ".join([f'"{c}"' for c in cols])
        sql = f"""
            SELECT {col_str}
            FROM (
                SELECT *, ROW_NUMBER() OVER () AS __rid__
                FROM __src__
            )
            QUALIFY ROW_NUMBER() OVER (PARTITION BY {col_str} ORDER BY __rid__) = 1
            ORDER BY __rid__
        """
        return rel.query(sql)

    # ========== 处理算子 ==========
    @staticmethod
    def op_replace(rel: DuckRelation, col: str, old: str, new: str):
        old_esc = old.replace("'", "''")
        new_esc = new.replace("'", "''")
        sql = f"SELECT REPLACE(\"{col}\", '{old_esc}', '{new_esc}') AS \"{col}\" FROM __src__"
        return rel.query(sql)

    @staticmethod
    def op_insert(rel: DuckRelation, col: str, prefix: str, suffix: str):
        prefix_esc = prefix.replace("'", "''")
        suffix_esc = suffix.replace("'", "''")
        sql = f"SELECT '{prefix_esc}' || \"{col}\" || '{suffix_esc}' AS \"{col}\" FROM __src__"
        return rel.query(sql)

    @staticmethod
    def op_upper(rel: DuckRelation, col: str):
        sql = f'SELECT UPPER("{col}") AS "{col}" FROM __src__'
        return rel.query(sql)

    @staticmethod
    def op_lower(rel: DuckRelation, col: str):
        sql = f'SELECT LOWER("{col}") AS "{col}" FROM __src__'
        return rel.query(sql)

    @staticmethod
    def op_add_serial(rel: DuckRelation, col: str, start: int = 1):
        sql = f"""
            SELECT (ROW_NUMBER() OVER () + {start - 1})::VARCHAR || "{col}" AS "{col}"
            FROM __src__
        """
        return rel.query(sql)

    @staticmethod
    def op_count_char(rel: DuckRelation, col: str, char: str) -> int:
        if not char:
            return 0
        char_esc = char.replace("'", "''")
        char_len = len(char)
        sql = f"""
            SELECT SUM(
                (LENGTH("{col}") - LENGTH(REPLACE("{col}", '{char_esc}', ''))) / {char_len}
            ) AS total_cnt
            FROM __src__
        """
        result = rel.query(sql).fetchone()[0]
        return int(result) if result is not None else 0

    @staticmethod
    def op_filter_contains(rel: DuckRelation, col: str, char: str):
        if not char:
            return rel
        char_esc = char.replace("'", "''")
        sql = f"SELECT * FROM __src__ WHERE CONTAINS(\"{col}\", '{char_esc}')"
        return rel.query(sql)

    # ========== 文件分割 ==========
    @staticmethod
    def split_file_by_lines(rel: DuckRelation, output_prefix: str, lines_per_file: int, sep: str = "\t"):
        total = DuckTools.count_rows(rel)
        if total == 0:
            return 0
        file_count = (total + lines_per_file - 1) // lines_per_file
        for i in range(file_count):
            offset = i * lines_per_file
            chunk_rel = rel.limit(lines_per_file, offset)
            out_path = f"{output_prefix}_第{i+1}份.txt"
            DuckTools.safe_to_csv(chunk_rel, out_path, sep)
        return file_count

    # ========== 匹配功能 ==========
    @staticmethod
    def match_single_mode(
        file_a: str, sep_a: str, enc_a: str, a_key_cols: list,
        file_b: str, sep_b: str, enc_b: str, b_key_cols: list,
        mode: str = "find"
    ) -> DuckRelation:
        conn = DuckTools._create_connection()
        cols_a = DuckTools._read_clean_view(conn, file_a, sep_a, enc_a, "view_a")
        cols_b = DuckTools._read_clean_view(conn, file_b, sep_b, enc_b, "view_b")

        key_a = DuckTools._build_key_sql(cols_a, a_key_cols)
        key_b = DuckTools._build_key_sql(cols_b, b_key_cols)
        join_type = "SEMI" if mode == "find" else "ANTI"
        col_select = ", ".join([f'a."{c}"' for c in cols_a])

        sql = f"""
            WITH a_numbered AS (
                SELECT *, ROW_NUMBER() OVER () AS __rid__
                FROM view_a
            ),
            b_keys AS (
                SELECT DISTINCT ({key_b}) AS __mk__
                FROM view_b
                WHERE TRIM({key_b}) != ''
            )
            SELECT {col_select}
            FROM a_numbered a
            {join_type} JOIN b_keys b
            ON ({key_a}) = b.__mk__
            ORDER BY a.__rid__
        """
        return DuckRelation(conn, conn.sql(sql))

    @staticmethod
    def union_rows(
        file_a: str, sep_a: str, enc_a: str,
        file_b: str, sep_b: str, enc_b: str
    ) -> DuckRelation:
        conn = DuckTools._create_connection()
        cols_a = DuckTools._read_clean_view(conn, file_a, sep_a, enc_a, "view_a")
        cols_b = DuckTools._read_clean_view(conn, file_b, sep_b, enc_b, "view_b")

        if len(cols_a) != len(cols_b):
            raise ValueError("两文件列数不一致，无法执行并集")

        col_str = ", ".join(cols_a)
        sql = f"""
            WITH all_data AS (
                SELECT *, ROW_NUMBER() OVER () AS __ord__, 1 AS __src__
                FROM view_a
                UNION ALL
                SELECT *, ROW_NUMBER() OVER () + (SELECT COUNT(*) FROM view_a) AS __ord__, 2 AS __src__
                FROM view_b
            )
            SELECT {col_str}
            FROM all_data
            QUALIFY ROW_NUMBER() OVER (PARTITION BY {col_str} ORDER BY __ord__) = 1
            ORDER BY __ord__
        """
        return DuckRelation(conn, conn.sql(sql))

    @staticmethod
    def advanced_compare(
        file_a: str, sep_a: str, enc_a: str, a_key_indices: list,
        file_b: str, sep_b: str, enc_b: str, b_key_indices: list,
        ignore_cols: list = None
    ):
        conn = DuckTools._create_connection()

        # 仅基础清洗+行号，跳过整行去重（后续按key去重，避免重复计算）
        cols_a = DuckTools._read_clean_view(conn, file_a, sep_a, enc_a, "view_a", dedup=False)
        cols_b = DuckTools._read_clean_view(conn, file_b, sep_b, enc_b, "view_b", dedup=False)

        # 列索引转实际列名，拼接联合key表达式
        a_key_names = [cols_a[i] for i in a_key_indices]
        b_key_names = [cols_b[i] for i in b_key_indices]
        a_key_expr = " || CHR(1) || ".join([f"TRIM(COALESCE(\"{c}\", ''))" for c in a_key_names])
        b_key_expr = " || CHR(1) || ".join([f"TRIM(COALESCE(\"{c}\", ''))" for c in b_key_names])

        col_a_str = ", ".join([f'"{c}"' for c in cols_a])
        col_b_str = ", ".join([f'"{c}"' for c in cols_b])

        # A表按key去重（保留第一次出现的行，性能优于ROW_NUMBER+QUALIFY）
        sql_a_dedup = f"""
            SELECT DISTINCT ON (__k__) {col_a_str}, __orig_row__, __k__
            FROM (
                SELECT {col_a_str}, __orig_row__, {a_key_expr} AS __k__
                FROM view_a
                WHERE TRIM({a_key_expr}) != ''
            )
            ORDER BY __k__, __orig_row__
        """
        conn.execute(f"CREATE TEMP VIEW a_dedup AS {sql_a_dedup}")

        # B表按key去重
        sql_b_dedup = f"""
            SELECT DISTINCT ON (__k__) {col_b_str}, __k__
            FROM (
                SELECT {col_b_str}, {b_key_expr} AS __k__
                FROM view_b
                WHERE TRIM({b_key_expr}) != ''
            )
            ORDER BY __k__
        """
        conn.execute(f"CREATE TEMP VIEW b_dedup AS {sql_b_dedup}")

        # 1. A有B无，保持原文件行序
        sql_a_only = f"""
            SELECT {col_a_str}
            FROM a_dedup
            WHERE __k__ NOT IN (SELECT __k__ FROM b_dedup)
            ORDER BY __orig_row__
        """
        rel_a_only = DuckRelation(conn, conn.sql(sql_a_only))

        # 2. B有A无
        sql_b_only = f"""
            SELECT {col_b_str}
            FROM b_dedup
            WHERE __k__ NOT IN (SELECT __k__ FROM a_dedup)
        """
        rel_b_only = DuckRelation(conn, conn.sql(sql_b_only))

        # 3. 双方共有，保持A表行序
        sql_both = f"""
            SELECT {col_a_str}
            FROM a_dedup
            WHERE __k__ IN (SELECT __k__ FROM b_dedup)
            ORDER BY __orig_row__
        """
        rel_both = DuckRelation(conn, conn.sql(sql_both))

        return rel_a_only, rel_b_only, rel_both

    # ========== 安全写入 ==========
    @staticmethod
    def safe_to_csv(
        rel: DuckRelation, file_path: str, sep: str,
        mode: str = "w", encoding: str = "utf-8-sig",
        header: bool = False
    ):
        conn = rel.conn
        native_rel = rel.native_rel

        if len(sep) != 1:
            cols = [c[0] for c in rel.description]
            concat_expr = f" || '{sep}' || ".join([f'\"{c}\"' for c in cols])
            native_rel = native_rel.select(f"{concat_expr} AS line")
            sep_out = "\n"
        else:
            sep_out = sep

        tmp_path = file_path + ".tmp"
        # 注册导出临时文件
        register_temp_file(tmp_path)
        view_name = f"__export_{DuckRelation._query_id}__"
        DuckRelation._query_id += 1
        conn.register(view_name, native_rel)

        # ✅ 关键修复：移除COPY命令中的ENCODING参数（DuckDB 1.0+ 仅支持读取时使用）
        # 统一先写入UTF-8临时文件，后续再转码到目标编码
        conn.execute(f"""
            COPY (SELECT * FROM {view_name})
            TO '{tmp_path.replace("'", "''")}'
            (FORMAT CSV, DELIMITER '{sep_out}', HEADER {str(header).lower()},
            QUOTE '', ESCAPE '')
        """)

        # ✅ 统一转码逻辑：不管目标编码是什么，都从UTF-8临时文件转换
        final_tmp = tmp_path + ".enc"
        register_temp_file(final_tmp)
        
        with open(tmp_path, "r", encoding="utf-8") as fr, \
            open(final_tmp, "w", encoding=encoding) as fw:
            while True:
                chunk = fr.read(8 * 1024 * 1024)
                if not chunk:
                    break
                fw.write(chunk)

        # 原子替换最终文件
        FileTools.atomic_rename(final_tmp, file_path)
        
        # ✅ 新增：立即删除所有临时文件，不用等到程序退出
        FileTools.safe_remove(tmp_path)
        FileTools.safe_remove(final_tmp)
        
        # 从全局临时集合中移除，避免退出时重复清理
        with _temp_files_lock:
            _temp_files_set.discard(os.path.abspath(tmp_path))
            _temp_files_set.discard(os.path.abspath(final_tmp))

    # ========== 辅助工具 ==========
    @staticmethod
    def count_rows(rel: DuckRelation) -> int:
        return rel.count("*").fetchone()[0]


class AppTools:
    @staticmethod
    def setup_windows_high_dpi():
        """Windows 高DPI屏幕适配，必须在创建Tk根窗口前调用，Mac/Linux自动跳过"""
        if not sys.platform.startswith("win"):
            return
        try:
            import ctypes
            # 优先：Per-Monitor V2 感知（Win10 1703+，多显示器不同缩放适配最佳）
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                # 降级：系统级DPI感知（兼容Win7/Win8）
                import ctypes
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                # 全部失败静默跳过，不影响主程序运行
                pass

    @staticmethod
    def get_font(size: int = 10, bold: bool = False) -> tuple:
        """跨平台自动适配中文字体，返回 (字体名, 字号, 字重) 元组"""
        if sys.platform == "darwin":
            font_name = "PingFang SC"
        elif sys.platform.startswith("linux"):
            font_name = "Noto Sans CJK SC"
        else:
            font_name = "Microsoft YaHei"
        
        weight = "bold" if bold else "normal"
        return (font_name, size, weight)
    

    @staticmethod
    def batch_update_widget_state(widgets: list, state: str):
        for widget in widgets:
            try:
                widget.config(state=state)
            except Exception:
                pass

    @staticmethod
    def parse_escape_string(text: str) -> str:
        """安全解析转义字符串，仅处理标准转义序列，中文/普通字符原样保留，无乱码风险"""
        if not text:
            return ""

        # 支持的转义字符映射表，可按需扩展
        escape_map = {
            'n': '\n',    # 换行
            't': '\t',    # 制表符
            'r': '\r',    # 回车
            '\\': '\\',   # 反斜杠本身
            '"': '"',     # 双引号
            "'": "'",     # 单引号
            '0': '\0',    # 空字节
        }

        def _replace_match(match: re.Match) -> str:
            char_after_backslash = match.group(1)
            # 映射表中有对应转义则替换，没有则保留原反斜杠+字符
            return escape_map.get(char_after_backslash, match.group(0))

        # 只匹配「反斜杠 + 单个字符」的组合，其余内容完全不动
        return re.sub(r'\\(.)', _replace_match, text)
    
    # ========== 新增：正整数校验 ==========
    @staticmethod
    def validate_positive_int(text: str) -> tuple[bool, int]:
        """校验输入是否为正整数，返回 (是否合法, 转换后数值)"""
        try:
            val = int(text)
            if val <= 0:
                return False, 0
            return True, val
        except (ValueError, TypeError):
            return False, 0

    # ========== 新增：输入框数字限制 ==========
    @staticmethod
    def restrict_numeric_input(entry_widget: tk.Entry):
        """给输入框绑定事件，仅允许输入数字字符"""
        def _validate(char):
            return char.isdigit() or char == ""
        vcmd = (entry_widget.register(_validate), '%S')
        entry_widget.config(validate="key", validatecommand=vcmd)


class LogComponent:
    def __init__(self, parent_frame: tk.Frame, max_count: int, root_widget):
        self.frame = parent_frame
        self.max_count = max_count
        self.log_labels = []
        self.root = root_widget
        self.top_widget = None  # 固定记录当前最顶部日志行容器

        # 全局统一配色，无硬编码
        self.color_info = AppConfig.COLOR_INFO
        self.color_success = AppConfig.COLOR_SUCCESS
        self.color_warning = AppConfig.COLOR_WARNING
        self.color_error = AppConfig.COLOR_ERROR

    def append(self, message: str, color: str = None, is_final: bool = False):
        # 颜色优先级：手动指定 > is_final成功色 > 默认信息蓝
        if color is None:
            color = self.color_success if is_final else self.color_info
        self.root.after(0, self._append_ui, message, color, is_final)

    def _append_ui(self, message: str, color: str, is_final: bool):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}"

        row_frame = tk.Frame(self.frame, bg="#F0F0F0")

        # 新日志插到最顶端
        if self.top_widget is not None:
            row_frame.pack(fill=tk.X, pady=1, before=self.top_widget)
        else:
            row_frame.pack(fill=tk.X, pady=1)
        self.top_widget = row_frame

        label = tk.Label(
            row_frame, text=full_msg, font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL),
            fg=color, bg="#F0F0F0", 
            anchor="w",
            justify=tk.LEFT,
            wraplength=610, 
            padx=0
        )
        label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.log_labels.insert(0, label)

        # 超出最大条数，删除最旧的一行
        if len(self.log_labels) > self.max_count:
            oldest_label = self.log_labels.pop()
            oldest_label.master.destroy()

        # 上一条非错误日志自动变为成功色（标识该步骤已完成）
        if len(self.log_labels) >= 2 and self.log_labels[1].cget("fg") != self.color_error:
            self.log_labels[1].config(fg=self.color_success)

        if is_final:
            label.config(fg=self.color_success, font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True))


class BaseTab(ttk.Frame):
    """所有标签页基类：统一后台任务调度、状态管理"""
    def __init__(self, parent):
        super().__init__(parent)
        self.is_loading = False
        self.disabled_widgets = []
        self.logger: Optional[LogComponent] = None
        # ✅ 新增：所有子类统一继承 root 属性，指向程序根窗口
        self.root = self.winfo_toplevel()

    def safe_ui_update(self, func: Callable):
        """线程安全的UI更新：窗口已销毁时跳过，避免退出时报错"""
        if not self.winfo_exists():
            return
        self.after(0, func)

    def run_background_task(self, task_func: Callable, start_msg: str):
        """统一后台任务入口：自动处理状态、线程、异常、收尾"""
        if self.is_loading:
            self.logger.append("正在处理中，请稍候...")
            return

        self.is_loading = True
        AppTools.batch_update_widget_state(self.disabled_widgets, tk.DISABLED)
        self.logger.append(start_msg)

        def _task():
            try:
                finish_msg = task_func()
                if finish_msg:
                    self.logger.append(finish_msg, is_final=True)
            except Exception as e:
                self.logger.append(f"处理失败：{str(e)}", AppConfig.COLOR_ERROR)
            finally:
                self.safe_ui_update(self._task_cleanup)

        threading.Thread(target=_task, daemon=True).start()

    def _task_cleanup(self):
        """任务收尾：主线程中恢复控件状态"""
        self.is_loading = False
        AppTools.batch_update_widget_state(self.disabled_widgets, tk.NORMAL)
        
class MatchBusiness:
    def __init__(self, log_callback):
        self.log = log_callback
        # 数据A（主数据）
        self.source_path: Optional[str] = None
        self.source_encoding: str = "utf-8-sig"
        self.source_sep: str = AppConfig.DEFAULT_SEPARATOR
        self.source_col_count: int = 0
        self.source_col_idx: int = 0  # A匹配列索引，默认第一列

        # 数据B（对比数据）
        self.match_path: Optional[str] = None
        self.match_encoding: str = "utf-8-sig"
        self.match_col_count: int = 0
        self.match_col_idx: int = 0  # B匹配列索引，默认第一列

        self.fuzzy_match = False
        # 数据A（主数据）
        self.source_first_row: list = []  # 第一行预览缓存
        # 数据B（对比数据）
        self.match_first_row: list = []   # 第一行预览缓存


    def load_source(self, file_path: str, sep: str):
        """加载数据A（主数据）"""
        load_info = FileTools.prepare_file_load(file_path)
        if not load_info["valid"]:
            raise RuntimeError(load_info["message"])

        self.source_path = file_path
        self.source_encoding = load_info["encoding"]
        self.source_sep = sep
        
        # 用轻量方法拿列数+第一行，大文件加载速度也会大幅提升
        first_row = DuckTools.peek_first_row(file_path, sep, load_info["encoding"])
        self.source_col_count = len(first_row)
        self.source_first_row = first_row
        # 兜底：重置匹配列为第一列，避免索引越界
        self.source_col_idx = 0
        
        self.log(f"数据A加载完成：{self.source_col_count} 列，DuckDB流式处理")

    def load_match(self, file_path: str, sep: str):
        """加载数据B（对比数据）"""
        load_info = FileTools.prepare_file_load(file_path)
        if not load_info["valid"]:
            raise RuntimeError(load_info["message"])

        self.match_path = file_path
        self.match_encoding = load_info["encoding"]
        
        # 同理，轻量读取
        first_row = DuckTools.peek_first_row(file_path, self.source_sep, load_info["encoding"])
        self.match_col_count = len(first_row)
        self.match_first_row = first_row
        # 兜底：重置匹配列为第一列，避免索引越界
        self.match_col_idx = 0
        
        self.log(f"数据B加载完成：{self.match_col_count} 列，DuckDB流式处理")

    # ========== 对外功能入口 ==========
    def run_find(self, output_path: str) -> int:
        """功能1：A中找B返回新A"""
        self.log("开始执行：A中找B")
        a_col = self.source_col_idx
        b_col = self.match_col_idx

        # 索引越界兜底
        if a_col >= self.source_col_count or b_col >= self.match_col_count:
            raise ValueError("匹配列索引超出文件实际列数，请重新选择匹配列")

        if self.fuzzy_match:
            # 模糊匹配：同一连接内加载两张表，避免跨连接报错
            conn = DuckTools._create_connection()
            cols_a = DuckTools._read_clean_view(conn, self.source_path, self.source_sep, self.source_encoding, "view_a")
            cols_b = DuckTools._read_clean_view(conn, self.match_path, self.source_sep, self.match_encoding, "view_b")
            
            a_col_name = cols_a[a_col]
            b_col_name = cols_b[b_col]
            col_select = ", ".join([f'"{c}"' for c in cols_a])
            
            sql = f"""
                WITH b_keys AS (
                    SELECT DISTINCT "{b_col_name}" AS kw
                    FROM view_b
                    WHERE "{b_col_name}" != ''
                ),
                a_numbered AS (
                    SELECT *, ROW_NUMBER() OVER () AS __rid__
                    FROM view_a
                )
                SELECT {col_select}
                FROM a_numbered a
                WHERE EXISTS (
                    SELECT 1 FROM b_keys b
                    WHERE CONTAINS(a."{a_col_name}", b.kw)
                )
                ORDER BY __rid__
            """
            rel_result = DuckRelation(conn, conn.sql(sql))
            DuckTools.safe_to_csv(rel_result, output_path, self.source_sep, encoding="utf-8-sig")
            return DuckTools.count_rows(rel_result)

        # 精确匹配：全量走DuckDB
        rel_result = DuckTools.match_single_mode(
            self.source_path, self.source_sep, self.source_encoding, [a_col],
            self.match_path, self.source_sep, self.match_encoding, [b_col],
            "find"
        )
        DuckTools.safe_to_csv(rel_result, output_path, self.source_sep, encoding="utf-8-sig")
        return DuckTools.count_rows(rel_result)

    def run_remove(self, output_path: str) -> int:
        """功能2：A中剔除B返回新A"""
        self.log("开始执行：A中剔除B")
        a_col = self.source_col_idx
        b_col = self.match_col_idx

        # 索引越界兜底
        if a_col >= self.source_col_count or b_col >= self.match_col_count:
            raise ValueError("匹配列索引超出文件实际列数，请重新选择匹配列")

        if self.fuzzy_match:
            # 模糊剔除：同一连接内加载两张表，避免跨连接报错
            conn = DuckTools._create_connection()
            cols_a = DuckTools._read_clean_view(conn, self.source_path, self.source_sep, self.source_encoding, "view_a")
            cols_b = DuckTools._read_clean_view(conn, self.match_path, self.source_sep, self.match_encoding, "view_b")
            
            a_col_name = cols_a[a_col]
            b_col_name = cols_b[b_col]
            col_select = ", ".join([f'"{c}"' for c in cols_a])
            
            sql = f"""
                WITH b_keys AS (
                    SELECT DISTINCT "{b_col_name}" AS kw
                    FROM view_b
                    WHERE "{b_col_name}" != ''
                ),
                a_numbered AS (
                    SELECT *, ROW_NUMBER() OVER () AS __rid__
                    FROM view_a
                )
                SELECT {col_select}
                FROM a_numbered a
                WHERE NOT EXISTS (
                    SELECT 1 FROM b_keys b
                    WHERE CONTAINS(a."{a_col_name}", b.kw)
                )
                ORDER BY __rid__
            """
            rel_result = DuckRelation(conn, conn.sql(sql))
            DuckTools.safe_to_csv(rel_result, output_path, self.source_sep, encoding="utf-8-sig")
            return DuckTools.count_rows(rel_result)

        # 精确匹配：全量走DuckDB
        rel_result = DuckTools.match_single_mode(
            self.source_path, self.source_sep, self.source_encoding, [a_col],
            self.match_path, self.source_sep, self.match_encoding, [b_col],
            "remove"
        )
        DuckTools.safe_to_csv(rel_result, output_path, self.source_sep, encoding="utf-8-sig")
        return DuckTools.count_rows(rel_result)

    def run_union(self, output_path: str) -> int:
        """功能5：A和B的并集（需列数相同）"""
        self.log("开始执行：A和B的并集")
        if self.source_col_count != self.match_col_count:
            raise ValueError(f"列数不一致！数据A {self.source_col_count} 列，数据B {self.match_col_count} 列")

        rel_result = DuckTools.union_rows(
            self.source_path, self.source_sep, self.source_encoding,
            self.match_path, self.source_sep, self.match_encoding
        )
        DuckTools.safe_to_csv(rel_result, output_path, self.source_sep, encoding="utf-8-sig")
        return DuckTools.count_rows(rel_result)

    def run_left_join(self, output_path: str) -> int:
        """功能3：A左连接B 返回A全量+补B字段"""
        self.log("开始执行：A左连接B 返回A全量+补B字段")
        a_col = self.source_col_idx
        b_col = self.match_col_idx

        # 索引越界兜底
        if a_col >= self.source_col_count or b_col >= self.match_col_count:
            raise ValueError("匹配列索引超出文件实际列数，请重新选择匹配列")

        conn = DuckTools._create_connection()
        cols_a = DuckTools._read_clean_view(conn, self.source_path, self.source_sep, self.source_encoding, "view_a")
        cols_b = DuckTools._read_clean_view(conn, self.match_path, self.source_sep, self.match_encoding, "view_b")

        a_col_name = cols_a[a_col]
        b_col_name = cols_b[b_col]

        # 列名重名兜底：B表所有列加 b_ 前缀，避免列名冲突
        a_select = ", ".join([f'a."{c}"' for c in cols_a])
        b_select = ", ".join([f'COALESCE(b."{c}", \'\') AS b_{c}' for c in cols_b])
        col_select = f"{a_select}, {b_select}"

        if self.fuzzy_match:
            # 模糊左连接：A匹配列包含B匹配列内容即命中
            sql = f"""
                WITH a_numbered AS (
                    SELECT *, ROW_NUMBER() OVER () AS __rid__
                    FROM view_a
                ),
                b_distinct AS (
                    SELECT DISTINCT *
                    FROM view_b
                    WHERE "{b_col_name}" != ''
                )
                SELECT {col_select}
                FROM a_numbered a
                LEFT JOIN b_distinct b
                ON CONTAINS(a."{a_col_name}", b."{b_col_name}")
                ORDER BY a.__rid__
            """
        else:
            # 精确左连接：等值匹配，B表按匹配列去重，避免一对多导致A行重复
            sql = f"""
                WITH a_numbered AS (
                    SELECT *, ROW_NUMBER() OVER () AS __rid__
                    FROM view_a
                ),
                b_dedup AS (
                    SELECT DISTINCT ON ("{b_col_name}") *
                    FROM view_b
                    WHERE "{b_col_name}" != ''
                )
                SELECT {col_select}
                FROM a_numbered a
                LEFT JOIN b_dedup b
                ON TRIM(a."{a_col_name}") = TRIM(b."{b_col_name}")
                ORDER BY a.__rid__
            """

        rel_result = DuckRelation(conn, conn.sql(sql))
        DuckTools.safe_to_csv(rel_result, output_path, self.source_sep, encoding="utf-8-sig")
        return DuckTools.count_rows(rel_result)

    def run_intersection(self, output_path: str) -> int:
        """功能2：A中找B 返回新A+B全部"""
        self.log("开始执行：A中找B 返回新A+B全部")
        a_col = self.source_col_idx
        b_col = self.match_col_idx

        # 索引越界兜底
        if a_col >= self.source_col_count or b_col >= self.match_col_count:
            raise ValueError("匹配列索引超出文件实际列数，请重新选择匹配列")

        conn = DuckTools._create_connection()
        cols_a = DuckTools._read_clean_view(conn, self.source_path, self.source_sep, self.source_encoding, "view_a")
        cols_b = DuckTools._read_clean_view(conn, self.match_path, self.source_sep, self.match_encoding, "view_b")

        a_col_name = cols_a[a_col]
        b_col_name = cols_b[b_col]

        # 列名重名兜底：B表所有列加 b_ 前缀，避免列名冲突
        a_select = ", ".join([f'a."{c}"' for c in cols_a])
        b_select = ", ".join([f'b."{c}" AS b_{c}' for c in cols_b])
        col_select = f"{a_select}, {b_select}"

        if self.fuzzy_match:
            # 模糊交集：A匹配列包含B匹配列内容即命中
            sql = f"""
                WITH a_numbered AS (
                    SELECT *, ROW_NUMBER() OVER () AS __rid__
                    FROM view_a
                ),
                b_distinct AS (
                    SELECT DISTINCT *
                    FROM view_b
                    WHERE "{b_col_name}" != ''
                )
                SELECT {col_select}
                FROM a_numbered a
                JOIN b_distinct b
                ON CONTAINS(a."{a_col_name}", b."{b_col_name}")
                ORDER BY a.__rid__
            """
        else:
            # 精确交集：等值匹配
            sql = f"""
                WITH a_numbered AS (
                    SELECT *, ROW_NUMBER() OVER () AS __rid__
                    FROM view_a
                ),
                b_distinct AS (
                    SELECT DISTINCT *
                    FROM view_b
                    WHERE "{b_col_name}" != ''
                )
                SELECT {col_select}
                FROM a_numbered a
                INNER JOIN b_distinct b
                ON TRIM(a."{a_col_name}") = TRIM(b."{b_col_name}")
                ORDER BY a.__rid__
            """

        rel_result = DuckRelation(conn, conn.sql(sql))
        DuckTools.safe_to_csv(rel_result, output_path, self.source_sep, encoding="utf-8-sig")
        return DuckTools.count_rows(rel_result)

    def run_compare(self, compare_mode: str, a_key_cols: list, b_key_cols: list,
                    ignore_cols: list = None) -> dict:
        """功能6：AB文件高级比较"""
        self.log("开始执行：AB文件高级比较")
        base_a = os.path.splitext(os.path.basename(self.source_path))[0]
        base_b = os.path.splitext(os.path.basename(self.match_path))[0]
        
        # 统一输出到程序根目录，命名规则和其他功能保持一致
        out_dir = FileTools.get_app_root_dir()
        time_str = datetime.datetime.now().strftime('%Y%m%d_%H时%M分')
        file_prefix = f"匹配_高级对比_{time_str}"

        outputs = {
            "a_only": os.path.join(out_dir, f"{file_prefix}_{base_a}_有_{base_b}_无.txt"),
            "b_only": os.path.join(out_dir, f"{file_prefix}_{base_b}_有_{base_a}_无.txt"),
            "both": os.path.join(out_dir, f"{file_prefix}_{base_a}_{base_b}_共同数据.txt")
        }

        # 列数校验
        if compare_mode in ["whole_row", "ignore_cols"]:
            if self.source_col_count != self.match_col_count:
                raise ValueError(f"列数不一致！数据A {self.source_col_count} 列，数据B {self.match_col_count} 列")

        if compare_mode == "whole_row":
            all_cols = list(range(self.source_col_count))
            a_key_cols = all_cols
            b_key_cols = all_cols
        elif compare_mode == "ignore_cols":
            a_key_cols = [i for i in range(self.source_col_count) if i not in ignore_cols]
            b_key_cols = [i for i in range(self.match_col_count) if i not in ignore_cols]

        self.log("正在读取数据并构建索引...")
        rel_a_only, rel_b_only, rel_both = DuckTools.advanced_compare(
            self.source_path, self.source_sep, self.source_encoding, a_key_cols,
            self.match_path, self.source_sep, self.match_encoding, b_key_cols,
            ignore_cols
        )

        self.log("正在写入「A有B无」结果...")
        DuckTools.safe_to_csv(rel_a_only, outputs["a_only"], self.source_sep)
        count_a = DuckTools.count_rows(rel_a_only)
        self.log(f"「A有B无」完成：共 {count_a:,} 条")

        self.log("正在写入「B有A无」结果...")
        DuckTools.safe_to_csv(rel_b_only, outputs["b_only"], self.source_sep)
        count_b = DuckTools.count_rows(rel_b_only)
        self.log(f"「B有A无」完成：共 {count_b:,} 条")

        self.log("正在写入「双方共有」结果...")
        DuckTools.safe_to_csv(rel_both, outputs["both"], self.source_sep)
        count_both = DuckTools.count_rows(rel_both)
        self.log(f"「双方共有」完成：共 {count_both:,} 条")

        self.log("AB文件高级比较全部完成")
        return outputs
    
class MatchTab(BaseTab):
    def __init__(self, parent):
        super().__init__(parent)
        self.selected_func = tk.StringVar(value="find")
        self.compare_mode = tk.StringVar(value="single_col")
        self._init_ui()
        self.business = MatchBusiness(self.logger.append)
        self._update_func_desc()

    def _init_ui(self):
        top_container = ttk.Frame(self)
        top_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=20)

        # ========== 左侧操作区 ==========
        left_area = ttk.Frame(top_container, width=400)
        left_area.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_area.pack_propagate(False)

        # 1. 导入数据A
        ttk.Label(left_area, text="1. 导入数据A（主数据）", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(6, 2), fill=tk.X, padx=5)
        sep_frame = ttk.Frame(left_area)
        sep_frame.pack(pady=2, fill=tk.X, padx=5)
        self.sep_entry = ttk.Entry(sep_frame, width=8)
        self.sep_entry.insert(0, AppConfig.DEFAULT_SEPARATOR)
        self.sep_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(sep_frame, text="分隔符（默认&&）").pack(side=tk.LEFT)

        # 快捷分隔符按钮
        btn_frame = ttk.Frame(left_area)
        btn_frame.pack(pady=2, fill=tk.X, padx=5)
        ttk.Button(btn_frame, text="竖线|", width=6, command=lambda: self.set_separator("|")).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text='双引"', width=6, command=lambda: self.set_separator('"')).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="逗号,", width=6, command=lambda: self.set_separator(",")).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="制表符", width=6, command=lambda: self.set_separator("\t")).pack(side=tk.LEFT, padx=2)

        # 数据A：选择文件 + 匹配列 同一行
        source_row = ttk.Frame(left_area)
        source_row.pack(pady=4, fill=tk.X, padx=5)
        self.btn_source = ttk.Button(source_row, text="选择数据A文件", command=self.load_source_file, width=16)
        self.btn_source.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(source_row, text="匹配列：").pack(side=tk.LEFT)
        self.col_combo_a = ttk.Combobox(source_row, state="disabled", width=18)
        self.col_combo_a.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.label_source = ttk.Label(left_area, text="未导入", foreground="#666666", anchor="w")
        self.label_source.pack(pady=2, anchor=tk.W, padx=5)

        # 2. 导入数据B
        ttk.Label(left_area, text="2. 导入数据B（对比数据）", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(10, 2), fill=tk.X, padx=5)
        
        # 数据B：选择文件 + 匹配列 同一行
        match_row = ttk.Frame(left_area)
        match_row.pack(pady=4, fill=tk.X, padx=5)
        self.btn_match = ttk.Button(match_row, text="选择数据B文件", command=self.load_match_file, width=16)
        self.btn_match.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(match_row, text="匹配列：").pack(side=tk.LEFT)
        self.col_combo_b = ttk.Combobox(match_row, state="disabled", width=18)
        self.col_combo_b.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.label_match = ttk.Label(left_area, text="未导入", foreground="#666666", anchor="w")
        self.label_match.pack(pady=2, anchor=tk.W, padx=5)

        # 3. 匹配模式
        ttk.Label(left_area, text="3. 匹配模式", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(10, 2), fill=tk.X, padx=5)
        self.fuzzy_var = tk.BooleanVar()
        self.fuzzy_check = ttk.Checkbutton(
            left_area, text="模糊匹配（子串包含）",
            variable=self.fuzzy_var, command=self.on_fuzzy_change
        )
        self.fuzzy_check.pack(pady=2, anchor=tk.W, padx=8)

        # 4. 选择功能
        ttk.Label(left_area, text="4. 选择功能", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(10, 2), fill=tk.X, padx=5)

        ttk.Radiobutton(left_area, text="功能1：A中找B 返回新A", variable=self.selected_func, value="find", command=self._on_func_change).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(left_area, text="功能2：A中找B 返回新A+B全部", variable=self.selected_func, value="intersection", command=self._on_func_change).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(left_area, text="功能3：A左连接B 返回A全量+补B字段", variable=self.selected_func, value="left_join", command=self._on_func_change).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(left_area, text="功能4：A中剔除B 返回新A", variable=self.selected_func, value="remove", command=self._on_func_change).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(left_area, text="功能5：A和B的并集（需列数相同）", variable=self.selected_func, value="union", command=self._on_func_change).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(left_area, text="功能6：AB文件高级比较", variable=self.selected_func, value="compare", command=self._on_func_change).pack(anchor="w", padx=8, pady=2)

        # 高级配置框
        self.compare_config_frame = ttk.LabelFrame(left_area, text="高级比较配置")
        self.compare_config_frame.pack(fill="x", padx=5, pady=5)
        self.compare_config_frame.pack_forget()

        # 5个模式单选按钮
        ttk.Radiobutton(self.compare_config_frame, text="单列主键对比", variable=self.compare_mode, value="single_col", command=self._switch_compare_mode).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(self.compare_config_frame, text="多列联合主键对比", variable=self.compare_mode, value="multi_col", command=self._switch_compare_mode).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(self.compare_config_frame, text="整行对比（需列数相同）", variable=self.compare_mode, value="whole_row", command=self._switch_compare_mode).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(self.compare_config_frame, text="忽略指定列对比", variable=self.compare_mode, value="ignore_cols", command=self._switch_compare_mode).pack(anchor="w", padx=8, pady=2)

        # 模式1：单列主键
        self.frame_single_col = ttk.Frame(self.compare_config_frame)
        row_a = ttk.Frame(self.frame_single_col)
        row_a.pack(fill="x", padx=8, pady=3)
        ttk.Label(row_a, text="A主键列：", width=10).pack(side="left")
        self.comp_a_col = ttk.Combobox(row_a, state="readonly")
        self.comp_a_col.pack(side="left", fill="x", expand=True)

        row_b = ttk.Frame(self.frame_single_col)
        row_b.pack(fill="x", padx=8, pady=3)
        ttk.Label(row_b, text="B主键列：", width=10).pack(side="left")
        self.comp_b_col = ttk.Combobox(row_b, state="readonly")
        self.comp_b_col.pack(side="left", fill="x", expand=True)

        # 模式2：多列联合主键
        self.frame_multi_col = ttk.Frame(self.compare_config_frame)
        list_container = ttk.Frame(self.frame_multi_col)
        list_container.pack(fill="x", padx=8, pady=2)

        box_a = ttk.Frame(list_container)
        box_a.pack(side="left", fill="both", expand=True, padx=2)
        ttk.Label(box_a, text="A主键列（多选）").pack(anchor="w")
        self.comp_a_list = tk.Listbox(box_a, selectmode="extended", height=4, exportselection=False, width=17)
        self.comp_a_list.pack(fill="both", expand=True, pady=2)

        box_b = ttk.Frame(list_container)
        box_b.pack(side="left", fill="both", expand=True, padx=2)
        ttk.Label(box_b, text="B主键列（多选）").pack(anchor="w")
        self.comp_b_list = tk.Listbox(box_b, selectmode="extended", height=4, exportselection=False, width=17)
        self.comp_b_list.pack(fill="both", expand=True, pady=2)

        ttk.Label(self.frame_multi_col, text="提示：Command多选，Shift连选",
                  foreground="#888888", font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL)).pack(pady=2, padx=8, anchor="w")

        # 模式3：整行对比
        self.frame_whole_row = ttk.Frame(self.compare_config_frame)
        tip_whole = "• 两文件列数、列顺序必须完全一致\n• 整行内容完全相同才判定为同一行"
        ttk.Label(self.frame_whole_row, text=tip_whole, foreground="#555555", justify="left").pack(pady=8, padx=10, anchor="w")

        # 模式4：忽略指定列
        self.frame_ignore_cols = ttk.Frame(self.compare_config_frame)
        ttk.Label(self.frame_ignore_cols, text="选择要忽略的列（可多选）：").pack(anchor="w", padx=8)
        self.comp_ignore_list = tk.Listbox(self.frame_ignore_cols, selectmode="extended", height=4, exportselection=False, width=45)
        self.comp_ignore_list.pack(fill="x", padx=8, pady=2)
        ttk.Label(self.frame_ignore_cols, text="两文件列数需完全一致，忽略列号一一对应",
                  foreground="#888888", font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL)).pack(anchor="w", padx=8, pady=2)

        # 执行按钮
        self.btn_run = ttk.Button(left_area, text="开始执行", command=self.start_run, state=tk.DISABLED, width=32)
        self.btn_run.pack(pady=12, anchor=tk.W, padx=5)

        # ========== 右侧信息区 ==========
        right_area = ttk.Frame(top_container)
        right_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(right_area, text="处理日志", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)

        # 日志区
        log_frame = tk.Frame(right_area, bg="#F0F0F0", height=580)
        log_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        log_frame.pack_propagate(False)
        self.logger = LogComponent(log_frame, AppConfig.LOG_MATCH_LIMIT, self)

        # 功能说明区
        ttk.Label(right_area, text="功能说明", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 3), fill=tk.X)
        desc_frame = tk.Frame(right_area, bg="#F5F5F5", bd=1, relief="solid", height=120)
        desc_frame.pack(fill=tk.BOTH, expand=True)
        desc_frame.pack_propagate(False)
        self.desc_label = tk.Label(
            desc_frame, text="", font=AppTools.get_font(12),
            fg="#333333", bg="#F5F5F5", anchor="nw", justify="left",
            wraplength=420, padx=8, pady=6
        )
        self.desc_label.pack(fill=tk.BOTH, expand=True)

        # 禁用控件列表
        self.disabled_widgets = [
            self.btn_source, self.btn_match, self.col_combo_a, self.col_combo_b,
            self.fuzzy_check, self.btn_run
        ]

    def set_separator(self, sep):
        self.sep_entry.delete(0, tk.END)
        self.sep_entry.insert(0, sep)
        self.logger.append(f"已选择分隔符：{repr(sep)}")

    def on_fuzzy_change(self):
        self.business.fuzzy_match = self.fuzzy_var.get()
        self.logger.append(f"模糊匹配{'启用' if self.business.fuzzy_match else '关闭'}")
        self._update_func_desc()

    def _on_func_change(self):
        func = self.selected_func.get()
        if func == "compare":
            self.compare_config_frame.pack(fill="x", padx=5, pady=5)
            self.fuzzy_check.config(state=tk.DISABLED)
            self._switch_compare_mode()
            if self.business.source_col_count > 0 or self.business.match_col_count > 0:
                self._fill_all_col_widgets()
        else:
            self.compare_config_frame.pack_forget()
            self.fuzzy_check.config(state=tk.NORMAL)

        if func in ["union", "compare"]:
            self.fuzzy_var.set(False)
            self.business.fuzzy_match = False
        self._update_func_desc()
        self._check_run_state()

    def _switch_compare_mode(self):
        """切换对比模式：只做显隐，强制重绘"""
        self.frame_single_col.pack_forget()
        self.frame_multi_col.pack_forget()
        self.frame_whole_row.pack_forget()
        self.frame_ignore_cols.pack_forget()

        mode = self.compare_mode.get()
        if mode == "single_col":
            self.frame_single_col.pack(fill="x", pady=2)
        elif mode == "multi_col":
            self.frame_multi_col.pack(fill="x", pady=2)
        elif mode == "whole_row":
            self.frame_whole_row.pack(fill="x", pady=2)
        elif mode == "ignore_cols":
            self.frame_ignore_cols.pack(fill="x", pady=2)

        self.update_idletasks()
        self._update_func_desc()

    def _fill_all_col_widgets(self):
        """导入文件后，统一填充所有列选择控件的数据"""
        source_ready = self.business.source_col_count > 0
        match_ready = self.business.match_col_count > 0

        if source_ready:
            try:
                self._fill_col_combo(self.comp_a_col, self.business.source_col_count, self.business.source_path)
            except:
                self.comp_a_col["values"] = [f"第{i+1}列" for i in range(self.business.source_col_count)]
            self.comp_a_col.current(0)

            self.comp_a_list.delete(0, tk.END)
            self.comp_ignore_list.delete(0, tk.END)
            try:
                self._fill_col_list(self.comp_a_list, self.business.source_col_count, self.business.source_path)
                self._fill_col_list(self.comp_ignore_list, self.business.source_col_count, self.business.source_path)
            except:
                for i in range(self.business.source_col_count):
                    self.comp_a_list.insert(tk.END, f"第{i+1}列")
                    self.comp_ignore_list.insert(tk.END, f"第{i+1}列")

        if match_ready:
            try:
                self._fill_col_combo(self.comp_b_col, self.business.match_col_count, self.business.match_path)
            except:
                self.comp_b_col["values"] = [f"第{i+1}列" for i in range(self.business.match_col_count)]
            self.comp_b_col.current(0)

            self.comp_b_list.delete(0, tk.END)
            try:
                self._fill_col_list(self.comp_b_list, self.business.match_col_count, self.business.match_path)
            except:
                for i in range(self.business.match_col_count):
                    self.comp_b_list.insert(tk.END, f"第{i+1}列")

    def _get_first_row_values(self, file_path: str, col_count: int) -> list:
        """直接从缓存取第一行预览，毫秒级返回，不再读磁盘"""
        if file_path == self.business.source_path:
            values = self.business.source_first_row
        elif file_path == self.business.match_path:
            values = self.business.match_first_row
        else:
            return [f"第{i+1}列" for i in range(col_count)]
        
        if not values:
            return [f"第{i+1}列" for i in range(col_count)]
        
        # 截断超长内容，保证下拉框美观
        return [v[:15] + "..." if len(v) > 15 else v for v in values]

    def _fill_col_combo(self, combo: ttk.Combobox, col_count: int, file_path: str):
        values = self._get_first_row_values(file_path, col_count)
        items = [f"第{i+1}列：{values[i]}" for i in range(col_count)]
        combo["values"] = items
        combo.current(0)

    def _fill_col_list(self, listbox: tk.Listbox, col_count: int, file_path: str):
        values = self._get_first_row_values(file_path, col_count)
        for i in range(col_count):
            listbox.insert(tk.END, f"第{i+1}列：{values[i]}")

    def _update_func_desc(self):
        func = self.selected_func.get()
        fuzzy = self.fuzzy_var.get()
        mode = self.compare_mode.get()

        desc = ""
        if func == "find":
            desc = (
                "【功能1：A中找B 返回新A】\n"
                "作用：从数据A中，筛选出「指定列的值在B指定列存在」的完整行\n"
                "适用：从全量数据里命中目标名单、提取指定用户\n"
                "列数要求：无要求，A、B列数可不同\n"
                "输出：和A格式、列数完全一致的新文件"
            )
            if fuzzy:
                desc += "\n提示：已开启模糊匹配，子串包含即判定匹配"

        elif func == "remove":
            desc = (
                "【功能4：A中剔除B 返回新A】\n"
                "作用：从数据A中，删除「指定列的值在B指定列存在」的完整行\n"
                "适用：剔除黑名单、过滤已处理数据、去重无效名单\n"
                "列数要求：无要求，A、B列数可不同\n"
                "输出：和A格式、列数完全一致的新文件"
            )
            if fuzzy:
                desc += "\n提示：已开启模糊匹配，子串包含即判定匹配"

        elif func == "union":
            desc = (
                "【功能3：A和B的并集】\n"
                "作用：两文件整行合并，自动去除完全重复的行，生成全量无重复数据\n"
                "适用：多日数据汇总、多渠道数据整合、全量数据库更新\n"
                "列数要求：强制列数、列顺序完全一致，否则无法执行\n"
                "输出：列数与AB一致的全量去重文件\n"
                "提示：整行内容完全一致才判定为重复，不支持模糊匹配"
            )

        elif func == "left_join":
            desc = (
                "【功能3：A左连接B 返回A全量+补B字段】\n"
                "作用：以A表为基准保留全部行，匹配成功则拼接B表所有字段，未匹配则B字段留空\n"
                "适用：给主名单补维度信息、打标签，A表行数和行序完全不变\n"
                "列数要求：无要求，A、B列数可不同\n"
                "输出：A所有列在前，B所有列在后（自动加b_前缀避免重名）"
            )
            if fuzzy:
                desc += "\n提示：已开启模糊匹配，子串包含即判定匹配"

        elif func == "intersection":
            desc = (
                "【功能2：A中找B 返回新A+B全部】\n"
                "作用：按指定列匹配，保留两边都命中的行，输出A全部列+B全部列\n"
                "适用：关联两份不同维度的数据、补全字段信息\n"
                "列数要求：无要求，A、B列数可不同\n"
                "输出：A表所有列在前，B表所有列在后（自动加b_前缀避免重名）"
            )
            if fuzzy:
                desc += "\n提示：已开启模糊匹配，子串包含即判定匹配"

        elif func == "compare":
            base_desc = (
                "【功能6：AB文件高级比较】\n"
                "作用：按指定规则对比两文件，一次性输出3个结果文件\n"
                "输出：①A有B无 ②B有A无 ③两者都有\n"
                "说明：100%精确匹配，不支持模糊模式\n\n"
            )
            if mode == "single_col":
                base_desc += "当前模式：单列主键对比\n逻辑：各选一列作为唯一标识，值相等即判定为同一行\n适用：手机号、工单号等单字段对比匹配"
            elif mode == "multi_col":
                base_desc += "当前模式：多列联合主键对比\n逻辑：各选多列按顺序拼接，作为唯一标识对比\n适用：日期+用户+姓名+地址等组合字段"
            elif mode == "whole_row":
                base_desc += "当前模式：整行对比\n逻辑：整行所有列内容完全一致，才判定为同一行\n要求：强制列数、列顺序完全一致\n适用：文件版本对比、查找所有新增/删除/修改行"
            elif mode == "ignore_cols":
                base_desc += "当前模式：指定列忽略对比\n逻辑：跳过选中的列，其余列一致即判定为同一行\n要求：强制列数、列顺序完全一致\n适用：忽略时间戳、导出序号等无关列"
            desc = base_desc

        self.desc_label.config(text=desc)

    def load_source_file(self):
        path = filedialog.askopenfilename(filetypes=[("文本文件", "*.txt;*.csv"), ("所有文件", "*.*")])
        if not path:
            return

        sep = self.sep_entry.get().strip() or AppConfig.DEFAULT_SEPARATOR
        self.logger.append(f"已选择数据A：{FileTools.get_file_full_path(path)}")

        def _task():
            self.business.load_source(path, sep)
            def update_ui():
                size_mb = FileTools.get_file_size_mb(path)
                self.label_source.config(
                    text=f"已导入：{self.business.source_col_count} 列数据（{size_mb:.1f}MB）",
                    foreground="#166534",
                    font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)
                )
                items = self._get_first_row_values(path, self.business.source_col_count)
                combo_items = [f"第{i+1}列：{items[i]}" for i in range(self.business.source_col_count)]
                self.col_combo_a["values"] = combo_items
                self.col_combo_a.current(0)
                self.col_combo_a.config(state="readonly")
                self._check_run_state()
            self.safe_ui_update(update_ui)
            return "数据A加载完成"

        self.run_background_task(_task, "正在加载和清洗数据A...")

    def load_match_file(self):
        path = filedialog.askopenfilename(filetypes=[("文本文件", "*.txt;*.csv"), ("所有文件", "*.*")])
        if not path:
            return

        sep = self.sep_entry.get().strip() or AppConfig.DEFAULT_SEPARATOR
        self.logger.append(f"已选择数据B：{FileTools.get_file_full_path(path)}")

        def _task():
            self.business.load_match(path, sep)
            def update_ui():
                size_mb = FileTools.get_file_size_mb(path)
                self.label_match.config(
                    text=f"已导入：{self.business.match_col_count} 列数据（{size_mb:.1f}MB）",
                    foreground="#166534",
                    font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)
                )
                items = self._get_first_row_values(path, self.business.match_col_count)
                combo_items = [f"第{i+1}列：{items[i]}" for i in range(self.business.match_col_count)]
                self.col_combo_b["values"] = combo_items
                self.col_combo_b.current(0)
                self.col_combo_b.config(state="readonly")
                self._check_run_state()
            self.safe_ui_update(update_ui)
            return "数据B加载完成"

        self.run_background_task(_task, "正在加载和清洗数据B...")

    def _check_run_state(self):
        source_ok = self.business.source_path is not None
        match_ok = self.business.match_path is not None
        state = tk.NORMAL if (source_ok and match_ok) else tk.DISABLED
        self.btn_run.config(state=state)
        if self.selected_func.get() == "compare":
            self._fill_all_col_widgets()

    def _get_selected_col_idx(self, combo: ttk.Combobox) -> int:
        return combo.current()

    def _get_list_selected_indices(self, listbox: tk.Listbox) -> list:
        return list(listbox.curselection())

    def start_run(self):
        func = self.selected_func.get()

        # 同步UI选择的匹配列到业务层
        a_col = self.col_combo_a.current()
        b_col = self.col_combo_b.current()
        if a_col < 0 or b_col < 0:
            messagebox.showwarning("提示", "请先选择A、B的匹配列")
            return
        self.business.source_col_idx = a_col
        self.business.match_col_idx = b_col

        # 按功能映射文件名前缀，清晰可辨
        func_name_map = {
            "find": "匹配_A中找B",
            "remove": "匹配_A中剔除B",
            "union": "匹配_AB并集",
            "left_join": "匹配_A左连接B",
            "intersection": "匹配_AB交集全字段"
        }
        file_tag = func_name_map.get(func, "匹配_处理结果")
        save_path = FileTools.get_save_file_path(file_tag)

        def _task():
            if func == "find":
                count = self.business.run_find(save_path)
                return f"匹配完成！共 {count:,} 条\n保存至：{save_path}"
            elif func == "remove":
                count = self.business.run_remove(save_path)
                return f"剔除完成！共 {count:,} 条\n保存至：{save_path}"
            elif func == "union":
                count = self.business.run_union(save_path)
                return f"并集完成！共 {count:,} 条\n保存至：{save_path}"
            elif func == "left_join":
                count = self.business.run_left_join(save_path)
                return f"左连接完成！共 {count:,} 条\n保存至：{save_path}"
            elif func == "intersection":
                count = self.business.run_intersection(save_path)
                return f"交集匹配完成！共 {count:,} 条\n保存至：{save_path}"
            elif func == "compare":
                return self._run_compare()

        self.run_background_task(_task, "开始处理...")

    def _run_compare(self):
        mode = self.compare_mode.get()
        a_keys = []
        b_keys = []
        ignore_cols = []

        if mode == "single_col":
            a_idx = self._get_selected_col_idx(self.comp_a_col)
            b_idx = self._get_selected_col_idx(self.comp_b_col)
            if a_idx < 0 or b_idx < 0:
                raise ValueError("请选择主键列")
            a_keys = [a_idx]
            b_keys = [b_idx]

        elif mode == "multi_col":
            a_keys = self._get_list_selected_indices(self.comp_a_list)
            b_keys = self._get_list_selected_indices(self.comp_b_list)
            if not a_keys or not b_keys:
                raise ValueError("请至少选择一列作为主键")
            if len(a_keys) != len(b_keys):
                raise ValueError("A和B选择的主键列数量必须一致")

        elif mode == "whole_row":
            pass

        elif mode == "ignore_cols":
            ignore_cols = self._get_list_selected_indices(self.comp_ignore_list)
            if not ignore_cols:
                raise ValueError("请至少选择一列要忽略的列")

        outputs = self.business.run_compare(
            mode, a_keys, b_keys, ignore_cols
        )
        return (
            f"对比完成！三个结果文件已生成：\n"
            f"1. A有B无：{os.path.basename(outputs['a_only'])}\n"
            f"2. B有A无：{os.path.basename(outputs['b_only'])}\n"
            f"3. 两者都有：{os.path.basename(outputs['both'])}\n"
            f"保存目录：{os.path.dirname(outputs['a_only'])}"
        )
class ProcessBusiness:
    def __init__(self):
        self.file_path: Optional[str] = None
        self.encoding: str = "utf-8-sig"
        self._sep = "\0"  # 单列固定分隔符

        # 操作栈：按顺序保存所有处理步骤，每次执行从头应用
        self._ops: list[tuple[str, dict]] = []
        # 状态通知标志：记录是否已经触发过“已修改”状态
        self._modified_notified: bool = False

        # ===================== 内存状态管理核心模块 =====================
        self._status_callback: Optional[Callable[[bool], None]] = None  # UI状态回调

    def register_status_callback(self, callback: Callable[[bool], None]):
        """注册UI状态回调，状态变化自动通知，线程安全"""
        self._status_callback = callback

    def _trigger_modified(self):
        """内部触发修改状态：仅在首次从「未修改」变为「已修改」时触发一次回调"""
        if not self._modified_notified:
            self._modified_notified = True
            if self._status_callback:
                self._status_callback(True)

    @property
    def is_memory_modified(self) -> bool:
        """只读属性：操作栈非空即为内存已修改"""
        return len(self._ops) > 0

    def reset_to_original(self) -> bool:
        """重置到原始文件状态，原子操作"""
        if not self.file_path:
            return False
        
        # 清空操作栈，重置通知标志，回调通知UI
        self._ops.clear()
        self._modified_notified = False
        if self._status_callback:
            self._status_callback(False)
        return True

    def load_init_state(self):
        """文件加载/切换数据源时统一初始化状态，唯一入口"""
        self._ops.clear()
        self._modified_notified = False
        if self._status_callback:
            self._status_callback(False)
    # ==============================================================

    def _build_current_rel(self) -> DuckRelation:
        """内部方法：当场构建当前完整数据集
        每次调用都新建连接，从头应用所有操作，彻底规避跨线程问题
        """
        if not self.file_path:
            raise RuntimeError("未加载数据源")
        rel = DuckTools.read_single_column(self.file_path, self.encoding)
        # 依次应用所有操作
        for op_name, params in self._ops:
            if op_name == "replace":
                rel = DuckTools.op_replace(rel, "col0", params["old"], params["new"])
            elif op_name == "insert":
                rel = DuckTools.op_insert(rel, "col0", params["prefix"], params["suffix"])
            elif op_name == "upper":
                rel = DuckTools.op_upper(rel, "col0")
            elif op_name == "lower":
                rel = DuckTools.op_lower(rel, "col0")
            elif op_name == "add_serial":
                rel = DuckTools.op_add_serial(rel, "col0", start=params["start"])
        return rel

    def preview_data(self, rows: int) -> list[str]:
        """预览前N行数据，直接返回字符串列表，完全不依赖pandas"""
        if not self.file_path:
            return []
        rel = self._build_current_rel()
        result = rel.limit(rows).native_rel.fetchall()
        return [row[0] if row[0] is not None else '' for row in result]

    def export_current(self, output_path: str, sep: str = "\t"):
        """导出当前数据到文件，统一封装"""
        rel = self._build_current_rel()
        DuckTools.safe_to_csv(rel, output_path, sep, encoding="utf-8-sig")

    # ========== 字符替换 ==========
    def replace_string(self, old: str, new: str) -> int:
        """内存模式：加入操作栈，修改当前数据集，触发修改标记"""
        self._trigger_modified()  # 先触发状态，再入栈
        self._ops.append(("replace", {"old": old, "new": new}))
        rel = self._build_current_rel()
        return DuckTools.count_rows(rel)

    # ========== 首尾插入 ==========
    def insert_suffix_prefix(self, prefix: str, suffix: str) -> int:
        """内存模式：加入操作栈，修改当前数据集，触发修改标记"""
        self._trigger_modified()
        self._ops.append(("insert", {"prefix": prefix, "suffix": suffix}))
        rel = self._build_current_rel()
        return DuckTools.count_rows(rel)

    # ========== 纯行首添加自增序号 ==========
    def add_serial_only(self) -> int:
        """内存模式：加入操作栈，修改当前数据集，触发修改标记"""
        self._trigger_modified()
        self._ops.append(("add_serial", {"start": 1}))
        rel = self._build_current_rel()
        return DuckTools.count_rows(rel)

    # ========== 英文字母大小写转换 ==========
    def convert_case(self, to_upper: bool) -> int:
        """内存模式：加入操作栈，修改当前数据集，触发修改标记"""
        op_name = "upper" if to_upper else "lower"
        self._trigger_modified()
        self._ops.append((op_name, {}))
        rel = self._build_current_rel()
        return DuckTools.count_rows(rel)
        
    # ========== 文件分割 ==========
    def split_file(self, mode: str, value: int, output_prefix: str) -> int:
        """统一文件分割，全量走DuckDB，不再返回DataFrame"""
        rel = self._build_current_rel()
        total = DuckTools.count_rows(rel)
        if mode == "parts":
            value = max(1, math.ceil(total / value))
        
        return DuckTools.split_file_by_lines(rel, output_prefix, value, sep="\t")


    # ========== 字符统计 ==========
    def count_char_times(self, char: str, log_callback=None) -> int:
        rel = self._build_current_rel()
        return DuckTools.op_count_char(rel, "col0", char)

    # ========== 字符筛选（不修改当前数据集，直接导出结果） ==========
    def filter_contain_char(self, char: str, output_path: str, log_callback=None) -> int:
        rel = self._build_current_rel()
        rel_result = DuckTools.op_filter_contains(rel, "col0", char)
        DuckTools.safe_to_csv(rel_result, output_path, "\t", encoding="utf-8-sig")
        return DuckTools.count_rows(rel_result)

    def switch_to_result(self, result_path: str):
        """切换数据源为结果文件，自动重新初始化状态"""
        self.file_path = result_path
        self.encoding = "utf-8-sig"
        # 切换数据源后清空操作栈和状态标志
        self._ops.clear()
        self._modified_notified = False
        self.load_init_state()
        
class ProcessTab(BaseTab):
    def __init__(self, parent):
        super().__init__(parent)
        self.business = ProcessBusiness()
        self.case_is_upper = False
        self.active_entry = None
        self.selected_opt = tk.StringVar(value="")
        self.split_mode = tk.StringVar(value="lines")
        self.export_preview = tk.BooleanVar(value=False)
        self.last_output_path = ""

        self.btn_reset_memory = None
        self.lbl_memory_status = None
        self.business.register_status_callback(self._update_memory_status_ui)
        
        self._init_ui()

    def _init_ui(self):
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        left_area = ttk.Frame(main_container, width=410)
        left_area.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        left_area.pack_propagate(False)
        self._init_left_operations(left_area)

        right_area = ttk.Frame(main_container)
        right_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self._init_right_preview_log(right_area)

    def _init_left_operations(self, parent):
        ttk.Label(parent, text="1. 导入待处理文件", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(6, 2), fill=tk.X, padx=5)
        self.btn_import = ttk.Button(parent, text="选择文件", command=self.load_process_file)
        self.btn_import.pack(pady=2, anchor=tk.W, padx=5)
        self.label_file = ttk.Label(parent, text="未导入", foreground="#666666", anchor="w")
        self.label_file.pack(pady=3, anchor=tk.W, padx=5)

        # ========== 完全修复：padx 移到 pack 中 ==========
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, padx=5, pady=(15, 5))

        ttk.Label(
            row_frame, 
            text="2. 选择要执行的操作", 
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)
        ).pack(side=tk.LEFT)

        self.btn_reset_memory = ttk.Button(
            row_frame, 
            text="重置缓存修改", 
            command=self.reset_memory_changes,
            state=tk.DISABLED
        )
        self.btn_reset_memory.pack(side=tk.LEFT, padx=(15, 10))

        self.lbl_memory_status = ttk.Label(
            row_frame,
            text="未加载文件",
            font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL, bold=True),
            foreground="#888888"
        )
        self.lbl_memory_status.pack(side=tk.LEFT, padx=5)
        # ==========================================================

        self.opt_frames = {}
        self._create_radio_button(parent, "预览前N行数据", "preview")
        self._create_radio_button(parent, "字符替换", "replace")
        self._create_radio_button(parent, "每行首尾插入字符", "insert")
        self._create_radio_button(parent, "分割文件", "split")
        self._create_radio_button(parent, "统计字符出现次数", "count")
        self._create_radio_button(parent, "筛选包含特定字符的行", "filter")
        self._build_all_opt_frames(parent)

        self.selected_opt.trace_add("write", lambda *args: self.switch_opt_frame())
        self.disabled_widgets = [self.btn_import, self.btn_reset_memory]

    def _create_radio_button(self, parent, text, value):
        rdo = ttk.Radiobutton(parent, text=text, variable=self.selected_opt, value=value)
        rdo.pack(pady=3, anchor=tk.W, padx=5)
        self.opt_frames[value] = ttk.Frame(parent, padding=(20, 5, 10, 15))

    def _build_all_opt_frames(self, parent):
        self._build_preview_frame()
        self._build_replace_frame()
        self._build_insert_frame()
        self._build_split_frame()
        self._build_count_frame()
        self._build_filter_frame()

    def _init_right_preview_log(self, parent):
        ttk.Label(parent, text="数据预览区", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)
        preview_container = ttk.Frame(parent, height=200)
        preview_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        preview_container.pack_propagate(False)

        v_scroll = ttk.Scrollbar(preview_container, orient=tk.VERTICAL)
        h_scroll = ttk.Scrollbar(preview_container, orient=tk.HORIZONTAL)
        self.preview_text = tk.Text(
            preview_container, font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL), wrap=tk.NONE, bg="#F5F5F5",
            yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set
        )
        v_scroll.config(command=self.preview_text.yview)
        h_scroll.config(command=self.preview_text.xview)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.preview_text.config(state=tk.DISABLED)

        ttk.Label(parent, text="处理日志", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)
        log_frame = tk.Frame(parent, bg="#F0F0F0", height=200)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        log_frame.pack_propagate(False)
        self.logger = LogComponent(log_frame, AppConfig.LOG_PROCESS_LIMIT, self)

    def _build_preview_frame(self):
        f = self.opt_frames["preview"]
        ttk.Label(f, text="预览行数：").grid(row=0, column=0, padx=(0, 5))
        self.preview_rows = ttk.Combobox(f, values=["100", "500", "1000", "5000", "10000"], width=10, state="readonly")
        self.preview_rows.current(2)
        self.preview_rows.grid(row=0, column=1)
        ttk.Checkbutton(f, text="同时导出预览结果到文件", variable=self.export_preview).grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))
        ttk.Button(f, text="执行预览", command=self.run_preview, width=15).grid(row=2, column=0, columnspan=2, pady=(10, 0))

    def _build_replace_frame(self):
        f = self.opt_frames["replace"]
        ttk.Label(f, text="被替换字符：").grid(row=0, column=0, sticky="w", pady=2)
        self.replace_old = ttk.Entry(f, width=18)
        self.replace_old.grid(row=0, column=1, pady=2, padx=10)
        ttk.Label(f, text="替换为：").grid(row=1, column=0, sticky="w", pady=2)
        self.replace_new = ttk.Entry(f, width=18)
        self.replace_new.grid(row=1, column=1, pady=2, padx=10)

        self.replace_old.bind("<FocusIn>", lambda e: setattr(self, "active_entry", self.replace_old))
        self.replace_new.bind("<FocusIn>", lambda e: setattr(self, "active_entry", self.replace_new))
        self.active_entry = self.replace_old

        ttk.Label(f, text="常用特殊字符：").grid(row=2, column=0, sticky="w", pady=(10, 2))
        char_frame = ttk.Frame(f)
        char_frame.grid(row=2, column=1, pady=(10, 2), sticky="w")
        for idx, (name, char) in enumerate(AppConfig.SPECIAL_CHARS.items()):
            ttk.Button(char_frame, text=name, width=5, command=lambda c=char: self.insert_special_char(c)).grid(
                row=idx//3, column=idx%3, padx=2, pady=1
            )

        self.btn_case_convert = ttk.Button(
            f, text="一键大小写转换",
            command=self.run_case_convert,
            width=12
        )
        self.btn_case_convert.grid(row=3, column=0, sticky="w", pady=(10, 0))

        ttk.Button(f, text="执行缓存替换", command=self.run_replace, width=15).grid(row=4, column=0, pady=(10, 0))
        self.btn_export_replace = ttk.Button(f, text="导出修改后数据", command=self.export_data, width=15)
        self.btn_export_replace.grid(row=4, column=1, pady=(10, 0))

    def _build_insert_frame(self):
        f = self.opt_frames["insert"]
        ttk.Label(f, text="开头插入：").grid(row=0, column=0, sticky="w", pady=2)
        self.insert_pre = ttk.Entry(f, width=18)
        self.insert_pre.grid(row=0, column=1, pady=2, padx=10)

        ttk.Label(f, text="末尾插入：").grid(row=1, column=0, sticky="w", pady=2)
        self.insert_suf = ttk.Entry(f, width=18)
        self.insert_suf.grid(row=1, column=1, pady=2, padx=10)

        self.btn_add_serial = ttk.Button(
            f, text="行首添加序号", command=self.run_add_serial, width=15
        )
        self.btn_add_serial.grid(row=2, column=0, pady=(15, 0))

        ttk.Button(f, text="执行缓存插入", command=self.run_insert, width=15).grid(row=3, column=0, pady=(15, 0))
        self.btn_export_insert = ttk.Button(f, text="导出修改后数据", command=self.export_data, width=15)
        self.btn_export_insert.grid(row=3, column=1, pady=(15, 0))

        self.insert_pre.bind("<FocusIn>", lambda e: setattr(self, "active_entry", self.insert_pre))
        self.insert_suf.bind("<FocusIn>", lambda e: setattr(self, "active_entry", self.insert_suf))
        self.active_entry = self.insert_pre

    def _build_split_frame(self):
        f = self.opt_frames["split"]
        ttk.Radiobutton(f, text="按行数分割（每个文件N行）", variable=self.split_mode, value="lines", command=self.update_split_mode).grid(row=0, column=0, sticky="w", pady=2)
        self.split_lines = ttk.Entry(f, width=12)
        self.split_lines.insert(0, "10000")
        self.split_lines.grid(row=0, column=1, pady=2, padx=10)
        ttk.Radiobutton(f, text="平均分成N份", variable=self.split_mode, value="parts", command=self.update_split_mode).grid(row=1, column=0, sticky="w", pady=2)
        self.split_parts = ttk.Entry(f, width=12, state=tk.DISABLED)
        self.split_parts.insert(0, "10")
        self.split_parts.grid(row=1, column=1, pady=2, padx=10)
        ttk.Button(f, text="执行分割", command=self.run_split, width=15).grid(row=2, column=0, columnspan=2, pady=(10, 0))
        AppTools.restrict_numeric_input(self.split_lines)
        AppTools.restrict_numeric_input(self.split_parts)

    def _build_count_frame(self):
        f = self.opt_frames["count"]
        ttk.Label(f, text="要统计的字符：").grid(row=0, column=0, sticky="w", pady=2)
        self.count_char = ttk.Entry(f, width=18)
        self.count_char.grid(row=0, column=1, pady=2, padx=10)
        
        self.count_result_label = ttk.Label(
            f, text="统计结果：", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True), foreground="#006400"
        )
        self.count_result_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        
        ttk.Button(f, text="执行统计", command=self.run_count, width=15).grid(row=2, column=0, columnspan=2, pady=(10, 0))

    def _build_filter_frame(self):
        f = self.opt_frames["filter"]
        ttk.Label(f, text="要筛选的字符：").grid(row=0, column=0, sticky="w", pady=2)
        self.filter_char = ttk.Entry(f, width=18)
        self.filter_char.grid(row=0, column=1, pady=2, padx=10)
        
        self.filter_result_label = ttk.Label(
            f, text="符合条件行数：", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True), foreground="#006400"
        )
        self.filter_result_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        
        ttk.Button(f, text="执行筛选", command=self.run_filter, width=15).grid(row=2, column=0, columnspan=2, pady=(10, 0))

    def switch_opt_frame(self):
        for frame in self.opt_frames.values():
            frame.pack_forget()
        current = self.selected_opt.get()
        if current in self.opt_frames:
            self.opt_frames[current].pack(fill=tk.X)
        self.update_idletasks()

    def update_split_mode(self):
        if self.split_mode.get() == "lines":
            self.split_lines.config(state=tk.NORMAL)
            self.split_parts.config(state=tk.DISABLED)
        else:
            self.split_lines.config(state=tk.DISABLED)
            self.split_parts.config(state=tk.NORMAL)

    def insert_special_char(self, char):
        if self.active_entry:
            self.active_entry.delete(0, tk.END)
            self.active_entry.insert(0, char)
            self.logger.append(f"插入特殊字符：{repr(char)}")

    def update_preview_display(self, lines: list[str], max_rows=AppConfig.PREVIEW_MAX_ROWS):
        def _update():
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            if lines:
                show_lines = lines[:max_rows]
                self.preview_text.insert(1.0, "\n".join(show_lines))
            self.preview_text.config(state=tk.DISABLED)
        self.after(0, _update)

    def _check_file_loaded(self) -> bool:
        if self.business.file_path is None:
            self.logger.append("处理失败：您未导入任何文件，请确认", AppConfig.COLOR_ERROR)
            return False
        return True

    def _after_memory_process(self):
        df = self.business.preview_data(AppConfig.PREVIEW_MAX_ROWS)
        self.update_preview_display(df)

    def _after_stream_process(self, save_path: str):
        # 直接读取文件前100行做预览
        lines = []
        try:
            enc = FileTools.detect_file_encoding(save_path)
            with open(save_path, 'r', encoding=enc) as f:
                for i, line in enumerate(f):
                    if i >= 100:
                        break
                    lines.append(line.rstrip('\n'))
        except:
            pass
        self.update_preview_display(lines)

        self.business.switch_to_result(save_path)
        self.logger.append("已自动切换数据源为本次结果文件，可继续链式处理")
        self.last_output_path = save_path

    def _update_memory_status_ui(self, modified: bool):
        def _update():
            if modified:
                self.lbl_memory_status.config(text="内存已修改", foreground=AppConfig.COLOR_ERROR)
            else:
                self.lbl_memory_status.config(text="内存未修改", foreground=AppConfig.COLOR_SUCCESS)
            self.btn_reset_memory.config(state=tk.NORMAL)
        self.safe_ui_update(_update)

    def reset_memory_changes(self):
        if self.is_loading:
            self.logger.append("正在处理中，请稍后再重置", AppConfig.COLOR_WARNING)
            return

        if self.business.reset_to_original():
            preview_df = self.business.preview_data(AppConfig.PREVIEW_MAX_ROWS)
            self.update_preview_display(preview_df)
            self.logger.append("内存数据已重置为原始文件状态", is_final=True)
        else:
            self.logger.append("当前模式无法重置内存", AppConfig.COLOR_WARNING)

    def load_process_file(self):
        path = filedialog.askopenfilename(filetypes=[("文本文件", "*.txt;*.csv"), ("所有文件", "*.*")])
        if not path:
            return

        full_path = FileTools.get_file_full_path(path)
        is_valid, check_msg = FileTools.check_file_valid(path)
        self.logger.append(f"已选择文件路径：{full_path}")

        if not is_valid:
            self.logger.append(f"文件校验失败：{check_msg}", AppConfig.COLOR_ERROR)
            messagebox.showerror("文件错误", f"{check_msg}\n请检查文件后重新选择！")
            return

        def _task():
            self.last_output_path = ""
            load_info = FileTools.prepare_file_load(path)
            if not load_info["valid"]:
                raise RuntimeError(load_info["message"])
            
            self.business.file_path = path
            self.business.encoding = load_info["encoding"]
            self.business.load_init_state()

            rel = DuckTools.read_single_column(path, load_info["encoding"])
            total_rows = DuckTools.count_rows(rel)
            raw_lines = FileTools.count_file_lines_fast(path)
            dedup = raw_lines - total_rows

            def update_ui():
                self.label_file.config(
                    text=f"已导入：{total_rows:,} 条数据",
                    foreground="#166534",
                    font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)
                )
                self.btn_export_replace.config(text="导出修改后数据")
                self.btn_export_insert.config(text="导出修改后数据")
            self.safe_ui_update(update_ui)
            if dedup > 0:
                self.logger.append(f"去除重复数据：{dedup:,} 条")
            preview_df = self.business.preview_data(AppConfig.PREVIEW_MAX_ROWS)
            self.update_preview_display(preview_df)
            return "文件加载完成"

        self.run_background_task(_task, "开始加载文件...")

    def export_data(self):
        if not self._check_file_loaded():
            return

        path = FileTools.get_save_file_path("处理_修改结果")
        try:
            self.business.export_current(path)
            self.last_output_path = path
            self.logger.append(f"导出成功：\n{path}", is_final=True)
        except Exception as e:
            self.logger.append(f"导出失败：{str(e)}", AppConfig.COLOR_ERROR)

    def run_preview(self):
        if not self._check_file_loaded():
            return
        valid, rows = AppTools.validate_positive_int(self.preview_rows.get())
        if not valid:
            messagebox.showerror("输入错误", "预览行数必须是大于0的正整数")
            return
        lines = self.business.preview_data(rows)
        self.update_preview_display(lines)
        if self.export_preview.get():
            save_path = FileTools.get_save_file_path("处理_预览结果")
            with open(save_path, 'w', encoding='utf-8-sig') as f:
                f.write('\n'.join(lines) + '\n')
        self.logger.append("预览完成", is_final=True)

    def run_replace(self):
        if not self._check_file_loaded():
            return
        old = AppTools.parse_escape_string(self.replace_old.get())
        new = AppTools.parse_escape_string(self.replace_new.get())

        if not old.strip() and not new.strip():
            messagebox.showwarning("提示", "替换内容和目标内容均为空，无需执行")
            return

        def _task():
            count = self.business.replace_string(old, new)
            return count

        def _ui_callback(count):
            self._after_memory_process()
            self.logger.append(f"替换完成！共 {count:,} 行（内存已更新，可继续修改后再导出）", is_final=True)
            self.update_idletasks()

        if self.is_loading:
            self.logger.append("正在处理中，请稍候...")
            return
        self.is_loading = True
        AppTools.batch_update_widget_state(self.disabled_widgets, tk.DISABLED)
        self.logger.append("开始执行替换...")

        def _bg_task():
            try:
                res = _task()
                self.safe_ui_update(lambda: _ui_callback(res))
            except Exception as e:
                self.logger.append(f"替换失败：{str(e)}", AppConfig.COLOR_ERROR)
            finally:
                self.safe_ui_update(self._task_cleanup)

        threading.Thread(target=_bg_task, daemon=True).start()

    def run_case_convert(self):
        if not self._check_file_loaded():
            return
        self.case_is_upper = not self.case_is_upper
        mode_text = "大写" if self.case_is_upper else "小写"

        def _task():
            count = self.business.convert_case(self.case_is_upper)
            return count

        def _ui_callback(count):
            self._after_memory_process()
            self.logger.append(f"已转换为全{mode_text}！共 {count:,} 行（内存已更新，可继续修改后再导出）", is_final=True)
            self.update_idletasks()

        if self.is_loading:
            self.logger.append("正在处理中，请稍候...")
            return
        self.is_loading = True
        AppTools.batch_update_widget_state(self.disabled_widgets, tk.DISABLED)
        self.logger.append(f"开始转换为全{mode_text}...")

        def _bg_task():
            try:
                res = _task()
                self.safe_ui_update(lambda: _ui_callback(res))
            except Exception as e:
                self.logger.append(f"转换失败：{str(e)}", AppConfig.COLOR_ERROR)
            finally:
                self.safe_ui_update(self._task_cleanup)

        threading.Thread(target=_bg_task, daemon=True).start()

    def run_insert(self):
        if not self._check_file_loaded():
            return
        prefix = AppTools.parse_escape_string(self.insert_pre.get())
        suffix = AppTools.parse_escape_string(self.insert_suf.get())

        if not prefix.strip() and not suffix.strip():
            messagebox.showwarning("提示", "开头和末尾均无插入内容，无需执行")
            return

        def _task():
            count = self.business.insert_suffix_prefix(prefix, suffix)
            self._after_memory_process()
            return f"首尾插入完成！共 {count:,} 行（内存已更新，可继续修改后再导出）"

        self.run_background_task(_task, "开始执行首尾插入...")

    def run_add_serial(self):
        if not self._check_file_loaded():
            return

        def _task():
            count = self.business.add_serial_only()
            self._after_memory_process()
            return f"行首序号添加完成！共 {count:,} 行（内存已更新，可继续修改后再导出）"

        self.run_background_task(_task, "开始在行首添加序号...")

    def run_split(self):
        if not self._check_file_loaded():
            return
        mode = self.split_mode.get()
        input_text = self.split_lines.get() if mode == "lines" else self.split_parts.get()
        valid, val = AppTools.validate_positive_int(input_text)
        if not valid:
            messagebox.showerror("输入错误", "请输入大于0的正整数")
            return

        output_prefix_raw = FileTools.get_save_file_path("处理_文件分割")
        save_dir = os.path.dirname(output_prefix_raw)
        base_name = os.path.splitext(os.path.basename(output_prefix_raw))[0]
        output_prefix = os.path.join(save_dir, base_name)

        def _task():
            file_count = self.business.split_file(mode, val, output_prefix)
            first_output_file = f"{output_prefix}_第1份.txt"
            self.last_output_path = first_output_file
            return f"分割完成！共生成 {file_count} 个文件\n保存至：{save_dir}"

        self.run_background_task(_task, "开始执行分割...")

    def run_count(self):
        if not self._check_file_loaded():
            return
        char = AppTools.parse_escape_string(self.count_char.get().strip())
        if not char:
            messagebox.showwarning("提示", "请输入要统计的字符")
            return

        def _task():
            cnt = self.business.count_char_times(char, self.logger.append)

            def update_ui():
                self.count_result_label.config(text=f"统计结果：字符 {repr(char)} 共出现 {cnt:,} 次")
            self.safe_ui_update(update_ui)
            return f"统计完成：字符{repr(char)}共出现 {cnt:,} 次"

        self.run_background_task(_task, "开始执行统计...")

    def run_filter(self):
        if not self._check_file_loaded():
            return
        char = AppTools.parse_escape_string(self.filter_char.get().strip())
        if not char:
            messagebox.showwarning("提示", "请输入要筛选的字符")
            return

        save_path = FileTools.get_save_file_path("处理_字符筛选")

        def _task():
            total = self.business.filter_contain_char(char, save_path, self.logger.append)

            def update_ui():
                self.filter_result_label.config(text=f"符合条件行数：{total:,} 行")
            self.safe_ui_update(update_ui)

            if total > 0:
                self._after_stream_process(save_path)
                return f"筛选完成！共{total}条\n已自动切换为结果文件，可继续链式处理"
            else:
                return "筛选完成！无符合条件的数据"

        self.run_background_task(_task, "开始执行筛选...")


class WpsTab(BaseTab):
    """WPS表格处理标签页"""
    def __init__(self, parent):
        super().__init__(parent)
        self.business = None
        self.high_fidelity_var = tk.BooleanVar(value=False)  # 默认快速模式
        self._init_ui()

    def _init_ui(self):
        top_container = ttk.Frame(self)
        top_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 左侧操作区
        left_area = ttk.Frame(top_container, width=470)
        left_area.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_area.pack_propagate(False)

        # ========== 第1行：导入文件 ==========
        ttk.Label(left_area, text="1. 导入表格文件", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(0, 4), fill=tk.X, padx=5)
        
        file_row = ttk.Frame(left_area)
        file_row.pack(fill=tk.X, padx=5, pady=2)

        # ========== 左列：选择按钮 + 文件状态文字 ==========
        left_col = ttk.Frame(file_row)
        left_col.pack(side=tk.LEFT, anchor="n")

        self.btn_load_file = ttk.Button(left_col, text="选择Excel文件", command=self.load_excel_file, width=14)
        self.btn_load_file.pack(side=tk.TOP, anchor="w")

        self.label_file_info = ttk.Label(
            left_col, text="未导入", foreground="#666666",
            anchor="w", wraplength=220, justify="left"
        )
        self.label_file_info.pack(side=tk.TOP, anchor="w", pady=(4, 0))

        # ========== 右列：模式选择 + 表头选项（绑定为一组，位置固定） ==========
        right_col = ttk.Frame(file_row)
        right_col.pack(side=tk.RIGHT, padx=(0, 60), anchor="n")

        # 两个模式单选（上下排列）
        self.radio_fast = ttk.Radiobutton(
            right_col, text="数据模式（速度快）",
            variable=self.high_fidelity_var, value=False,
            command=self._on_mode_change
        )
        self.radio_fast.pack(side=tk.TOP, anchor="w")
        self.radio_hifi = ttk.Radiobutton(
            right_col, text="格式模式（速度慢，保留格式）",
            variable=self.high_fidelity_var, value=True,
            command=self._on_mode_change
        )
        self.radio_hifi.pack(side=tk.TOP, anchor="w")

        # 首行作为表头：跟在模式下方，同属右列，永远绑定不动
        self.header_var = tk.BooleanVar(value=True)
        self.chk_header = ttk.Checkbutton(
            right_col, text="首行作为表头导出",
            variable=self.header_var, command=self._on_header_change
        )
        self.chk_header.pack(side=tk.TOP, anchor="w", pady=(6, 0))

        # 自动列名提示：复选框正下方
        self.label_auto_col = ttk.Label(right_col, text="系统自动分配列名", foreground="#d97706")
        self.label_auto_col.pack(side=tk.TOP, anchor="w")
        self.label_auto_col.pack_forget()

        # ========== 第2部分：Sheet列表 ==========
        ttk.Label(left_area, text="Sheet列表（可多选）", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(10, 4), fill=tk.X, padx=5)
        sheet_frame = ttk.Frame(left_area)
        sheet_frame.pack(fill=tk.X, padx=5, pady=2)
        self.sheet_listbox = tk.Listbox(
            sheet_frame, selectmode="extended", height=8,
            exportselection=False, bg="#F8F8F8"
        )
        scroll_sheet = ttk.Scrollbar(sheet_frame, orient=tk.VERTICAL, command=self.sheet_listbox.yview)
        self.sheet_listbox.config(yscrollcommand=scroll_sheet.set)
        scroll_sheet.pack(side=tk.RIGHT, fill=tk.Y)
        self.sheet_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ========== 第3行：三个按钮同行 ==========
        btn_row1 = ttk.Frame(left_area)
        btn_row1.pack(fill=tk.X, padx=5, pady=(10, 4))
        
        self.btn_split_all = ttk.Button(btn_row1, text="拆分所有SHEET", command=self.run_split_all, state=tk.DISABLED)
        self.btn_split_all.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        
        self.btn_split_selected = ttk.Button(btn_row1, text="拆分选中SHEET", command=self.run_split_selected, state=tk.DISABLED)
        self.btn_split_selected.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        
        self.btn_concat = ttk.Button(btn_row1, text="同列拼接所有SHEET", command=self.run_concat, state=tk.DISABLED)
        self.btn_concat.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # ========== 第4部分：高级筛选区域 ==========
        filter_title_label = ttk.Label(
            left_area,
            text="单表高级筛选拆分",
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)
        )
        filter_frame = ttk.LabelFrame(left_area, labelwidget=filter_title_label)
        filter_frame.pack(fill="x", padx=5, pady=(10, 4))

        # 主键列名行
        key_row = ttk.Frame(filter_frame)
        key_row.pack(fill="x", padx=8, pady=(6, 4))
        ttk.Label(key_row, text="主键列名").pack(side="left", padx=(0, 4))
        self.entry_key_col = ttk.Entry(key_row, width=16)
        self.entry_key_col.pack(side="left", padx=(0, 8))
        ttk.Label(key_row, text="留空默认取第一列", foreground="#666").pack(side="left")

        # ---- 条件组1 ----
        ttk.Label(filter_frame, text="满足以下所有条件（AND），最多支持10个条件判定", font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL, bold=True)).pack(
            anchor="w", padx=8, pady=(4, 2)
        )

        # 条件展示区：横向排列，自动换行
        row_list1 = ttk.Frame(filter_frame)
        row_list1.pack(fill="x", padx=8, pady=2)
        self.text_filter1 = tk.Text(row_list1, height=3, wrap="word", bg="#F8F8F8", state="disabled", relief="solid", borderwidth=1)
        self.text_filter1.pack(fill="x", expand=True)

        # 条件输入行：列、操作符、值输入框
        input_row1 = ttk.Frame(filter_frame)
        input_row1.pack(fill="x", padx=8, pady=(4, 2))

        self.combo_col1 = ttk.Combobox(input_row1, state="disabled", width=8)
        self.combo_col1.pack(side="left", padx=(0, 4))

        self.combo_op1 = ttk.Combobox(input_row1, state="readonly", width=9)
        self.combo_op1["values"] = [
            "等于", "大于", "小于", "大于等于", "小于等于", "不等于",
            "包含", "不包含", "开头是", "结尾是",
            "为空", "非空",
            "长度等于", "长度大于", "长度小于",
            "属于", "不属于",
            "介于"
        ]
        self.combo_op1.current(0)
        self.combo_op1.pack(side="left", padx=4)
        self.combo_op1.bind("<<ComboboxSelected>>", self._on_op_change1)

        self.entry_val1_1 = ttk.Entry(input_row1)
        self.entry_val1_1.pack(side="left", padx=4, fill="x", expand=True)

        self.entry_val1_2 = ttk.Entry(input_row1, width=8)
        self.entry_val1_2.pack(side="left", padx=4)
        self.entry_val1_2.pack_forget()  # 默认隐藏第二个值输入框

        # 按钮行：添加、删除、添加OR组
        btn_row_filter1 = ttk.Frame(filter_frame)
        btn_row_filter1.pack(fill="x", padx=8, pady=(2, 6))

        self.btn_add_cond1 = ttk.Button(btn_row_filter1, text="添加条件", command=self._add_condition1, width=8)
        self.btn_add_cond1.pack(side="left", padx=(0, 4))

        self.btn_del_cond1 = ttk.Button(btn_row_filter1, text="删除条件", command=self._del_condition1, width=10)
        self.btn_del_cond1.pack(side="left", padx=4)

        self.btn_add_or_group = ttk.Button(btn_row_filter1, text="＋ 添加另一组AND条件，OR关系", command=self._enable_group2)
        self.btn_add_or_group.pack(side="left", padx=(8, 0))

        # ---- 条件组2（默认隐藏） ----
        self.group2_frame = ttk.Frame(filter_frame)

        ttk.Label(self.group2_frame, text="━ 或 OR ━", foreground="#d97706", font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL, bold=True)).pack(
            anchor="center", pady=(0, 2)
        )

        group2_header = ttk.Frame(self.group2_frame)
        group2_header.pack(fill="x", padx=8)
        ttk.Label(group2_header, text="满足以下所有条件（AND）", font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL, bold=True)).pack(side="left")
        self.btn_del_group2 = ttk.Button(group2_header, text="删除本组", command=self._disable_group2, width=8)
        self.btn_del_group2.pack(side="right")

        # 条件展示区2
        row_list2 = ttk.Frame(self.group2_frame)
        row_list2.pack(fill="x", padx=8, pady=2)
        self.text_filter2 = tk.Text(row_list2, height=3, wrap="word", bg="#F8F8F8", state="disabled", relief="solid", borderwidth=1)
        self.text_filter2.pack(fill="x", expand=True)

        # 条件输入行2
        input_row2 = ttk.Frame(self.group2_frame)
        input_row2.pack(fill="x", padx=8, pady=(4, 2))

        self.combo_col2 = ttk.Combobox(input_row2, state="disabled", width=8)
        self.combo_col2.pack(side="left", padx=(0, 4))

        self.combo_op2 = ttk.Combobox(input_row2, state="readonly", width=9)
        self.combo_op2["values"] = self.combo_op1["values"]
        self.combo_op2.current(0)
        self.combo_op2.pack(side="left", padx=4)
        self.combo_op2.bind("<<ComboboxSelected>>", self._on_op_change2)

        self.entry_val2_1 = ttk.Entry(input_row2)
        self.entry_val2_1.pack(side="left", padx=4, fill="x", expand=True)

        self.entry_val2_2 = ttk.Entry(input_row2, width=8)
        self.entry_val2_2.pack(side="left", padx=4)
        self.entry_val2_2.pack_forget()

        # 按钮行2
        btn_row_filter2 = ttk.Frame(self.group2_frame)
        btn_row_filter2.pack(fill="x", padx=8, pady=(2, 6))

        self.btn_add_cond2 = ttk.Button(btn_row_filter2, text="添加条件", command=self._add_condition2, width=8)
        self.btn_add_cond2.pack(side="left", padx=(0, 4))

        self.btn_del_cond2 = ttk.Button(btn_row_filter2, text="删除最后一条", command=self._del_condition2, width=10)
        self.btn_del_cond2.pack(side="left", padx=4)

        # 导出按钮行：三个按钮并排 + 导出全部列复选框
        btn_export_row = ttk.Frame(filter_frame)
        btn_export_row.pack(fill="x", padx=8, pady=(6, 8))

        self.btn_inner_join = ttk.Button(btn_export_row, text="执行内连接导出", command=self.run_inner_join_export, state=tk.DISABLED)
        self.btn_inner_join.pack(side="left", fill="x", expand=True, padx=(0, 3))

        self.btn_separate = ttk.Button(btn_export_row, text="分别导出结果", command=self.run_separate_export, state=tk.DISABLED)
        self.btn_separate.pack(side="left", fill="x", expand=True, padx=3)

        self.btn_left_join = ttk.Button(btn_export_row, text="左连接Sheet1导出", command=self.run_left_join_export, state=tk.DISABLED)
        self.btn_left_join.pack(side="left", fill="x", expand=True, padx=(3, 10))

        self.var_export_all_cols = tk.BooleanVar(value=False)
        self.chk_export_all_cols = ttk.Checkbutton(btn_export_row, text="导出全部列", variable=self.var_export_all_cols)
        self.chk_export_all_cols.pack(side="right", padx=(0, 2))

        # 右侧区域：预览区 + 功能说明 + 日志区
        right_area = ttk.Frame(top_container)
        right_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 预览区
        ttk.Label(right_area, text="数据预览（前100行）", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)
        preview_frame = ttk.Frame(right_area)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.preview_tree = ttk.Treeview(preview_frame, show="tree headings")
        self.preview_tree.heading("#0", text="序号")
        self.preview_tree.column("#0", width=60, anchor="center", stretch=False)

        scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ========== 功能说明区 ==========
        ttk.Label(right_area, text="功能说明", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 3), fill=tk.X)
        desc_frame = tk.Frame(right_area, bg="#F5F5F5", bd=1, relief="solid", height=190)
        desc_frame.pack(fill=tk.X, pady=(0, 8))
        desc_frame.pack_propagate(False)
        
        tk.Label(
            desc_frame, 
            text="数据模式纯数据处理速度快；格式模式保留字体、颜色、布局、合并单元格\n等格式，速度较慢，适合需要保留原表格样式的导出场景",
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#b45309", bg="#F5F5F5", 
            anchor="w", justify="left", 
            wraplength=520
        ).pack(fill=tk.X, padx=8, pady=(4, 2))

        desc_style = ttk.Style()
        desc_style.configure("Desc.TNotebook.Tab", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True), padding=[21, 5])
        
        desc_notebook = ttk.Notebook(desc_frame, style="Desc.TNotebook")
        desc_notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 4))

        # 标签页1：基础操作
        tab_basic = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_basic, text="  基础操作  ")
        basic_text = """• Sheet列表：按住Ctrl可多选Sheet，按住Shift可连续选中多行
• 拆分全部/拆分选中：将每个Sheet独立导出为单独的Excel文件
• 同列拼接：仅支持标准二维表，按首行列名纵向合并，同名列自动对齐
• 高级筛选：仅支持标准二维表，含合并单元格的Sheet不参与筛选运算"""
        tk.Label(
            tab_basic, text=basic_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 标签页2：筛选逻辑
        tab_logic = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_logic, text="  筛选逻辑  ")
        logic_text = """• 组内规则：同一组内所有条件为AND关系，必须同时满足
• 组间规则：两个条件组之间为OR关系，满足任意一整组即命中
• 数量限制：最多支持2个AND组，不需要OR逻辑请勿点击添加OR组
• 主键列名：多表连接的匹配依据列，留空默认取第一列"""
        tk.Label(
            tab_logic, text=logic_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 标签页3：算子格式
        tab_operator = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_operator, text="  算子格式  ")
        op_text = """• 介于：用波浪号分隔，例：20~100
• 属于/不属于：多个值用英文逗号分隔，例：茂南,电白,高州
• 为空/非空：无需输入筛选值，直接添加条件即可
• 包含/不包含：输入子串内容，匹配字段中包含该内容的行"""
        tk.Label(
            tab_operator, text=op_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 标签页4：导出模式
        tab_export = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_export, text="  导出模式  ")
        export_text = """• 内连接导出：仅保留所有Sheet主键列都匹配成功的行
• 分别导出结果：每个Sheet独立筛选，分Sheet页存入同一文件
• 左连接Sheet1：以第一个Sheet为基准，未匹配到的列补“-”
• 导出全部列：勾选后导出所有列，不勾选仅导出条件列+主键列"""
        tk.Label(
            tab_export, text=export_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 日志区
        ttk.Label(right_area, text="处理日志", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)
        log_frame = tk.Frame(right_area, bg="#F0F0F0")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.logger = LogComponent(log_frame, AppConfig.LOG_MATCH_LIMIT, self)

        # 禁用控件列表
        self.disabled_widgets = [
            self.btn_load_file, self.chk_header,
            self.btn_split_all, self.btn_split_selected,
            self.btn_inner_join, self.btn_separate, self.btn_left_join,
            self.btn_concat
        ]
        # 选中Sheet自动切换预览
        self.sheet_listbox.bind("<<ListboxSelect>>", self._on_sheet_select)

        # 初始化筛选数据
        self._init_filter_data()

    # ========== 模式切换联动 ==========
    def _on_mode_change(self):
        is_hifi = self.high_fidelity_var.get()
        # 高保真模式禁用表头选项
        self.chk_header.state(["disabled"] if is_hifi else ["!disabled"])
        
        # 刷新Sheet列表状态
        if self.business and self.business.file_path:
            self._refresh_sheet_list()
            
            # 切到数据模式：自动取消所有复杂Sheet的选中
            if not is_hifi:
                complex_set = set(self.business.complex_sheet_reasons.keys())
                selected = list(self.sheet_listbox.curselection())
                cleared_count = 0
                # 倒序取消，避免索引错乱
                for idx in reversed(selected):
                    sheet_name = self.sheet_listbox.get(idx)
                    if sheet_name in complex_set:
                        self.sheet_listbox.selection_clear(idx)
                        cleared_count += 1
                if cleared_count > 0:
                    self.logger.append(f"已切换为数据模式，自动取消 {cleared_count} 个复杂格式Sheet的选中")
                    self.logger.append("提示：含合并单元格的Sheet仅支持保格式拆分，不支持拼接与筛选")
            else:
                self.logger.append("已切换为格式模式，所有非空Sheet均可参与拆分与合并")

    # ========== 高级筛选交互方法 ==========
    def _init_filter_data(self):
        """初始化条件数据"""
        self.filter_conds1 = []
        self.filter_conds2 = []
        self.group2_enabled = False

        self._refresh_filter_display(self.text_filter1, self.filter_conds1)
        self._refresh_filter_display(self.text_filter2, self.filter_conds2)

    def _refresh_col_combos(self):
        """刷新所有列名下拉框，仅基于数据模式兼容的Sheet生成"""
        if not self.business or not self.business.data_compatible_sheets:
            return

        cols = self.business.get_filter_column_options()

        for combo in [self.combo_col1, self.combo_col2]:
            combo["values"] = cols
            if cols:
                combo.current(0)
                combo.config(state="readonly")
            else:
                combo.config(state="disabled")

    def _on_op_change1(self, event):
        self._update_val_entry_state(self.combo_op1.get(), self.entry_val1_1, self.entry_val1_2)

    def _on_op_change2(self, event):
        self._update_val_entry_state(self.combo_op2.get(), self.entry_val2_1, self.entry_val2_2)

    def _update_val_entry_state(self, op_text, entry1, entry2):
        """根据算子动态调整值输入框状态"""
        zero_val_ops = {"为空", "非空"}

        if op_text in zero_val_ops:
            entry1.delete(0, tk.END)
            entry1.config(state=tk.DISABLED)
            entry2.pack_forget()
        else:
            entry1.config(state=tk.NORMAL)
            entry2.pack_forget()
            if entry1.get() in ("最小值", "最大值"):
                entry1.delete(0, tk.END)

    def _refresh_filter_display(self, text_widget: tk.Text, cond_list: list[dict]):
        """刷新横向条件展示"""
        text_widget.config(state="normal")
        text_widget.delete("1.0", tk.END)
        if cond_list:
            display_text = "；".join([self._cond_to_display(c) for c in cond_list])
            text_widget.insert("1.0", display_text)
        text_widget.config(state="disabled")

    def _cond_to_display(self, cond: dict) -> str:
        """条件转可读文本"""
        op = cond["op"]
        if op in ("为空", "非空"):
            return f"{cond['col']} {op}"
        elif op == "介于":
            return f"{cond['col']} 介于 {cond['val']} ~ {cond['val2']}"
        elif op in ("属于", "不属于"):
            return f"{cond['col']} {op} [{cond['val']}]"
        else:
            return f"{cond['col']} {op} {cond['val']}"

    def _add_condition1(self):
        self._add_condition(
            self.combo_col1, self.combo_op1,
            self.entry_val1_1, self.entry_val1_2,
            self.filter_conds1, self.text_filter1
        )

    def _add_condition2(self):
        self._add_condition(
            self.combo_col2, self.combo_op2,
            self.entry_val2_1, self.entry_val2_2,
            self.filter_conds2, self.text_filter2
        )

    def _add_condition(self, col_combo, op_combo, entry1, entry2, cond_list, text_widget):
        col = col_combo.get().strip()
        op = op_combo.get().strip()
        val1 = entry1.get().strip()

        if not col:
            messagebox.showwarning("提示", "请选择列名")
            return
        if op not in ("为空", "非空") and not val1:
            messagebox.showwarning("提示", "请输入筛选值")
            return

        val2 = ""
        if op == "介于":
            if "~" not in val1:
                messagebox.showwarning("提示", "介于格式为 最小值~最大值，例如 5~8")
                return
            parts = val1.split("~")
            if len(parts) != 2 or not parts[0].strip() or not parts[1].strip():
                messagebox.showwarning("提示", "介于格式错误，请输入 最小值~最大值")
                return
            val1 = parts[0].strip()
            val2 = parts[1].strip()

        if val1 in ("最小值", "最大值"):
            val1 = ""
        if val2 in ("最小值", "最大值"):
            val2 = ""

        if len(cond_list) >= 10:
            messagebox.showwarning("提示", "每组最多添加10个筛选条件")
            return

        cond = {"col": col, "op": op, "val": val1, "val2": val2}
        cond_list.append(cond)
        self._refresh_filter_display(text_widget, cond_list)

        if op not in ("为空", "非空"):
            entry1.delete(0, tk.END)

    def _del_condition1(self):
        self._del_condition(self.filter_conds1, self.text_filter1)

    def _del_condition2(self):
        self._del_condition(self.filter_conds2, self.text_filter2)

    def _del_condition(self, cond_list, text_widget):
        if not cond_list:
            messagebox.showwarning("提示", "暂无已添加的条件")
            return
        cond_list.pop()
        self._refresh_filter_display(text_widget, cond_list)

    def _enable_group2(self):
        """启用OR条件组2"""
        if self.group2_enabled:
            return
        self.group2_frame.pack(fill="x", before=self.btn_inner_join.master)
        self.group2_enabled = True
        self.btn_add_or_group.config(state=tk.DISABLED)
        self._refresh_col_combos()

    def _disable_group2(self):
        self.group2_frame.pack_forget()
        self.filter_conds2.clear()
        self._refresh_filter_display(self.text_filter2, self.filter_conds2)
        self.group2_enabled = False
        self.btn_add_or_group.config(state=tk.NORMAL)

    def _get_key_col(self) -> str:
        """获取主键列名，用户输入为空则取默认第一列"""
        key_col = self.entry_key_col.get().strip()
        if key_col:
            return key_col
        if self.business and self.business.data_compatible_sheets:
            first_sheet = self.business.data_compatible_sheets[0]
            first_cols = ExcelFileTools.get_sheet_column_names(
                self.business.file_path, first_sheet, header=True
            )
            if first_cols:
                return first_cols[0]
        return ""

    def run_inner_join_export(self):
        """执行内连接筛选导出"""
        groups = []
        if self.filter_conds1:
            groups.append(self.filter_conds1)
        if self.group2_enabled and self.filter_conds2:
            groups.append(self.filter_conds2)

        if not groups:
            messagebox.showwarning("提示", "请至少添加一个筛选条件")
            return

        key_col = self._get_key_col()
        if not key_col:
            messagebox.showwarning("提示", "无法确定主键列名")
            return

        save_path = os.path.join(
            self.business._get_output_dir(),
            f"{self.business._get_base_name()}_内连接筛选结果.xlsx"
        )
        export_all = self.var_export_all_cols.get()

        def _task():
            col_count = self.business.filter_inner_join_export(groups, save_path, key_col, export_all)
            return f"内连接筛选完成！共 {col_count} 列\n保存至：{save_path}"

        self.run_background_task(_task, "正在执行内连接筛选导出...")

    def run_separate_export(self):
        """执行分Sheet筛选导出"""
        groups = []
        if self.filter_conds1:
            groups.append(self.filter_conds1)
        if self.group2_enabled and self.filter_conds2:
            groups.append(self.filter_conds2)

        if not groups:
            messagebox.showwarning("提示", "请至少添加一个筛选条件")
            return

        key_col = self._get_key_col()
        if not key_col:
            messagebox.showwarning("提示", "无法确定主键列名")
            return

        save_path = os.path.join(
            self.business._get_output_dir(),
            f"{self.business._get_base_name()}_分Sheet筛选结果.xlsx"
        )
        export_all = self.var_export_all_cols.get()

        def _task():
            sheet_count = self.business.filter_separate_export(groups, save_path, key_col, export_all)
            return f"分Sheet筛选完成！共 {sheet_count} 个Sheet\n保存至：{save_path}"

        self.run_background_task(_task, "正在执行分Sheet筛选导出...")

    def run_left_join_export(self):
        """执行左连接Sheet1筛选导出"""
        groups = []
        if self.filter_conds1:
            groups.append(self.filter_conds1)
        if self.group2_enabled and self.filter_conds2:
            groups.append(self.filter_conds2)

        if not groups:
            messagebox.showwarning("提示", "请至少添加一个筛选条件")
            return

        key_col = self._get_key_col()
        if not key_col:
            messagebox.showwarning("提示", "无法确定主键列名")
            return

        save_path = os.path.join(
            self.business._get_output_dir(),
            f"{self.business._get_base_name()}_左连接筛选结果.xlsx"
        )
        export_all = self.var_export_all_cols.get()

        def _task():
            col_count = self.business.filter_left_join_export(groups, save_path, key_col, export_all)
            return f"左连接筛选完成！共 {col_count} 列\n保存至：{save_path}"

        self.run_background_task(_task, "正在执行左连接筛选导出...")

    # ========== 通用交互方法 ==========
    def _on_header_change(self):
        if self.header_var.get():
            self.label_auto_col.pack_forget()
        else:
            self.label_auto_col.pack(side=tk.TOP, anchor="w")

        if self.business:
            self.business.export_first_row_as_header = self.header_var.get()

    def _on_sheet_select(self, event):
        # 数据模式下：自动取消复杂Sheet的选中
        if self.high_fidelity_var.get() == False and self.business:
            complex_set = set(self.business.complex_sheet_reasons.keys())
            selected = list(self.sheet_listbox.curselection())
            need_clear = []
            for idx in selected:
                sheet_name = self.sheet_listbox.get(idx)
                if sheet_name in complex_set:
                    need_clear.append(idx)
            # 倒序取消
            for idx in reversed(need_clear):
                self.sheet_listbox.selection_clear(idx)

        # 正常切换预览
        selected = self._get_selected_sheets()
        if selected and self.business:
            self._preview_sheet(selected[0])

    def _refresh_sheet_list(self):
        """刷新Sheet列表：全量显示所有非空Sheet，数据模式下复杂Sheet置灰"""
        if not self.business or not self.business.file_path:
            self.sheet_listbox.delete(0, tk.END)
            return

        self.sheet_listbox.delete(0, tk.END)
        is_data_mode = (self.high_fidelity_var.get() == False)
        complex_set = set(self.business.complex_sheet_reasons.keys())

        for idx, sheet_name in enumerate(self.business.all_non_empty_sheets):
            self.sheet_listbox.insert(tk.END, sheet_name)
            if is_data_mode and sheet_name in complex_set:
                # 数据模式下复杂格式Sheet置灰
                self.sheet_listbox.itemconfig(idx, fg="#999999")

    def _preview_sheet(self, sheet_name):
        # 清空旧预览
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = []
        
        if not self.business or not sheet_name:
            return
        
        try:
            conn = duckdb.connect(":memory:")
            DuckExcelEngine._ensure_excel_extension(conn)
            sheet_param = f", sheet = '{sheet_name}'"
            
            rel = conn.sql(f"""
                SELECT *
                FROM read_xlsx('{self.business.file_path.replace(chr(39), chr(39)*2)}' {sheet_param}, header = false, all_varchar = true)
                LIMIT 100
            """)
            raw_cols = [c[0] for c in rel.description]
            rows = rel.fetchall()
            conn.close()

            # 生成自动列名
            sheet_idx = self.business.sheet_index_map.get(sheet_name, 0)
            display_cols = ExcelFileTools.generate_auto_column_names(sheet_idx, len(raw_cols))
            
            self.preview_tree["columns"] = display_cols
            for col in display_cols:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=100, anchor="w", stretch=False)
            
            for idx, row in enumerate(rows, 1):
                self.preview_tree.insert("", tk.END, text=str(idx), values=[str(v) if v is not None else "" for v in row])
        except Exception:
            pass

    def load_excel_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx;*.xlsm"), ("所有文件", "*.*")]
        )
        if not path:
            return

        self.logger.append(f"已选择文件：{os.path.basename(path)}")

        def _task():
            self.business = WpsExcelBusiness(self.logger.append)
            self.business.export_first_row_as_header = self.header_var.get()
            info = self.business.load_file(path)

            def update_ui():
                total = len(info["all_non_empty_sheets"])
                data_ok = len(info["valid_sheets"])
                complex_cnt = total - data_ok
                self.label_file_info.config(
                    text=f"共 {total} 个非空Sheet\n数据模式可用 {data_ok} 个",
                    foreground="#166534"
                )

                # 清空历史筛选条件
                self._init_filter_data()
                # 刷新列名下拉框
                self._refresh_col_combos()
                # 启用所有操作按钮
                for btn in [self.btn_split_all, self.btn_split_selected, self.btn_inner_join, self.btn_separate, self.btn_left_join, self.btn_concat]:
                    btn.config(state=tk.NORMAL)

                # 自动填充默认主键列名
                if info["valid_sheets"]:
                    first_sheet = info["valid_sheets"][0]
                    first_cols = ExcelFileTools.get_sheet_column_names(
                        self.business.file_path, first_sheet, header=True
                    )
                    if first_cols:
                        self.entry_key_col.delete(0, tk.END)
                        self.entry_key_col.insert(0, first_cols[0])

                # 填充Sheet列表
                self._refresh_sheet_list()

                # 自动预览第一个Sheet
                if info["all_non_empty_sheets"]:
                    self._preview_sheet(info["all_non_empty_sheets"][0])

            self.safe_ui_update(update_ui)
            return "文件加载完成"

        self.run_background_task(_task, "正在解析Excel文件...")

    def _get_selected_sheets(self) -> list[str]:
        indices = self.sheet_listbox.curselection()
        return [self.sheet_listbox.get(i) for i in indices]

    def run_split_all(self):
        def _task():
            count = self.business.split_all_sheets(high_fidelity=self.high_fidelity_var.get())
            return f"拆分完成！共生成 {count} 个文件\n保存至：{self.business._get_output_dir()}"
        self.run_background_task(_task, "开始拆分全部Sheet...")

    def run_split_selected(self):
        selected = self._get_selected_sheets()
        if not selected:
            messagebox.showwarning("提示", "请先在列表中选择要拆分的Sheet")
            return

        def _task():
            count = self.business.split_selected_sheets(selected, high_fidelity=self.high_fidelity_var.get())
            return f"拆分完成！共生成 {count} 个文件\n保存至：{self.business._get_output_dir()}"
        self.run_background_task(_task, f"开始拆分选中的 {len(selected)} 个Sheet...")

    def run_concat(self):
        selected = self._get_selected_sheets()
        # 有选中则拼接选中，无选中则拼接全部数据兼容Sheet
        target_sheets = selected if selected else self.business.data_compatible_sheets
        
        # 业务层会强制过滤，这里提前提示
        complex_in_selected = [s for s in target_sheets if s in self.business.complex_sheet_reasons]
        if complex_in_selected:
            self.logger.append(f"纵向拼接仅支持标准二维表，自动跳过 {len(complex_in_selected)} 个含复杂格式的Sheet")

        if len(target_sheets) - len(complex_in_selected) < 2:
            messagebox.showwarning("提示", "至少需要2个标准二维表Sheet才能拼接\n含合并单元格的Sheet不支持列对齐拼接")
            return

        save_path = os.path.join(
            self.business._get_output_dir(),
            f"{self.business._get_base_name()}_Sheet拼接结果.xlsx"
        )

        def _task():
            success, skipped = self.business.concat_same_column_sheets(save_path, target_sheets)
            msg = f"拼接完成！成功 {success} 个Sheet"
            if skipped > 0:
                msg += f"，跳过 {skipped} 个（格式不兼容/列数不一致）"
            msg += f"\n保存至：{save_path}"
            return msg
        self.run_background_task(_task, f"开始拼接 {len(target_sheets)} 个Sheet...")

class MultiSheetTab(BaseTab):
    """多表处理标签页：多文件管理、批量加载、后续多表功能扩展入口"""
    def __init__(self, parent):
        super().__init__(parent)
        self.business: Optional[MultiSheetBusiness] = None
        self.high_fidelity_var = tk.BooleanVar(value=False)  # 默认快速模式

        # ========== Sheet名称筛选相关变量（必须放在_init_ui之前）==========
        self.sheet_filter_keyword = tk.StringVar(value="")
        self.filter_mode = tk.StringVar(value="contains")  # 默认选中包含模式

        self._init_ui()

    def _init_ui(self):
        top_container = ttk.Frame(self)
        top_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 左侧操作区（宽度与单表完全对齐）
        left_area = ttk.Frame(top_container, width=450)
        left_area.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_area.pack_propagate(False)
        self._init_left_panel(left_area)

        # 右侧区域：预览 + 功能说明 + 日志
        right_area = ttk.Frame(top_container)
        right_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self._init_right_panel(right_area)

        # 禁用控件列表
        self.disabled_widgets = [
            self.btn_add_file, self.btn_remove_selected, self.btn_clear_all,
            self.chk_header,
            self.btn_confirm_load,
            self.btn_merge_all, self.btn_merge_selected
        ]

    def _init_left_panel(self, parent):
        # 1. 文件选择区
        ttk.Label(parent, text="1. 选择表格文件", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(0, 4), fill=tk.X, padx=5)
        
        btn_row1 = ttk.Frame(parent)
        btn_row1.pack(fill=tk.X, padx=5, pady=2)
        self.btn_add_file = ttk.Button(btn_row1, text="添加Excel文件", command=self._on_add_files, width=14)
        self.btn_add_file.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_remove_selected = ttk.Button(btn_row1, text="移除选中", command=self._on_remove_selected)
        self.btn_remove_selected.pack(side=tk.LEFT, padx=4)
        self.btn_clear_all = ttk.Button(btn_row1, text="清空全部", command=self._on_clear_all)
        self.btn_clear_all.pack(side=tk.LEFT, padx=4)

        self.label_file_count = ttk.Label(parent, text="已选择 0 个文件", foreground="#666666", anchor="w")
        self.label_file_count.pack(pady=2, anchor=tk.W, padx=5)

        # 2. 已选文件列表
        ttk.Label(parent, text="已选文件列表", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(10, 4), fill=tk.X, padx=5)
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.X, padx=5, pady=2)
        self.file_listbox = tk.Listbox(
            list_frame, selectmode="extended", height=8,
            exportselection=False, bg="#F8F8F8"
        )
        scroll_file = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=scroll_file.set)
        scroll_file.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        # 选中文件切换预览
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        # 3. 全局参数区
        ttk.Label(parent, text="全局参数（对所有文件生效）", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(10, 4), fill=tk.X, padx=5)
        param_frame = ttk.Frame(parent)
        param_frame.pack(fill=tk.X, padx=5, pady=2)
        # 处理模式：上下排列
        mode_frame = ttk.Frame(param_frame)
        mode_frame.pack(side=tk.LEFT, padx=(0, 20))
        self.radio_fast = ttk.Radiobutton(
            mode_frame, text="数据模式（速度快）", 
            variable=self.high_fidelity_var, value=False, 
            command=self._on_mode_change
        )
        self.radio_fast.pack(side=tk.TOP, anchor="w")
        self.radio_hifi = ttk.Radiobutton(
            mode_frame, text="格式模式（速度慢，保留格式）", 
            variable=self.high_fidelity_var, value=True, 
            command=self._on_mode_change
        )
        self.radio_hifi.pack(side=tk.TOP, anchor="w")

        # 首行作为表头：放在右侧
        self.header_var = tk.BooleanVar(value=True)
        self.chk_header = ttk.Checkbutton(
            param_frame, text="首行作为表头导出",
            variable=self.header_var, command=self._on_header_change
        )
        self.chk_header.pack(side=tk.LEFT, padx=(10, 0))

        # 4. 确认加载按钮
        self.btn_confirm_load = ttk.Button(
            parent, text="确认加载文件", command=self._on_confirm_load,
            state=tk.DISABLED
        )
        self.btn_confirm_load.pack(pady=(12, 8), fill=tk.X, padx=5)

        # 5. 功能操作区（已改造：左按钮 + 右筛选）
        ttk.Label(parent, text="多表操作功能（留空默认不筛选Sheet名）", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(10, 4), fill=tk.X, padx=5)
        func_frame = ttk.Frame(parent)
        func_frame.pack(fill=tk.X, padx=5, pady=2)

        # --- 左侧：功能按钮区（上下排列，固定宽度，缩短长度）---
        btn_frame = ttk.Frame(func_frame)
        btn_frame.pack(side=tk.LEFT, padx=(0, 16))

        self.btn_merge_all = ttk.Button(btn_frame, text="合并所有表", width=12, command=self._on_merge_all, state=tk.DISABLED)
        self.btn_merge_all.pack(side=tk.TOP, pady=2)
        self.btn_merge_selected = ttk.Button(btn_frame, text="合并选中表", width=12, command=self._on_merge_selected, state=tk.DISABLED)
        self.btn_merge_selected.pack(side=tk.TOP, pady=2)

        # --- 右侧：Sheet名称筛选区 ---
        filter_frame = ttk.Frame(func_frame)
        filter_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 第一行：关键词输入区
        input_row = ttk.Frame(filter_frame)
        input_row.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        ttk.Label(input_row, text="Sheet名称：").pack(side=tk.LEFT)
        ttk.Entry(input_row, textvariable=self.sheet_filter_keyword, width=24).pack(side=tk.LEFT, padx=6)

        # 第二、三行：匹配模式单选（2行3列，共6种，严格区分大小写）
        mode_row = ttk.Frame(filter_frame)
        mode_row.pack(side=tk.TOP, anchor="w")

        mode_list = [
            ("等于", "equal"),
            ("不等于", "not_equal"),
            ("包含", "contains"),
            ("不包含", "not_contains"),
            ("开头是", "start_with"),
            ("结尾是", "end_with")
        ]
        for idx, (text, mode_val) in enumerate(mode_list):
            row = idx // 3
            col = idx % 3
            ttk.Radiobutton(
                mode_row, text=text, variable=self.filter_mode, value=mode_val
            ).grid(row=row, column=col, sticky="w", padx=8, pady=1)

    def _init_right_panel(self, parent):
        # 预览区（与单表样式完全一致）
        ttk.Label(parent, text="数据预览（前100行）", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)
        preview_frame = ttk.Frame(parent)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.preview_tree = ttk.Treeview(preview_frame, show="tree headings")
        self.preview_tree.heading("#0", text="序号")
        self.preview_tree.column("#0", width=60, anchor="center", stretch=False)

        scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 功能说明区
        ttk.Label(parent, text="功能说明", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 3), fill=tk.X)
        desc_frame = tk.Frame(parent, bg="#F5F5F5", bd=1, relief="solid", height=200)
        desc_frame.pack(fill=tk.X, pady=(0, 8))
        desc_frame.pack_propagate(False)

        tk.Label(
            desc_frame, 
            text="快速模式纯数据处理速度快；高保真模式保留字体、颜色、布局、合并单元格\n等格式，速度较慢，适合需要保留原表格样式的导出场景",
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#b45309", bg="#F5F5F5", 
            anchor="w", justify="left", 
            wraplength=520
        ).pack(fill=tk.X, padx=8, pady=(4, 2))

        desc_style = ttk.Style()
        desc_style.configure("MultiDesc.TNotebook.Tab", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True), padding=[21, 5])
        desc_notebook = ttk.Notebook(desc_frame, style="MultiDesc.TNotebook")
        desc_notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 4))

        # 标签页1：基础操作
        tab_basic = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_basic, text="  基础操作  ")
        basic_text = """• 添加文件：支持Ctrl/Shift多选，可多次追加，自动去重
• 移除/清空：可选中删除部分文件，或一键清空列表
• 确认加载：选完文件后点击，批量校验所有文件有效性
• 加载完成后自动启用下方功能按钮，支持后续多表操作"""
        tk.Label(
            tab_basic, text=basic_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 标签页2：纵向拼接
        tab_concat = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_concat, text="  纵向拼接  ")
        concat_text = """• 功能：待开发"""
        tk.Label(
            tab_concat, text=concat_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 标签页3：横向关联
        tab_join = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_join, text="  横向关联  ")
        join_text = """• 功能：待开发"""
        tk.Label(
            tab_join, text=join_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 标签页4：差异对比
        tab_compare = ttk.Frame(desc_notebook)
        desc_notebook.add(tab_compare, text="  差异对比  ")
        compare_text = """• 功能：待开发"""
        tk.Label(
            tab_compare, text=compare_text,
            font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD),
            fg="#333333", bg="#ffffff",
            anchor="nw", justify="left",
            wraplength=500, padx=8, pady=6
        ).pack(fill=tk.BOTH, expand=True)

        # 日志区（完全独立，与单表互不干扰）
        ttk.Label(parent, text="处理日志", font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)
        log_frame = tk.Frame(parent, bg="#F0F0F0")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.logger = LogComponent(log_frame, AppConfig.LOG_MATCH_LIMIT, self)

    # ========== 模式切换联动 ==========
    def _on_mode_change(self):
        is_hifi = self.high_fidelity_var.get()
        # 高保真模式禁用表头选项
        self.chk_header.state(["disabled"] if is_hifi else ["!disabled"])

    # ========== 内部工具：Sheet名称筛选（全边界兜底，异常走日志）==========
    def _filter_sheet_names(self, original_sheets: list) -> list:
        """
        根据关键词和匹配模式过滤Sheet名称列表
        全程严格区分大小写，空关键词直接返回原列表
        所有异常均输出日志并兜底返回原列表，绝不中断主流程
        """
        try:
            keyword = self.sheet_filter_keyword.get().strip()
            mode = self.filter_mode.get()

            # 边界1：空关键词 → 不筛选，直接返回
            if not keyword:
                return original_sheets

            # 边界2：原列表为空 → 直接返回
            if not original_sheets:
                return []

            result = []
            for sheet_name in original_sheets:
                name_str = str(sheet_name)

                # 6种匹配模式
                if mode == "equal":
                    match = (name_str == keyword)
                elif mode == "not_equal":
                    match = (name_str != keyword)
                elif mode == "contains":
                    match = (keyword in name_str)
                elif mode == "not_contains":
                    match = (keyword not in name_str)
                elif mode == "start_with":
                    match = name_str.startswith(keyword)
                elif mode == "end_with":
                    match = name_str.endswith(keyword)
                else:
                    # 未知模式 → 兜底全匹配，输出日志
                    self.logger.append(f"未知筛选模式[{mode}]，已跳过Sheet名称筛选")
                    return original_sheets

                if match:
                    result.append(sheet_name)

            # 筛选后为空输出日志，由上层判断是否继续
            if not result:
                self.logger.append(f"Sheet名称筛选后无匹配项（关键词：{keyword}，模式：{mode}）")
            
            return result

        except Exception as e:
            self.logger.append(f"Sheet名称筛选异常，已自动跳过筛选，错误：{str(e)}")
            return original_sheets

    # ========== 交互事件 ==========
    def _on_add_files(self):
        """添加文件：支持多选，自动追加去重"""
        paths = filedialog.askopenfilenames(
            title="选择Excel文件（可多选）",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm"), ("所有文件", "*.*")]
        )
        if not paths:
            return

        try:
            if not self.business:
                self.business = MultiSheetBusiness(self.logger.append)
            added, total = self.business.add_files(paths)
        except RuntimeError as e:
            messagebox.showwarning("提示", str(e))
            return

        self._refresh_file_list()
        self.label_file_count.config(text=f"已选择 {total} 个文件")
        self.btn_confirm_load.config(state=tk.NORMAL if total > 0 else tk.DISABLED)
        self.logger.append(f"新增 {added} 个文件，当前共 {total} 个")

    def _on_remove_selected(self):
        """移除选中文件"""
        indices = list(self.file_listbox.curselection())
        if not indices:
            messagebox.showwarning("提示", "请先选择要移除的文件")
            return
        remain = self.business.remove_files(indices)
        self._refresh_file_list()
        self.label_file_count.config(text=f"已选择 {remain} 个文件")
        self.btn_confirm_load.config(state=tk.NORMAL if remain > 0 else tk.DISABLED)
        self.logger.append(f"移除 {len(indices)} 个文件，剩余 {remain} 个")

    def _on_clear_all(self):
        """清空所有文件"""
        if not self.business or not self.business.file_paths:
            return
        if not messagebox.askokcancel("确认", "确定要清空所有已选文件吗？"):
            return
        self.business.clear_all()
        self._refresh_file_list()
        self.label_file_count.config(text="已选择 0 个文件")
        self.btn_confirm_load.config(state=tk.DISABLED)
        # 清空预览
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = []
        self.logger.append("已清空所有文件")

    def _refresh_file_list(self):
        """刷新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        if not self.business:
            return
        for path in self.business.file_paths:
            self.file_listbox.insert(tk.END, os.path.basename(path))

    # ========== 通用交互方法 ==========
    def _on_header_change(self):
        """首行作为表头导出：仅影响最终导出，不影响预览和运算（与单表原则一致）"""
        if self.business:
            # 属性名和单表统一，避免混淆
            self.business.export_first_row_as_header = self.header_var.get()

    def _on_file_select(self, event):
        """选中文件切换预览"""
        if not self.business or not self.business.file_meta:
            return
        selected = self.file_listbox.curselection()
        if selected:
            self._preview_file(selected[0])

    def _preview_file(self, file_index: int):
        """预览指定索引文件的第一个有效Sheet"""
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = []

        file_path, sheet_name = self.business.get_preview_target(file_index)
        if not file_path or not sheet_name:
            return

        try:
            conn = duckdb.connect(":memory:")
            DuckExcelEngine._ensure_excel_extension(conn)
            
            rel = conn.sql(f"""
                SELECT *
                FROM read_xlsx('{file_path.replace(chr(39), chr(39)*2)}', sheet = '{sheet_name}', header = false, all_varchar = true)
                LIMIT 100
            """)
            raw_cols = [c[0] for c in rel.description]
            rows = rel.fetchall()
            conn.close()

            # 自动列名，与单表规则一致
            display_cols = ExcelFileTools.generate_auto_column_names(0, len(raw_cols))
            self.preview_tree["columns"] = display_cols
            for col in display_cols:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=100, anchor="w", stretch=False)
            
            for idx, row in enumerate(rows, 1):
                self.preview_tree.insert("", tk.END, text=str(idx), values=[str(v) if v is not None else "" for v in row])
        except Exception:
            pass

    def _on_merge_all(self):
        """合并所有表按钮事件"""
        if not self.business or not self.business.file_paths:
            messagebox.showwarning("提示", "没有可合并的文件，请先加载文件")
            return

        output_dir = self.business._get_output_dir()
        import datetime
        time_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        save_path = os.path.join(output_dir, f"多表合并_全部_{time_str}.xlsx")

        # 读取当前筛选条件
        filter_keyword = self.sheet_filter_keyword.get().strip()
        filter_mode = self.filter_mode.get()

        def _task():
            total_files = len(self.business.file_paths)
            file_cnt, sheet_cnt, skip_cnt = self.business.merge_all_files(
                save_path,
                sheet_filter_keyword=filter_keyword,
                sheet_filter_mode=filter_mode,
                high_fidelity=self.high_fidelity_var.get()
            )

            msg = f"合并完成！\n成功处理 {file_cnt} 个文件，共 {sheet_cnt} 个Sheet"
            skipped_files = total_files - file_cnt
            if skipped_files > 0 or skip_cnt > 0:
                msg += f"\n\n⚠️  跳过 {skipped_files} 个无效文件、{skip_cnt} 个无效Sheet\n详情请查看下方处理日志"
            msg += f"\n\n保存路径：{save_path}"
            return msg

        self.run_background_task(_task, "开始合并所有文件...")

    def _on_merge_selected(self):
        """合并选中表按钮事件"""
        selected = list(self.file_listbox.curselection())
        if not selected:
            messagebox.showwarning("提示", "请先在文件列表中选择要合并的文件")
            return

        output_dir = self.business._get_output_dir()
        import datetime
        time_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        save_path = os.path.join(output_dir, f"多表合并_选中_{time_str}.xlsx")

        # 读取当前筛选条件
        filter_keyword = self.sheet_filter_keyword.get().strip()
        filter_mode = self.filter_mode.get()

        def _task():
            total_selected = len(selected)
            file_cnt, sheet_cnt, skip_cnt = self.business.merge_selected_files(
                selected,
                save_path,
                sheet_filter_keyword=filter_keyword,
                sheet_filter_mode=filter_mode,
                high_fidelity=self.high_fidelity_var.get()
            )

            msg = f"合并完成！\n成功处理 {file_cnt} 个文件，共 {sheet_cnt} 个Sheet"
            skipped_files = total_selected - file_cnt
            if skipped_files > 0 or skip_cnt > 0:
                msg += f"\n\n⚠️  跳过 {skipped_files} 个无效文件、{skip_cnt} 个无效Sheet\n详情请查看下方处理日志"
            msg += f"\n\n保存路径：{save_path}"
            return msg

        self.run_background_task(_task, f"开始合并选中的 {len(selected)} 个文件...")

    def _on_confirm_load(self):
        """确认加载：后台线程批量处理"""
        if not self.business or not self.business.file_paths:
            messagebox.showwarning("提示", "请先添加至少一个文件")
            return

        self.business.export_first_row_as_header = self.header_var.get()

        def _task():
            stats = self.business.load_all_files()
            def update_ui():
                # 启用功能按钮
                for btn in [self.btn_merge_all, self.btn_merge_selected]:
                    btn.config(state=tk.NORMAL)
                # 默认预览第一个文件
                if stats["valid"] > 0:
                    self._preview_file(0)
                    self.file_listbox.selection_set(0)
            self.safe_ui_update(update_ui)
            return f"加载完成：有效 {stats['valid']} 个，失败 {stats['invalid']} 个，合计 {stats['total_sheets']} 个有效Sheet"

        self.run_background_task(_task, f"开始批量加载 {len(self.business.file_paths)} 个文件...")

class DataMatcherApp:
    def __init__(self, root):
        self.root = root
        self._setup_window()
        self._setup_global_style()
        self._build_notebook()
        # 绑定窗口关闭事件：点X时先通知任务停止，再销毁窗口
        self.root.protocol("WM_DELETE_WINDOW", self._on_safe_exit)

    def _setup_window(self):
        self.root.title(AppConfig.APP_TITLE)

        # ========== 跨平台固定窗口核心修复 ==========
        # 1. 先获取屏幕可用尺寸，避免小屏 Mac 窗口超出边界
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        # 预留菜单栏、Dock 空间，计算实际可用最大窗口尺寸
        max_w = min(AppConfig.MIN_SIZE[0], int(screen_w * 0.92))
        max_h = min(AppConfig.MIN_SIZE[1], int(screen_h * 0.85))

        # 2. 统一用 minsize = maxsize 锁定窗口尺寸，跨平台行为完全一致
        # 替代 resizable(False,False)，避免 Mac 下自动收缩的问题
        self.root.minsize(max_w, max_h)
        self.root.maxsize(max_w, max_h)

        # 3. 设置初始大小和居中显示
        self.root.geometry(f"{max_w}x{max_h}")
        # 窗口居中
        x = (screen_w - max_w) // 2
        y = (screen_h - max_h) // 2 - 20  # 向上微调，避开菜单栏
        self.root.geometry(f"{max_w}x{max_h}+{x}+{y}")

        # 4. 彻底移除 resizable(False,False)，避免 Mac 端副作用
        # 保留允许缩放的能力，小屏用户可自行调整；如需完全禁止可恢复，但不推荐
        # self.root.resizable(False, False)
    

    def _setup_global_style(self):
        style = ttk.Style()
        # 强制使用clam主题，全平台外观/尺寸100%一致，消除Mac/Windows布局差异
        style.theme_use("clam")
        
        # 全局基础样式
        style.configure(".", font=AppTools.get_font(AppConfig.FONT_SIZE_MAIN))
        style.configure("TButton", font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL))
        style.configure("TButton", focuscolor=style.lookup("TButton", "background"))
        
        # 统一控件内边距，优化clam主题视觉紧凑度
        style.configure("TButton", padding=(10, 4))
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)
        style.configure("TNotebook.Tab", padding=(15, 6))

    def _build_notebook(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.match_tab = MatchTab(notebook)
        self.process_tab = ProcessTab(notebook)
        # 新增加多表处理Tab
        self.multi_sheet_tab = MultiSheetTab(notebook)
        # 新增下面两行
        self.wps_tab = WpsTab(notebook)
        notebook.add(self.match_tab, text="文本匹配")
        notebook.add(self.process_tab, text="文本处理")
        notebook.add(self.wps_tab, text="单表格处理")
        notebook.add(self.multi_sheet_tab, text="多表格处理")
        # 实时模块：开关开启且加载成功才显示
        if ENABLE_REALTIME_MODULE:
            try:
                from realtime_module import create_realtime_tab
                self.realtime_tab = create_realtime_tab(notebook)
                notebook.add(self.realtime_tab, text="归属\黑名单查询")
            except Exception:
                # 任何加载异常都静默忽略，不影响主程序启动
                pass

    def has_active_tasks(self) -> bool:
        """判断是否有标签页正在运行后台任务"""
        return self.match_tab.is_loading or self.process_tab.is_loading

    def _on_safe_exit(self):
        if self.has_active_tasks():
            if not messagebox.askokcancel(
                "确认退出",
                "后台有任务正在处理中，退出会中断当前操作，结果文件不会生成。\n确定要强制退出吗？"
            ):
                return

        self.root.withdraw()
        try:
            cleanup_all_temp_files()
        except Exception:
            pass

        # 新增：强制终止当前进程及其所有子进程（终极杀残留）
        import os
        import signal
        # Windows系统直接杀进程树
        if os.name == 'nt':
            os.system(f"taskkill /F /T /PID {os.getpid()}")
        # 其他系统用标准信号
        else:
            os.killpg(os.getpgid(os.getpid()), signal.SIGKILL)


if __name__ == "__main__":
    # Windows高DPI适配，必须在创建根窗口前调用
    AppTools.setup_windows_high_dpi()
    root = tk.Tk()
    app = DataMatcherApp(root)
    root.mainloop()
