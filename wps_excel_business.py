import os
import duckdb
from typing import Callable, Optional ,Dict, List
from excel_tools import ExcelFileTools, DuckExcelEngine, FilterBuilder, get_high_fidelity_engine, ExcelStandardChecker, SheetCheckResult

class WpsExcelBusiness:
    """WPS表格处理业务层：拆分、筛选、拼接，统一清洗规则"""
    def __init__(self, log_callback: Callable[[str], None]):
        self.log = log_callback
        self.file_path: Optional[str] = None
        self.all_non_empty_sheets: list[str] = []  # 所有非空Sheet全集（格式模式可用）
        self.data_compatible_sheets: list[str] = []  # 数据模式兼容Sheet子集
        self.complex_sheet_reasons: dict[str, str] = {}  # 复杂Sheet不兼容原因
        self.all_sheets: list[str] = []  # 向后兼容：等价于data_compatible_sheets
        self.non_empty_sheets: list[str] = []  # 向后兼容：等价于all_non_empty_sheets
        self.merged_info: dict = {}
        self.sheet_index_map: dict[str, int] = {}  # sheet名 → 序号（0开始，有效sheet排序）
        self._force_header = False
        self.export_first_row_as_header = False

    def load_file(self, file_path: str) -> dict:
        """加载Excel文件，检测Sheet兼容性分级，返回结构化信息"""
        if not ExcelFileTools.is_excel_file(file_path):
            raise RuntimeError("文件格式不支持，仅支持 .xlsx / .xlsm")
        self.file_path = file_path
        self.log(f"正在解析文件：{os.path.basename(file_path)}")

        # 1. 获取所有非空Sheet
        self.all_non_empty_sheets = ExcelFileTools.get_non_empty_sheets(file_path)
        self.non_empty_sheets = self.all_non_empty_sheets  # 向后兼容
        if not self.all_non_empty_sheets:
            raise RuntimeError("文件中无有效数据Sheet")

        # 2. 统一检测数据模式兼容性
        self.log("正在检测Sheet格式兼容性...")
        compat_info = ExcelFileTools.check_sheet_data_compatibility(file_path)
        self.merged_info = {name: not info["compatible"] for name, info in compat_info.items()}

        # 3. 分级分类
        self.data_compatible_sheets = []
        self.complex_sheet_reasons = {}
        for name, info in compat_info.items():
            if name not in self.all_non_empty_sheets:
                continue
            if info["compatible"]:
                self.data_compatible_sheets.append(name)
            else:
                self.complex_sheet_reasons[name] = info["reason"]

        self.all_sheets = self.data_compatible_sheets  # 向后兼容

        # 4. 建立sheet序号映射（基于数据兼容Sheet，用于自动列名分配）
        self.sheet_index_map = {name: idx for idx, name in enumerate(self.data_compatible_sheets)}

        # 5. 输出明细日志
        total = len(self.all_non_empty_sheets)
        data_ok = len(self.data_compatible_sheets)
        complex_count = len(self.complex_sheet_reasons)
        self.log(f"加载完成：共 {total} 个非空Sheet，其中 {data_ok} 个支持数据模式，{complex_count} 个仅支持保格式模式")

        if self.complex_sheet_reasons:
            self.log("仅支持保格式模式的Sheet明细：")
            for name, reason in self.complex_sheet_reasons.items():
                self.log(f"  - {name}：{reason}")

        return {
            "valid_sheets": self.data_compatible_sheets,
            "all_non_empty_sheets": self.all_non_empty_sheets,
            "complex_sheets": list(self.complex_sheet_reasons.keys()),
            "merged_sheets": list(self.complex_sheet_reasons.keys()),
            "total_sheets": total
        }

    def _get_output_dir(self) -> str:
        """输出目录默认与源文件同目录"""
        return os.path.dirname(os.path.abspath(self.file_path))

    def _get_base_name(self) -> str:
        """获取不带后缀的主文件名"""
        return os.path.splitext(os.path.basename(self.file_path))[0]
    
    def _process_export_header(self, conn: duckdb.DuckDBPyConnection, rel: duckdb.DuckDBPyRelation, header_row: tuple = None) -> duckdb.DuckDBPyRelation:
        """
        根据导出选项处理表头
        :param header_row: 预先提取的表头行，筛选场景传入；为None则从结果第一行提取
        """
        if not self.export_first_row_as_header:
            return rel
        # 优先使用传入的表头行，没有则从结果第一行提取
        if header_row is not None:
            first_row = header_row
        else:
            first_row = rel.limit(1).fetchone()
        
        if first_row is None:
            return rel
        old_cols = rel.columns
        # 生成安全列名：空值用默认名，重名自动加后缀
        new_names = []
        name_counter = {}
        for idx, val in enumerate(first_row):
            raw_name = str(val).strip() if val is not None else ""
            if not raw_name:
                raw_name = f"列{idx+1}"
            # 处理重名
            if raw_name in name_counter:
                name_counter[raw_name] += 1
                raw_name = f"{raw_name}_{name_counter[raw_name]}"
            else:
                name_counter[raw_name] = 0
            new_names.append(raw_name)
        # 创建临时视图，重命名列并跳过数据首行
        rel.create_view("__export_header_tmp__")
        rename_parts = [
            f'"{old}" AS "{new}"'
            for old, new in zip(old_cols, new_names)
        ]
        col_str = ", ".join(rename_parts)
        # 传入表头行的场景，数据里已经不含表头行，不用OFFSET；否则OFFSET 1
        offset_sql = "OFFSET 1" if header_row is None else ""
        return conn.sql(f"""
            SELECT {col_str}
            FROM __export_header_tmp__
            {offset_sql}
        """)

    def _clean_and_export_sheet(
        self,
        conn: duckdb.DuckDBPyConnection,
        sheet_name: str,
        output_path: str
    ):
        """单个Sheet清洗+导出：去空格、去空行、去重、原子写入"""
        sheet_idx = self.sheet_index_map[sheet_name]
        raw_cols = ExcelFileTools.get_sheet_column_names(
            self.file_path, sheet_name, self._force_header
        )
        # 统一使用自动分配列名：a1、a2、a3...
        custom_cols = ExcelFileTools.generate_auto_column_names(sheet_idx, len(raw_cols))
        cols = DuckExcelEngine.read_clean_view(
            conn, self.file_path, sheet_name, self._force_header, "__tmp_clean__",
            custom_column_names=custom_cols
        )
        col_str = ", ".join([f'"{c}"' for c in cols])
        rel = conn.sql(f"SELECT {col_str} FROM __tmp_clean__")
        
        # 按选项处理导出表头
        final_rel = self._process_export_header(conn, rel)
        DuckExcelEngine.safe_to_excel(conn, final_rel, output_path, sheet_name=sheet_name, header=True)

    # ========== 功能1：拆分全部非空Sheet ==========
    def split_all_sheets(self, high_fidelity: bool = False) -> int:
        """全部Sheet拆分为独立文件，按模式自动选择可用范围"""
        if not self.file_path:
            raise RuntimeError("请先加载文件")
        output_dir = self._get_output_dir()
        base_name = self._get_base_name()

        if high_fidelity:
            target_sheets = self.all_non_empty_sheets
            self.log(f"启用保格式模式，全部 {len(target_sheets)} 个非空Sheet参与拆分")
            engine = get_high_fidelity_engine()
            count = 0
            for idx, sheet_name in enumerate(target_sheets, 1):
                out_filename = f"{base_name}-{sheet_name}.xlsx"
                engine.split_sheets(
                    self.file_path,
                    [sheet_name],
                    output_dir,
                    name_template=out_filename
                )
                count += 1
                if idx % max(1, len(target_sheets) // 5) == 0 or idx == len(target_sheets):
                    self.log(f"拆分进度：{idx}/{len(target_sheets)}")
            self.log(f"全部Sheet拆分完成，共生成 {count} 个文件")
            return count

        # 数据模式：仅使用兼容Sheet
        target_sheets = self.data_compatible_sheets
        skipped = len(self.all_non_empty_sheets) - len(target_sheets)
        self.log(f"数据模式，共 {len(target_sheets)} 个Sheet参与拆分，自动跳过 {skipped} 个复杂格式Sheet")
        if skipped > 0:
            self.log("跳过的Sheet：")
            for name in self.complex_sheet_reasons:
                self.log(f"  - {name}：{self.complex_sheet_reasons[name]}")

        conn = duckdb.connect(":memory:")
        count = 0
        try:
            for idx, sheet_name in enumerate(target_sheets, 1):
                out_file = os.path.join(output_dir, f"{base_name}-{sheet_name}.xlsx")
                self._clean_and_export_sheet(conn, sheet_name, out_file)
                count += 1
                if idx % max(1, len(target_sheets) // 5) == 0 or idx == len(target_sheets):
                    self.log(f"拆分进度：{idx}/{len(target_sheets)}")
            self.log(f"全部Sheet拆分完成，共生成 {count} 个文件")
            return count
        finally:
            conn.close()

    # ========== 功能2：拆分选中的Sheet ==========
    def split_selected_sheets(self, selected_sheets: list[str], high_fidelity: bool = False) -> int:
        """拆分用户选中的指定Sheet，按模式过滤有效项"""
        if not self.file_path:
            raise RuntimeError("请先加载文件")
        if not selected_sheets:
            raise ValueError("请至少选择一个Sheet")

        if high_fidelity:
            # 格式模式：校验是否存在，存在即可用
            invalid = [s for s in selected_sheets if s not in self.all_non_empty_sheets]
            if invalid:
                raise RuntimeError(f"以下Sheet不存在：{', '.join(invalid)}")
            target_sheets = selected_sheets
            self.log(f"启用保格式模式，选中的 {len(target_sheets)} 个Sheet参与拆分")
        else:
            # 数据模式：仅兼容Sheet可用，自动过滤
            target_sheets = [s for s in selected_sheets if s in self.data_compatible_sheets]
            skipped = [s for s in selected_sheets if s not in self.data_compatible_sheets]
            if not target_sheets:
                raise RuntimeError("选中的Sheet均不支持数据模式，请切换到保格式模式重试")
            self.log(f"数据模式，选中的 {len(selected_sheets)} 个Sheet中 {len(target_sheets)} 个有效，自动跳过 {len(skipped)} 个复杂格式Sheet")
            if skipped:
                self.log("跳过的Sheet：")
                for name in skipped:
                    self.log(f"  - {name}：{self.complex_sheet_reasons.get(name, '格式不兼容')}")

        output_dir = self._get_output_dir()
        base_name = self._get_base_name()

        if high_fidelity:
            engine = get_high_fidelity_engine()
            count = 0
            for sheet_name in target_sheets:
                out_filename = f"{base_name}-{sheet_name}.xlsx"
                engine.split_sheets(
                    self.file_path,
                    [sheet_name],
                    output_dir,
                    name_template=out_filename
                )
                count += 1
            self.log(f"选中Sheet拆分完成，共生成 {count} 个文件")
            return count

        # 数据模式
        conn = duckdb.connect(":memory:")
        count = 0
        try:
            for sheet_name in target_sheets:
                out_file = os.path.join(output_dir, f"{base_name}-{sheet_name}.xlsx")
                self._clean_and_export_sheet(conn, sheet_name, out_file)
                count += 1
            self.log(f"选中Sheet拆分完成，共生成 {count} 个文件")
            return count
        finally:
            conn.close()

    def get_filter_column_options(self) -> list[str]:
        """
        获取高级筛选下拉框的所有可选列名
        仅基于数据模式兼容的Sheet生成，复杂格式Sheet不参与筛选
        """
        if not self.data_compatible_sheets:
            return []
        
        total_col = 0
        for sheet_name in self.data_compatible_sheets:
            raw_cols = ExcelFileTools.get_sheet_column_names(
                self.file_path, sheet_name, self._force_header
            )
            total_col += len(raw_cols)
        
        return [f"a{i+1}" for i in range(total_col)]

    # ========== 功能3：高级筛选导出 ==========
    def _split_conditions_by_sheet(self, filter_groups: list[list[dict]]) -> dict[str, list[list[dict]]]:
        """
        条件拆分逻辑：仅针对数据兼容的Sheet
        1. UI的aN → 对应第一个Sheet的真实列名
        2. 每个Sheet检查自身是否存在该真实列名，存在则保留该条件，不存在则跳过
        3. 返回每个Sheet对应的条件组列表
        """
        if not self.data_compatible_sheets:
            return {}

        # 1. 建立UI列名 → 真实列名的映射（以第一个数据兼容Sheet为准）
        first_sheet = self.data_compatible_sheets[0]
        first_cols = ExcelFileTools.get_sheet_column_names(
            self.file_path, first_sheet, header=True
        )
        ui_to_real = {}
        for idx, real_col in enumerate(first_cols):
            ui_col = f"a{idx+1}".lower()
            ui_to_real[ui_col] = real_col

        # 2. 将所有UI条件转换为真实列名条件
        real_groups = []
        for group in filter_groups:
            real_group = []
            for cond in group:
                ui_col = cond["col"].strip().lower()
                if ui_col not in ui_to_real:
                    continue
                real_cond = cond.copy()
                real_cond["col"] = ui_to_real[ui_col]
                real_group.append(real_cond)
            if real_group:
                real_groups.append(real_group)

        # 3. 给每个数据兼容Sheet分配条件
        sheet_conditions = {}
        for sheet_name in self.data_compatible_sheets:
            sheet_cols = ExcelFileTools.get_sheet_column_names(
                self.file_path, sheet_name, header=True
            )
            sheet_col_set = set(sheet_cols)
            sheet_groups = []
            for group in real_groups:
                valid_conds = [c for c in group if c["col"] in sheet_col_set]
                if valid_conds:
                    sheet_groups.append(valid_conds)
            sheet_conditions[sheet_name] = sheet_groups
        return sheet_conditions

    def _filter_single_sheet(self, conn: duckdb.DuckDBPyConnection, sheet_name: str,
                            sheet_filter_groups: list[list[dict]], key_col: str,
                            export_all_cols: bool) -> tuple:
        """
        单个Sheet按真实列名执行筛选
        返回 (筛选结果Relation, 导出列名列表)，找不到主键返回(None, [])
        """
        real_cols = ExcelFileTools.get_sheet_column_names(
            self.file_path, sheet_name, header=True
        )
        if key_col not in real_cols:
            return None, []
        base_sql = f"""
            SELECT *
            FROM read_xlsx(
                '{self.file_path.replace(chr(39), chr(39)*2)}',
                sheet = '{sheet_name}',
                header = true,
                all_varchar = true
            )
        """
        rel = conn.sql(base_sql)
        # 应用筛选条件
        if sheet_filter_groups and any(sheet_filter_groups):
            where_sql = FilterBuilder.build_groups_sql(sheet_filter_groups)
            rel = rel.filter(where_sql)
        # 确定导出列：主键永远强制保留
        if export_all_cols:
            export_cols = real_cols.copy()
        else:
            used_cols = {key_col}
            for group in sheet_filter_groups:
                for cond in group:
                    used_cols.add(cond["col"])
            export_cols = [c for c in real_cols if c in used_cols]
        select_cols = ", ".join([f'"{c}"' for c in export_cols])
        final_rel = rel.select(select_cols)
        return final_rel, export_cols

    # ========== 模式1：内连接导出 ==========
    def filter_inner_join_export(self, filter_groups: list[list[dict]], save_path: str, key_col: str, export_all_cols: bool = False) -> int:
        """各Sheet分别筛选后，按主键列值内连接拼接导出，仅支持数据兼容Sheet"""
        if not self.data_compatible_sheets:
            raise RuntimeError("无支持数据模式的有效Sheet，无法执行高级筛选")
        if not filter_groups or not any(filter_groups):
            raise ValueError("请至少添加一个筛选条件")
        self.log("高级筛选仅支持标准二维表（无合并单元格），默认首行为真实列名")
        self.log(f"模式：内连接导出，主键列：{key_col}")
        self.log(f"共 {len(self.data_compatible_sheets)} 个Sheet参与筛选，复杂格式Sheet自动跳过")

        conn = duckdb.connect(":memory:")
        DuckExcelEngine._ensure_excel_extension(conn)
        try:
            sheet_conditions = self._split_conditions_by_sheet(filter_groups)
            sheet_results = {}
            valid_sheets = []
            for sheet_name in self.data_compatible_sheets:
                rel, cols = self._filter_single_sheet(conn, sheet_name, sheet_conditions[sheet_name], key_col, export_all_cols)
                if rel is None:
                    self.log(f"Sheet[{sheet_name}]未找到主键列，跳过连接运算")
                    continue
                sheet_results[sheet_name] = rel
                valid_sheets.append(sheet_name)
            if len(valid_sheets) == 0:
                raise RuntimeError("所有Sheet均未找到主键列，无法连接")
            if len(valid_sheets) == 1:
                final_rel = list(sheet_results.values())[0]
            else:
                base_name = valid_sheets[0]
                base_rel = sheet_results[base_name]
                base_rel.create_view("t0")
                current_view = "t0"
                current_data_cols = [c for c in base_rel.columns if c != key_col]
                for i in range(1, len(valid_sheets)):
                    name = valid_sheets[i]
                    rel = sheet_results[name]
                    rel.create_view(f"t{i}")
                    next_view = f"join_{i}"
                    new_data_cols = [c for c in rel.columns if c != key_col]
                    select_parts = [f'a."{key_col}"']
                    if current_data_cols:
                        select_parts.append(", ".join([f'a."{c}"' for c in current_data_cols]))
                    if new_data_cols:
                        select_parts.append(", ".join([f'b."{c}" AS "{name}_{c}"' for c in new_data_cols]))
                    select_str = ", ".join(select_parts)
                    join_sql = f"""
                        CREATE TEMP VIEW {next_view} AS
                        SELECT {select_str}
                        FROM {current_view} a
                        INNER JOIN t{i} b ON TRIM(a."{key_col}"::VARCHAR) = TRIM(b."{key_col}"::VARCHAR)
                    """
                    conn.execute(join_sql)
                    current_view = next_view
                    current_data_cols.extend([f"{name}_{c}" for c in new_data_cols])
                final_rel = conn.sql(f"SELECT * FROM {current_view}")
            # 空值补-
            data_cols = final_rel.columns
            final_select = ", ".join([f'COALESCE("{c}", \'-\') AS "{c}"' for c in data_cols])
            final_rel = conn.sql(f"SELECT {final_select} FROM final_rel")
            # 处理导出表头
            if not self.export_first_row_as_header:
                rename_parts = [f'"{col}" AS "a{idx+1}"' for idx, col in enumerate(data_cols)]
                final_rel = final_rel.select(", ".join(rename_parts))
            DuckExcelEngine.safe_to_excel(conn, final_rel, save_path, sheet_name="筛选结果", header=True)
            self.log(f"内连接筛选完成！结果已导出：{os.path.basename(save_path)}")
            return len(data_cols)
        finally:
            conn.close()

    # ========== 模式2：分Sheet分别导出 ==========
    def filter_separate_export(self, filter_groups: list[list[dict]], save_path: str,
                           key_col: str, export_all_cols: bool = False) -> int:
        """各Sheet分别筛选，合并为一个多Sheet的Excel，仅支持数据兼容Sheet"""
        if not self.data_compatible_sheets:
            raise RuntimeError("无支持数据模式的有效Sheet，无法执行高级筛选")
        if not filter_groups or not any(filter_groups):
            raise ValueError("请至少添加一个筛选条件")
        self.log("高级筛选仅支持标准二维表（无合并单元格），默认首行为真实列名")
        self.log("模式：各Sheet分别筛选，分Sheet导出")
        self.log(f"共 {len(self.data_compatible_sheets)} 个Sheet参与筛选，复杂格式Sheet自动跳过")

        conn = duckdb.connect(":memory:")
        DuckExcelEngine._ensure_excel_extension(conn)
        temp_files = []
        temp_sheet_names = []
        try:
            sheet_conditions = self._split_conditions_by_sheet(filter_groups)
            output_dir = os.path.dirname(save_path)
            base_name = os.path.splitext(os.path.basename(save_path))[0]
            for sheet_name in self.data_compatible_sheets:
                real_cols = ExcelFileTools.get_sheet_column_names(
                    self.file_path, sheet_name, header=True
                )
                has_key = key_col in real_cols
                # 执行筛选
                if has_key:
                    rel, cols = self._filter_single_sheet(conn, sheet_name, sheet_conditions[sheet_name],
                                                        key_col, export_all_cols)
                    if rel is None:
                        self.log(f"Sheet[{sheet_name}]未找到主键列，跳过")
                        continue
                else:
                    # 无主键时独立筛选
                    base_sql = f"""
                        SELECT *
                        FROM read_xlsx(
                            '{self.file_path.replace(chr(39), chr(39)*2)}',
                            sheet = '{sheet_name}',
                            header = true,
                            all_varchar = true
                        )
                    """
                    rel = conn.sql(base_sql)
                    if sheet_conditions[sheet_name] and any(sheet_conditions[sheet_name]):
                        where_sql = FilterBuilder.build_groups_sql(sheet_conditions[sheet_name])
                        rel = rel.filter(where_sql)
                    if export_all_cols:
                        cols = real_cols.copy()
                    else:
                        used_cols = set()
                        for group in sheet_conditions[sheet_name]:
                            for cond in group:
                                used_cols.add(cond["col"])
                        cols = [c for c in real_cols if c in used_cols]
                        if not cols:
                            cols = real_cols.copy()
                    select_cols = ", ".join([f'"{c}"' for c in cols])
                    rel = rel.select(select_cols)
                # 空值补-
                final_select = ", ".join([f'COALESCE("{c}", \'-\') AS "{c}"' for c in cols])
                data_rel = conn.sql(f"SELECT {final_select} FROM rel")
                # 导出为临时单Sheet文件
                temp_path = os.path.join(output_dir, f"{base_name}_tmp_{sheet_name}.xlsx")
                DuckExcelEngine.safe_to_excel(conn, data_rel, temp_path, sheet_name=sheet_name, header=True)
                temp_files.append(temp_path)
                temp_sheet_names.append(sheet_name)
            if not temp_files:
                raise RuntimeError("没有可导出的有效Sheet")
            # 合并为一个多Sheet文件
            ExcelFileTools.merge_excel_to_multi_sheet(temp_files, save_path, temp_sheet_names)
            exported_count = len(temp_files)
            self.log(f"分Sheet筛选完成！共导出 {exported_count} 个Sheet，保存至：{os.path.basename(save_path)}")
            return exported_count
        finally:
            conn.close()
            for f in temp_files:
                try:
                    if os.path.exists(f):
                        os.remove(f)
                except Exception:
                    pass

    # ========== 模式3：左连接Sheet1导出 ==========
    def filter_left_join_export(self, filter_groups: list[list[dict]], save_path: str, key_col: str, export_all_cols: bool = False) -> int:
        """以第一个Sheet为左表，按主键左连接其他Sheet，仅支持数据兼容Sheet"""
        if not self.data_compatible_sheets:
            raise RuntimeError("无支持数据模式的有效Sheet，无法执行高级筛选")
        if not filter_groups or not any(filter_groups):
            raise ValueError("请至少添加一个筛选条件")
        self.log("高级筛选仅支持标准二维表（无合并单元格），默认首行为真实列名")
        self.log(f"模式：以Sheet1为基准左连接，主键列：{key_col}")
        self.log(f"共 {len(self.data_compatible_sheets)} 个Sheet参与筛选，复杂格式Sheet自动跳过")

        conn = duckdb.connect(":memory:")
        DuckExcelEngine._ensure_excel_extension(conn)
        try:
            sheet_conditions = self._split_conditions_by_sheet(filter_groups)
            sheet_results = {}
            valid_sheets = []
            for sheet_name in self.data_compatible_sheets:
                rel, cols = self._filter_single_sheet(conn, sheet_name, sheet_conditions[sheet_name], key_col, export_all_cols)
                if rel is None:
                    self.log(f"Sheet[{sheet_name}]未找到主键列，跳过连接运算")
                    continue
                sheet_results[sheet_name] = rel
                valid_sheets.append(sheet_name)
            if len(valid_sheets) == 0:
                raise RuntimeError("所有Sheet均未找到主键列，无法连接")
            if len(valid_sheets) == 1:
                final_rel = list(sheet_results.values())[0]
            else:
                base_name = valid_sheets[0]
                base_rel = sheet_results[base_name]
                base_rel.create_view("t0")
                current_view = "t0"
                current_data_cols = [c for c in base_rel.columns if c != key_col]
                for i in range(1, len(valid_sheets)):
                    name = valid_sheets[i]
                    rel = sheet_results[name]
                    rel.create_view(f"t{i}")
                    next_view = f"join_{i}"
                    new_data_cols = [c for c in rel.columns if c != key_col]
                    select_parts = [f'a."{key_col}"']
                    if current_data_cols:
                        select_parts.append(", ".join([f'a."{c}"' for c in current_data_cols]))
                    if new_data_cols:
                        select_parts.append(", ".join([f'b."{c}" AS "{name}_{c}"' for c in new_data_cols]))
                    select_str = ", ".join(select_parts)
                    join_sql = f"""
                        CREATE TEMP VIEW {next_view} AS
                        SELECT {select_str}
                        FROM {current_view} a
                        LEFT JOIN t{i} b ON TRIM(a."{key_col}"::VARCHAR) = TRIM(b."{key_col}"::VARCHAR)
                    """
                    conn.execute(join_sql)
                    current_view = next_view
                    current_data_cols.extend([f"{name}_{c}" for c in new_data_cols])
                final_rel = conn.sql(f"SELECT * FROM {current_view}")
            # 空值补-
            data_cols = final_rel.columns
            final_select = ", ".join([f'COALESCE("{c}", \'-\') AS "{c}"' for c in data_cols])
            final_rel = conn.sql(f"SELECT {final_select} FROM final_rel")
            # 处理导出表头
            if not self.export_first_row_as_header:
                rename_parts = [f'"{col}" AS "a{idx+1}"' for idx, col in enumerate(data_cols)]
                final_rel = final_rel.select(", ".join(rename_parts))
            DuckExcelEngine.safe_to_excel(conn, final_rel, save_path, sheet_name="筛选结果", header=True)
            self.log(f"左连接筛选完成！结果已导出：{os.path.basename(save_path)}")
            return len(data_cols)
        finally:
            conn.close()

    # ========== 功能4：按列名对齐纵向合并 ==========
    def concat_same_column_sheets(self, output_path: str, sheet_list: list[str] = None) -> tuple[int, int]:
        """
        按真实列名对齐纵向合并，强制仅支持数据模式兼容的Sheet
        sheet_list：指定要拼接的Sheet列表，为None时默认拼接全部数据兼容Sheet
        返回：(成功拼接Sheet数, 跳过Sheet数)
        """
        if not self.file_path:
            raise RuntimeError("请先加载文件")
        
        if sheet_list is None:
            sheet_list = self.data_compatible_sheets.copy()
        else:
            # 强制过滤，仅保留数据兼容的Sheet
            original_count = len(sheet_list)
            sheet_list = [s for s in sheet_list if s in self.data_compatible_sheets]
            skipped = original_count - len(sheet_list)
            if skipped > 0:
                self.log(f"纵向拼接仅支持标准二维表，自动跳过 {skipped} 个含复杂格式的Sheet")
                self.log("跳过的Sheet：")
                for name in sheet_list:
                    if name not in self.data_compatible_sheets:
                        self.log(f"  - {name}：{self.complex_sheet_reasons.get(name, '格式不兼容')}")

        if len(sheet_list) < 2:
            raise RuntimeError("有效Sheet不足2个，无法执行拼接。带合并单元格的Sheet不支持列对齐拼接")
        
        self.log("纵向拼接强制使用标准二维表结构，首行作为列名进行对齐合并")
        conn = duckdb.connect(":memory:")
        DuckExcelEngine._ensure_excel_extension(conn)
        try:
            self.log("正在按列名对齐合并Sheet...")
            
            # 1. 收集所有Sheet的列名，按出现顺序生成并集列
            all_columns = []
            sheet_col_map = {}
            valid_sheets = []
            
            for sheet_name in sheet_list:
                try:
                    cols = ExcelFileTools.get_sheet_column_names(
                        self.file_path, sheet_name, header=True
                    )
                    sheet_col_map[sheet_name] = cols
                    valid_sheets.append(sheet_name)
                    # 新增列追加到末尾，保持第一个Sheet的列顺序优先
                    for col in cols:
                        if col not in all_columns:
                            all_columns.append(col)
                except Exception:
                    continue
            
            if len(valid_sheets) < 2:
                raise RuntimeError("有效Sheet不足2个，无法拼接")
            
            # 2. 每个Sheet生成拼接SQL：补全缺失列、加来源Sheet标识
            union_sql_parts = []
            file_esc = self.file_path.replace("'", "''")
            
            for sheet_name in valid_sheets:
                sheet_cols = sheet_col_map[sheet_name]
                select_parts = []
                
                # 第一列：来源Sheet名称
                sheet_name_esc = sheet_name.replace("'", "''")
                select_parts.append(f"'{sheet_name_esc}' AS 来源Sheet")
                
                # 按并集列顺序生成列，缺失的补'-'
                for col in all_columns:
                    if col in sheet_cols:
                        select_parts.append(f'"{col}"')
                    else:
                        select_parts.append(f"'-' AS \"{col}\"")
                
                select_str = ", ".join(select_parts)
                sheet_esc = sheet_name.replace("'", "''")
                
                sheet_sql = f"""
                    SELECT {select_str}
                    FROM read_xlsx('{file_esc}', sheet='{sheet_esc}', header=true, all_varchar=true)
                """
                union_sql_parts.append(sheet_sql)
            
            # 3. UNION ALL 拼接 + 整行去重
            union_all_sql = " UNION ALL ".join(union_sql_parts)
            dedup_sql = f"""
                SELECT DISTINCT *
                FROM ({union_all_sql}) t
                ORDER BY 来源Sheet
            """
            final_rel = conn.sql(dedup_sql)
            # 4. 导出Excel
            DuckExcelEngine.safe_to_excel(conn, final_rel, output_path, sheet_name="拼接结果", header=True)
            
            success_count = len(valid_sheets)
            skipped_count = len(sheet_list) - success_count
            self.log(f"列名对齐合并完成，成功 {success_count} 个Sheet，跳过 {skipped_count} 个")
            return success_count, skipped_count
        finally:
            conn.close()