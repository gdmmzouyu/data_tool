from __future__ import annotations
import os
import sys
import hashlib
import concurrent.futures
import time
import json
import html
import requests
import duckdb
import datetime
import threading
import pyarrow as pa
from typing import Optional, List
from cryptography.fernet import Fernet


# ===================== 模块内部常量 =====================
_REALTIME_TABLE_NAME = "dat_tmp"
_REALTIME_USER_SEP = "|"
_REALTIME_FIELD_SEP = ","
_INSERT_BATCH_SIZE = 680000
_DEFAULT_FETCH_WORKERS = 4  # 默认并行线程数
_MIN_WORKERS = 1
_MAX_WORKERS = 8
_API_REQUEST_TIMEOUT = 300
_API_CONNECT_TIMEOUT = 10         # 新增：TCP连接建立超时，内网快速失败
_API_READ_TIMEOUT = 120           # 新增：数据读取超时，给足大数据量传输时间
_FUTURE_TIMEOUT = 1800  # 单任务最大超时时间，略大于接口超时

# ========== 新增：测试专用True/False - 本机MAC旁路放行开关，正式环境务必保持False ==========
_DEBUG_BYPASS_MAC_WHITELIST = True

# ========== 新增：展示与日志内部开关 ==========
_ENABLE_NODE_VIEW = True        # True=UI显示节点进度图，隐藏原日志面板；False=回退原日志模式
_ENABLE_LOG_FILE = False         # True=落地日志txt文件到程序目录，用于调试；False=不生成文件
_LOG_FILE_ENCODING = "utf-8"

# ===================== 内部：线程本地连接池 =====================
_thread_local = threading.local()

def _get_thread_session() -> requests.Session:
    """获取当前线程专属的requests会话，复用TCP连接池，线程安全"""
    if not hasattr(_thread_local, 'session'):
        session = requests.Session()
        # 配置连接池参数，适配多分区并发
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=8,
            pool_maxsize=8,
            max_retries=0  # 重试逻辑由业务层统一控制
        )
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        _thread_local.session = session
    return _thread_local.session

# ===================== 内部：密钥配置加载（原逻辑完整保留）=====================
_sec_config = None
_module_available = False

def _get_app_root_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _load_sec_config():
    global _sec_config, _module_available
    try:
         # 直接复用主程序已定义的密钥路径，和打包逻辑完全对齐
        from __main__ import KEY_PATH
        master_key_path = KEY_PATH
        if not os.path.exists(master_key_path):
            return

        with open(master_key_path, "rb") as f:
            master_key = f.read()
        fernet = Fernet(master_key)

        import sec_encrypted
        dec_text = fernet.decrypt(sec_encrypted.ENCRYPTED_BLOB).decode("utf-8")
        scope = {}
        exec(dec_text, {}, scope)

        class SecWrapper:
            pass
        _sec_config = SecWrapper()
        _sec_config.API_CONFIG = scope["API_CONFIG"]
        _sec_config.ENCRYPT_SALT = scope["ENCRYPT_SALT"]
        _sec_config.BASE62_CHARS = scope["BASE62_CHARS"]
        _sec_config.KEY_EXPIRE_DATE = scope["KEY_EXPIRE_DATE"]

        _module_available = True
    except Exception:
        _module_available = False
        _sec_config = None

_load_sec_config()

def is_module_available() -> bool:
    return _module_available

# ===================== 内部：加密工具（原逻辑完整保留）=====================
class EncryptTools:
    if _module_available:
        BASE62_CHARS = _sec_config.BASE62_CHARS
        SALT = _sec_config.ENCRYPT_SALT
    else:
        BASE62_CHARS = ""
        SALT = ""

    AREA_MAP = {
        "1": "茂南",
        "2": "电白",
        "3": "高州",
        "4": "化州",
        "5": "信宜"
    }

    @staticmethod
    def mobile_to_md5_code(mobile: str) -> Optional[str]:
        if not mobile or len(mobile) != 11 or not mobile.startswith("1") or not mobile.isdigit():
            return None
        if not EncryptTools.BASE62_CHARS:
            return None
        raw_str = mobile + EncryptTools.SALT
        md5_hex = hashlib.md5(raw_str.encode("utf-8")).hexdigest()
        segments = [
            md5_hex[0:3], md5_hex[3:6], md5_hex[6:9], md5_hex[9:12],
            md5_hex[12:15], md5_hex[15:18], md5_hex[18:21], md5_hex[21:24],
            md5_hex[24:27], md5_hex[27:30]
        ]
        code = []
        for seg in segments:
            idx = int(seg, 16) % 62
            code.append(EncryptTools.BASE62_CHARS[idx])
        return "".join(code)

    @staticmethod
    def batch_mobile_to_code(mobile_list: List[str]) -> List[tuple[str, str]]:
        result = []
        for mobile in mobile_list:
            code = EncryptTools.mobile_to_md5_code(mobile.strip())
            if code:
                result.append((mobile.strip(), code))
        return result

    @staticmethod
    def mac_bind_hash(mac: str, salt: str) -> str:
        raw = mac + salt
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()

# ===================== 内部：本机MAC获取（Windows已修复）=====================
def _get_local_mac() -> str:
    """获取本机首个活跃有线物理网卡MAC，无线/虚拟网卡一律不支持"""
    try:
        if sys.platform == "win32":
            import ctypes
            from ctypes import wintypes

            AF_UNSPEC = 0
            IF_TYPE_ETHERNET_CSMACD = 6
            IfOperStatusUp = 1
            ERROR_BUFFER_OVERFLOW = 111
            GAA_FLAG_DEFAULT = 0

            class SOCKADDR(ctypes.Structure):
                _fields_ = [
                    ("sa_family", wintypes.USHORT),
                    ("sa_data", wintypes.CHAR * 14)
                ]

            class IP_ADAPTER_ADDRESSES_LH(ctypes.Structure):
                pass

            PIP_ADAPTER_ADDRESSES_LH = ctypes.POINTER(IP_ADAPTER_ADDRESSES_LH)
            # 修复：补全结构体开头字段，内存布局与Windows API完全一致
            IP_ADAPTER_ADDRESSES_LH._fields_ = [
                ("Length", wintypes.ULONG),
                ("IfIndex", wintypes.DWORD),
                ("Next", PIP_ADAPTER_ADDRESSES_LH),
                ("AdapterName", wintypes.LPSTR),
                ("FirstUnicastAddress", ctypes.c_void_p),
                ("FirstAnycastAddress", ctypes.c_void_p),
                ("FirstMulticastAddress", ctypes.c_void_p),
                ("FirstDnsServerAddress", ctypes.c_void_p),
                ("DnsSuffix", wintypes.LPWSTR),
                ("Description", wintypes.LPWSTR),
                ("FriendlyName", wintypes.LPWSTR),
                ("PhysicalAddress", wintypes.BYTE * 8),
                ("PhysicalAddressLength", wintypes.DWORD),
                ("Flags", wintypes.DWORD),
                ("Mtu", wintypes.DWORD),
                ("IfType", wintypes.DWORD),
                ("OperStatus", ctypes.c_uint),
                ("Ipv6IfIndex", wintypes.DWORD),
                ("ZoneIndices", wintypes.DWORD * 16),
            ]

            GetAdaptersAddresses = ctypes.windll.iphlpapi.GetAdaptersAddresses
            GetAdaptersAddresses.argtypes = [
                wintypes.ULONG, wintypes.ULONG, ctypes.c_void_p,
                PIP_ADAPTER_ADDRESSES_LH, ctypes.POINTER(wintypes.ULONG)
            ]
            GetAdaptersAddresses.restype = wintypes.DWORD

            # 修复：标准两次调用流程，自适应缓冲区大小
            buf_size = wintypes.ULONG(0)
            ret = GetAdaptersAddresses(
                AF_UNSPEC, GAA_FLAG_DEFAULT,
                None, None, ctypes.byref(buf_size)
            )
            if ret != ERROR_BUFFER_OVERFLOW:
                raise RuntimeError(f"系统网卡信息读取失败，错误码：{ret}")

            buf = ctypes.create_string_buffer(buf_size.value)
            p_adapters = ctypes.cast(buf, PIP_ADAPTER_ADDRESSES_LH)

            ret = GetAdaptersAddresses(
                AF_UNSPEC, GAA_FLAG_DEFAULT,
                None, p_adapters, ctypes.byref(buf_size)
            )
            if ret != 0:
                raise RuntimeError(f"系统网卡信息读取失败，错误码：{ret}")

            current = p_adapters
            while current:
                adapter = current.contents
                if (adapter.IfType == IF_TYPE_ETHERNET_CSMACD
                        and adapter.OperStatus == IfOperStatusUp
                        and adapter.PhysicalAddressLength == 6):
                    mac_bytes = bytes(adapter.PhysicalAddress[:6])
                    mac_hex = mac_bytes.hex().upper()
                    return mac_hex
                current = adapter.Next

            raise RuntimeError("未检测到活跃的有线物理网卡，本功能仅支持有线网络环境使用，不支持WiFi及虚拟网卡")

        elif sys.platform == "darwin":
            import subprocess
            try:
                output = subprocess.check_output(
                    ["networksetup", "-listallhardwareports"],
                    stderr=subprocess.DEVNULL,
                    text=True
                )
            except Exception:
                raise RuntimeError("系统网卡信息读取失败")

            ethernet_devices = []
            port_blocks = output.strip().split("\n\n")
            for block in port_blocks:
                lines = block.strip().splitlines()
                port_name = ""
                device = ""
                mac = ""
                for line in lines:
                    line = line.strip()
                    if line.startswith("Hardware Port:"):
                        port_name = line.split(":", 1)[1].strip()
                    elif line.startswith("Device:"):
                        device = line.split(":", 1)[1].strip()
                    elif line.startswith("Ethernet Address:"):
                        mac = line.split(":", 1)[1].strip()

                if (port_name and device and mac
                        and "Ethernet" in port_name
                        and "Wi-Fi" not in port_name
                        and "Bluetooth" not in port_name
                        and "Thunderbolt Bridge" not in port_name):
                    ethernet_devices.append((device, mac))

            if not ethernet_devices:
                raise RuntimeError("未检测到有线物理网卡，本功能仅支持有线网络环境使用，不支持WiFi及虚拟网卡")

            for dev, mac in ethernet_devices:
                try:
                    ifconfig_out = subprocess.check_output(
                        ["ifconfig", dev],
                        stderr=subprocess.DEVNULL,
                        text=True
                    )
                    if "status: active" in ifconfig_out:
                        mac_clean = mac.replace(":", "").upper()
                        if len(mac_clean) == 12:
                            return mac_clean
                except Exception:
                    continue
            raise RuntimeError("未检测到活跃的有线物理网卡，本功能仅支持有线网络环境使用，不支持WiFi及虚拟网卡")

        elif sys.platform.startswith("linux"):
            net_dir = "/sys/class/net"
            if not os.path.isdir(net_dir):
                raise RuntimeError("系统网卡信息读取失败")

            for iface in sorted(os.listdir(net_dir)):
                iface_path = os.path.join(net_dir, iface)
                if iface == "lo":
                    continue
                device_path = os.path.join(iface_path, "device")
                if not os.path.isdir(device_path):
                    continue
                wireless_path = os.path.join(iface_path, "wireless")
                if os.path.isdir(wireless_path):
                    continue
                try:
                    with open(os.path.join(iface_path, "type"), "r") as f:
                        if_type = f.read().strip()
                    if if_type != "1":
                        continue
                    with open(os.path.join(iface_path, "operstate"), "r") as f:
                        operstate = f.read().strip()
                    if operstate != "up":
                        continue
                    with open(os.path.join(iface_path, "address"), "r") as f:
                        mac = f.read().strip()
                    mac_clean = mac.replace(":", "").upper()
                    if len(mac_clean) == 12:
                        return mac_clean
                except Exception:
                    continue
            raise RuntimeError("未检测到活跃的有线物理网卡，本功能仅支持有线网络环境使用，不支持WiFi及虚拟网卡")

        else:
            raise RuntimeError("不支持的操作系统，本功能仅支持Windows、MacOS、Linux")

    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"设备身份校验失败：{str(e)}")

def _get_system_resource_info() -> tuple[int, int]:
    """
    跨平台获取系统资源信息，零强依赖
    优先使用psutil精准探测，无则自动降级为标准库估算，探测失败返回保守默认值
    :return: (可用内存字节数, CPU逻辑核心数)
    """
    cpu_count = os.cpu_count() or 2
    default_mem = 2 * 1024 * 1024 * 1024  # 探测失败默认按2GB可用内存保守估算

    # 优先尝试psutil，精度最高
    try:
        import psutil
        mem = psutil.virtual_memory()
        return int(mem.available), cpu_count
    except ImportError:
        pass

    # 标准库降级实现
    try:
        if sys.platform == "win32":
            import ctypes
            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]
            ms = MEMORYSTATUSEX()
            ms.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms))
            available_mem = int(ms.ullAvailPhys * 0.8)
            return available_mem, cpu_count

        elif sys.platform == "darwin":
            import subprocess
            output = subprocess.check_output(["sysctl", "-n", "hw.memsize"], text=True).strip()
            total_mem = int(output)
            available_mem = int(total_mem * 0.5)
            return available_mem, cpu_count

        elif sys.platform.startswith("linux"):
            mem_info = {}
            with open("/proc/meminfo", "r") as f:
                for line in f:
                    parts = line.split()
                    if len(parts) >= 2:
                        mem_info[parts[0].rstrip(":")] = int(parts[1]) * 1024
            available_mem = mem_info.get("MemAvailable", 
                mem_info.get("MemFree", 0) + mem_info.get("Buffers", 0) + mem_info.get("Cached", 0))
            return int(available_mem * 0.9), cpu_count
    except Exception:
        pass

    # 所有探测失败，返回保守默认值
    return default_mem, cpu_count

# ===================== 业务核心类：纯内存实时比对 =====================
class RealtimeBusiness:
    def __init__(self, log_callback, node_callback=None):
        self.log = log_callback
        # 节点进度回调，为空则不执行，兼容旧调用
        self._node_update = node_callback if node_callback is not None else lambda *_args, **_kwargs: None
        self.table_name = _REALTIME_TABLE_NAME

        if not _module_available:
            raise RuntimeError("缺少授权文件，实时数据模块无法使用")
        self.api_config = _sec_config.API_CONFIG.copy()

        self.api_method_name = self.api_config["method_name"]
        self.data_field = self.api_config["data_field"]
        self.check_field = self.api_config["check_field"]
        self.busi_prefix = self.api_config["busi_prefix"]
        self.key_expire_date = _sec_config.KEY_EXPIRE_DATE
        if "mac_busi_type" not in self.api_config:
            raise RuntimeError("授权配置缺失：缺少MAC白名单业务类型配置")
        self.mac_busi_type = self.api_config["mac_busi_type"]

        self.business_types = [f"{self.busi_prefix}{i}" for i in range(1, 9)]
        self.max_retries = 5
        self.retry_delay = 10

        # 分区一预拉取缓存，单次调用内有效
        self._cached_partition_one: Optional[list] = None
        # 运行时自适应状态
        self._degrade_level = 0  # 0-正常 1-一级预警 2-二级减压 3-三级限流
        self._degrade_notified = False  # 降级提示仅输出一次
        self._system_mem_available = 0
        self._system_cpu_count = 0
        # 动态运行参数（运行时可随降级调整）
        self._flush_threshold = _INSERT_BATCH_SIZE
        self._batch_write_size = 100000
        self._runtime_fetch_workers = _DEFAULT_FETCH_WORKERS
        self._duckdb_threads = 4
        self._duckdb_mem_limit = "2GB"

    def _detect_system_resources(self):
        self._node_update("n2", "running")
        """探测系统资源并输出适配提示，用户无专业感知"""
        avail_mem, cpu_cnt = _get_system_resource_info()
        self._system_mem_available = avail_mem
        self._system_cpu_count = cpu_cnt

        avail_mem_mb = avail_mem // (1024 * 1024)
        self.log(f"【系统适配】检测到可用内存约{avail_mem_mb:,}MB / {cpu_cnt}核CPU，已自动优化配置")
        self._node_update("n2", "success", f"可用内存{avail_mem_mb/1024:.1f}G/{cpu_cnt}核")

        # 低配设备兼容模式提示
        if avail_mem_mb < 1024 or cpu_cnt <= 2:
            self.log("【系统适配】检测到设备配置较低，已自动开启兼容模式，运行会稍慢，请耐心等待")

    def _calc_adaptive_params(self, user_fetch_workers: Optional[int] = None):
        """
        根据系统资源计算所有自适应运行参数
        :param user_fetch_workers: 用户手动指定的拉取线程数，None则全自动适配
        """
        self._node_update("n4", "running")
        avail_mem_mb = self._system_mem_available // (1024 * 1024)
        cpu_cnt = self._system_cpu_count

        # 1. DuckDB内存配额：可用内存35%，硬上限2GB，硬下限512MB
        duckdb_mem_mb = int(avail_mem_mb * 0.35)
        duckdb_mem_mb = max(512, min(duckdb_mem_mb, 2048))
        self._duckdb_mem_limit = f"{duckdb_mem_mb}MB"

        # 2. DuckDB内部计算线程：留1核给UI，硬上限4，硬下限1
        duck_threads = max(1, cpu_cnt - 1)
        duck_threads = min(duck_threads, 4)
        self._duckdb_threads = duck_threads

        # 3. 拉取并发线程数：用户手动指定优先，否则自动计算
        if user_fetch_workers is not None:
            fetch_workers = user_fetch_workers
        else:
            if cpu_cnt <= 2:
                fetch_workers = 2
            elif cpu_cnt <= 4:
                fetch_workers = 3
            else:
                fetch_workers = 4
        fetch_workers = max(_MIN_WORKERS, min(fetch_workers, _MAX_WORKERS))
        self._runtime_fetch_workers = fetch_workers
        self._base_fetch_workers = fetch_workers  # 保存基准值，降级恢复用

        # 4. 缓冲区刷写阈值：内存配额MB × 100，硬上限对齐单分区最大数据量，硬下限20万
        flush_threshold = duckdb_mem_mb * 100
        flush_threshold = max(200000, min(flush_threshold, _INSERT_BATCH_SIZE))
        self._flush_threshold = flush_threshold
        self._base_flush_threshold = flush_threshold  # 保存基准值，降级恢复用

        # 5. 单批次写入条数：缓冲区阈值的1/2，硬上限40万，硬下限5万
        batch_size = flush_threshold // 4
        batch_size = max(50000, min(batch_size, 400000))
        self._batch_write_size = batch_size

        # 6. 节点成功状态
        self._node_update("n4", "success", f"全速{self._runtime_fetch_workers}线程")

    def _check_memory_watermark(self):
        """检查进程内存水位，分级降级+自动恢复，等级变化时同步更新日志与节点状态"""
        # 无psutil则降级为强制GC，不做精准水位判断
        try:
            import psutil
            process = psutil.Process()
            current_mem = process.memory_info().rss
        except ImportError:
            import gc
            gc.collect()
            return

        mem_limit_bytes = int(float(self._duckdb_mem_limit.replace("MB", "")) * 1024 * 1024 * 1.5)
        mem_ratio = current_mem / mem_limit_bytes
        mem_percent = f"{mem_ratio * 100:.0f}%"
        prev_level = self._degrade_level

        # 三级限流：超过95%，串行拉取，极小缓冲区
        if mem_ratio >= 0.95:
            self._degrade_level = 3
        # 二级减压：超过85%，降并发，缩小缓冲区
        elif mem_ratio >= 0.85:
            self._degrade_level = 2
        # 一级预警：超过70%，强制GC
        elif mem_ratio >= 0.7:
            self._degrade_level = 1
        # 水位回落，逐步恢复性能
        elif mem_ratio < 0.5:
            self._degrade_level = max(0, self._degrade_level - 1)

        # 等级发生变化时，同步更新运行参数、日志、节点状态
        if self._degrade_level != prev_level:
            old_workers = self._runtime_fetch_workers

            if self._degrade_level == 0:
                # 恢复到基准值
                self._runtime_fetch_workers = self._base_fetch_workers
                self._flush_threshold = self._base_flush_threshold
                self.log(f"【自适应调整】内存水位{mem_percent}，已恢复全速运行，并发线程调整为 {self._runtime_fetch_workers} 线程")
                self._node_update("n4", "success", f"全速{self._runtime_fetch_workers}线程")

            elif self._degrade_level == 1:
                # 一级：仅GC，不调整运行参数
                self.log(f"【自适应调整】内存水位{mem_percent}，触发一级预警，保持 {self._runtime_fetch_workers} 线程，执行内存回收")
                self._node_update("n4", "success", f"1级预警 {self._runtime_fetch_workers}线程")

            elif self._degrade_level == 2:
                # 二级：并发数-1，缓冲区减半
                self._runtime_fetch_workers = max(_MIN_WORKERS, self._base_fetch_workers - 1)
                self._flush_threshold = max(200000, self._base_flush_threshold // 2)
                self.log(f"【自适应调整】内存水位{mem_percent}，触发二级减压，并发线程从 {old_workers} 调整为 {self._runtime_fetch_workers} 线程，缓冲区已缩小")
                self._node_update("n4", "success", f"2级减压 {self._runtime_fetch_workers}线程")

            elif self._degrade_level == 3:
                # 三级：串行拉取，缓冲区最小
                self._runtime_fetch_workers = _MIN_WORKERS
                self._flush_threshold = 200000
                self.log(f"【自适应调整】内存水位{mem_percent}，触发三级限流，并发线程从 {old_workers} 调整为 {self._runtime_fetch_workers} 线程，已降至最低负载")
                self._node_update("n4", "success", f"3级限流 {self._runtime_fetch_workers}线程")

            import gc
            gc.collect()
            return

        # 一级预警强制GC
        if self._degrade_level >= 1:
            import gc
            gc.collect()

    @property
    def fixed_day(self) -> str:
        yesterday = datetime.date.today() - datetime.timedelta(days=1)
        return yesterday.strftime("%Y%m%d")

    # ========== 三层安全校验（原逻辑完整保留）==========
    def _verify_mac_whitelist(self, access_token: str, pre_fetch_data: Optional[list] = None):
        self._node_update("n3_2", "running")
        self.log("正在校验设备授权...")
        local_mac = _get_local_mac()

        # 优先使用预拉取数据，未传入则回退原逻辑
        if pre_fetch_data is None:
            resp_data = self._call_single_business(access_token, self.mac_busi_type, with_day=True, custom_day="20991231")
        else:
            resp_data = pre_fetch_data

        if not resp_data or len(resp_data) == 0:
            self._node_update("n3_2", "error")
            raise RuntimeError("设备授权校验失败")

        whitelist = set()
        for item in resp_data:
            raw_str = item.get("cnt_99", "")
            if not raw_str:
                continue
            mac_list = raw_str.split(_REALTIME_USER_SEP)
            for mac in mac_list:
                mac = mac.strip().upper().replace(":", "").replace("-", "")
                if len(mac) == 12 and all(c in "0123456789ABCDEF" for c in mac):
                    whitelist.add(mac)

        if not whitelist:
            self._node_update("n3_2", "error", "无有效授权")
            raise RuntimeError("无有效授权设备列表")
        if local_mac not in whitelist:
            # 测试模式旁路放行，全流程正常执行
            if _DEBUG_BYPASS_MAC_WHITELIST:
                self.log("【测试模式】本机MAC已旁路放行，白名单校验流程执行完成")
            else:
                self._node_update("n3_2", "error", "设备未授权")
                raise RuntimeError("当前设备未在授权列表中，无法使用该功能")
        self.log("设备授权校验通过")
        self._node_update("n3_2", "success", "允许使用本版块")

    def _get_access_token(self) -> Optional[str]:
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.api_config["client_id"],
            "client_secret": self.api_config["client_secret"]
        }
        try:
            resp = requests.post(self.api_config["token_url"], data=payload, timeout=_API_REQUEST_TIMEOUT)
            if resp.status_code == 200:
                token_data = json.loads(html.unescape(resp.text))
                if 'access_token' in token_data:
                    return token_data['access_token']
            self.log("连接获取失败")
            return None
        except Exception:
            self.log("连接获取异常")
            return None

    def _check_key_expire(self, access_token: str, pre_fetch_data: Optional[list] = None):
        self._node_update("n3_1", "running")
        self.log("正在校验授权有效期...")

        # 优先使用预拉取数据，未传入则回退原逻辑
        if pre_fetch_data is None:
            resp_data = self._call_single_business(access_token, self.business_types[0])
        else:
            resp_data = pre_fetch_data

        if not resp_data or len(resp_data) == 0:
            self._node_update("n3_1", "error", "请求超时")
            raise RuntimeError("请求超时，请稍后再试")

        first_item = resp_data[0]
        server_date = str(first_item.get(self.check_field, "")).strip()
        if not server_date or len(server_date) != 8 or not server_date.isdigit():
            self._node_update("n3_1", "error", "校验异常")
            raise RuntimeError("授权校验异常，请联系管理员")

        if server_date <= self.key_expire_date:
            self.log(f"校验通过，授权有效期至{self.key_expire_date}")
            self._node_update("n3_1", "success", f"有效期至{self.key_expire_date}")
        else:
            self._node_update("n3_1", "error", "授权已过期")
            raise RuntimeError("授权已过期，请联系管理员索取新的授权文件")
        

    def _parallel_fetch_verify_data(self, access_token: str) -> tuple[list, list]:
        """
        双线程并行拉取：MAC白名单数据 + 分区一有效期数据
        :return: (mac_resp_data, partition_one_resp_data)
        """
        self.log("正在校验设备与鉴权是否有效...")
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
            future_mac = executor.submit(
                self._call_single_business,
                access_token, self.mac_busi_type,
                with_day=True, custom_day="20991231"
            )
            future_part1 = executor.submit(
                self._call_single_business,
                access_token, self.business_types[0]
            )
            future_map = {
                future_mac: "MAC白名单",
                future_part1: "分区一有效期"
            }
            mac_data = None
            part1_data = None
            for future in concurrent.futures.as_completed(future_map, timeout=_FUTURE_TIMEOUT):
                task_name = future_map[future]
                try:
                    data = future.result()
                    if not data:
                        raise RuntimeError(f"{task_name}接口返回空数据")
                    if task_name == "MAC白名单":
                        mac_data = data
                    else:
                        part1_data = data
                except Exception as e:
                    raise RuntimeError(f"{task_name}校验拉取失败: {str(e)}")
        return mac_data, part1_data

    # ========== 接口调用（支持自定义日期，连接池复用）==========
    def _call_single_business(self, access_token: str, business_type: str, with_day: bool = True, custom_day: str = None) -> Optional[list]:
        session = _get_thread_session()
        start_time = time.time()
        current_token = access_token
        token_refreshed = False  # 标记是否已经续过一次token，避免死循环

        for attempt in range(1, self.max_retries + 1):
            # 总耗时兜底，复用原有全局总超时常量
            if time.time() - start_time > _FUTURE_TIMEOUT:
                self.log(f"分区{business_type} 总耗时已达上限，终止重试")
                return None

            try:
                params = {
                    "access_token": current_token,
                    "method": self.api_method_name,
                    "format": "json",
                    "version": "1.0.0"
                }
                headers = {
                    "thirdPlatform-appid": self.api_config["app_id"],
                    "thirdPlatform-appKey": self.api_config["app_key"],
                    "Content-Type": "application/json"
                }
                data = {
                    "ReqMethod": f"com.gmcc.mm.dsc.{self.api_method_name}",
                    "MesFormat": "json",
                    "YW_TYPE": business_type
                }
                if with_day:
                    data["DAY"] = custom_day if custom_day is not None else self.fixed_day

                # 拆分超时：直接引用顶部外部常量
                resp = session.post(
                    self.api_config["api_url"],
                    params=params,
                    headers=headers,
                    json=data,
                    timeout=(_API_CONNECT_TIMEOUT, _API_READ_TIMEOUT)
                )

                # ========== HTTP状态码分类处理 ==========
                if resp.status_code == 200:
                    # 兜底：响应体非JSON异常，服务端网关报错时触发
                    try:
                        result = json.loads(html.unescape(resp.text))
                    except json.JSONDecodeError:
                        if attempt < self.max_retries:
                            wait_sec = min(2 ** attempt, 10)
                            self.log(f"分区{business_type} 第{attempt}次响应格式异常，{wait_sec}秒后重试")
                            time.sleep(wait_sec)
                            continue
                        else:
                            self.log(f"分区{business_type} 响应格式持续异常，重试耗尽")
                            return None

                    resp_data = result.get('BodyResp', {}).get('RespData', [])

                    # 兜底：首次空数据软重试，避免偶发缓存抖动
                    if not resp_data and attempt == 1:
                        self.log(f"分区{business_type} 返回空数据，执行一次软重试")
                        time.sleep(2)
                        continue

                    return resp_data

                # Token失效自动续期兜底：仅重试一次，不计入重试次数
                elif resp.status_code in (401, 403) and not token_refreshed:
                    self.log(f"分区{business_type} 鉴权失效，正在重新获取连接...")
                    new_token = self._get_access_token()
                    if new_token:
                        current_token = new_token
                        token_refreshed = True
                        continue
                    # 获取失败则走正常错误逻辑

                # 客户端参数错误，重试无效，直接终止
                elif 400 <= resp.status_code < 500 and resp.status_code != 429:
                    self.log(f"分区{business_type} 调用失败，状态码{resp.status_code}（客户端错误，终止重试）")
                    return None

                # 服务端错误/限流，进入指数退避重试
                if attempt < self.max_retries:
                    wait_sec = min(2 ** attempt, 10)
                    self.log(f"分区{business_type} 第{attempt}次调用失败，状态码{resp.status_code}，{wait_sec}秒后重试")
                    time.sleep(wait_sec)

            # ========== 异常分类处理 ==========
            except requests.exceptions.ConnectionError:
                if attempt < self.max_retries:
                    wait_sec = min(2 ** attempt, 10)
                    self.log(f"分区{business_type} 第{attempt}次连接异常，{wait_sec}秒后重试")
                    time.sleep(wait_sec)

            except requests.exceptions.Timeout:
                if attempt < self.max_retries:
                    wait_sec = min(2 ** attempt, 10)
                    self.log(f"分区{business_type} 第{attempt}次请求超时，{wait_sec}秒后重试")
                    time.sleep(wait_sec)

            except Exception:
                if attempt < self.max_retries:
                    wait_sec = min(2 ** attempt, 10)
                    self.log(f"分区{business_type} 第{attempt}次调用异常，{wait_sec}秒后重试")
                    time.sleep(wait_sec)

        # 所有重试耗尽
        self.log(f"分区{business_type} 重试{self.max_retries}次后全部失败")
        return None

    # ========== 数据解析（原逻辑完整保留）==========
    def _parse_raw_data(self, raw_list: list) -> list[tuple]:
        parsed = []
        sep_user = _REALTIME_USER_SEP
        sep_field = _REALTIME_FIELD_SEP
        for item in raw_list:
            raw_str = item.get(self.data_field, "")
            if not raw_str:
                continue
            users = raw_str.split(sep_user)
            for user_str in users:
                user_str = user_str.strip()
                if not user_str:
                    continue
                fields = user_str.split(sep_field)
                if len(fields) != 3:
                    continue
                area, code, if_mg = fields
                area = area.strip()
                code = code.strip()
                if_mg = if_mg.strip().upper()
                if len(code) == 10 and area in EncryptTools.AREA_MAP and if_mg in ("Y", "N"):
                    parsed.append((area, code, if_mg))
        return parsed

    # ========== pyarrow内存批量写入：最终稳定通用版 ==========
    def _memory_batch_insert(self, conn: duckdb.DuckDBPyConnection, data_list: list, table_name: str, columns: Optional[List[str]] = None):
        """
        PyArrow 零拷贝批量写入（通用动态列版）
        :param columns: 数据对应的列名列表，不传则默认按顺序使用 data_list 中元组的索引作为列名
        """
        if not data_list:
            return
        
        batch_size = self._batch_write_size
        col_count = len(data_list[0])
        
        # 未传列名则自动生成，保证兼容性
        if columns is None:
            columns = [f"col{i}" for i in range(col_count)]
        if len(columns) != col_count:
            raise ValueError(f"列名数量 {len(columns)} 与数据列数 {col_count} 不匹配")

        conn.execute("BEGIN TRANSACTION")
        try:
            for start_idx in range(0, len(data_list), batch_size):
                batch = data_list[start_idx: start_idx + batch_size]
                # 按传入列名动态构造 Arrow 表，支持任意列数
                cols = list(zip(*batch))
                arrow_dict = {
                    col_name: list(cols[i])
                    for i, col_name in enumerate(columns)
                }
                arrow_table = pa.table(arrow_dict)
                
                conn.from_arrow(arrow_table).insert_into(table_name)
            
            conn.execute("COMMIT")
        except Exception as e:
            conn.execute("ROLLBACK")
            err_short = str(e).splitlines()[0][:300]
            raise RuntimeError(f"批量写入失败: {err_short}") from e
        
    # ========== 核心：全内存实时比对入口 ==========
    def realtime_compare(self, mobile_list: List[str], fetch_workers: int = _DEFAULT_FETCH_WORKERS) -> List[dict]:
        """
        一键执行：授权校验 → 自适应配置 → 多线程拉取全量 → 内存建库 → 实时比对 → 自动释放
        :param mobile_list: 待查询手机号列表
        :param fetch_workers: 拉取并发线程数，传入None则全自动适配
        :return: 比对结果列表，包含手机号、归属地、是否敏感
        """
        # 前置号码清洗与去重
        if not mobile_list:
            raise ValueError("待查询手机号列表为空")

        original_total = len(mobile_list)
        mobile_list = list(dict.fromkeys(m.strip() for m in mobile_list if m.strip()))
        dedup_count = original_total - len(mobile_list)
        if dedup_count > 0:
            self.log(f"待查询号码前置去重，移除重复号码 {dedup_count:,} 条，剩余有效号码 {len(mobile_list):,} 条")

        if not mobile_list:
            raise ValueError("去重后无有效手机号可查询")

        # 1. 系统资源探测 + 自适应参数计算
        self._detect_system_resources()
        self._calc_adaptive_params(user_fetch_workers=fetch_workers)
        self.log(f"本次使用 {self._runtime_fetch_workers} 线程并行拉取")

        # 2. 获取访问令牌
        self.log("正在获取连接...")
        token = self._get_access_token()
        if not token:
            raise RuntimeError("访问令牌获取失败，无法拉取数据")

        # 3. 双线程并行拉取校验数据 + 顺序执行校验
        self._node_update("n3_1", "running")
        self._node_update("n3_2", "running")
        mac_data, part1_data = self._parallel_fetch_verify_data(token)
        self._verify_mac_whitelist(token, pre_fetch_data=mac_data)
        self._check_key_expire(token, pre_fetch_data=part1_data)

        # 缓存分区一数据，正式拉取时复用
        self._cached_partition_one = part1_data

        total_biz = len(self.business_types)
        mem_conn = None
        try:
            # 4. 创建纯内存数据库，应用自适应配置
            mem_conn = duckdb.connect(':memory:', config={
                'memory_limit': self._duckdb_mem_limit,
                'wal_autocheckpoint': '-1',
                'threads': str(self._duckdb_threads),
            })
            mem_conn.execute(f"""
                CREATE TABLE {self.table_name} (
                    area CHAR(1),
                    md5_code CHAR(10),
                    if_mg CHAR(1)
                )
            """)

            success_count = 0
            total_parsed = 0
            batch_buffer = []
            _n6_running_set = False  # 标记节点6是否已开启运行态
            write_lock = threading.RLock()  # 缓冲区与数据库写入互斥锁

            def _flush_buffer():
                nonlocal batch_buffer, total_parsed, _n6_running_set
                if not batch_buffer:
                    return
                with write_lock:
                    # 双重检查，避免加锁前已被其他线程刷空
                    if not batch_buffer:
                        return
                    # 第一次刷写内存库时，开启节点6运行态
                    if not _n6_running_set:
                        self._node_update("n6", "running")
                        _n6_running_set = True
                    self._memory_batch_insert(mem_conn, batch_buffer, self.table_name, columns=["area", "md5_code", "if_mg"])
                    total_parsed += len(batch_buffer)
                    batch_buffer.clear()
                # 内存水位检查放在锁外，不占用锁时间
                self._check_memory_watermark()

            # 5. 优先处理缓存的分区一数据，跳过重复请求
            self._node_update("n5_1", "running")
            self.log(f"复用预拉取的分区一数据，跳过重复请求...")
            part_start = time.time()
            parsed = self._parse_raw_data(self._cached_partition_one)
            part_cost = time.time() - part_start
            if parsed:
                batch_buffer.extend(parsed)
                self.log(f"分区一 解析完成，有效数据 {len(parsed):,} 条")
                self._node_update("n5_1", "success", f"解析完成，{len(parsed):,}条")
            else:
                self.log(f"分区一 无有效数据，跳过")
                self._node_update("n5_1", "success", "无有效数据")
            success_count += 1
            progress = int(success_count / total_biz * 100)
            self.log(f"整体进度：{success_count}/{total_biz}（{progress}%）")

            # 释放原始数据引用，尽早回收内存
            del self._cached_partition_one
            self._cached_partition_one = None
            self._check_memory_watermark()

            # 6. 剩余7个分区动态并发拉取（支持运行时降级调优+异常兜底）
            remaining_biz = self.business_types[1:]
            pending_tasks = list(remaining_biz)  # 待提交任务队列
            self.log(f"开始并行拉取剩余 {len(remaining_biz)} 个业务分区...")

            # 线程池设为最大上限，实际并发由动态提交逻辑控制
            with concurrent.futures.ThreadPoolExecutor(max_workers=_MAX_WORKERS) as executor:
                future_map = {}
                running_count = 0

                # 初始提交第一批任务，数量为当前自适应并发数
                initial_batch = min(self._runtime_fetch_workers, len(pending_tasks))
                for _ in range(initial_batch):
                    bt = pending_tasks.pop(0)
                    part_num = bt.replace(self.busi_prefix, "")
                    node_id = f"n5_{part_num}"
                    self._node_update(node_id, "running")
                    future = executor.submit(self._call_single_business, token, bt)
                    future_map[future] = bt
                    running_count += 1

                try:
                    while future_map:
                        # 等待任意一个任务完成，设置单轮超时避免永久挂起
                        done_futures = concurrent.futures.as_completed(future_map, timeout=_FUTURE_TIMEOUT)
                        try:
                            future = next(done_futures)
                        except StopIteration:
                            break

                        bt = future_map.pop(future)
                        running_count -= 1
                        part_num = bt.replace(self.busi_prefix, "")
                        node_id = f"n5_{part_num}"

                        try:
                            data = future.result(timeout=_FUTURE_TIMEOUT)
                            if not data:
                                raise RuntimeError(f"分区{bt}返回空数据")
                            self.log(f"分区{bt} 获取完成，开始解析...")

                            parsed = self._parse_raw_data(data)
                            # 立即释放原始数据引用
                            del data

                            if not parsed:
                                self.log(f"分区{bt} 无有效数据，跳过")
                                self._node_update(node_id, "success", "无有效数据")
                                success_count += 1
                            else:
                                # 加锁追加数据并判断是否需要刷写
                                with write_lock:
                                    batch_buffer.extend(parsed)
                                    need_flush = len(batch_buffer) >= self._flush_threshold
                                
                                self.log(f"分区{bt} 解析完成，有效数据 {len(parsed):,} 条")
                                self._node_update(node_id, "success", f"解析完成，{len(parsed):,}条")
                                
                                if need_flush:
                                    self.log(f"缓冲区满 {len(batch_buffer):,} 条，开始写入内存库...")
                                    start_time = time.time()
                                    _flush_buffer()
                                    cost = time.time() - start_time
                                    self.log(f"写入完成，耗时 {cost:.1f} 秒")

                                success_count += 1

                            progress = int(success_count / total_biz * 100)
                            self.log(f"整体进度：{success_count}/{total_biz}（{progress}%）")

                            # 内存水位检查放在锁外，可能触发降级调整并发数
                            self._check_memory_watermark()

                            # 尝试提交下一个任务：读取最新的并发限制
                            if pending_tasks and running_count < self._runtime_fetch_workers:
                                next_bt = pending_tasks.pop(0)
                                next_part_num = next_bt.replace(self.busi_prefix, "")
                                next_node_id = f"n5_{next_part_num}"
                                self._node_update(next_node_id, "running")
                                next_future = executor.submit(self._call_single_business, token, next_bt)
                                future_map[next_future] = next_bt
                                running_count += 1

                        except Exception as e:
                            self._node_update(node_id, "error")
                            raise RuntimeError(f"分区{bt} 执行失败: {str(e)}")

                except Exception:
                    # 异常兜底：取消所有未启动的任务，快速释放资源
                    for f in future_map.keys():
                        f.cancel()
                    pending_tasks.clear()
                    raise

            # 刷写最后剩余的缓冲区数据
            if batch_buffer:
                self.log(f"刷写剩余 {len(batch_buffer):,} 条数据...")
                start_time = time.time()
                _flush_buffer()
                cost = time.time() - start_time
                self.log(f"剩余数据写入完成，耗时 {cost:.1f} 秒")

            if success_count != total_biz:
                raise RuntimeError(f"部分分区拉取失败，成功{success_count}/{total_biz}")
            if total_parsed == 0:
                self._node_update("n6", "error", "无有效数据")
                raise RuntimeError("未获取到任何有效数据")
            
            self._node_update("n6", "running")
            
            # 写入完成后整理列式碎片 + 更新统计信息
            try:
                mem_conn.execute(f"ANALYZE {self.table_name}")
                self._node_update("n6", "success", f"共{total_parsed:,}条数据")
            except Exception as e:
                self._node_update("n6", "error", "整理失败")
                raise RuntimeError(f"数据合并整理失败: {str(e)}") from e

            # 7. 建索引加速比对
            # self.log("正在构建内存查询索引...")
            # mem_conn.execute(f"CREATE INDEX idx_mem_md5 ON {self.table_name}(md5_code)")
            self.log(f"全量数据加载完成，共 {total_parsed:,} 条")

            # 8. 待查询号码转码 + 实时比对
            self._node_update("n7", "running")
            try:
                self.log(f"正在处理 {len(mobile_list):,} 条待查询号码...")
                code_pairs = EncryptTools.batch_mobile_to_code(mobile_list)
                invalid_count = len(mobile_list) - len(code_pairs)
                if invalid_count > 0:
                    self.log(f"过滤无效手机号：{invalid_count:,} 条")
                if not code_pairs:
                    raise RuntimeError("无有效11位手机号可查询")

                mem_conn.execute("CREATE TEMP TABLE temp_query (mobile VARCHAR, md5_code VARCHAR)")
                # 显式传入列名，适配PyArrow动态列批量写入，与临时表字段一一对应
                self._memory_batch_insert(mem_conn, code_pairs, "temp_query", columns=["mobile", "md5_code"])

                self.log("正在执行实时匹配...")
                sql = f"""
                    SELECT 
                        t.mobile AS 手机号,
                        COALESCE(
                            CASE d.area
                                WHEN '1' THEN '茂南'
                                WHEN '2' THEN '电白'
                                WHEN '3' THEN '高州'
                                WHEN '4' THEN '化州'
                                WHEN '5' THEN '信宜'
                                ELSE '未匹配'
                            END, '未匹配'
                        ) AS 归属地,
                        COALESCE(d.if_mg, 'N') AS 是否敏感
                    FROM temp_query t
                    LEFT JOIN {self.table_name} d ON t.md5_code = d.md5_code
                """
                result_rows = mem_conn.execute(sql).fetchall()

                # 组装结果
                result_list = []
                for row in result_rows:
                    result_list.append({
                        "手机号": row[0],
                        "归属地": row[1],
                        "是否敏感": row[2]
                    })

                match_count = sum(1 for r in result_list if r["归属地"] != "未匹配")
                self.log(f"比对完成，共命中 {match_count:,} 条")
                self._node_update("n7", "success", f"命中{match_count:,}条")
                return result_list

            except Exception:
                # 匹配阶段异常，同步更新节点状态，避免界面卡住
                self._node_update("n7", "error", "匹配失败")
                raise

        finally:
            # 强制释放内存 + 清空缓存 + 重置运行状态，零残留
            self._cached_partition_one = None
            self._degrade_level = 0
            self._degrade_notified = False
            if mem_conn:
                try:
                    self._node_update("n8", "running")
                    mem_conn.close()
                    self.log("内存数据已释放，无本地残留")
                    self._node_update("n8", "success", "内存数据已释放")
                except Exception:
                    pass

# ===================== UI标签页（支持xlsx，Windows专项适配）=====================
def create_realtime_tab(parent):
    from __main__ import BaseTab, LogComponent, AppTools, FileTools, AppConfig, DuckTools
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class NodeProgressView:
        """树形节点进度图，基于Canvas原生绘制，支持全局缩放适配全平台"""
        def __init__(self, parent, app_config):
            self.app_config = app_config
            self.canvas = tk.Canvas(parent, bg="#FFFFFF", highlightthickness=0)
            self.canvas.pack(fill=tk.BOTH, expand=True)

            # ========== 全局缩放系数（唯一需要调的参数）==========
            # Windows 建议 1.0；macOS 建议 0.75 ~ 0.8
            self.scale_factor = 1

            # 节点定义：id, 名称, 层级(从上到下), 同层序号, 同层总数
            self.nodes = [
                {"id": "n1",   "name": "号码文件读取",   "level": 0, "index": 0, "total": 1},
                {"id": "n2",   "name": "环境分析",       "level": 1, "index": 0, "total": 1},
                {"id": "n4",   "name": "资源自适应配置",       "level": 2, "index": 0, "total": 1},
                {"id": "n3_1", "name": "密钥有效期鉴权", "level": 3, "index": 0, "total": 2},
                {"id": "n3_2", "name": "设备白名单鉴权", "level": 3, "index": 1, "total": 2},
                {"id": "n5_1", "name": "分区1数据",      "level": 4, "index": 0, "total": 8},
                {"id": "n5_2", "name": "分区2数据",      "level": 4, "index": 1, "total": 8},
                {"id": "n5_3", "name": "分区3数据",      "level": 4, "index": 2, "total": 8},
                {"id": "n5_4", "name": "分区4数据",      "level": 4, "index": 3, "total": 8},
                {"id": "n5_5", "name": "分区5数据",      "level": 4, "index": 4, "total": 8},
                {"id": "n5_6", "name": "分区6数据",      "level": 4, "index": 5, "total": 8},
                {"id": "n5_7", "name": "分区7数据",      "level": 4, "index": 6, "total": 8},
                {"id": "n5_8", "name": "分区8数据",      "level": 4, "index": 7, "total": 8},
                {"id": "n6",   "name": "分区持续写入内存库",   "level": 5, "index": 0, "total": 1},
                {"id": "n7",   "name": "匹配目标号码",       "level": 6, "index": 0, "total": 1},
                {"id": "n8",   "name": "释放内存",       "level": 7, "index": 0, "total": 1},
            ]

            # 状态：pending待处理(灰) / running运行中(蓝) / success成功(绿) / error错误(红)
            self.node_state = {n["id"]: "pending" for n in self.nodes}
            self.node_text = {n["id"]: "就绪" for n in self.nodes}

            # ========== 所有尺寸统一乘以缩放系数，等比例缩放 ==========
            self.node_height = int(52 * self.scale_factor)
            self.level_gap = int(28 * self.scale_factor)
            self.partition_row_gap = int(0 * self.scale_factor)
            self.node_width = int(160 * self.scale_factor)

            # 内部细节尺寸，同步缩放
            self._y_start = int(20 * self.scale_factor)
            self._extra_padding = int(20 * self.scale_factor)
            self._text_title_offset = int(16 * self.scale_factor)
            self._text_desc_offset = int(36 * self.scale_factor)
            self._line_width = int(2 * self.scale_factor)
            self._arrow_offset = int(6 * self.scale_factor)
            self._arrow_size = int(10 * self.scale_factor)
            self._result_text_gap = int(40 * self.scale_factor)

            self.result_text = ""  # 底部结果路径文本，空字符串则不显示
            self._draw_all()
            self.canvas.bind("<Configure>", lambda e: self._redraw_all())

        def _get_color(self, state):
            if state == "running":
                return self.app_config.COLOR_INFO
            elif state == "success":
                return self.app_config.COLOR_SUCCESS
            elif state == "error":
                return self.app_config.COLOR_ERROR
            else:  # pending
                return "#838181"

        def _draw_all(self):
            self.canvas.delete("all")
            w = self.canvas.winfo_width()
            if w < 100:
                w = 800
            # 按层级分组
            levels = {}
            for node in self.nodes:
                lv = node["level"]
                if lv not in levels:
                    levels[lv] = []
                levels[lv].append(node)

            self.node_pos = {}
            y_start = self._y_start
            sorted_levels = sorted(levels.keys())
            # 预计算每层顶部Y坐标（累计高度，适配分区层多行）
            level_ys = {}
            current_y = y_start
            for lv in sorted_levels:
                level_ys[lv] = current_y
                if lv == 4:
                    # 分区层：4行垂直排列，计算总高度（使用分区专属行间距）
                    row_count = 4
                    level_height = row_count * self.node_height + (row_count - 1) * self.partition_row_gap + self._extra_padding
                else:
                    level_height = self.node_height
                current_y += level_height + self.level_gap
            self.level_ys = level_ys

            # 逐个层级计算节点位置
            for lv in sorted_levels:
                nodes = levels[lv]
                y_top = level_ys[lv]

                if lv == 4:
                    # 分区层：左右双列布局，每列4个垂直排列
                    col_gap = 0  # 两列之间的间距，可微调
                    total_w = 2 * self.node_width + col_gap
                    # 整体水平居中
                    left_col_x = (w - total_w) / 2
                    right_col_x = left_col_x + self.node_width + col_gap

                    # 左列：分区1-4
                    left_nodes = nodes[:4]
                    for idx, node in enumerate(left_nodes):
                        x1 = left_col_x
                        x2 = x1 + self.node_width
                        y1 = y_top + idx * (self.node_height + self.partition_row_gap)
                        y2 = y1 + self.node_height
                        self.node_pos[node["id"]] = (x1, y1, x2, y2)
                        self._draw_node(node["id"], node["name"])

                    # 右列：分区5-8
                    right_nodes = nodes[4:]
                    for idx, node in enumerate(right_nodes):
                        x1 = right_col_x
                        x2 = x1 + self.node_width
                        y1 = y_top + idx * (self.node_height + self.partition_row_gap)
                        y2 = y1 + self.node_height
                        self.node_pos[node["id"]] = (x1, y1, x2, y2)
                        self._draw_node(node["id"], node["name"])
                else:
                    # 普通层：横向均匀分布
                    total = len(nodes)
                    gap = (w - total * self.node_width) / (total + 1)
                    for idx, node in enumerate(nodes):
                        x1 = gap + idx * (self.node_width + gap)
                        x2 = x1 + self.node_width
                        y1 = y_top
                        y2 = y1 + self.node_height
                        self.node_pos[node["id"]] = (x1, y1, x2, y2)
                        self._draw_node(node["id"], node["name"])

            # 绘制层间连接线
            for lv in sorted_levels:
                if lv + 1 not in levels:
                    continue
                self._draw_layer_connector(lv, lv + 1)

            # 绘制底部结果路径（仅非空时显示，居中纯文字）
            if self.result_text:
                n8_bottom_y = self.node_pos["n8"][3]  # 取释放内存节点底部
                text_y = n8_bottom_y + self._result_text_gap  # 和节点保持间距
                self.canvas.create_text(
                    w / 2, text_y,
                    text=self.result_text,
                    fill="#374151",
                    font=AppTools.get_font(AppConfig.FONT_SIZE_BASE),
                    anchor="center"
                )

        def _draw_layer_connector(self, upper_lv, lower_lv):
            # 计算上下层的边界Y坐标
            if upper_lv == 4:
                u_height = 4 * self.node_height + 3 * self.partition_row_gap
            else:
                u_height = self.node_height
            u_y_bottom = self.level_ys[upper_lv] + u_height
            l_y_top = self.level_ys[lower_lv]

            # ========== 强制优先级：分区层 → 写入内存层，笔直居中直线 ==========
            if upper_lv == 4 and lower_lv == 5:
                # 分区块左右边界，算几何中心
                block_left = self.node_pos["n5_1"][0]
                block_right = self.node_pos["n5_8"][2]
                center_x = (block_left + block_right) / 2
                # 一条垂直线直接连到底，无任何拐弯
                self.canvas.create_line(
                    center_x, u_y_bottom,
                    center_x, l_y_top - self._arrow_offset,
                    fill="#8E8D8D", width=self._line_width
                )
                self._draw_arrow(center_x, l_y_top)
                return

            # 鉴权层 → 分区层：汇合后指向分区块顶部中心
            if upper_lv == 3 and lower_lv == 4:
                mid_y = (u_y_bottom + l_y_top) / 2
                ux_left = (self.node_pos["n3_1"][0] + self.node_pos["n3_1"][2]) / 2
                ux_right = (self.node_pos["n3_2"][0] + self.node_pos["n3_2"][2]) / 2
                center_x = (ux_left + ux_right) / 2

                self.canvas.create_line(ux_left, u_y_bottom, center_x, mid_y, fill="#8E8D8D", width=self._line_width)
                self.canvas.create_line(ux_right, u_y_bottom, center_x, mid_y, fill="#8E8D8D", width=self._line_width)

                block_left = self.node_pos["n5_1"][0]
                block_right = self.node_pos["n5_8"][2]
                block_center_x = (block_left + block_right) / 2

                self.canvas.create_line(center_x, mid_y, block_center_x, mid_y, fill="#8E8D8D", width=self._line_width)
                self.canvas.create_line(block_center_x, mid_y, block_center_x, l_y_top - self._arrow_offset, fill="#8E8D8D", width=self._line_width)
                self._draw_arrow(block_center_x, l_y_top)
                return

            # ========== 以下为普通层通用连线逻辑 ==========
            upper_nodes = [n for n in self.nodes if n["level"] == upper_lv]
            lower_nodes = [n for n in self.nodes if n["level"] == lower_lv]

            if len(upper_nodes) == 1 and len(lower_nodes) == 1:
                ux = (self.node_pos[upper_nodes[0]["id"]][0] + self.node_pos[upper_nodes[0]["id"]][2]) / 2
                lx = (self.node_pos[lower_nodes[0]["id"]][0] + self.node_pos[lower_nodes[0]["id"]][2]) / 2
                self.canvas.create_line(ux, u_y_bottom, lx, l_y_top - self._arrow_offset, fill="#8E8D8D", width=self._line_width)
                self._draw_arrow(lx, l_y_top)
            elif len(upper_nodes) == 1 and len(lower_nodes) > 1:
                ux = (self.node_pos[upper_nodes[0]["id"]][0] + self.node_pos[upper_nodes[0]["id"]][2]) / 2
                mid_y = (u_y_bottom + l_y_top) / 2
                self.canvas.create_line(ux, u_y_bottom, ux, mid_y, fill="#8E8D8D", width=self._line_width)
                for n in lower_nodes:
                    lx = (self.node_pos[n["id"]][0] + self.node_pos[n["id"]][2]) / 2
                    self.canvas.create_line(ux, mid_y, lx, mid_y, fill="#8E8D8D", width=self._line_width)
                    self.canvas.create_line(lx, mid_y, lx, l_y_top - self._arrow_offset, fill="#8E8D8D", width=self._line_width)
                    self._draw_arrow(lx, l_y_top)
            elif len(upper_nodes) > 1 and len(lower_nodes) == 1:
                lx = (self.node_pos[lower_nodes[0]["id"]][0] + self.node_pos[lower_nodes[0]["id"]][2]) / 2
                mid_y = (u_y_bottom + l_y_top) / 2
                self.canvas.create_line(lx, mid_y, lx, l_y_top - self._arrow_offset, fill="#8E8D8D", width=self._line_width)
                self._draw_arrow(lx, l_y_top)
                for n in upper_nodes:
                    ux = (self.node_pos[n["id"]][0] + self.node_pos[n["id"]][2]) / 2
                    self.canvas.create_line(ux, u_y_bottom, ux, mid_y, fill="#8E8D8D", width=self._line_width)
                    self.canvas.create_line(ux, mid_y, lx, mid_y, fill="#8E8D8D", width=self._line_width)

        def _draw_arrow(self, x, y):
            half = self._arrow_size // 2
            self.canvas.create_polygon(
                x, y,
                x - half, y - self._arrow_size,
                x + half, y - self._arrow_size,
                fill="#8E8D8D", outline=""
            )

        def _draw_node(self, nid, name):
            x1, y1, x2, y2 = self.node_pos[nid]
            state = self.node_state[nid]
            color = self._get_color(state)
            # 节点边框
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=self._line_width, fill="#FAFAFA")
            # 节点名称（加粗）
            self.canvas.create_text(
                (x1 + x2) / 2, y1 + self._text_title_offset,
                text=f"【{name}】",
                fill=color,
                font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL, bold=True)
            )
            # 状态说明
            self.canvas.create_text(
                (x1 + x2) / 2, y1 + self._text_desc_offset,
                text=self.node_text[nid],
                fill=color,
                font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL - 1)
            )

        def _redraw_all(self):
            self._draw_all()

        def update_node(self, node_id, state, text="处理中"):
            """外部调用：更新单个节点状态与文本"""
            if node_id not in self.node_state:
                return
            self.node_state[node_id] = state
            self.node_text[node_id] = text
            node_info = next(n for n in self.nodes if n["id"] == node_id)
            self._draw_node(node_id, node_info["name"])

        def set_result_text(self, text: str):
            """设置底部结果路径，传空字符串则隐藏"""
            self.result_text = text
            self._redraw_all()


    class RealtimeTab(BaseTab):
        def __init__(self, parent):
            super().__init__(parent)
            import queue
            self._log_queue = queue.Queue()
            self._node_update_queue = queue.Queue()
            self._log_refresh_running = False
            self._log_cache = []  # 日志缓存，用于文件落盘
            self._init_ui()
            self._start_log_refresh()

            if not is_module_available():
                self._safe_log("警告：缺少授权文件，实时数据功能不可用", AppConfig.COLOR_ERROR)
                AppTools.batch_update_widget_state(self.disabled_widgets, tk.DISABLED)
                return
            # 传入节点更新回调
            self.business = RealtimeBusiness(self._safe_log, node_callback=self._update_node_safe)
            self.query_file_path = None

        def _update_node_safe(self, node_id: str, state: str, text: str = "处理中"):
            """线程安全的节点状态更新，放入队列由主线程刷新"""
            if _ENABLE_NODE_VIEW:
                self._node_update_queue.put((node_id, state, text))

        def _safe_log(self, msg: str, color: str = None):
            """线程安全的日志入口：收集缓存用于落盘，统一入队后按模式分发输出"""
            # 收集日志缓存，用于文件落盘
            if _ENABLE_LOG_FILE:
                self._log_cache.append(msg)
            # 所有日志统一入队，由刷新逻辑按模式分发到对应日志框
            self._log_queue.put((msg, color))

        def _start_log_refresh(self):
            """启动200ms一次的批量刷新，兼顾节点图与日志"""
            if self._log_refresh_running:
                return
            self._log_refresh_running = True

            def _refresh():
                # 刷新节点进度图
                if _ENABLE_NODE_VIEW and hasattr(self, 'node_view'):
                    while not self._node_update_queue.empty():
                        try:
                            nid, state, text = self._node_update_queue.get_nowait()
                            self.node_view.update_node(nid, state, text)
                        except Exception:
                            break

                # 批量取出日志
                msgs = []
                while not self._log_queue.empty():
                    try:
                        msgs.append(self._log_queue.get_nowait())
                    except Exception:
                        break

                if msgs:
                    if _ENABLE_NODE_VIEW:
                        # 节点模式：输出到左侧精简日志框
                        for msg, color in msgs:
                            self.mini_logger.append(msg, color)
                    else:
                        # 原日志模式：输出到右侧完整日志面板
                        for msg, color in msgs:
                            self.logger.append(msg, color)

                self.after(200, _refresh)

            _refresh()

        def _save_log_file(self):
            """任务结束后将缓存的日志落地为txt文件"""
            if not _ENABLE_LOG_FILE or not self._log_cache:
                return
            try:
                time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                file_name = f"日志-{time_str}.txt"
                save_path = os.path.join(_get_app_root_dir(), file_name)
                with open(save_path, "w", encoding=_LOG_FILE_ENCODING) as f:
                    f.write("\n".join(self._log_cache))
            except Exception:
                pass

        def _init_ui(self):
            self.root.report_callback_exception = self._handle_main_thread_exception
            top_container = ttk.Frame(self)
            top_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=20)

            left_area = ttk.Frame(top_container, width=420)
            left_area.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
            left_area.pack_propagate(False)

            # 功能区标题
            ttk.Label(left_area, text="内存实时比对（零文件残留）",
                      font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(pady=(6, 10), fill=tk.X, padx=5)

            # 1. 选择文件
            self.btn_select_file = ttk.Button(left_area, text="选择手机号文件（单列txt文本）", command=self.select_query_file)
            self.btn_select_file.pack(pady=4, anchor=tk.W, padx=5)

            self.label_file_status = ttk.Label(left_area, text="未选择文件", foreground="#666666", anchor="w")
            self.label_file_status.pack(pady=2, anchor=tk.W, padx=8)

            # 2. 线程数选择
            thread_row = ttk.Frame(left_area)
            thread_row.pack(pady=10, fill=tk.X, padx=5)
            ttk.Label(thread_row, text="并发线程数\n（默认自动最优）：", width=15).pack(side=tk.LEFT)
            self.combo_workers = ttk.Combobox(
                thread_row,
                values=["自动（推荐）"] + [str(i) for i in range(1, 9)],
                state="readonly",
                width=12
            )
            self.combo_workers.current(0)  # 默认自动模式
            self.combo_workers.pack(side=tk.LEFT)

            # 3. 开始比对
            self.btn_start_compare = ttk.Button(
                left_area, text="开始实时比对",
                command=self.start_realtime_compare,
                state=tk.DISABLED,
                width=20
            )
            self.btn_start_compare.pack(pady=12, anchor=tk.W, padx=5)

            # 安全说明
            safety_frame = ttk.LabelFrame(left_area, text="数据安全说明")
            safety_frame.pack(fill="x", padx=5, pady=(15, 0))

            safety_text = (
                "⚠️ 郑重提示\n\n"
                "       根据您电脑的性能配置，每次耗时大约5-10分钟，中途退出直接关闭工具即可。\n\n"
                "       全程内存运算，关闭工具即释放内存，无敏感信息残留。\n\n"
                "       仅允许内网环境使用。\n"
                "       仅允许有线物理连接使用。\n"
                "       仅允许被授权的办公电脑使用。\n"
                "       接口密钥存在有效期，到期后自动失效。\n"
            )

            safety_label = tk.Label(
                safety_frame,
                text=safety_text,
                font=AppTools.get_font(AppConfig.FONT_SIZE_SMALL),
                fg="#374151",
                bg="#f9fafb",
                anchor="nw",
                justify="left",
                wraplength=380,
                padx=8,
                pady=6
            )
            safety_label.pack(fill="both", expand=True)

            # 节点模式精简日志框：常驻显示运行日志与错误信息
            mini_log_frame = ttk.LabelFrame(left_area, text="运行提示")
            mini_log_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))
            log_inner = tk.Frame(mini_log_frame, bg="#F0F0F0", height=130)
            log_inner.pack(fill="both", expand=True, padx=2, pady=2)
            log_inner.pack_propagate(False)
            self.mini_logger = LogComponent(log_inner, AppConfig.LOG_MATCH_LIMIT_REAL, self)

            # 右侧展示区：根据开关切换节点进度图 / 原日志
            right_area = ttk.Frame(top_container)
            right_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

            title_text = "运行进度" if _ENABLE_NODE_VIEW else "运行日志"
            ttk.Label(right_area, text=title_text,
                      font=AppTools.get_font(AppConfig.FONT_SIZE_BOLD, bold=True)).pack(anchor=tk.NW, pady=(0, 5), fill=tk.X)

            # 先初始化日志组件，保证基类调用永远有效（节点模式下隐藏不显示）
            log_frame = tk.Frame(right_area, bg="#F0F0F0", height=580)
            self.logger = LogComponent(log_frame, AppConfig.LOG_MATCH_LIMIT_REAL, self)

            if _ENABLE_NODE_VIEW:
                # 节点进度图模式：显示节点图，隐藏日志框
                view_frame = tk.Frame(right_area, bg="#FFFFFF", height=580)
                view_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                view_frame.pack_propagate(False)
                self.node_view = NodeProgressView(view_frame, AppConfig)
            else:
                # 原日志模式：显示日志框
                log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                log_frame.pack_propagate(False)

            self.disabled_widgets = [
                self.btn_select_file, self.btn_start_compare, self.combo_workers
            ]

        def select_query_file(self):
            path = filedialog.askopenfilename(
                filetypes=[
                    ("文本文件", "*.txt;*.csv"),
                    ("所有文件", "*.*")
                ]
            )
            if not path:
                return
            is_valid, msg = FileTools.check_file_valid(path)
            if not is_valid:
                messagebox.showerror("文件错误", msg)
                return
            self.query_file_path = path
            self.label_file_status.config(
                text=f"已选择：{os.path.basename(path)}",
                foreground="#166534"
            )
            self.btn_start_compare.config(state=tk.NORMAL)
            self._safe_log(f"已选择比对文件：{FileTools.get_file_full_path(path)}")

        def start_realtime_compare(self):
            # 新增：节点进入运行状态
            self._update_node_safe("n1", "running")
            if not self.query_file_path:
                messagebox.showwarning("提示", "请先选择手机号文件")
                return

            import re

            try:
                # 仅保留 txt/csv 读取逻辑
                enc = FileTools.detect_file_encoding(self.query_file_path)
                rel = DuckTools.read_single_column(self.query_file_path, enc)
                all_rows = rel.native_rel.fetchall()
                raw_mobile_list = [str(row[0]).strip() for row in all_rows if str(row[0]).strip()]

                # 统一清洗：提取纯数字，过滤空格、横杠、全角等所有非数字字符
                mobile_list = []
                for val in raw_mobile_list:
                    digits = re.sub(r"\D", "", str(val))
                    if len(digits) == 11 and digits.startswith("1"):
                        mobile_list.append(digits)

                # 导入即输出统计日志，提前校验避免无效拉取
                self._safe_log(f"文件读取完成，原始数据共 {len(raw_mobile_list):,} 行，有效11位手机号 {len(mobile_list):,} 条")
                # 更新节点1：号码文件读取
                self._update_node_safe("n1", "success", f"有效号码{len(mobile_list):,}条")

            except Exception as e:
                messagebox.showerror("读取失败", f"文件读取失败：{str(e)}")
                return

            if not mobile_list:
                messagebox.showwarning("提示", "文件中未识别到有效11位手机号，请检查文件格式")
                self._safe_log("未识别到有效11位手机号，任务已终止", AppConfig.COLOR_ERROR)
                return

            # 获取线程数：自动模式传None，手动模式传具体数值
            selected = self.combo_workers.get()
            if selected == "自动（推荐）":
                workers = None
            else:
                try:
                    workers = int(selected)
                except:
                    workers = None

            # 后台执行
            def _task():
                try:
                    result_list = self.business.realtime_compare(mobile_list, fetch_workers=workers)

                    # 询问保存结果
                    save_path = FileTools.get_save_file_path("实时比对结果")
                    if save_path:
                        with open(save_path, "w", encoding="utf-8-sig") as f:
                            f.write("手机号|归属地|是否敏感\n")
                            for item in result_list:
                                f.write(f"{item['手机号']}|{item['归属地']}|{item['是否敏感']}\n")
                        # 更新节点7：补上保存路径
                        match_count = sum(1 for r in result_list if r['归属地']!='未匹配')
                        self._update_node_safe("n7", "success", f"命中{match_count:,}条")
                        # 新增：显示底部结果路径
                        self.after(0, lambda: self.node_view.set_result_text(f"\n处理结果：\n{save_path}"))
                        
                        # 保存完成后清理当前目录临时文件
                        import glob
                        for tmp_file in glob.glob("*.tmp") + glob.glob("*.duckdb.*"):
                            try:
                                os.remove(tmp_file)
                            except Exception:
                                pass
                        return f"比对完成！命中 {sum(1 for r in result_list if r['归属地']!='未匹配'):,} 条\n结果已保存至：{save_path}"
                    else:
                        # 未保存也清理临时文件
                        import glob
                        for tmp_file in glob.glob("*.tmp") + glob.glob("*.duckdb.*"):
                            try:
                                os.remove(tmp_file)
                            except Exception:
                                pass
                        return f"比对完成！命中 {sum(1 for r in result_list if r['归属地']!='未匹配'):,} 条\n未保存结果文件"
                
                finally:
                    # 无论成功失败，统一落地日志
                    self._save_log_file()

            # 重置所有节点为初始待执行状态：灰色 + 就绪
            if _ENABLE_NODE_VIEW and hasattr(self, "node_view"):
                for nid in self.node_view.node_state.keys():
                    self.node_view.update_node(nid, "pending", "就绪")
                    # 新增：清空上一次结果路径
                self.node_view.set_result_text("")

            self.run_background_task(_task, "开始执行全内存实时比对，根据主机性能，全程大约5-10分钟，请耐心等待...")

        def _handle_main_thread_exception(self, exc_type, exc_value, exc_traceback):
            error_msg = f"主线程异常：{exc_type.__name__}: {exc_value}"
            self._safe_log(error_msg, AppConfig.COLOR_ERROR)
            messagebox.showerror("程序错误", f"发生未捕获的错误：\n{exc_value}\n请查看日志详情")

    return RealtimeTab(parent)