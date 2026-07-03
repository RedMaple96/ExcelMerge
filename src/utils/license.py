"""软件许可证（密匙）模块 — 一机一码。

提供机器码生成、密匙生成/校验、机器绑定与激活状态持久化。

一机一码流程：
1. 用户启动软件，未激活时对话框显示本机「机器码」(XXXX-XXXX-XXXX)。
2. 用户将机器码发给开发者。
3. 开发者运行 ``python -m src.utils.license <机器码>`` 生成该机器专属密匙。
4. 用户输入密匙，校验通过且机器码匹配本机后激活。

密匙格式: ``XXXX-XXXX-XXXX-XXXX-XXXX``（5 组 4 位十六进制，不区分大小写）。
  - 前 3 组：目标机器的机器码。
  - 后 2 组：对机器码的 HMAC-SHA256 截断签名（8 位十六进制）。
校验方式: 离线自校验，无需联网。密匙内嵌机器码，无法在其他机器使用。
激活状态: 写入用户配置目录下的 ``license.dat``，绑定本机 machine_id，
         复制激活文件到其他机器无效。

开发者可通过 ``python -m src.utils.license <机器码>`` 生成专属密匙。
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import platform
import subprocess
import sys
import uuid
from datetime import datetime
from typing import Optional

# 内嵌签名密钥（用于离线校验，不要外泄）
_SECRET = b"ExcelMerge-License-v1-7f3a9c2e"

_APP_DIR_NAME = "ExcelMerge"
_LICENSE_FILE = "license.dat"

_HEXCHARS = set("0123456789ABCDEF")


# ---------------------------------------------------------------------- #
# 辅助
# ---------------------------------------------------------------------- #
def _normalize_hex(text: str) -> str:
    """提取文本中的十六进制字符并大写（去除连字符/空格等分隔符）。"""
    return "".join(c for c in text.upper() if c in _HEXCHARS)


def _sign(payload: str) -> str:
    """对载荷取 HMAC-SHA256 的前 8 位十六进制作为签名（大写）。"""
    digest = hmac.new(
        _SECRET, payload.encode("utf-8"), hashlib.sha256
    ).hexdigest()
    return digest[:8].upper()


# ---------------------------------------------------------------------- #
# 机器标识
# ---------------------------------------------------------------------- #
def _raw_machine_id() -> str:
    """获取操作系统级别的稳定机器标识（跨平台）。"""
    sysname = platform.system()
    try:
        if sysname == "Darwin":
            out = subprocess.run(
                ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            for line in out.stdout.splitlines():
                if "IOPlatformUUID" in line:
                    return line.split("=", 1)[1].strip().strip('"')
        elif sysname == "Windows":
            import winreg  # type: ignore

            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography",
            ) as regkey:
                guid, _ = winreg.QueryValueEx(regkey, "MachineGuid")
                return str(guid)
        elif sysname == "Linux":
            for path in ("/etc/machine-id", "/var/lib/dbus/machine-id"):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        content = f.read().strip()
                        if content:
                            return content
                except OSError:
                    continue
    except Exception:  # noqa: BLE001
        pass
    # 兜底：uuid.getnode()（通常为 MAC 地址）
    return f"uuid-{uuid.getnode()}"


def get_machine_id() -> str:
    """返回当前机器的稳定哈希标识（32 位十六进制，内部防篡改用）。"""
    raw = _raw_machine_id()
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]


def get_machine_code() -> str:
    """返回当前机器的机器码（XXXX-XXXX-XXXX，12 位十六进制，对外展示）。

    机器码是本机硬件标识的哈希截断，用户可将其发送给开发者以换取专属密匙。
    """
    raw = _raw_machine_id()
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()
    code = digest[:12].upper()
    return f"{code[0:4]}-{code[4:8]}-{code[8:12]}"


# ---------------------------------------------------------------------- #
# 密匙生成与校验
# ---------------------------------------------------------------------- #
def generate_key(machine_code: str) -> str:
    """为指定机器码生成专属密匙（一机一码）。

    machine_code 可含连字符/空格，本函数会归一化。
    返回 ``XXXX-XXXX-XXXX-XXXX-XXXX``：前 3 组为机器码，后 2 组为签名。
    """
    mc = _normalize_hex(machine_code)
    if len(mc) != 12:
        raise ValueError("机器码必须为 12 位十六进制（格式 XXXX-XXXX-XXXX）")
    sig = _sign(mc)
    return f"{mc[0:4]}-{mc[4:8]}-{mc[8:12]}-{sig[0:4]}-{sig[4:8]}"


def validate_key(key: str) -> bool:
    """校验密匙格式与签名（不校验机器绑定，绑定请用 key_matches_machine）。"""
    if not key:
        return False
    hexstr = _normalize_hex(key)
    if len(hexstr) != 20:
        return False
    mc = hexstr[:12]
    sig = hexstr[12:]
    expected = _sign(mc)
    return hmac.compare_digest(expected, sig)


def key_matches_machine(key: str, machine_code: str) -> bool:
    """判断密匙是否绑定指定机器码（密匙内嵌的机器码须与之一致）。"""
    if not validate_key(key):
        return False
    key_mc = _normalize_hex(key)[:12]
    target_mc = _normalize_hex(machine_code)
    if len(target_mc) != 12:
        return False
    return key_mc == target_mc


# ---------------------------------------------------------------------- #
# 激活状态持久化
# ---------------------------------------------------------------------- #
def _config_dir() -> str:
    """返回跨平台的用户配置目录。"""
    if sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    elif os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser(
            "~/.config"
        )
    return os.path.join(base, _APP_DIR_NAME)


def _license_path() -> str:
    return os.path.join(_config_dir(), _LICENSE_FILE)


def is_activated() -> bool:
    """检查本机是否已用合法密匙激活（且密匙绑定当前机器）。"""
    path = _license_path()
    if not os.path.exists(path):
        return False
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return False
    machine_id = data.get("machine_id")
    key = data.get("key")
    if not machine_id or not key:
        return False
    # 防复制：激活文件中的 machine_id 须匹配本机
    if machine_id != get_machine_id():
        return False
    if not validate_key(key):
        return False
    # 一机一码：密匙内嵌的机器码须匹配本机
    return key_matches_machine(key, get_machine_code())


def activate(key: str) -> bool:
    """用密匙激活本机。成功返回 True 并持久化，失败返回 False。

    密匙必须合法且绑定当前机器（一机一码），否则拒绝激活。
    """
    if not validate_key(key):
        return False
    if not key_matches_machine(key, get_machine_code()):
        return False
    config_dir = _config_dir()
    try:
        os.makedirs(config_dir, exist_ok=True)
    except OSError:
        return False
    data = {
        "machine_id": get_machine_id(),
        "machine_code": get_machine_code(),
        "key": _normalize_hex(key),
        "activated_at": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        with open(_license_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        return False
    return True


def deactivate() -> bool:
    """移除本机激活状态（删除激活文件）。返回是否已不存在激活文件。"""
    path = _license_path()
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            return False
    return True


def activation_info() -> Optional[dict]:
    """返回激活信息字典（未激活或读取失败返回 None）。

    仅读取文件内容，不校验有效性；如需校验请配合 is_activated()。
    """
    path = _license_path()
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def _format_key(hex_str: str, groups: int = 5) -> str:
    """将连续十六进制串按 4 位分组并用连字符连接（默认 5 组）。"""
    parts = [hex_str[i : i + 4] for i in range(0, len(hex_str), 4)]
    return "-".join(parts[:groups])


if __name__ == "__main__":
    if len(sys.argv) > 1:
        try:
            print(generate_key(sys.argv[1]))
        except ValueError as exc:
            print(f"错误: {exc}", file=sys.stderr)
            sys.exit(1)
    else:
        print("用法: python -m src.utils.license <机器码>")
        print(f"本机机器码: {get_machine_code()}")
        sys.exit(1)
