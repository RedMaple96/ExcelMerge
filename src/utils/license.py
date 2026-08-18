"""ExcelMerge 在线授权模块。

授权流程与独立 activation-server 配合：
1. 客户端生成稳定设备指纹；
2. 用户可自助申请 30 天授权码；
3. 在线激活后服务端返回设备绑定 token；
4. 客户端保存签名后的授权记录并定期在线校验；
5. 网络不可用时允许 7 天离线宽限期。
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import platform
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid
from datetime import datetime
from typing import Any, Dict, Optional

from src import __version__ as APP_VERSION

DEFAULT_ACTIVATION_SERVER_URL = "http://49.232.103.38:6005"
ACTIVATION_SERVER_URL = os.environ.get(
    "EXCEL_MERGE_ACTIVATION_SERVER", DEFAULT_ACTIVATION_SERVER_URL
).rstrip("/")

_APP_DIR_NAME = "ExcelMerge"
_LICENSE_FILE = "license.json"
_LAST_KNOWN_TIME_FILE = "last-known-time"
_REQUEST_TIMEOUT_SECONDS = 12
_GRACE_PERIOD_SECONDS = 7 * 24 * 60 * 60
_TIME_ROLLBACK_TOLERANCE_SECONDS = 60

# 本地文件防普通篡改。桌面客户端无法保存绝对安全的私密密钥；
# 授权有效性最终仍由远程服务端控制。
_SECRET_PARTS = ("ExcelMerge", "License", "v2", "2026")
_LICENSE_SECRET = "-".join(_SECRET_PARTS).encode("utf-8")


class ActivationError(Exception):
    """远程激活服务返回的可展示错误。"""

    def __init__(self, message: str, code: str = "SERVER_ERROR") -> None:
        super().__init__(message)
        self.code = code


class NetworkError(Exception):
    """无法连接激活服务。"""


def _raw_machine_id() -> str:
    """获取操作系统级稳定机器标识。"""
    system = platform.system()
    try:
        if system == "Darwin":
            output = subprocess.run(
                ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            for line in output.stdout.splitlines():
                if "IOPlatformUUID" in line:
                    return line.split("=", 1)[1].strip().strip('"')
        elif system == "Windows":
            import winreg  # type: ignore

            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography",
            ) as key:
                value, _ = winreg.QueryValueEx(key, "MachineGuid")
                return str(value)
        elif system == "Linux":
            for path in ("/etc/machine-id", "/var/lib/dbus/machine-id"):
                try:
                    with open(path, "r", encoding="utf-8") as file:
                        content = file.read().strip()
                        if content:
                            return content
                except OSError:
                    continue
    except Exception:  # noqa: BLE001
        pass
    return f"uuid-{uuid.getnode()}"


def get_machine_fingerprint() -> str:
    """返回 32 位十六进制设备指纹。"""
    seed = f"{_raw_machine_id()}::{uuid.getnode()}"
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()[:32]


def get_machine_label() -> str:
    return f"{platform.system()}/{socket.gethostname()}"


def _memory_gb() -> int:
    try:
        if platform.system() == "Darwin":
            output = subprocess.run(
                ["sysctl", "-n", "hw.memsize"],
                capture_output=True,
                text=True,
                timeout=3,
            )
            return round(int(output.stdout.strip()) / 1024**3)
        if platform.system() == "Linux":
            with open("/proc/meminfo", "r", encoding="utf-8") as file:
                for line in file:
                    if line.startswith("MemTotal:"):
                        return round(int(line.split()[1]) / 1024**2)
        if platform.system() == "Windows":
            import ctypes

            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                ]

            status = MEMORYSTATUSEX()
            status.dwLength = ctypes.sizeof(status)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                return round(status.ullTotalPhys / 1024**3)
    except Exception:  # noqa: BLE001
        pass
    return 0


def get_device_info() -> Dict[str, Any]:
    """返回服务端后台展示所需的脱敏设备信息。"""
    return {
        "platform": platform.system(),
        "osVersion": platform.version(),
        "hostname": socket.gethostname(),
        "arch": platform.machine(),
        "cpuCores": os.cpu_count() or 0,
        "memoryGB": _memory_gb(),
        "appVersion": APP_VERSION,
        "pythonVersion": platform.python_version(),
    }


def _config_dir() -> str:
    if sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    elif os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, _APP_DIR_NAME)


def _license_path() -> str:
    return os.path.join(_config_dir(), _LICENSE_FILE)


def _last_known_time_path() -> str:
    return os.path.join(_config_dir(), _LAST_KNOWN_TIME_FILE)


def _canonical_record(record: Dict[str, Any]) -> str:
    return json.dumps(record, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _sign_record(record: Dict[str, Any]) -> str:
    return hmac.new(
        _LICENSE_SECRET,
        _canonical_record(record).encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()


def _atomic_write(path: str, content: str) -> None:
    directory = os.path.dirname(path)
    os.makedirs(directory, exist_ok=True)
    temporary_path = f"{path}.tmp"
    with open(temporary_path, "w", encoding="utf-8") as file:
        file.write(content)
        file.flush()
        os.fsync(file.fileno())
    os.chmod(temporary_path, 0o600)
    os.replace(temporary_path, path)


def _write_last_known_time(timestamp: int) -> None:
    try:
        _atomic_write(_last_known_time_path(), str(timestamp))
    except OSError:
        pass


def _read_last_known_time() -> Optional[int]:
    try:
        with open(_last_known_time_path(), "r", encoding="utf-8") as file:
            value = int(file.read().strip())
        return value if value > 0 else None
    except (OSError, ValueError):
        return None


def _normalize_expires_at(value: Any) -> Optional[int]:
    """兼容服务端 JavaScript Date.now() 的毫秒时间戳。"""
    if value is None:
        return None
    timestamp = int(value)
    if timestamp > 10_000_000_000:
        timestamp = timestamp // 1000
    return timestamp


def _write_license(record: Dict[str, Any]) -> bool:
    try:
        signed = dict(record)
        signed["signature"] = _sign_record(record)
        _atomic_write(
            _license_path(), json.dumps(signed, ensure_ascii=False, indent=2)
        )
        _write_last_known_time(int(time.time()))
        return True
    except OSError:
        return False


def read_license() -> Optional[Dict[str, Any]]:
    """读取并验签本地授权记录，旧 license.dat 永远不会被读取。"""
    try:
        with open(_license_path(), "r", encoding="utf-8") as file:
            signed = json.load(file)
        if not isinstance(signed, dict):
            return None
        signature = signed.get("signature")
        if not isinstance(signature, str):
            return None
        record = dict(signed)
        record.pop("signature", None)
        expected = _sign_record(record)
        if not hmac.compare_digest(signature, expected):
            return None
        return record
    except (OSError, ValueError):
        return None


def _post_json(path: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        f"{ACTIVATION_SERVER_URL}{path}",
        data=body,
        headers={
            "Content-Type": "application/json",
            "User-Agent": f"ExcelMerge/{APP_VERSION}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=_REQUEST_TIMEOUT_SECONDS) as response:
            raw = response.read(1024 * 1024)
    except urllib.error.HTTPError as exc:
        try:
            raw = exc.read(1024 * 1024)
            data = json.loads(raw.decode("utf-8"))
        except Exception:  # noqa: BLE001
            data = {}
        if exc.code >= 500:
            raise NetworkError(f"激活服务错误({exc.code})") from exc
        raise ActivationError(
            str(data.get("error") or f"激活服务错误({exc.code})"),
            str(data.get("code") or f"HTTP_{exc.code}"),
        ) from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise NetworkError(f"无法连接激活服务: {exc}") from exc

    try:
        data = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, ValueError) as exc:
        raise NetworkError("激活服务返回格式错误") from exc
    if not isinstance(data, dict):
        raise NetworkError("激活服务返回格式错误")
    return data


def request_code() -> Dict[str, Any]:
    """为当前设备自助申请 30 天授权码。"""
    data = _post_json(
        "/api/request-code",
        {
            "machineFingerprint": get_machine_fingerprint(),
            "machineLabel": get_machine_label(),
            "deviceInfo": get_device_info(),
        },
    )
    if data.get("ok") and data.get("licenseCode"):
        return {"ok": True, "licenseCode": data["licenseCode"]}
    raise ActivationError(
        str(data.get("error") or "获取授权码失败"),
        str(data.get("code") or "REQUEST_CODE_FAILED"),
    )


def activate(license_code: str) -> Dict[str, Any]:
    """在线激活当前设备并保存服务端 token。"""
    code = license_code.strip()
    if not code:
        raise ActivationError("请输入授权码", "EMPTY_CODE")
    fingerprint = get_machine_fingerprint()
    data = _post_json(
        "/api/activate",
        {
            "licenseCode": code,
            "machineFingerprint": fingerprint,
            "machineLabel": get_machine_label(),
            "deviceInfo": get_device_info(),
            "appVersion": APP_VERSION,
        },
    )
    token = data.get("token")
    if not data.get("ok") or not isinstance(token, str) or not token:
        raise ActivationError(
            str(data.get("error") or "激活失败"),
            str(data.get("code") or "ACTIVATE_FAILED"),
        )
    now = int(time.time())
    record = {
        "licenseCode": code,
        "machineFingerprint": fingerprint,
        "token": token,
        "activatedAt": now,
        "expiresAt": _normalize_expires_at(data.get("expiresAt")),
        "lastVerifiedAt": now,
    }
    if not _write_license(record):
        raise ActivationError("无法写入本地授权文件，请检查文件权限", "WRITE_FAILED")
    return {"ok": True, "token": token, "expiresAt": data.get("expiresAt")}


def _offline_check(record: Dict[str, Any]) -> Dict[str, Any]:
    now = int(time.time())
    last_known = _read_last_known_time()
    if last_known is not None and now < last_known - _TIME_ROLLBACK_TOLERANCE_SECONDS:
        return {"ok": False, "reason": "grace-expired"}
    expires_at = record.get("expiresAt")
    if expires_at is not None and int(expires_at) < now:
        return {"ok": False, "reason": "expired"}
    if now - int(record["lastVerifiedAt"]) < _GRACE_PERIOD_SECONDS:
        return {"ok": True, "reason": "offline-grace"}
    return {"ok": False, "reason": "grace-expired"}


def check_license() -> Dict[str, Any]:
    """校验本地授权并联系远程服务确认撤销、过期与设备绑定状态。"""
    record = read_license()
    if record is None:
        return {"ok": False, "reason": "no-license"}
    fingerprint = get_machine_fingerprint()
    if record.get("machineFingerprint") != fingerprint:
        return {"ok": False, "reason": "machine-mismatch"}
    try:
        data = _post_json(
            "/api/verify",
            {
                "token": record.get("token"),
                "machineFingerprint": fingerprint,
                "appVersion": APP_VERSION,
                "deviceInfo": get_device_info(),
            },
        )
    except NetworkError:
        return _offline_check(record)
    except ActivationError as exc:
        if exc.code == "RATE_LIMITED":
            return _offline_check(record)
        return {"ok": False, "reason": "network-error", "detail": str(exc)}

    if data.get("valid"):
        record["lastVerifiedAt"] = int(time.time())
        if "expiresAt" in data:
            record["expiresAt"] = _normalize_expires_at(data.get("expiresAt"))
        if not _write_license(record):
            return {"ok": False, "reason": "write-failed"}
        return {"ok": True, "reason": "verified"}

    reason = str(data.get("reason") or "revoked")
    if reason == "LICENSE_EXPIRED":
        return {"ok": False, "reason": "expired", "detail": reason}
    if reason == "VERSION_TOO_LOW":
        return {
            "ok": False,
            "reason": "version-too-low",
            "detail": str(data.get("message") or reason),
        }
    return {"ok": False, "reason": "revoked", "detail": reason}


def is_activated() -> bool:
    """仅检查本地记录是否可在离线宽限期内使用。"""
    record = read_license()
    if record is None:
        return False
    if record.get("machineFingerprint") != get_machine_fingerprint():
        return False
    return _offline_check(record).get("ok", False)


def clear_local_activation() -> bool:
    """清除本机授权缓存；服务端设备席位需在管理后台释放。"""
    try:
        os.remove(_license_path())
        os.remove(_last_known_time_path())
    except FileNotFoundError:
        return True
    except OSError:
        return False
    return True


def activation_info() -> Optional[Dict[str, Any]]:
    """返回本地授权摘要。"""
    record = read_license()
    if record is None:
        return None
    result = dict(record)
    result.pop("token", None)
    result["activatedAtText"] = datetime.fromtimestamp(
        int(record["activatedAt"])
    ).isoformat(timespec="seconds")
    result["lastVerifiedAtText"] = datetime.fromtimestamp(
        int(record["lastVerifiedAt"])
    ).isoformat(timespec="seconds")
    result["expiresAtText"] = (
        datetime.fromtimestamp(
            _normalize_expires_at(record.get("expiresAt"))
        ).isoformat(timespec="seconds")
        if record.get("expiresAt") is not None
        else "永久"
    )
    result["serverUrl"] = ACTIVATION_SERVER_URL
    return result
