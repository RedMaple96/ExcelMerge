"""license 模块单元测试 — 一机一码。

覆盖：机器码生成、密匙生成/校验往返、非法密匙拒绝、大小写不敏感、
机器绑定（密匙内嵌机器码）、跨机器密匙失效、激活持久化、解除激活。
"""

import json
import os
import shutil
import tempfile
import unittest
from unittest import mock

from src.utils import license


class MachineCodeTest(unittest.TestCase):
    """机器码。"""

    def test_machine_code_format(self):
        code = license.get_machine_code()
        parts = code.split("-")
        self.assertEqual(len(parts), 3)
        for p in parts:
            self.assertEqual(len(p), 4)
            self.assertTrue(all(c in "0123456789ABCDEF" for c in p))

    def test_machine_code_stable(self):
        """同一进程内多次调用应返回相同机器码。"""
        self.assertEqual(license.get_machine_code(), license.get_machine_code())

    def test_machine_code_derived_from_raw(self):
        """机器码应随底层硬件标识变化。"""
        with mock.patch.object(
            license, "_raw_machine_id", return_value="hardware-A"
        ):
            code_a = license.get_machine_code()
        with mock.patch.object(
            license, "_raw_machine_id", return_value="hardware-B"
        ):
            code_b = license.get_machine_code()
        self.assertNotEqual(code_a, code_b)


class LicenseKeyTest(unittest.TestCase):
    """密匙生成与校验。"""

    def setUp(self):
        self.mc = license.get_machine_code()

    def test_generate_key_is_valid(self):
        for _ in range(100):
            key = license.generate_key(self.mc)
            self.assertTrue(license.validate_key(key), f"生成的密匙应合法: {key}")

    def test_generate_key_format(self):
        key = license.generate_key(self.mc)
        parts = key.split("-")
        self.assertEqual(len(parts), 5)
        for p in parts:
            self.assertEqual(len(p), 4)
            self.assertTrue(all(c in "0123456789ABCDEF" for c in p))

    def test_generate_key_embeds_machine_code(self):
        """密匙前 3 组应等于机器码。"""
        key = license.generate_key(self.mc)
        parts = key.split("-")
        self.assertEqual("-".join(parts[:3]), self.mc)

    def test_case_insensitive(self):
        key = license.generate_key(self.mc)
        self.assertTrue(license.validate_key(key.lower()))
        self.assertTrue(license.validate_key(key.upper()))

    def test_whitespace_tolerant(self):
        key = license.generate_key(self.mc)
        self.assertTrue(license.validate_key(f"  {key}  "))

    def test_reject_empty(self):
        self.assertFalse(license.validate_key(""))

    def test_reject_wrong_format(self):
        self.assertFalse(license.validate_key("ABCD-EFGH-IJKL-MNOP-QRST"))
        self.assertFalse(license.validate_key("1234-5678-90AB-CDEF"))
        self.assertFalse(
            license.validate_key("1234-5678-90AB-CDEF-1234-5678")
        )
        self.assertFalse(license.validate_key("not-a-key-at-all"))

    def test_reject_tampered_signature(self):
        """篡改签名段应校验失败。"""
        key = license.generate_key(self.mc)
        parts = key.split("-")
        first = parts[4][0]
        parts[4] = ("0" if first != "0" else "1") + parts[4][1:]
        self.assertFalse(license.validate_key("-".join(parts)))

    def test_reject_tampered_machine_code(self):
        """篡改机器码段应校验失败（签名不再匹配）。"""
        key = license.generate_key(self.mc)
        parts = key.split("-")
        first = parts[0][0]
        parts[0] = ("0" if first != "0" else "1") + parts[0][1:]
        self.assertFalse(license.validate_key("-".join(parts)))

    def test_generate_key_rejects_bad_machine_code(self):
        with self.assertRaises(ValueError):
            license.generate_key("ABCD-EFGH")  # 不足 12 位
        with self.assertRaises(ValueError):
            license.generate_key("ZZZZ-ZZZZ-ZZZZ")  # 非十六进制


class MachineBindingTest(unittest.TestCase):
    """一机一码：密匙与机器绑定。"""

    def setUp(self):
        self.mc = license.get_machine_code()

    def test_key_matches_its_machine(self):
        key = license.generate_key(self.mc)
        self.assertTrue(license.key_matches_machine(key, self.mc))

    def test_key_does_not_match_other_machine(self):
        """为机器 A 生成的密匙在机器 B 上应不匹配。"""
        key_a = license.generate_key(self.mc)
        other_mc = "0000-0000-0000"
        self.assertNotEqual(other_mc, self.mc)
        self.assertFalse(license.key_matches_machine(key_a, other_mc))

    def test_key_matches_machine_ignores_separators(self):
        """机器码带不带连字符都应匹配。"""
        key = license.generate_key(self.mc)
        mc_no_sep = self.mc.replace("-", "")
        self.assertTrue(license.key_matches_machine(key, mc_no_sep))


class LicenseActivationTest(unittest.TestCase):
    """激活状态持久化与机器绑定。"""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.addCleanup(self._cleanup)
        # 将配置目录重定向到临时目录，避免污染真实用户配置
        self._patch_config = mock.patch.object(
            license, "_config_dir", return_value=self.tmpdir
        )
        self._patch_config.start()
        self.addCleanup(self._patch_config.stop)
        self.mc = license.get_machine_code()

    def _cleanup(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_not_activated_initially(self):
        self.assertFalse(license.is_activated())

    def test_activate_round_trip(self):
        key = license.generate_key(self.mc)
        self.assertTrue(license.activate(key))
        self.assertTrue(license.is_activated())
        info = license.activation_info()
        self.assertIsNotNone(info)
        self.assertEqual(info["key"], key.replace("-", ""))
        self.assertEqual(info["machine_id"], license.get_machine_id())
        self.assertEqual(info["machine_code"], self.mc)
        self.assertIn("activated_at", info)

    def test_activate_invalid_key_fails(self):
        self.assertFalse(license.activate("ABCD-EFGH-IJKL-MNOP-QRST"))
        self.assertFalse(license.is_activated())

    def test_activate_other_machine_key_fails(self):
        """其他机器的合法密匙无法在本机激活（一机一码核心）。"""
        other_mc = "0000-0000-0000"
        self.assertNotEqual(other_mc, self.mc)
        other_key = license.generate_key(other_mc)
        # 密匙本身合法
        self.assertTrue(license.validate_key(other_key))
        # 但不绑定本机，激活被拒
        self.assertFalse(license.activate(other_key))
        self.assertFalse(license.is_activated())

    def test_deactivate(self):
        key = license.generate_key(self.mc)
        license.activate(key)
        self.assertTrue(license.deactivate())
        self.assertFalse(license.is_activated())
        self.assertIsNone(license.activation_info())

    def test_deactivate_when_not_activated(self):
        self.assertTrue(license.deactivate())

    def test_machine_binding_via_machine_id(self):
        """激活文件复制到其他机器（machine_id 不同）应失效。"""
        key = license.generate_key(self.mc)
        license.activate(key)
        with mock.patch.object(
            license, "get_machine_id", return_value="another-machine-id"
        ):
            self.assertFalse(license.is_activated())

    def test_machine_binding_via_machine_code(self):
        """即便 machine_id 相同，密匙机器码不匹配也应失效。"""
        key = license.generate_key(self.mc)
        license.activate(key)
        with mock.patch.object(
            license, "get_machine_code", return_value="1111-2222-3333"
        ):
            self.assertFalse(license.is_activated())

    def test_corrupted_license_file(self):
        """损坏的激活文件应判定为未激活。"""
        os.makedirs(self.tmpdir, exist_ok=True)
        with open(license._license_path(), "w", encoding="utf-8") as f:
            f.write("{ not valid json")
        self.assertFalse(license.is_activated())
        self.assertIsNone(license.activation_info())

    def test_tampered_key_in_file(self):
        """激活文件中的密匙被篡改为非法值后应失效。"""
        key = license.generate_key(self.mc)
        license.activate(key)
        path = license._license_path()
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data["key"] = "ABCD-EFGH-IJKL-MNOP-QRST"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        self.assertFalse(license.is_activated())

    def test_tampered_machine_code_in_file(self):
        """激活文件中密匙的机器码段被篡改为他机后应失效。"""
        key = license.generate_key(self.mc)
        license.activate(key)
        path = license._license_path()
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # 构造一个合法但绑定其他机器的密匙写入文件
        other_key = license.generate_key("9999-8888-7777")
        data["key"] = other_key.replace("-", "")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        # 文件 machine_id 仍为本机，但密匙机器码非本机 -> 失效
        self.assertFalse(license.is_activated())


if __name__ == "__main__":
    unittest.main()
