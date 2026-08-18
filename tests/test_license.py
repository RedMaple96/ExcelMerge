"""在线授权客户端回归测试。"""

import json
import os
import shutil
import tempfile
import time
import unittest
from unittest import mock

from src.utils import license


class OnlineLicenseTest(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmpdir, ignore_errors=True)
        patcher = mock.patch.object(license, "_config_dir", return_value=self.tmpdir)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_machine_fingerprint_is_stable_and_hashed(self):
        fingerprint = license.get_machine_fingerprint()
        self.assertEqual(fingerprint, license.get_machine_fingerprint())
        self.assertEqual(len(fingerprint), 32)
        self.assertTrue(all(char in "0123456789abcdef" for char in fingerprint))

    def test_machine_fingerprint_changes_with_hardware_id(self):
        with mock.patch.object(license, "_raw_machine_id", return_value="hardware-A"):
            first = license.get_machine_fingerprint()
        with mock.patch.object(license, "_raw_machine_id", return_value="hardware-B"):
            second = license.get_machine_fingerprint()
        self.assertNotEqual(first, second)

    def test_request_code_uses_remote_service(self):
        with mock.patch.object(
            license,
            "_post_json",
            return_value={"ok": True, "licenseCode": "EM-TEST"},
        ) as post:
            result = license.request_code()
        self.assertEqual(result["licenseCode"], "EM-TEST")
        self.assertEqual(post.call_args.args[0], "/api/request-code")
        self.assertIn("machineFingerprint", post.call_args.args[1])

    def test_activate_writes_signed_local_record(self):
        response = {"ok": True, "token": "token-1", "expiresAt": 2000000000}
        with mock.patch.object(license, "_post_json", return_value=response):
            license.activate("EM-TEST")
        record = license.read_license()
        self.assertIsNotNone(record)
        self.assertEqual(record["licenseCode"], "EM-TEST")
        self.assertEqual(record["token"], "token-1")
        self.assertTrue(license.is_activated())

    def test_old_offline_license_file_is_ignored(self):
        with open(os.path.join(self.tmpdir, "license.dat"), "w", encoding="utf-8") as file:
            file.write("{}")
        self.assertIsNone(license.read_license())
        self.assertFalse(license.is_activated())

    def test_tampered_license_file_is_rejected(self):
        response = {"ok": True, "token": "token-1", "expiresAt": None}
        with mock.patch.object(license, "_post_json", return_value=response):
            license.activate("EM-TEST")
        path = license._license_path()
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
        data["expiresAt"] = 9999999999
        with open(path, "w", encoding="utf-8") as file:
            json.dump(data, file)
        self.assertIsNone(license.read_license())

    def test_verify_updates_last_verified_time(self):
        response = {"ok": True, "token": "token-1", "expiresAt": None}
        with mock.patch.object(license, "_post_json", return_value=response):
            license.activate("EM-TEST")
        old_time = int(time.time()) - 1000
        record = license.read_license()
        record["lastVerifiedAt"] = old_time
        self.assertTrue(license._write_license(record))
        with mock.patch.object(
            license, "_post_json", return_value={"ok": True, "valid": True, "expiresAt": None}
        ):
            result = license.check_license()
        self.assertTrue(result["ok"])
        self.assertGreater(license.read_license()["lastVerifiedAt"], old_time)

    def test_offline_grace_period(self):
        now = int(time.time())
        response = {"ok": True, "token": "token-1", "expiresAt": now + 3600}
        with mock.patch.object(license, "_post_json", return_value=response):
            license.activate("EM-TEST")
        with mock.patch.object(
            license,
            "_post_json",
            side_effect=license.NetworkError("offline"),
        ):
            result = license.check_license()
        self.assertTrue(result["ok"])
        self.assertEqual(result["reason"], "offline-grace")

    def test_offline_grace_expires_after_seven_days(self):
        response = {"ok": True, "token": "token-1", "expiresAt": None}
        with mock.patch.object(license, "_post_json", return_value=response):
            license.activate("EM-TEST")
        record = license.read_license()
        record["lastVerifiedAt"] = int(time.time()) - license._GRACE_PERIOD_SECONDS - 1
        license._write_license(record)
        with mock.patch.object(
            license, "_post_json", side_effect=license.NetworkError("offline")
        ):
            result = license.check_license()
        self.assertFalse(result["ok"])
        self.assertEqual(result["reason"], "grace-expired")

    def test_time_rollback_rejects_offline_grace(self):
        response = {"ok": True, "token": "token-1", "expiresAt": None}
        with mock.patch.object(license, "_post_json", return_value=response):
            license.activate("EM-TEST")
        license._write_last_known_time(int(time.time()) + 86400)
        with mock.patch.object(
            license, "_post_json", side_effect=license.NetworkError("offline")
        ):
            result = license.check_license()
        self.assertFalse(result["ok"])
        self.assertEqual(result["reason"], "grace-expired")

    def test_remote_revocation_invalidates_license(self):
        response = {"ok": True, "token": "token-1", "expiresAt": None}
        with mock.patch.object(license, "_post_json", return_value=response):
            license.activate("EM-TEST")
        revoked = {"ok": True, "valid": False, "reason": "ACTIVATION_REVOKED"}
        with mock.patch.object(license, "_post_json", return_value=revoked):
            result = license.check_license()
        self.assertFalse(result["ok"])
        self.assertEqual(result["reason"], "revoked")


if __name__ == "__main__":
    unittest.main()
