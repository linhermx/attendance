from __future__ import annotations

import json
import time
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import attendance_launcher as launcher


def write_cache(path: Path, *, tag: str, url: str, checked_at: float) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "version": list(launcher.parse_version(tag)),
                "tag": tag,
                "url": url,
                "asset_name": launcher.ASSET_NAME,
                "checked_at": checked_at,
            }
        ),
        encoding="utf-8",
    )


class LauncherReleaseResolutionTests(unittest.TestCase):
    def test_resolve_latest_release_uses_fresh_cache_without_network(self) -> None:
        with TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "state" / "launcher_state.json"
            write_cache(
                cache_path,
                tag="v2.1.0",
                url="https://example.invalid/attendance_windows.zip",
                checked_at=time.time(),
            )

            with patch("attendance_launcher.get_latest_release") as mocked:
                release_info, refreshed_live = launcher.resolve_latest_release(cache_path)

            self.assertFalse(refreshed_live)
            self.assertEqual(release_info["tag"], "v2.1.0")
            mocked.assert_not_called()

    def test_resolve_latest_release_refreshes_stale_cache_from_remote(self) -> None:
        with TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "state" / "launcher_state.json"
            write_cache(
                cache_path,
                tag="v2.1.0",
                url="https://example.invalid/old.zip",
                checked_at=time.time() - launcher.CACHE_MAX_AGE_SECONDS - 60,
            )

            with patch(
                "attendance_launcher.get_latest_release",
                return_value=((2, 1, 1), "v2.1.1", "https://example.invalid/new.zip", launcher.ASSET_NAME),
            ) as mocked:
                release_info, refreshed_live = launcher.resolve_latest_release(
                    cache_path,
                    timeout_seconds=launcher.STARTUP_LATEST_TIMEOUT_SECONDS,
                )

            self.assertTrue(refreshed_live)
            self.assertEqual(release_info["tag"], "v2.1.1")
            self.assertEqual(
                launcher.load_cached_release(cache_path)["tag"],
                "v2.1.1",
            )
            mocked.assert_called_once_with(timeout_seconds=launcher.STARTUP_LATEST_TIMEOUT_SECONDS)

    def test_resolve_latest_release_returns_stale_cache_when_remote_fails(self) -> None:
        with TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "state" / "launcher_state.json"
            write_cache(
                cache_path,
                tag="v2.1.0",
                url="https://example.invalid/old.zip",
                checked_at=time.time() - launcher.CACHE_MAX_AGE_SECONDS - 60,
            )

            with patch(
                "attendance_launcher.get_latest_release",
                side_effect=RuntimeError("network down"),
            ):
                release_info, refreshed_live = launcher.resolve_latest_release(cache_path)

            self.assertFalse(refreshed_live)
            self.assertEqual(release_info["tag"], "v2.1.0")

    def test_resolve_latest_release_raises_without_cache_when_remote_fails(self) -> None:
        with TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "state" / "launcher_state.json"

            with patch(
                "attendance_launcher.get_latest_release",
                side_effect=RuntimeError("network down"),
            ):
                with self.assertRaisesRegex(RuntimeError, "network down"):
                    launcher.resolve_latest_release(cache_path)


if __name__ == "__main__":
    unittest.main()
