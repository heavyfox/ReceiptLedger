"""Verify local release archives without extracting or executing their contents."""
import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import zipfile


def verify(folder):
    entries = {}
    for line in (folder / "SHA256SUMS.txt").read_text(encoding="ascii").splitlines():
        expected, name = line.split("  ", 1)
        if Path(name).name != name or name in entries:
            raise ValueError("Invalid checksum filename")
        entries[name] = expected
    if len(entries) != 3:
        raise ValueError("Expected Windows, app source and dependency source archives")
    for name, expected in entries.items():
        with (folder / name).open("rb") as stream:
            if hashlib.file_digest(stream, "sha256").hexdigest() != expected:
                raise ValueError(f"Checksum mismatch: {name}")
        with zipfile.ZipFile(folder / name) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)) or any(
                PurePosixPath(n).is_absolute() or ".." in PurePosixPath(n).parts
                or "\\" in n or ":" in n for n in names
            ):
                raise ValueError(f"Invalid archive paths: {name}")
            if archive.testzip() is not None:
                raise ValueError(f"Damaged ZIP: {name}")
        print(f"SHA-256 / CRC / paths: OK ({name})")

    source_name = next(n for n in entries if n.endswith("-source.zip"))
    with zipfile.ZipFile(folder / source_name) as source:
        prefix = "ReceiptLedger-GitHub/"
        required = ("LICENSE", "README.md", "version.py", "THIRD_PARTY_NOTICES.md",
                    "THIRD_PARTY_SOURCES.json", "requirements-lock.txt", "vendor/SHA256SUMS.txt",
                    ".github/workflows/windows.yml", ".github/ISSUE_TEMPLATE/bug_report.md",
                    "docs/RELEASE_NOTES.md", "test_connection_settings.py", "test_config.py",
                    "diagnostics.py", "receipt_panel.py", "settings_panel.py", "excel_validation.py",
                    "test_diagnostics.py", "test_excel_validation.py", "docs/ARCHITECTURE.md")
        for name in required:
            if prefix + name not in source.namelist():
                raise ValueError(f"Missing public source file: {name}")
        for name in source.namelist():
            relative = name.removeprefix(prefix)
            parts = PurePosixPath(relative).parts
            if parts[0] in ("data", "build", "dist", ".git", "qa", "logs") or parts[0].startswith(".venv"):
                raise ValueError(f"Private/development folder: {relative}")
            if parts[-1] == "settings.json" or parts[-1].startswith("draft-session"):
                raise ValueError(f"User state in release: {relative}")
            suffix = PurePosixPath(relative).suffix.lower()
            if suffix in (".heic", ".heif", ".jpg", ".jpeg", ".png", ".webp", ".xlsx", ".xls"):
                if relative not in ("tests/fixtures/receipt.heic", "assets/receipt-ledger.png",
                                    "docs/screenshots/light.png", "docs/screenshots/dark.png") and not (
                                        parts[0] == "templates" and suffix == ".xlsx"):
                    raise ValueError(f"Unexpected image/workbook: {relative}")
        wheel_hash, wheel = source.read(prefix + "vendor/SHA256SUMS.txt").decode().strip().split(None, 1)
        if hashlib.sha256(source.read(prefix + "vendor/" + wheel)).hexdigest() != wheel_hash:
            raise ValueError("Vendored HEIC wheel checksum mismatch")
    print("Public source contents / HEIC wheel: OK")

    windows_name = next(n for n in entries if n.endswith("-windows-x64.zip"))
    manifest = json.loads((folder / "WINDOWS-MANIFEST.json").read_text(encoding="utf-8"))
    with zipfile.ZipFile(folder / windows_name) as archive:
        expected_names = {"ReceiptLedger/" + entry["file"] for entry in manifest}
        if set(archive.namelist()) != expected_names:
            raise ValueError("Windows manifest file list mismatch")
        for entry in manifest:
            raw = archive.read("ReceiptLedger/" + entry["file"])
            if len(raw) != entry["bytes"] or hashlib.sha256(raw).hexdigest() != entry["sha256"]:
                raise ValueError(f"Windows manifest mismatch: {entry['file']}")
    print("Windows manifest: OK")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument("folder", type=Path)
    verify(parser.parse_args().folder)
