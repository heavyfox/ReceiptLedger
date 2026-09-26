# HEIC decoder wheel

`pillow_heif-1.8.0+receiptledger1-cp312-cp312-win_amd64.whl` is built specifically for ReceiptLedger on Windows x64 / CPython 3.12.

- Python binding and loader modification: BSD-3-Clause.
- Bundled heif.dll (libheif 1.23.4) and libde265.dll (1.1.3): LGPL-3.0-or-later, dynamically linked. Complete license texts are also inside the wheel.
- x265 / x264 encoders and the MinGW runtime are not present.
- Source archives and their hashes: `../THIRD_PARTY_SOURCES.json`.
- Rebuild instructions and modifications: `../tools/build_heif.py` and `../docs/BUILD.md`.
- Corresponding sources are provided in the release's `ReceiptLedger-v1.4.1-third-party-sources.zip`.

Verify the wheel using `SHA256SUMS.txt` before use. If rebuilding or modifying the wheel, update that checksum file for your build. Runtime library replacement is permitted.
