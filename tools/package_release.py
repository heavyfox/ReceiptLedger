"""Prepare local GitHub/release artifacts. This script never uploads anything."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from version import APP_VERSION
from core import ExcelLedger


def digest(path):
    with path.open('rb') as handle:
        return hashlib.file_digest(handle, 'sha256').hexdigest()


def archive_tree(folder, target):
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(folder.rglob('*')):
            if path.is_file():
                archive.write(path, path.relative_to(folder.parent).as_posix())


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--sources', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    output, sources = args.output.resolve(), args.sources.resolve()
    # Refuse to mix a new release with any old files or user data.
    if output.exists() and any(output.iterdir()):
        raise SystemExit('Choose an empty output directory; existing files are never deleted.')
    output.mkdir(parents=True, exist_ok=True)
    binary = ROOT / 'dist/ReceiptLedger'
    if not (binary / 'ReceiptLedger.exe').is_file():
        raise SystemExit('Run build.ps1 first.')
    forbidden = ('x265', 'x264', 'libgcc', 'libstdc++', 'opengl32sw', 'qt6quick', 'qt6qml',
                 'qt6pdf', 'qt6virtualkeyboard', '_avif')
    for path in binary.rglob('*'):
        if path.is_file() and any(x in path.name.lower() for x in forbidden):
            raise RuntimeError(f'Unexpected binary: {path.name}')
    # Ensure the shipped template is actually empty.
    for template in (ROOT / 'templates').glob('*.xlsx'):
        if ExcelLedger(template).load():
            raise RuntimeError('Template contains receipt records')
    source_entries = json.loads((ROOT / 'THIRD_PARTY_SOURCES.json').read_text(encoding='utf-8'))
    for entry in source_entries:
        archive = sources / entry['file']
        if not archive.is_file() or digest(archive) != entry['sha256']:
            raise RuntimeError(f"Missing or incorrect source archive: {entry['file']}")

    repo = output / 'ReceiptLedger-GitHub'
    repo.mkdir()
    names = ['.gitignore', 'LICENSE', 'README.md', 'UPDATE.md', 'THIRD_PARTY_NOTICES.md',
             'THIRD_PARTY_SOURCES.json', 'DEPENDENCIES.json', 'requirements.txt',
             'requirements-lock.txt', 'build.ps1', 'ReceiptLedger.spec']
    names += [p.name for p in ROOT.glob('*.py')]
    for name in sorted(set(names)):
        shutil.copy2(ROOT / name, repo / name)
    # Explicit directory allowlist: no application data, build logs or private QA.
    for directory in ('.github', 'assets', 'docs', 'hooks', 'licenses', 'templates', 'tests', 'tools', 'vendor'):
        shutil.copytree(ROOT / directory, repo / directory,
                        ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))

    windows = output / 'ReceiptLedger'
    shutil.copytree(binary, windows)
    for name in ('LICENSE', 'README.md', 'UPDATE.md', 'THIRD_PARTY_NOTICES.md', 'DEPENDENCIES.json', 'THIRD_PARTY_SOURCES.json'):
        shutil.copy2(repo / name, windows / name)
    for directory in ('licenses', 'docs', 'templates', 'assets'):
        shutil.copytree(repo / directory, windows / directory)

    src_tree = output / 'ReceiptLedger-ThirdPartySources'
    src_tree.mkdir()
    for entry in source_entries:
        shutil.copy2(sources / entry['file'], src_tree / entry['file'])
    shutil.copy2(ROOT / 'THIRD_PARTY_SOURCES.json', src_tree / 'SOURCES.json')
    shutil.copy2(ROOT / 'tools/build_heif.py', src_tree / 'build_heif.py')
    shutil.copy2(ROOT / 'docs/BUILD.md', src_tree / 'BUILD.md')
    (src_tree / 'README.txt').write_text(
        'Corresponding sources for ReceiptLedger ' + APP_VERSION + '\n'
        'Qt Base / PySide / Shiboken: unmodified official 6.11.2 binaries.\n'
        'HEIC: libheif 1.23.4 + libde265 1.1.3 dynamically linked, encoders disabled.\n'
        'pillow-heif loader/packaging modifications: build_heif.py.\n'
        'Original copyright/license texts are inside each source archive.\n'
        'Source URLs and SHA-256 values: SOURCES.json.\n'
        'Also obtain ReceiptLedger-v' + APP_VERSION + '-source.zip for the application and complete build/package instructions.\n',
        encoding='utf-8')
    files = []
    for folder, suffix in ((windows, 'windows-x64'), (repo, 'source'), (src_tree, 'third-party-sources')):
        target = output / f'ReceiptLedger-v{APP_VERSION}-{suffix}.zip'
        archive_tree(folder, target)
        files.append(target)
    (output / 'SHA256SUMS.txt').write_text(''.join(f'{digest(p)}  {p.name}\n' for p in files), encoding='ascii')
    binary_manifest = [{'file':p.relative_to(windows).as_posix(), 'bytes':p.stat().st_size, 'sha256':digest(p)}
                       for p in sorted(windows.rglob('*')) if p.is_file()]
    (output / 'WINDOWS-MANIFEST.json').write_text(json.dumps(binary_manifest, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({'version':APP_VERSION, 'expanded_bytes':sum(x['bytes'] for x in binary_manifest),
                      'archives':{p.name:p.stat().st_size for p in files}}, indent=2))


if __name__ == '__main__':
    main()
