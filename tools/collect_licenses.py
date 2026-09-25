"""Collect actual dependency notices and corresponding-source metadata.

Run with the release build's Python and the source archives used for that release.
"""
from __future__ import annotations
import argparse
import hashlib
import importlib.metadata as metadata
import json
from pathlib import Path, PurePosixPath
import shutil
import sys
import tarfile
from urllib.request import urlopen

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument("--sources", type=Path, required=True)
    args = parser.parse_args()
    dest = ROOT / "licenses"
    dest.mkdir(exist_ok=True)
    manifest = []
    for dist in sorted(metadata.distributions(), key=lambda d: d.metadata['Name'].lower()):
        name, version = dist.metadata['Name'], dist.version
        if name.lower() == 'pip':
            continue
        copies = []
        for member in dist.files or []:
            parts = PurePosixPath(str(member)).parts
            if not any(p.endswith('.dist-info') for p in parts):
                continue
            if '/licenses/' not in str(member).lower() and not any(word in member.name.lower() for word in ('license', 'licence', 'copying', 'notice')):
                continue
            origin = Path(dist.locate_file(member))
            if not origin.is_file():
                continue
            target = dest / 'python-packages' / f'{name}-{version}' / Path(*parts[1:])
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(origin, target)
            copies.append(target.relative_to(ROOT).as_posix())
        manifest.append({'name':name, 'version':version, 'notices':copies})
    (dest / 'Python').mkdir(exist_ok=True)
    shutil.copyfile(Path(sys.base_prefix) / 'LICENSE.txt', dest / 'Python/LICENSE.txt')
    for archive in sorted(args.sources.glob('*.tar.*')):
        with tarfile.open(archive) as source:
            for member in source:
                if not member.isfile():
                    continue
                rel = PurePosixPath(member.name)
                if rel.is_absolute() or '..' in rel.parts:
                    raise ValueError('Unsafe source archive member')
                lower = member.name.lower()
                selected = ('/licenses/' in lower or
                            any(word in rel.name.lower() for word in ('license', 'licence', 'copying', 'copyright', 'notice', 'qt_attribution')))
                if not selected or rel.suffix.lower() in ('.png', '.jpg', '.cpp', '.h', '.py', '.rst', '.qdoc', '.cmake'):
                    continue
                target = dest / 'native-sources' / Path(*rel.parts)
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(source.extractfile(member).read())
    # Licenses of runtime libraries provided by the Python/Windows packages.
    extra = {
        'OpenSSL/LICENSE.txt': 'https://raw.githubusercontent.com/openssl/openssl/openssl-3.5.8/LICENSE.txt',
        'OpenSSL/AUTHORS.md': 'https://raw.githubusercontent.com/openssl/openssl/openssl-3.5.8/AUTHORS.md',
        'libffi/LICENSE.txt': 'https://raw.githubusercontent.com/libffi/libffi/v3.4.4/LICENSE',
        'Microsoft/VC-Runtime.docx': 'https://visualstudio.microsoft.com/wp-content/uploads/2021/09/Visual-C-Runtime-2015-2022-License-1.docx',
    }
    for name, url in extra.items():
        target = dest / name
        if not target.exists():
            with urlopen(url, timeout=60) as response:
                contents = response.read()
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(contents)
    (ROOT / 'DEPENDENCIES.json').write_text(json.dumps({
        'python':sys.version.split()[0], 'packages':manifest,
        'note':'Includes build dependencies. Qt/Pillow source license directories include notices for upstream optional components; these do not imply all components are in the Windows app.',
        'extra_notice_sources':extra,
    }, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Collected {sum(1 for p in dest.rglob("*") if p.is_file())} license/notice files')


if __name__ == '__main__':
    main()
