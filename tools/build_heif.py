"""Build a decoder-only, dynamically linked HEIC wheel on Windows x64.

Requires Python 3.12, CMake and Visual Studio 2022 C++ Build Tools.
All source archives are verified before extraction. No system install is made.
"""
from __future__ import annotations

import argparse
import hashlib
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tarfile
from urllib.request import urlopen

SOURCES = {
    "libde265-1.1.3": ("https://github.com/strukturag/libde265/releases/download/v1.1.3/libde265-1.1.3.tar.gz", "554228bd17788c99a7e63b37ab5634722190e6e2bf60c1dcb01cef328e133905"),
    "libheif-1.23.4": ("https://github.com/strukturag/libheif/releases/download/v1.23.4/libheif-1.23.4.tar.gz", "d0c02b4b0e978f34a1974b6f3eea7975a537bf7a9195ffeea38e7242ff316fdd"),
    "pillow_heif-1.8.0": ("https://files.pythonhosted.org/packages/bb/4c/d5319a1f276c70528ff97893afc42a300ff28029e27ca8de89bb3b271680/pillow_heif-1.8.0.tar.gz", "e47c27432c6fd3d66c22f0de9f27fd379383b646c947520bc485854ce72060d0"),
}


def run(*args, **kwargs):
    print("RUN", *map(str, args), flush=True)
    subprocess.run(list(map(str, args)), check=True, **kwargs)


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument("--work", type=Path, required=True)
    parser.add_argument("--archives", type=Path, required=True)
    parser.add_argument("--wheel-dir", type=Path, required=True)
    args = parser.parse_args()
    root, archives, wheels = (p.resolve() for p in (args.work, args.archives, args.wheel_dir))
    for p in (root, archives, wheels):
        p.mkdir(parents=True, exist_ok=True)
    for name, (url, digest) in SOURCES.items():
        archive = archives / (name + ".tar.gz")
        if not archive.exists():
            with urlopen(url, timeout=120) as response:
                archive.write_bytes(response.read())
        if hashlib.sha256(archive.read_bytes()).hexdigest() != digest:
            raise RuntimeError(f"Source checksum mismatch: {archive}")
        if not (root / name).exists():
            with tarfile.open(archive) as src:
                src.extractall(root, filter="data")
    prefix = root / "install"
    # Configure MSVC directly so builds do not depend on per-user SDK registry
    # entries. The compiler and SDK remain installed at their normal locations.
    vswhere = Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")) / "Microsoft Visual Studio/Installer/vswhere.exe"
    vs = Path(subprocess.check_output([str(vswhere), "-latest", "-products", "*", "-requires",
             "Microsoft.VisualStudio.Component.VC.Tools.x86.x64", "-property", "installationPath"], text=True).strip())
    vc = sorted((vs / "VC/Tools/MSVC").iterdir())[-1]
    sdk = Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")) / "Windows Kits/10"
    sdk_version = sorted((sdk / "Include").iterdir())[-1].name
    os.environ["PATH"] = os.pathsep.join(map(str, [Path(sys.executable).parent, vc / "bin/Hostx64/x64", sdk / "bin" / sdk_version / "x64"])) + os.pathsep + os.environ["PATH"]
    os.environ["INCLUDE"] = os.pathsep.join(map(str, [vc / "include", *[sdk / "Include" / sdk_version / p for p in ("ucrt", "shared", "um", "winrt")]]))
    os.environ["LIB"] = os.pathsep.join(map(str, [vc / "lib/x64", sdk / "Lib" / sdk_version / "ucrt/x64", sdk / "Lib" / sdk_version / "um/x64"]))
    os.environ["DISTUTILS_USE_SDK"] = "1"
    os.environ["CL"] = "/utf-8 " + os.environ.get("CL", "")

    def build(name, options):
        folder = root / (name + "-ninja")
        run("cmake", "-S", root / name, "-B", folder, "-G", "Ninja",
            "-DCMAKE_BUILD_TYPE=Release", "-DCMAKE_C_COMPILER=cl", "-DCMAKE_CXX_COMPILER=cl",
            f"-DCMAKE_INSTALL_PREFIX={prefix}", "-DBUILD_SHARED_LIBS=ON", *options)
        run("cmake", "--build", folder, "--config", "Release", "--parallel", "4")
        run("cmake", "--install", folder, "--config", "Release")

    build("libde265-1.1.3", ["-DENABLE_SDL=OFF", "-DENABLE_DECODER=OFF", "-DENABLE_ENCODER=OFF"])
    disabled = ["X265", "X264", "KVAZAAR", "UVG266", "VVDEC", "VVENC", "OpenH264_DECODER",
                "OpenH264_ENCODER", "DAV1D", "AOM_DECODER", "AOM_ENCODER", "SvtEnc", "RAV1E",
                "JPEG_DECODER", "JPEG_ENCODER", "OpenJPEG_DECODER", "OpenJPEG_ENCODER",
                "FFMPEG_DECODER", "OPENJPH_DECODER", "OPENJPH_ENCODER", "LIBSHARPYUV",
                "HEADER_COMPRESSION", "EXAMPLES", "GDK_PIXBUF"]
    build("libheif-1.23.4", [f"-DCMAKE_PREFIX_PATH={prefix}", "-DWITH_LIBDE265=ON",
          "-DWITH_LIBDE265_PLUGIN=OFF", "-DENABLE_PLUGIN_LOADING=OFF", "-DBUILD_TESTING=OFF",
          "-DCMAKE_DISABLE_FIND_PACKAGE_JPEG=ON", "-DCMAKE_DISABLE_FIND_PACKAGE_PNG=ON",
          *[f"-DWITH_{name}=OFF" for name in disabled]])
    # Upstream's Windows setup expects the MSYS import-library filename. The
    # MSVC import library has the same C ABI and is built from these exact sources.
    shutil.copy2(prefix / "lib/heif.lib", prefix / "lib/libheif.lib")
    source = root / "pillow_heif-1.8.0"
    package = source / "pillow_heif"
    shutil.copy2(root / "libheif-1.23.4/COPYING", source / "LICENSE.libheif.txt")
    shutil.copy2(root / "libde265-1.1.3/COPYING", source / "LICENSE.libde265.txt")
    (package / "_version.py").write_text('__version__ = "1.8.0+receiptledger1"\n', encoding="utf-8")
    init = package / "__init__.py"
    original = init.read_text(encoding="utf-8")
    marker = "# ReceiptLedger decoder-only DLL search path"
    if marker not in original:
        init.write_text(marker + '\nimport os as _os\nfrom pathlib import Path as _Path\n'
                        '_dll_directory = _os.add_dll_directory(str(_Path(__file__).resolve().parent))\n'
                        + original, encoding="utf-8")
    for dll in (prefix / "bin").glob("*.dll"):
        shutil.copy2(dll, package / dll.name)
    cfg = source / "setup.cfg"
    contents = cfg.read_text(encoding="utf-8")
    if "[options.package_data]" not in contents:
        cfg.write_text(contents + "\n[options.package_data]\npillow_heif = *.dll\n", encoding="utf-8")
    (source / "LICENSES_bundled.txt").write_text(
        "ReceiptLedger decoder-only build of pillow-heif 1.8.0.\n"
        "pillow-heif and the DLL-loader modification: BSD-3-Clause (LICENSE.txt).\n"
        "libheif 1.23.4 and libde265 1.1.3: LGPL-3.0-or-later, dynamically linked.\n"
        "No x265, HEIC encoder, or MinGW runtime is bundled.\n"
        "Rebuild: ReceiptLedger tools/build_heif.py with the matching source archives.\n",
        encoding="utf-8")
    env = dict(os.environ, MSYS2_PREFIX=str(prefix))
    run(sys.executable, "-m", "pip", "wheel", "--no-cache-dir", "--disable-pip-version-check", "--no-deps", "--no-build-isolation",
        "--wheel-dir", wheels, source, env=env)
    wheel = wheels / "pillow_heif-1.8.0+receiptledger1-cp312-cp312-win_amd64.whl"
    (wheels / "SHA256SUMS.txt").write_text(hashlib.sha256(wheel.read_bytes()).hexdigest() + "  " + wheel.name + "\n", encoding="ascii")


if __name__ == "__main__":
    main()
