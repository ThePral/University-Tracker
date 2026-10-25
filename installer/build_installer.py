"""Build the Qt app and a self-contained Windows installer from any working directory."""
from __future__ import annotations

import argparse
import ast
import hashlib
from importlib import metadata
import os
from pathlib import Path
import shutil
import struct
import subprocess
import sys
import urllib.request

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build_installer"
PAYLOAD = BUILD / "app" / "UniversityTrackerQt"
COMPILER_VERSION = "7.1.0"
COMPILER_SHA256 = "0362a383ed217d4c4239b5933866dd96d3eb2102737da92f80f6057a4b40df2f"
COMPILER_URL = (
    "https://github.com/jrsoftware/issrc/releases/download/is-7_1_0/"
    "innosetup-7.1.0-x64.exe"
)


def run(arguments: list[str]) -> None:
    environment = os.environ.copy()
    environment["PYINSTALLER_CONFIG_DIR"] = str(BUILD / "cache")
    subprocess.run(arguments, cwd=ROOT, env=environment, check=True)


def app_version() -> str:
    tree = ast.parse((ROOT / "university_tracker_qt.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "APP_VERSION" for target in node.targets
        ):
            version = ast.literal_eval(node.value)
            if isinstance(version, str) and len(version.split(".")) == 3 and all(
                part.isdigit() and 0 <= int(part) <= 65535 for part in version.split(".")
            ):
                return version
    raise RuntimeError("Set APP_VERSION in university_tracker_qt.py to a three-part version such as 1.0.0.")


def find_compiler(explicit: str | None) -> Path | None:
    configured = explicit or os.environ.get("INNO_SETUP_COMPILER")
    if configured:
        candidate = Path(configured).expanduser().resolve()
        if not candidate.is_file():
            raise RuntimeError(f"Inno Setup compiler was not found at {candidate}")
        return candidate
    candidates = [ROOT / ".installer-tools" / "Inno Setup 7" / "ISCC.exe"]
    on_path = shutil.which("ISCC.exe")
    if on_path:
        candidates.append(Path(on_path))
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if base:
            candidates.extend(Path(base) / f"Inno Setup {major}" / "ISCC.exe" for major in (7, 6))
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.extend(Path(local) / "Programs" / f"Inno Setup {major}" / "ISCC.exe" for major in (7, 6))
    return next((candidate for candidate in candidates if candidate.is_file()), None)


def bootstrap_compiler() -> Path:
    """Use Inno Setup's official portable mode; no system installation or registry entry."""
    cache = ROOT / ".installer-tools"
    cache.mkdir(exist_ok=True)
    download = cache / f"innosetup-{COMPILER_VERSION}-x64.exe"
    if not download.is_file() or hashlib.sha256(download.read_bytes()).hexdigest() != COMPILER_SHA256:
        print(f"Downloading the official Inno Setup {COMPILER_VERSION} compiler...", flush=True)
        request = urllib.request.Request(COMPILER_URL, headers={"User-Agent": "UniversityTracker-Builder"})
        try:
            with urllib.request.urlopen(request, timeout=60) as response, download.open("wb") as output:
                shutil.copyfileobj(response, output)
        except OSError as exc:
            raise RuntimeError(
                "Could not download the installer compiler. Check your internet connection, "
                "or install Inno Setup from https://jrsoftware.org/isdl.php and run this build again. "
                "A custom compiler path can be supplied with --compiler or INNO_SETUP_COMPILER."
            ) from exc
    if hashlib.sha256(download.read_bytes()).hexdigest() != COMPILER_SHA256:
        raise RuntimeError("The Inno Setup download failed its SHA-256 check. It will not be run.")
    compiler_folder = cache / "Inno Setup 7"
    print("Preparing the portable compiler inside .installer-tools...", flush=True)
    run([
        str(download), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CURRENTUSER",
        "/PORTABLE=1", "/NOICONS", "/TASKS=", f"/DIR={compiler_folder}", f"/LOG={cache / 'inno-setup.log'}",
    ])
    compiler = compiler_folder / "ISCC.exe"
    if not compiler.is_file():
        raise RuntimeError("The portable Inno Setup compiler was not created.")
    return compiler


def version_resource(version: str) -> Path:
    numbers = tuple(int(part) for part in version.split(".")) + (0,)
    resource = BUILD / "version_info.txt"
    resource.write_text(
        "VSVersionInfo(\n"
        f"  ffi=FixedFileInfo(filevers={numbers!r}, prodvers={numbers!r},\n"
        "    mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),\n"
        "  kids=[StringFileInfo([StringTable('040904B0', [\n"
        "    StringStruct('CompanyName', 'University Tracker'),\n"
        "    StringStruct('FileDescription', 'University Tracker'),\n"
        f"    StringStruct('FileVersion', '{version}'),\n"
        "    StringStruct('InternalName', 'UniversityTrackerQt'),\n"
        "    StringStruct('OriginalFilename', 'UniversityTrackerQt.exe'),\n"
        "    StringStruct('ProductName', 'University Tracker'),\n"
        f"    StringStruct('ProductVersion', '{version}')\n"
        "  ])]), VarFileInfo([VarStruct('Translation', [1033, 1200])])]\n"
        ")\n",
        encoding="utf-8",
    )
    return resource


def copy_dependency_licenses() -> None:
    licenses = PAYLOAD / "ThirdPartyLicenses"
    licenses.mkdir(exist_ok=True)
    inventory = []
    for package in ("openpyxl", "et-xmlfile", "PySide6", "PySide6_Essentials", "PySide6_Addons", "shiboken6",
                    "numpy", "Pillow", "psutil"):
        try:
            distribution = metadata.distribution(package)
        except metadata.PackageNotFoundError:
            if package in ("numpy", "Pillow", "psutil"):
                continue
            raise
        inventory.append(f"{distribution.metadata['Name']} {distribution.version}")
        for entry in distribution.files or ():
            if (".dist-info/" in str(entry).replace("\\", "/") and
                    ("licen" in str(entry).lower() or entry.name.lower().startswith(("copying", "notice")))):
                source = Path(distribution.locate_file(entry))
                if source.is_file():
                    destination = licenses / package / Path(*entry.parts[1:])
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, destination)
    python_license = Path(sys.base_prefix) / "LICENSE.txt"
    if python_license.is_file():
        shutil.copy2(python_license, licenses / "Python-LICENSE.txt")
    (licenses / "DEPENDENCIES.txt").write_text("\n".join(inventory) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--compiler", help="Path to an existing Inno Setup ISCC.exe")
    parser.add_argument("--skip-dependencies", action="store_true", help="Use Python dependencies already installed")
    options = parser.parse_args()
    if sys.platform != "win32" or struct.calcsize("P") != 8:
        raise RuntimeError("Build this installer with 64-bit Python on Windows.")
    if sys.version_info < (3, 11):
        raise RuntimeError("Python 3.11 or newer is required.")
    version = app_version()
    compiler = find_compiler(options.compiler) or bootstrap_compiler()
    if not options.skip_dependencies:
        run([sys.executable, "-m", "pip", "install", "-r", str(ROOT / "requirements_qt.txt")])
    BUILD.mkdir(exist_ok=True)
    print(f"Building University Tracker {version} for the installer...", flush=True)
    run([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed", "--onedir",
        "--noupx", "--name", "UniversityTrackerQt", "--distpath", str(BUILD / "app"),
        "--workpath", str(BUILD / "pyinstaller"), "--specpath", str(BUILD),
        "--version-file", str(version_resource(version)), str(ROOT / "university_tracker_qt.py"),
    ])
    if not (PAYLOAD / "UniversityTrackerQt.exe").is_file():
        raise RuntimeError("PyInstaller did not produce UniversityTrackerQt.exe.")
    if (PAYLOAD / "Excel_Sheets").exists() or any(PAYLOAD.rglob("universities.xlsx")):
        raise RuntimeError("The installer payload contains a workbook. Remove it before distributing.")
    copy_dependency_licenses()
    print("Compiling the setup wizard...", flush=True)
    run([str(compiler), "/Qp", f"/DAppVersion={version}", f"/DAppSource={PAYLOAD}",
         str(ROOT / "installer" / "university_tracker.iss")])
    artifact = ROOT / "dist_installer" / f"UniversityTracker-Setup-{version}.exe"
    if not artifact.is_file():
        raise RuntimeError("The installer compiler did not produce the expected setup executable.")
    digest = hashlib.sha256(artifact.read_bytes()).hexdigest()
    artifact.with_suffix(".exe.sha256").write_text(f"{digest}  {artifact.name}\n", encoding="ascii")
    print(f"\nInstaller ready: {artifact}\nSize: {artifact.stat().st_size / 1024 / 1024:.1f} MB", flush=True)
    print("Share this setup executable. Recipients do not need Python or the build tools.", flush=True)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        print(f"\nBuild failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
