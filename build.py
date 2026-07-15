"""
NebaxusBeta — all-in-one build script.
Run on Windows:     python build.py
Or double-click:    build.bat

Does everything:
  1. Generate nebaxus.ico from nebaxus.svg
  2. Compile desktop.py → NebaxusBeta.exe via Nuitka
  3. Package → NebaxusBeta_Setup.exe via Inno Setup
"""

import os
import sys
import subprocess
import shutil
import io

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(PROJECT_DIR, 'dist')
BUILD_DIR = os.path.join(PROJECT_DIR, 'build')

ICON_PATH = os.path.join(PROJECT_DIR, 'nebaxus.ico')
SVG_PATH = os.path.join(PROJECT_DIR, 'nebaxus.svg')


def log(msg):
    print(f'[build] {msg}')


def ensure_icon():
    if os.path.exists(ICON_PATH):
        log(f'Icon exists: {ICON_PATH}')
        return True
    if not os.path.exists(SVG_PATH):
        log('No SVG found — skipping icon')
        return False
    try:
        import cairosvg
        from PIL import Image
    except ImportError:
        log('cairosvg/pillow not installed — skipping icon')
        log('Install: pip install cairosvg pillow')
        return False
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = []
    for s in sizes:
        png_data = cairosvg.svg2png(url=SVG_PATH, output_width=s, output_height=s)
        images.append(Image.open(io.BytesIO(png_data)))
    images[0].save(
        ICON_PATH, format='ICO',
        sizes=[(s, s) for s in sizes],
        append_images=images[1:],
    )
    log(f'Icon generated: nebaxus.ico ({len(sizes)} sizes)')
    return True


def run_nuitka():
    log('Compiling with Nuitka (10-20 min)...')
    cmd = [
        sys.executable, '-m', 'nuitka',
        '--standalone',
        '--onefile',
        '--windows-console-mode=disable',
        '--output-dir=' + DIST_DIR,
        '--include-data-dir=app/templates=app/templates',
        '--include-data-dir=static=static',
        '--include-data-dir=migrations=migrations',
        '--include-package=app',
        '--include-package=alembic',
        '--nofollow-import-to=psycopg2',
        '--nofollow-import-to=webview.platforms.android',
        '--nofollow-import-to=gunicorn',
        '--nofollow-import-to=pytest',
        '--include-module=webview.platforms.win32',
        '--no-deployment-flag=excluded-module-usage',
    ]
    if os.path.exists(ICON_PATH):
        cmd.append('--windows-icon-from-ico=' + ICON_PATH)
        cmd.append('--include-data-files=nebaxus.ico=nebaxus.ico')
    cmd.append(os.path.join(PROJECT_DIR, 'desktop.py'))

    result = subprocess.run(cmd, cwd=PROJECT_DIR)
    if result.returncode != 0:
        log(f'Nuitka failed (code {result.returncode})')
        sys.exit(result.returncode)
    log('Nuitka OK')


def run_innosetup():
    iscc_paths = [
        r'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
        r'C:\Program Files\Inno Setup 6\ISCC.exe',
        r'C:\Program Files (x86)\Inno Setup\ISCC.exe',
        r'C:\Program Files\Inno Setup\ISCC.exe',
    ]
    for path in iscc_paths:
        if os.path.exists(path):
            log(f'Running Inno Setup: {path}')
            result = subprocess.run([path, os.path.join(PROJECT_DIR, 'installer.iss')], cwd=PROJECT_DIR)
            if result.returncode != 0:
                log(f'Inno Setup failed (code {result.returncode})')
                sys.exit(result.returncode)
            log(f'Installer: {os.path.join(DIST_DIR, "NebaxusBeta_Setup.exe")}')
            return
    log('Inno Setup not found — install from https://jrsoftware.org/isdl.php')
    log('Then open installer.iss and click Build → Compile')


def main():
    log('=== NebaxusBeta Build ===')
    log(f'Python: {sys.version.split()[0]}')

    os.makedirs(DIST_DIR, exist_ok=True)
    os.makedirs(BUILD_DIR, exist_ok=True)

    ensure_icon()
    run_nuitka()
    run_innosetup()

    log('Done!')
    log(f'EXE:      {os.path.join(DIST_DIR, "NebaxusBeta.exe")}')
    log(f'Installer: {os.path.join(DIST_DIR, "NebaxusBeta_Setup.exe")}')


if __name__ == '__main__':
    main()
