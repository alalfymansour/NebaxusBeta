import os
import sys
import threading
import random
import logging
import time
import shutil
from datetime import datetime

os.environ.setdefault('FLASK_ENV', 'sqlite')
os.environ.setdefault('SECRET_KEY', 'nebaxus-desktop-trial-secret-key-2026')
os.environ.setdefault('ADMIN_USERNAME', 'nebaxusbeta')
os.environ.setdefault('ADMIN_PASSWORD', 'nebaxusbetapassword')
os.environ.setdefault('NEBAXUS_MODE', 'trial')
os.environ.setdefault('FAST_SPLASH', '1')
os.environ.setdefault('FAST_LOGIN', '1')
os.environ.setdefault('SKIP_STARTUP_SEED', '0')

_base = os.environ.get('LOCALAPPDATA') or os.path.join(os.path.expanduser('~'), 'AppData', 'Local')
INSTANCE_DIR = os.path.join(_base, 'NebaxusBeta')
os.environ.setdefault('NEBAXUS_INSTANCE_PATH', os.path.join(INSTANCE_DIR, 'instance'))
os.environ.setdefault('NEBAXUS_EXPORT_PATH', os.path.join(INSTANCE_DIR, 'exports'))
os.environ.setdefault('NEBAXUS_BACKUP_PATH', os.path.join(INSTANCE_DIR, 'backups'))

for d in ('instance', 'logs', 'backups', 'exports'):
    os.makedirs(os.path.join(INSTANCE_DIR, d), exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(
            os.path.join(INSTANCE_DIR, 'logs', f'{datetime.now():%Y-%m-%d}.log'),
            encoding='utf-8',
        ),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger('nebaxus')

PORT = random.randint(49152, 65535)


def daily_backup():
    while True:
        time.sleep(86400)
        try:
            src = os.path.join(os.environ['NEBAXUS_INSTANCE_PATH'], 'dukana.db')
            if not os.path.exists(src):
                continue
            date_str = datetime.now().strftime('%Y-%m-%d')
            dst = os.path.join(os.environ['NEBAXUS_BACKUP_PATH'], f'{date_str}.db')
            shutil.copy2(src, dst)
            backups = sorted(os.listdir(os.environ['NEBAXUS_BACKUP_PATH']))
            while len(backups) > 30:
                os.remove(os.path.join(os.environ['NEBAXUS_BACKUP_PATH'], backups.pop(0)))
            logger.info('Daily backup: %s', date_str)
        except Exception as e:
            logger.error('Backup failed: %s', e)


def do_startup_backup():
    src = os.path.join(os.environ['NEBAXUS_INSTANCE_PATH'], 'dukana.db')
    if not os.path.exists(src):
        return
    date_str = datetime.now().strftime('%Y-%m-%d')
    dst = os.path.join(os.environ['NEBAXUS_BACKUP_PATH'], f'{date_str}.db')
    if not os.path.exists(dst):
        shutil.copy2(src, dst)
        logger.info('Startup backup: %s', date_str)


def start_flask():
    from app import create_app
    app = create_app('sqlite')
    app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False)


def wait_for_flask():
    import urllib.request
    url = f'http://127.0.0.1:{PORT}/ping'
    for _ in range(100):
        try:
            urllib.request.urlopen(url, timeout=1)
            return
        except Exception:
            time.sleep(0.1)
    raise RuntimeError('Flask failed to start')


if __name__ == '__main__':
    logger.info('Starting NebaxusBeta on port %d', PORT)
    logger.info('Instance path: %s', os.environ['NEBAXUS_INSTANCE_PATH'])

    t = threading.Thread(target=start_flask, daemon=True)
    t.start()
    wait_for_flask()
    logger.info('Flask ready')

    do_startup_backup()
    b = threading.Thread(target=daily_backup, daemon=True)
    b.start()

    import webview
    webview.create_window(
        'NebaxusBeta',
        f'http://127.0.0.1:{PORT}/splash',
        width=1280,
        height=800,
        resizable=True,
    )
    webview.start()
