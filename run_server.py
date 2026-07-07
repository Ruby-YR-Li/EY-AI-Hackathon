"""质检Agent本地Web服务入口，供源码运行和PyInstaller打包使用。"""

from __future__ import annotations

import logging
from datetime import datetime

import uvicorn

from src.app_paths import APP_VERSION, LOG_DIR
from web.app import app


def main():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"server-{datetime.now():%Y%m%d}.log"
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )
    logging.info("启动质检Agent %s", APP_VERSION)
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")


if __name__ == "__main__":
    main()
