import logging
import logging.handlers
from pathlib import Path


def setup_root_logger(verbose):
    # root loggerの設定
    rootLogger = logging.getLogger()

    rootLogger.setLevel(logging.DEBUG if verbose else logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s (%(process)d)[%(levelname)s] %(message)s"
    )

    streamHandler = logging.StreamHandler()
    streamHandler.setLevel(logging.DEBUG)
    streamHandler.setFormatter(formatter)

    appLogHandler = logging.handlers.RotatingFileHandler(
        Path(__file__) / "../../../logs/info.log",
        encoding="utf-8",
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
    )
    appLogHandler.setLevel(logging.DEBUG)
    appLogHandler.setFormatter(formatter)

    errorLogHandler = logging.handlers.RotatingFileHandler(
        Path(__file__) / "../../../logs/error.log",
        encoding="utf-8",
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
    )
    errorLogHandler.setLevel(logging.WARN)
    errorLogHandler.setFormatter(formatter)

    rootLogger.addHandler(streamHandler)
    rootLogger.addHandler(appLogHandler)
    rootLogger.addHandler(errorLogHandler)
