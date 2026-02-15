"""Logging configuration."""

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logging(config: dict) -> logging.Logger:
    """Configure logging based on config settings.

    Args:
        config: The full config dict (uses config["logging"]).

    Returns:
        The root logger for the application.
    """
    log_config = config["logging"]
    logger = logging.getLogger("outlook_gcal_sync")
    logger.setLevel(getattr(logging, log_config["level"].upper(), logging.INFO))

    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Console handler
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    logger.addHandler(console)

    # File handler (rotating)
    log_file = log_config.get("file")
    if log_file:
        Path(log_file).parent.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=log_config.get("max_bytes", 5_242_880),
            backupCount=log_config.get("backup_count", 3),
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger
