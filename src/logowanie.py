# ============================================================================ #
#                                  LOGOWANIE                                   
# ============================================================================ #
import builtins
import logging
import os
import sys
from datetime import datetime
from pathlib import Path


def setup_logging(
    log_dir: str = "..\logs",
    level: int = logging.INFO,
    echo_to_console: bool = True,
    capture_print: bool = True,
    LOG_PREFIX: str = "LOG_",
) -> str:
    """
    Logi do pliku logs/<PREFIX>YYYY-MM-DD.log (append) + opcjonalnie na konsolę.
    Przechwytuje print() -> logger.info() bez ręcznego echo na stdout (brak duplikatów).
    """
    date_str = datetime.now().strftime(f"{LOG_PREFIX}%Y-%m-%d")
    log_path = Path(log_dir) / f"{date_str}.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)

    handlers: list[logging.Handler] = [
        logging.FileHandler(log_path, mode="a", encoding="utf-8")
    ]
    if echo_to_console:
        handlers.append(logging.StreamHandler(sys.stdout))

    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        handlers=handlers,
        force=True,
    )

    if capture_print:
        _orig_print = builtins.print

        def print_to_logger(*args, **kwargs):
            file = kwargs.get("file")
            if file not in (None, sys.stdout, sys.stderr):
                return _orig_print(*args, **kwargs)
            sep = kwargs.get("sep", " ")
            msg = sep.join(str(a) for a in args)
            logging.getLogger().info(msg)

        builtins.print = print_to_logger

    return str(log_path)
