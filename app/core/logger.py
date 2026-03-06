import sys
import os
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from functools import wraps

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")


def setup_logging():
    path = "C:/tmp"

    # Check if the directory exists
    if not os.path.exists(path):
        # Create the directory
        os.makedirs(path)
        print(f"Directory {path} created.")
    else:
        print(f"Directory {path} already exists.")

    LOG_FILE = os.path.join(path, f'test_station_interface_{timestamp}.log')

    # Custom formatter that includes function name
    class ContextFormatter(logging.Formatter):
        def format(self, record):
            # Only add function name if it's not already there
            if not hasattr(record, 'func_name'):
                record.func_name = "Internal_Function_driver"
            return super().format(record)

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # File handler
    file_handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=1024 * 1024,
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # Formatter with function name
    formatter = ContextFormatter(
        '%(asctime)s - %(name)s - %(levelname)s - [%(func_name)s] - %(message)s'
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    # Decorator to add function name to log records
    def log_function(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = logging.getLogger(func.__module__)
            logger.debug(f"Entering {func.__name__}", extra={'func_name': func.__name__})
            try:
                result = func(*args, **kwargs)
                logger.debug(f"Exiting {func.__name__}", extra={'func_name': func.__name__})
                return result
            except Exception as e:
                logger.error(f"Error in {func.__name__}: {str(e)}",
                             exc_info=True,
                             extra={'func_name': func.__name__})
                raise

        return wrapper

    return logger, log_function


# Initialize logging
logger, log_function = setup_logging()
