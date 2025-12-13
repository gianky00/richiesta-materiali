import logging
import sys

def setup_logging(name="RDA_Bot"):
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # Stream Handler (Console)
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(formatter)
        logger.addHandler(ch)

        # File Handler could be added here if needed

    return logger

logger = setup_logging()
