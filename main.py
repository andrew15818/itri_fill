import sys
import asyncio
import logging
import docx
import pandas as pd
from datetime import datetime

from config import EXPERT_LIST, CONTRACT_DOC, RECEIPT_DOC, SIGNATURE_DOC


def init_logger(filename: str = "logs/logs.log") -> logging.RootLogger:
    logger = logging.getLogger(__name__)
    logging.basicConfig(filename=filename, encoding="utf-8", level=logging.DEBUG)
    return logger


def edit_contract(data: pd.DataFrame):
    """
    Edit 切結書, fill in date, number, and expert information
    Args:
        data (pd.DataFrame): dataframe from which to fill columns.
    """
    doc = docx.Document(CONTRACT_DOC)
    pass


def edit_receipt():
    pass


def edit_signature_sheet():
    pass


def main():
    """Initial function called on startup.
    Process the Excel document and fill in the fields on
    the required word docs."""
    logger = init_logger()
    logger.info(f"Starting program at {datetime.now()}")

    expert_info = pd.read_excel(EXPERT_LIST)
    logger.info(f"Loaded {EXPERT_LIST}")

    edit_contract(expert_info)


if __name__ == "__main__":
    main()
