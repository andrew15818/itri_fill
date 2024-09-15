import sys
import asyncio
import logging
import docx
import pandas as pd
from datetime import datetime
from typing import List, Dict

from config import EXPERT_LIST, CONTRACT_DOC, RECEIPT_DOC, SIGNATURE_DOC, CONTRACT_WORDS


def init_logger(filename: str = "logs/logs.log") -> logging.RootLogger:
    logger = logging.getLogger(__name__)
    logging.basicConfig(filename=filename, encoding="utf-8", level=logging.DEBUG)
    return logger


def format_contract_fill_in_data(data: pd.DataFrame) -> List[Dict[str, str]]:
    """
    Organize the dict of values we search for and the data we fill in.
    Args:
        data (pd.DataFrame): User data
    Returns:
        List of dictionaries with values to search for and fill in for each user.
    """
    formatted = []
    for index, row in data.iterrows():
        formatted_row = {
            "意於年月日": row["會議日期"].strftime("%Y-%m-%d"),
            "申請案號：案號": f"申請案號：{row['案號']}",
            "課程名稱：課程全名": f"課程名稱：{row['課程名稱']}",
            "單位名稱：單位全名": f"單位名稱：{row['單位名稱']}",
            "立切結書人：姓名": f"立切結書人：{row['姓名']}",
            "身分證統一編號：身分證字號": f"身分證統一編號：{row['身分證字號']}",
            "中華民國Date": f"中華民國{row['會議日期']}",
        }
        formatted.append(formatted_row)
    return formatted


def format_receipt_fill_in_data(data: pd.DataFrame) -> List[Dict[str, str]]:
    """
    Organize the dictionary values to search for and the values to replace for receipt file.
    Args:
        data (pd.DataFrame): User data
    Returns:
        List of dictionaries with values to search for and fill in for each user.
    """
    formatted = []
    for index, row in data.iterrows():
        formatted_row = {
            "領款人簽章(正楷)：姓名": f"領款人簽章(正楷)：{row['姓名']}",
            "聯絡電話：電話": f"聯絡電話：{row['手機']}",
            "Date": f"{row['會議日期']}",
            "中華民國國籍：身分證統一編號　ID": f"中華民國國籍：身分證統一編號 {row['身分證字號']}",
        }
        formatted.append(formatted_row)
    return formatted


def convert_date_to_chinese(date: datetime) -> str:
    """
    Take a DateTime object and format it to Taiwanese format.
    Args:
        date (datetime.TimeStamp): DateTime object with needed date
    Returns:
        (str) containing the time converted to Taiwanese time (e.g 1998/03/29 -> 87年03月29日)
    """
    return ""


def search_and_replace_expert_info(
    filename: str,
    user_data: List[Dict[str, str]],
) -> List[docx.Document]:
    """
    Search for keywords in the document's paragraphs and replace it with the matching data.
    Args:
        filename (str): Name of the Word file to open
        user_data (List[Dict[str, str]]): Formatted information for each user, phrase
            to search for and value to replace by.
    Returns:
        List of documents containing the modified word documents.
    """
    modified_docs = []
    for idx, user in enumerate(user_data):
        doc = docx.Document(filename)
        for par in doc.paragraphs:
            for keyword, value in user.items():
                if keyword in par.text:
                    par.text = par.text.replace(keyword, value)
        modified_docs.append(doc)
    return modified_docs


def edit_contract(data: pd.DataFrame):
    """
    Edit 切結書, fill in date, number, and expert information
    Args:
        data (pd.DataFrame): dataframe from which to fill columns.
    """

    formatted = format_contract_fill_in_data(data)
    documents = search_and_replace_expert_info(CONTRACT_DOC, formatted)
    for idx, doc in enumerate(documents):
        doc.save(f"data/切結書_{data.iloc[idx]['姓名']}.docx")


def edit_receipt(data: pd.DataFrame):
    """
    Edit the receipt file  領據 with expert information.
    """
    formatted = format_receipt_fill_in_data(data)
    docs = search_and_replace_expert_info(RECEIPT_DOC, formatted)
    for idx, doc in enumerate(docs):
        doc.save(f"data/領據_{data.iloc[idx]["姓名"]}.docx")


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

    # edit_contract(expert_info)
    edit_receipt(expert_info)


if __name__ == "__main__":
    main()
