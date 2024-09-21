import os
from dotenv import load_dotenv

load_dotenv()

EXPERT_LIST = os.environ["EXPERT_LIST"]
CONTRACT_DOC = os.environ["CONTRACT_DOC"]
RECEIPT_DOC = os.environ["RECEIPT_DOC"]
SIGNATURE_DOC = os.environ["SIGNATURE_DOC"]
TAIWAN_DATE_OFFSET = int(os.environ["TAIWAN_DATE_OFFSET"])

TARGET_DIR = os.environ["TARGET_DIR"]
TARGET_FILE = os.environ["TARGET_FILE"]
