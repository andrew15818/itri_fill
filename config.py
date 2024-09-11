import os
from dotenv import load_dotenv

load_dotenv()

EXPERT_LIST = os.environ["EXPERT_LIST"]
CONTRACT_DOC = os.environ["CONTRACT_DOC"]
RECEIPT_DOC = os.environ["RECEIPT_DOC"]
SIGNATURE_DOC = os.environ["SIGNATURE_DOC"]

# Maybe just the parts after the colon?
CONTRACT_WORDS = ["意於年月日", "申請案號：案號", "課程名稱：課程全名", "單位名稱：單位全名"]
