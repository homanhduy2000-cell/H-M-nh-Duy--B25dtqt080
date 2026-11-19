import sys
import os

# Đảm bảo thư mục gốc của project được thêm vào sys.path để import app.py và các file dữ liệu
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from app import app
