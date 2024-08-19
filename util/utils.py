import uuid
from datetime import datetime

def generate_unique_urlid():
    """
    生成一个独一无二的urlid，前面加上当前的日期（格式：YYYYMMDD）。
    """
    current_date = datetime.now().strftime('%Y%m%d')
    unique_id = str(uuid.uuid4())
    return f"{current_date}-{unique_id}"


def get_uploaded_file_path():
    """
    获取当前日期的 uploaded_files.txt 文件路径。
    """
    current_date = datetime.now().strftime('%Y-%m-%d')
    return f'uploaded_files_{current_date}.txt'
