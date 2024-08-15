import uuid
from datetime import datetime

def generate_unique_urlid():
    """
    生成一个独一无二的urlid，前面加上当前的日期（格式：YYYYMMDD）。
    """
    current_date = datetime.now().strftime('%Y%m%d')
    unique_id = str(uuid.uuid4())
    return f"{current_date}-{unique_id}"

