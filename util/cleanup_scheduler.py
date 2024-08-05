import os
import shutil
import schedule
import time
from datetime import datetime, timedelta

def delete_previous_day_directories():
    """
    删除前一天的URL目录。
    """
    previous_day = (datetime.now() - timedelta(1)).strftime('%Y%m%d')
    script_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前脚本所在目录
    base_dir = os.path.dirname(script_dir)  # 获取上一级目录
    print(f"Checking for directories to delete for date: {previous_day} in base directory: {base_dir}")

    # 遍历基础目录中的所有目录
    for dir_name in os.listdir(base_dir):
        dir_path = os.path.join(base_dir, dir_name)
        # 如果目录名以前一天的日期开头，并且是一个目录，则删除它
        if os.path.isdir(dir_path) and dir_name.startswith(previous_day):
            try:
                shutil.rmtree(dir_path)
                print(f"Deleted directory: {dir_path}")
            except Exception as e:
                print(f"Failed to delete {dir_path}: {e}")
        else:
            print(f"Skipping directory: {dir_path}")

# # 手动调用删除函数以确认其工作正常
# delete_previous_day_directories()


# 安排每天凌晨1点运行删除任务
schedule.every().day.at("01:00").do(delete_previous_day_directories)

print("Scheduled daily directory cleanup at 01:00 AM.")

# 保持程序运行，以便调度器可以工作
while True:
    # 运行所有调度的任务
    schedule.run_pending()
    # 打印当前系统时间
    print("Current system time:", datetime.now())
    time.sleep(1)
