import os
import zipfile
import shutil


# def clear_directory(directory):
#     """
#     清空指定目录中的所有文件和子目录。
#
#     参数:
#         directory (str): 要清空的目录路径。
#     """
#     if os.path.exists(directory):
#         for root, dirs, files in os.walk(directory, topdown=False):
#             for name in files:
#                 os.remove(os.path.join(root, name))
#             for name in dirs:
#                 shutil.rmtree(os.path.join(root, name))


def check_and_extract_archive(zip_path, extract_to):
    """
    解压并检查ZIP文件内容是否包含.md文件。

    参数:
        zip_path (str): ZIP文件的路径。
        extract_to (str): 解压文件的目标目录。

    返回:
        bool: 如果ZIP文件包含至少一个.md文件，则返回True；否则返回False。
    """
    clear_directory(extract_to)
    os.makedirs(extract_to, exist_ok=True)
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        extracted_contents = zip_ref.namelist()
        has_md_file = any(name.endswith('.md') for name in extracted_contents)
        if has_md_file:
            zip_ref.extractall(extract_to)
        return has_md_file


def get_subdirs(directory):
    '''
    获取指定目录下的所有子目录名称。

    参数:
    directory (str): 目录路径。

    返回:
    list: 子目录名称列表。
    '''
    return [name for name in os.listdir(directory) if os.path.isdir(os.path.join(directory, name))]


# 获取指定目录下的所有子目录路径。
def get_all_subdirs(directory):
    """
    获取指定目录下的所有子目录路径。

    参数:
        directory (str): 要搜索的目录路径。

    返回:
        list: 目录下所有子目录的路径列表。
    """
    subdirs = []  # 用于存储子目录路径的列表
    for root, dirs, _ in os.walk(directory):
        # 遍历当前目录中的所有子目录
        for dir in dirs:
            subdir_path = os.path.join(root, dir)  # 获取子目录的完整路径
            if os.path.exists(subdir_path):  # 检查子目录是否存在
                subdirs.append(subdir_path)  # 将子目录路径添加到列表中
    return subdirs  # 返回子目录路径列表


def clear_directory(directory_path):
    """
    清空指定目录的所有内容。
    """
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')
