from flask import Flask, request, jsonify, send_file, render_template, url_for, after_this_request
import os
from templates import config
import logging
from logging.handlers import RotatingFileHandler  # 日志文件旋转处理器
from flask_cors import CORS  # 跨域资源共享
from util.file_operations import get_all_subdirs, clear_directory, check_and_extract_archive, get_subdirs
from util.markdown_operations import convert_markdown_to_pdf_use, convert_markdown_to_pdf_basic, convert_markdown_to_html, convert_md_to_docx_with_toc_and_template
from util.utils import generate_unique_urlid, get_uploaded_file_path
from util.generate import generate_latex_document_pdf, generate_latex_document_no_header_footer, generate_parameter, create_template_with_headers
import shutil
from datetime import datetime, timedelta  # 日期和时间处理
from werkzeug.utils import secure_filename  # 文件名安全处理
import schedule  # 任务调度
import time
import threading  # 线程处理
import portalocker
import traceback


# 创建Flask应用实例，指定静态文件和模板文件的目录
app = Flask(__name__, static_folder="templates/assets", template_folder="templates")
CORS(app)  # 允许跨域资源共享

# 存储上传的 Markdown 文件名
uploaded_md_filename = {}

# 配置日志记录
if not os.path.exists('logs'):  # 如果日志目录不存在，创建日志目录
    os.makedirs('logs')

# 获取当前日期
current_date = datetime.now().strftime('%Y-%m-%d')

# 配置日志文件旋转处理器，包含日期信息
log_handler = RotatingFileHandler(f'logs/application_{current_date}.log', maxBytes=1000000, backupCount=5)
log_handler.setLevel(logging.INFO)  # 设置日志记录级别为INFO
log_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')  # 设置日志格式
log_handler.setFormatter(log_format)
app.logger.addHandler(log_handler)  # 将处理器添加到Flask应用的日志记录器中


def setup_logger(name):
    """
    创建和配置单独的日志记录器。
    """
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    handler = RotatingFileHandler(f'logs/{name}_{current_date}.log', maxBytes=1000000, backupCount=5)
    handler.setFormatter(log_format)
    logger.addHandler(handler)
    return logger

# 创建多个日志记录器，用于不同的日志记录
index_logger = setup_logger('index')
upload_logger = setup_logger('upload')
convert_logger = setup_logger('convert')
download_logger = setup_logger('download')


@app.route('/')
def index():
    """
    渲染主页模板。
    """
    index_logger.info("Rendering index page")
    return render_template('index.html')


def add_uploaded_file_record(urlid, md_filename):
    try:
        file_path = get_uploaded_file_path()
        with open(file_path, 'a') as f:
            portalocker.lock(f, portalocker.LOCK_EX)  # 排他锁
            f.write(f"{urlid},{md_filename}\n")
            portalocker.unlock(f)  # 释放锁
    except Exception as e:
        upload_logger.error(f"Error while adding uploaded file record: {e}")
        upload_logger.error(traceback.format_exc())


def get_md_filename(urlid):
    md_filename = None
    try:
        current_date = datetime.now().strftime('%Y-%m-%d')
        file_path = f'uploaded_files_{current_date}.txt'
        with open(file_path, 'r') as f:
            portalocker.lock(f, portalocker.LOCK_SH)  # 共享锁
            lines = f.readlines()  # 读取所有行
            for line in lines:
                line_urlid, line_md_filename = line.strip().split(',')
                if line_urlid == urlid:
                    md_filename = line_md_filename
                    break
            f.seek(0)  # 重置文件指针位置
            portalocker.unlock(f)  # 释放锁
    except Exception as e:
        convert_logger.error(f"Error while getting markdown filename: {e}")
        convert_logger.error(traceback.format_exc())
    return md_filename


@app.route('/upload', methods=['POST'])
def upload_file():
    """
    处理文件上传，解压缩文件并返回唯一标识符（urlid）。

    请求:
        POST /upload

    返回:
        包含上传和解压状态及唯一标识符的 JSON 响应。
    """
    if 'file' not in request.files:
        upload_logger.error("No file part in the request")
        return jsonify({"error": "未指定文件"}), 400

    file = request.files['file']

    if file.filename == '':
        upload_logger.error("No selected file")
        return jsonify({"error": "未选择文件"}), 400

    urlid = request.form.get('urlid', generate_unique_urlid())  # 获取或生成唯一标识符
    base_dir = os.path.join(os.getcwd(), 'trans_docx')  # 基础目录为 trans_docx/
    # base_dir = os.getcwd()
    extract_to = os.path.join(base_dir, urlid)  # 解压目标路径

    if not os.path.exists(extract_to):
        os.makedirs(extract_to, exist_ok=True)
    else:
        clear_directory(extract_to)  # 清空目标目录

    temp_dir = os.path.join(os.getcwd(), 'trans_docx', 'temp')  # 临时目录
    # temp_dir = os.path.join(os.getcwd(), 'temp')  # 临时目录

    os.makedirs(temp_dir, exist_ok=True)

    zip_path = os.path.join(temp_dir, secure_filename(file.filename))  # 安全处理后的文件路径
    file.save(zip_path)  # 保存上传文件

    result = check_and_extract_archive(zip_path, extract_to)  # 解压文件
    os.remove(zip_path)  # 删除临时压缩文件

    if result:
        try:
            md_file_name = next(file for file in os.listdir(extract_to) if file.endswith('.md'))  # 获取Markdown文件名
            uploaded_md_filename[urlid] = md_file_name
            str_name = uploaded_md_filename[urlid].split(".")
            upload_logger.info(f"File uploaded and extracted successfully: {md_file_name}, urlid: {urlid}")

            add_uploaded_file_record(urlid=urlid, md_filename=md_file_name)  # 记录上传的文件信息

            return jsonify({"success": f"文件已上传并解压至 {extract_to}", "urlid": urlid, "name": str_name[0]}), 200
        except StopIteration:
            upload_logger.error("No valid .md file found in the archive")
            return jsonify({"error": "未找到有效的 .md 文件"}), 400
    else:
        upload_logger.error("File extraction failed")
        return jsonify({"error": "解压失败"}), 400


@app.route('/convert', methods=['POST'])
def convert_file():
    """
    根据请求将 Markdown 文件转换为指定格式（pdf、html、docx）。

    请求:
        POST /convert

    返回:
        生成的文件或错误信息。
    """
    try:
        if 'output_format' not in request.form:
            convert_logger.error("No output format specified")
            return jsonify({"error": "未指定格式"}), 400

        output_format = request.form['output_format']

        if output_format not in ['pdf', 'html', 'docx']:
            convert_logger.error("Invalid format specified")
            return jsonify({"error": "格式无效"}), 400

        # 获取是否使用模板的变量
        use_template = request.form.get('use_template', '1')  # 默认值为 '1'

        title = request.form.get('title', 'Document Title')  # 获取文档标题
        version = request.form.get('version', '版本号: 1.0')  # 获取版本号
        statement = request.form.get('statement', '')  # 获取声明
        left_header = request.form.get('left_header', 'Left Header')  # 获取左侧页眉
        right_header = request.form.get('right_header', 'Right Header')  # 获取右侧页眉
        cover_footer = request.form.get('cover_footer', 'Cover Footer')  # 获取封面页脚

        urlid = request.form.get('urlid')

        base_dir = os.path.join(os.getcwd(), 'trans_docx')  # 基础目录为 trans_docx/
        # base_dir = os.getcwd()
        extract_to = os.path.join(base_dir, urlid)  # 解压目录
        output_directory = os.path.join(base_dir, f'{urlid}_out')  # 输出目录
        template_directory = os.path.join(base_dir, f'{urlid}_template')  # 模板目录
        os.makedirs(output_directory, exist_ok=True)
        os.makedirs(template_directory, exist_ok=True)

        resource_paths = get_all_subdirs(extract_to)  # 获取所有子目录
        resource_paths.append(os.path.abspath(extract_to))
        imgs_dir = get_subdirs(extract_to)

        if imgs_dir:
            resource_paths.append(os.path.join(extract_to, imgs_dir[0]))

        md_filename = get_md_filename(urlid=urlid)
        if md_filename is None:
            convert_logger.error("No markdown file found for the given URLID")
            return jsonify({"error": "未找到与urlid相关的Markdown文件"}), 400

        input_file = os.path.join(extract_to, md_filename)  # 输入文件路径
        output_file = os.path.join(output_directory, os.path.basename(input_file).replace(".md", f".{output_format}"))  # 输出文件路径

        
        logo_file = request.files.get('logo')  # 获取Logo文件
        logo_path = None
        if logo_file:
            logo_path = os.path.join(extract_to, 'logo.png')
            logo_file.save(logo_path)
            logo_path = logo_path.replace("\\", "/")

        parameter = generate_parameter(title=title, version=version, statement=statement)  # 生成参数

        if output_format == "pdf":
            if use_template == '1':
                tex_path = generate_latex_document_pdf(
                    left_header=left_header,
                    right_header=right_header,
                    cover_footer=cover_footer,
                    urlid=template_directory,
                )
                # print(os.path.join(os.getcwd(), tex_path))
                convert_markdown_to_pdf_use(
                    urlid=urlid,
                    input_file=input_file,
                    title=parameter["title"],
                    version=parameter["version"],
                    date=parameter["date"],
                    output_file=output_file,
                    header_file=os.path.join(os.getcwd(), tex_path),
                    logo_path=logo_path,
                    resource_paths=resource_paths,
                    statement=parameter["statement"]
                )
            else:
                tex_path = generate_latex_document_no_header_footer(
                    urlid=urlid
                )
                convert_markdown_to_pdf_basic(
                    urlid=urlid,
                    input_file=input_file,
                    output_file=output_file,
                    header_file=os.path.join(os.getcwd(), tex_path),
                    resource_paths=resource_paths,
                )
            
        elif output_format == "html":
            convert_markdown_to_html(
                input_file=input_file,
                output_file=output_file,
                resource_paths=resource_paths,
                title=parameter["title"]
            )
        elif output_format == "docx":
            template_file_path = os.path.join(template_directory, 'template_with_headers.docx')
            create_template_with_headers(
                template_path=template_file_path,
                left_header=left_header,
                right_header=right_header,
            )
            convert_md_to_docx_with_toc_and_template(
                md_file_path=input_file,
                docx_file_path=output_file,
                template_file_path=template_file_path,
                title=title,
                version=version,
                date=datetime.now().strftime("%Y-%m-%d"),
                left_header=left_header,
                right_header=right_header,
                statement=statement,
                resource_paths=resource_paths,
                logo_path=logo_path
            )

        if not os.path.exists(output_file):
            convert_logger.error(f"{output_format.upper()} file not created")
            return jsonify({"error": f"{output_format.upper()} 文件未创建"}), 500

        download_link = url_for('download_file', urlid=urlid, filename=os.path.basename(output_file), _external=True)  # 生成下载链接
        convert_logger.info(f"File converted successfully: {output_file}")
        return jsonify({"download_link": download_link}), 200

    except Exception as e:
        convert_logger.error(f"Internal server error: {e}")
        return jsonify({"error": "内部服务器错误"}), 500


@app.route('/download/<urlid>/<filename>')
def download_file(urlid, filename):
    """
    下载指定文件。

    请求:
        GET /download/<urlid>/<filename>

    返回:
        下载文件。
    """
    # os.remove('temp')
    output_directory = os.path.join(os.getcwd(), 'trans_docx', f'{urlid}_out')
    # output_directory = os.path.join(os.getcwd(), f'{urlid}_out')

    file_path = os.path.join(output_directory, filename)

    if os.path.exists(file_path):
        download_logger.info(f"File downloaded: {file_path}")
        return send_file(file_path, as_attachment=True)
    else:
        download_logger.error(f"File not found: {file_path}")
        return jsonify({"error": "文件未找到"}), 404


@app.route('/download_file')
def download_user_tools():
    """
    提供 intewell_user_tools.pdf 文件的下载链接。
    """
    # 获取文件的绝对路径
    file_path = os.path.join(os.getcwd(), 'intewell_user_tools.pdf')

    # 打印文件路径以调试
    print(f"Trying to download file from: {file_path}")

    try:
        # 检查文件是否存在
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            print("File not found")
            return jsonify({"error": "File not found"}), 404
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        return jsonify({"error": str(e)}), 500



def delete_previous_day_directories():
    """
    删除前一天的URL目录。
    """
    previous_day = (datetime.now() - timedelta(1)).strftime('%Y%m%d')
    script_dir = os.path.join(os.getcwd(), 'trans_docx')
    # script_dir = os.getcwd()

    # print(script_dir)
    app.logger.info(f"Checking for directories to delete for date: {previous_day} in base directory: {script_dir}")

    # 遍历基础目录中的所有目录
    for dir_name in os.listdir(script_dir):
        dir_path = os.path.join(script_dir, dir_name)
        # 如果目录名以前一天的日期开头，并且是一个目录，则删除它
        if os.path.isdir(dir_path) and dir_name.startswith(previous_day):
            try:
                shutil.rmtree(dir_path)
                app.logger.info(f"Deleted directory: {dir_path}")
            except Exception as e:
                app.logger.error(f"Failed to delete {dir_path}: {e}")
        else:
            app.logger.info(f"Skipping directory: {dir_path}")


def delete_old_logs():
    """
    删除超过3天的日志文件。
    """
    logs_dir = os.path.join(os.getcwd(), 'logs')
    current_time = datetime.now()

    for log_file in os.listdir(logs_dir):
        log_file_path = os.path.join(logs_dir, log_file)
        if os.path.isfile(log_file_path):
            # 提取日志文件中的日期部分
            try:
                log_date_str = log_file.split('_')[-1].replace('.log', '')
                log_date = datetime.strptime(log_date_str, '%Y-%m-%d')
                # 计算日志文件的年龄
                file_age = (current_time - log_date).days
                if file_age >= 3:
                    os.remove(log_file_path)
                    app.logger.info(f"Deleted old log file: {log_file_path}")
            except Exception as e:
                app.logger.error(f"Error while deleting old log file {log_file_path}: {e}")


def delete_previous_day_uploaded_files():
    """
    删除前一天的 uploaded_files.txt 文件。
    """
    previous_day = (datetime.now() - timedelta(1)).strftime('%Y-%m-%d')
    file_path = f'uploaded_files_{previous_day}.txt'

    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            app.logger.info(f"Deleted uploaded files record: {file_path}")
        except Exception as e:
            app.logger.error(f"Failed to delete {file_path}: {e}")
    else:
        app.logger.info(f"No uploaded files record found for {previous_day} to delete.")


def schedule_tasks(stop_event):
    """
    安排定时任务。
    """
    schedule.every().day.at("03:00").do(delete_previous_day_directories)  # 每天凌晨1点删除前一天的目录
    schedule.every().day.at("03:00").do(delete_previous_day_uploaded_files)  # 每天凌晨1点10分删除前一天的 uploaded_files 文件
    schedule.every().day.at("03:00").do(delete_old_logs)  # 每天凌晨1点30分删除超过3天的日志文件

    app.logger.info("Scheduled daily directory and files cleanup at 03:00 AM.")

    while not stop_event.is_set():
        schedule.run_pending()
        time.sleep(1)

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)

    # 创建一个停止事件
    stop_event = threading.Event()

    # 启动后台线程运行定时任务
    task_thread = threading.Thread(target=schedule_tasks, args=(stop_event,), daemon=True)
    task_thread.start()

    try:
        app.run(host=config.HOST, port=config.PORT, debug=config.DEBUG, use_reloader=False)
    finally:
        stop_event.set()  # 停止后台线程
        task_thread.join()  # 确保后台线程在应用关闭时正确退出
