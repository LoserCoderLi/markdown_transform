from flask import Flask, request, jsonify, send_file, render_template
import os
import shutil
from datetime import datetime
import logging
from flask_cors import CORS
from util.file_operations import get_all_subdirs, clear_directory, check_and_extract_archive, get_subdirs
from util.markdown_operations import convert_markdown_to_pdf, convert_markdown_to_html, \
    convert_md_to_docx_with_toc_and_template
from util.utils import generate_unique_urlid
from util.generate import generate_latex_document_pdf, generate_parameter, create_template_with_headers

app = Flask(__name__)
CORS(app)

# 存储上传的 Markdown 文件名
uploaded_md_filename = {}


@app.route('/')
def index():
    """
    渲染主页模板。
    """
    return render_template('index.html')


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
        return jsonify({"error": "未指定文件"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "未选择文件"}), 400

    # 生成或接收唯一 URL 标识符
    urlid = request.form.get('urlid', generate_unique_urlid())
    extract_to = os.path.join(os.getcwd(), urlid)

    # 创建目录或清空现有目录
    if not os.path.exists(extract_to):
        os.makedirs(extract_to, exist_ok=True)
    else:
        clear_directory(extract_to)  # 清空已有目录内容

    # 暂存上传的文件
    zip_path = os.path.join('/tmp', file.filename)
    file.save(zip_path)

    # 检查并解压文件
    result = check_and_extract_archive(zip_path, extract_to)
    os.remove(zip_path)

    if result:
        # 找到上传的 Markdown 文件名并保存
        md_file_name = next(file for file in os.listdir(extract_to) if file.endswith('.md'))
        uploaded_md_filename[urlid] = md_file_name
        return jsonify({"success": f"文件已上传并解压至 {extract_to}", "urlid": urlid}), 200
    else:
        return jsonify({"error": "解压失败或未找到有效的 .md 文件"}), 400


@app.route('/cleanup', methods=['POST'])
def cleanup():
    """
    处理清理请求，删除与特定 URL 标识符（urlid）相关的目录。

    请求:
        POST /cleanup

    返回:
        包含清理状态的 JSON 响应。
    """
    try:
        data = request.get_json()
        urlid = data.get('urlid')
    except Exception as e:
        return jsonify({"error": "解析请求数据失败"}), 400

    if not urlid:
        return jsonify({"error": "未指定urlid"}), 400

    directory_path = os.path.join(os.getcwd(), urlid)
    directory_path_out = os.path.join(os.getcwd(), urlid) + "_out"
    directory_path_template = os.path.join(os.getcwd(), urlid) + "_template"

    if os.path.exists(directory_path):
        try:
            # 删除目录
            shutil.rmtree(directory_path)
            shutil.rmtree(directory_path_out)
            shutil.rmtree(directory_path_template)
            # 从全局字典中删除相关条目
            if urlid in uploaded_md_filename:
                del uploaded_md_filename[urlid]
            return jsonify({"success": f"与 {urlid} 相关的目录已删除"}), 200
        except Exception as e:
            return jsonify({"error": f"删除与 {urlid} 相关的目录失败: {str(e)}"}), 500
    else:
        return jsonify({"error": f"与 {urlid} 相关的目录不存在"}), 400


@app.route('/convert', methods=['POST'])
def convert_file():
    """
    根据请求将 Markdown 文件转换为指定格式（pdf、html、docx）。

    请求:
        POST /convert

    返回:
        生成的文件或错误信息。
    """
    if 'format' not in request.form:
        return jsonify({"error": "未指定格式"}), 400

    output_format = request.form['format']

    if output_format not in ['pdf', 'html', 'docx']:
        return jsonify({"error": "格式无效"}), 400

    # 获取请求参数
    title = request.form.get('title', 'Document Title')
    author = request.form.get('author', 'Author Name')
    statement = request.form.get('statement', '')
    left_header = request.form.get('left_header', 'Left Header')
    right_header = request.form.get('right_header', 'Right Header')
    cover_footer = request.form.get('cover_footer', 'Cover Footer')

    urlid = request.form.get('urlid')
    extract_to = os.path.join(os.getcwd(), urlid)
    output_directory = os.path.join(os.getcwd(), f'{urlid}_out')
    template_directory = os.path.join(os.getcwd(), f'{urlid}_template')
    os.makedirs(output_directory, exist_ok=True)
    os.makedirs(template_directory, exist_ok=True)

    # 获取资源路径
    resource_paths = get_all_subdirs(extract_to)
    resource_paths.append(os.path.abspath(extract_to))
    imgs_dir = get_subdirs(extract_to)

    if imgs_dir:
        resource_paths.append(os.path.join(extract_to, imgs_dir[0]))

    # 找到输入的 Markdown 文件
    input_file = os.path.join(extract_to, uploaded_md_filename.get(urlid))
    output_file = os.path.join(output_directory, os.path.basename(input_file).replace(".md", f".{output_format}"))

    # 设置 logo 路径
    logo_file = request.files.get('logo')
    logo_path = None
    print(logo_file)
    if logo_file:
        logo_path = os.path.join(extract_to, 'logo.png')
        logo_file.save(logo_path)

    # 生成参数
    parameter = generate_parameter(title=title, author=author, statement=statement)

    # 根据指定格式转换 Markdown 文件
    if output_format == "pdf":
        tex_path = generate_latex_document_pdf(
            left_header=left_header,
            right_header=right_header,
            cover_footer=cover_footer,
            urlid=template_directory,
        )
        convert_markdown_to_pdf(
            input_file=input_file,
            title=parameter["title"],
            author=parameter["author"],
            date=parameter["date"],
            output_file=output_file,
            header_file=os.path.join(os.getcwd(), tex_path),
            logo_path=logo_path,
            resource_paths=resource_paths,
            statement=parameter["statement"]
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
            title=title,
            author=author,
            date=datetime.now().strftime("%Y-%m-%d"),
            statement=statement,
        )
        convert_md_to_docx_with_toc_and_template(
            md_file_path=input_file,
            docx_file_path=output_file,
            template_file_path=template_file_path,
            title=title,
            author=author,
            date=datetime.now().strftime("%Y-%m-%d"),
            left_header=left_header,
            right_header=right_header,
            statement=statement,
            resource_paths=resource_paths,
            logo_path=logo_path
        )

    # 返回生成的文件
    if os.path.exists(output_file):
        return send_file(output_file, as_attachment=True, mimetype=f'application/{output_format}'), 200
    else:
        return jsonify({"error": f"{output_format.upper()} 文件未创建"}), 500


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    app.run(debug=True)
