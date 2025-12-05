#!/usr/bin/env python3
# compare_tool.py
# Requirements: python3, Flask, pandas, openpyxl

import os
import sys
import tempfile
import uuid
import webbrowser
import re
from datetime import datetime
from flask import (
    Flask, request, redirect, url_for, send_file,
    render_template, flash, make_response, jsonify
)
import pandas as pd
from pandas import ExcelFile

# 计算运行时模板/静态资源路径（兼容开发环境与 PyInstaller onefile）
def get_resource_path(relative_path: str) -> str:
    """
    Return absolute path to resource, whether running normally or frozen by PyInstaller.
    For onefile builds, PyInstaller extracts files to sys._MEIPASS.
    """
    if getattr(sys, "_MEIPASS", None):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

# 使用辅助函数来设置 Flask 的 template_folder 与 static_folder
template_folder = get_resource_path("templates")
static_folder = get_resource_path("static")

# --- diagnostic: print resource locations and existence (临时，用于排查 onefile 问题) ---
print("DEBUG: sys._MEIPASS =", getattr(sys, "_MEIPASS", None))
print("DEBUG: template_folder resolved to:", template_folder, " exists:", os.path.exists(template_folder))
print("DEBUG: static_folder resolved to:", static_folder, " exists:", os.path.exists(static_folder))
# ---------------------------------------------------------------------------------------


# 然后初始化 Flask（替换原来的 app = Flask(...)）：
app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
app.secret_key = "replace-this-with-random-if-needed"
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "py_excel_compare")
os.makedirs(UPLOAD_DIR, exist_ok=True)


# Simplified i18n
TRANSLATIONS = {
    'zh': {
        'title': '本地化库存比对工具',
        'local_run': '(本地运行)',
        'copyright': '© kuriyamasss',
        'upload_old': '上传 - 旧 库存表（Excel 或 CSV）：',
        'upload_new': '上传 - 新 库存表（Excel 或 CSV）：',
        'upload_hint': '两表上传后，会读取 sheet 列表（若为 Excel）并在下一步呈现 sheet 与表头选择。',
        'upload_button': '上传并读取文件信息',
        'step2_title': '第二步：选择要比对的 Sheet 与 欄位 (Key)',
        'sheet_old_label': '旧表 Sheet（若为 CSV 则无选项）：',
        'sheet_new_label': '新表 Sheet：',
        'header_identify': '表头识别：',
        'header_auto': '自动检测（建议）',
        'header_manual': '手动指定表头列（0 起算）',
        'header_none': '无表头（整表为数据）',
        'header_row_index': '若手动指定，请输入表头行索引（0 起算）：',
        'prepare_fields': '读取栏位并产生可选键值',
        'step3_title': '第三步：选择用作比对的键值（Key）',
        'select_key': '可选栏位（两表共有）：',
        'dup_strategy': '重复键处理策略：',
        'dup_last': "保留最后出现（drop_duplicates keep='last')",
        'dup_first': "保留第一次出现（keep='first')",
        'dup_error': "若发现重复则停止并提示（需要唯一键）",
        'compare_button': '比对并产生报表',
        'download_ready': '完成： 比对文件已生成，点击下载结果 Excel：',
        'large_hint': '提示：若资料量很大（数十万列），建议先用 CSV 并确保可用内存充足。',
        'processing': '数据处理中，请稍候…',
        'no_common_columns': '两表没有共同栏位，请确认 sheet 与 header 设置。',
        'no_files': '请同时上传旧表与新表。',
        'save_failed': '存档失败: ',
        'read_failed': '读取表格失败: ',
        'temp_missing': '暂存档案不存在，请重新上傳。',
        'key_missing': '所选键值在上传的表格中不存在，请重新上传并选择正确栏位。',
        'dup_found': '发现重复键值： 请先在原始档处理唯一键或改变重复策略。',
        'generate_failed': '生成结果档失败: ',
        'download_failed': '下载失败：',
        'lang_label': '语言',
        'download_button': '下载结果',
        'reset_button': '重新处理数据',
        'back_button': '上一步',
        'close_button': '关闭程序'
    },
    'vi': {
        'title': 'Công cụ đối chiếu tồn kho cục bộ',
        'local_run': '(Chạy cục bộ)',
        'copyright': '© Yining.li',
        'upload_old': 'Tải lên - Bảng tồn kho cũ (Excel hoặc CSV):',
        'upload_new': 'Tải lên - Bảng tồn kho mới (Excel hoặc CSV):',
        'upload_hint': 'Sau khi tải lên, sẽ đọc danh sách sheet (nếu là Excel) và hiển thị lựa chọn sheet / header ở bước tiếp theo.',
        'upload_button': 'Tải lên và đọc thông tin tệp',
        'step2_title': 'Bước 2: Chọn Sheet và trường để đối chiếu (Key)',
        'sheet_old_label': 'Sheet bảng cũ (nếu là CSV thì không có):',
        'sheet_new_label': 'Sheet bảng mới:',
        'header_identify': 'Phát hiện header:',
        'header_auto': 'Tự động (khuyến nghị)',
        'header_manual': 'Chỉ định thủ công dòng header (đếm từ 0)',
        'header_none': 'Không có header (toàn bộ là dữ liệu)',
        'header_row_index': 'Nếu chỉ định thủ công, nhập chỉ số dòng header (bắt đầu từ 0):',
        'prepare_fields': 'Đọc trường và tạo danh sách key có thể chọn',
        'step3_title': 'Bước 3: Chọn khóa (Key) để đối chiếu',
        'select_key': 'Các trường có thể chọn (xuất hiện ở cả 2 bảng):',
        'dup_strategy': 'Chiến lược xử lý key trùng:',
        'dup_last': "Giữ lần xuất hiện cuối cùng (drop_duplicates keep='last')",
        'dup_first': "Giữ lần xuất hiện đầu tiên (keep='first')",
        'dup_error': "Nếu phát hiện trùng thì dừng và báo lỗi (yêu cầu khóa duy nhất)",
        'compare_button': 'Đối chiếu và tạo báo cáo',
        'download_ready': 'Hoàn thành: Tệp đối chiếu đã được tạo, bấm để tải Excel kết quả:',
        'large_hint': 'Gợi ý: nếu dữ liệu rất lớn (hàng trăm nghìn), khuyến nghị dùng CSV và đảm bảo đủ RAM.',
        'processing': 'Đang xử lý dữ liệu, vui lòng chờ…',
        'no_common_columns': 'Hai bảng không có cột chung, vui lòng kiểm tra sheet và thiết lập header.',
        'no_files': 'Vui lòng tải cả bảng cũ và bảng mới.',
        'save_failed': 'Lưu tệp thất bại: ',
        'read_failed': 'Đọc bảng thất bại: ',
        'temp_missing': 'Tệp tạm thời không tồn tại, vui lòng tải lại.',
        'key_missing': 'Khóa đã chọn không tồn tại trong bảng đã tải lên, vui lòng kiểm tra lại.',
        'dup_found': 'Phát hiện khóa trùng: Vui lòng xử lý khóa duy nhất ở nguồn hoặc thay đổi chiến lược.',
        'generate_failed': 'Tạo tệp kết quả thất bại: ',
        'download_failed': 'Tải xuống thất bại:',
        'lang_label': 'Ngôn ngữ',
        'download_button': 'Tải xuống kết quả',
        'reset_button': 'Xử lý lại dữ liệu',
        'back_button': 'Quay lại',
        'close_button': 'Đóng chương trình'
    }
}



def get_lang_from_request():
    lang = request.args.get('lang')
    if lang and lang in TRANSLATIONS:
        return lang
    lang = request.cookies.get('lang')
    if lang and lang in TRANSLATIONS:
        return lang
    return 'zh'

def t(key):
    return TRANSLATIONS.get(get_lang_from_request(), TRANSLATIONS['zh']).get(key, key)

def safe_save_upload(file_storage, prefix):
    ext = os.path.splitext(file_storage.filename)[1].lower()
    fname = f"{prefix}_{uuid.uuid4().hex}{ext}"
    path = os.path.join(UPLOAD_DIR, fname)
    file_storage.save(path)
    return path

def detect_header_row_from_df(df_preview, min_nonnull_ratio=0.4):
    nonnull_counts = df_preview.notna().sum(axis=1)
    max_cols = df_preview.shape[1]
    threshold = max(1, int(max_cols * min_nonnull_ratio))
    for idx, cnt in enumerate(nonnull_counts):
        if cnt >= threshold:
            return idx
    return 0

def read_table(path, sheet_name=0, header_mode='auto', header_row_index=None, preview_rows=10):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in ('.xls', '.xlsx'):
            preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=preview_rows, engine='openpyxl')
            if header_mode == 'auto':
                detected = detect_header_row_from_df(preview)
            elif header_mode == 'manual' and header_row_index is not None:
                detected = int(header_row_index)
            else:
                detected = None

            if detected is None:
                df = pd.read_excel(path, sheet_name=sheet_name, dtype=str, engine='openpyxl')
            else:
                header_block = preview.iloc[: detected + 1].fillna(method='ffill', axis=1).astype(str)
                combined_header = []
                for col in range(header_block.shape[1]):
                    parts = [str(header_block.iloc[r, col]).strip() for r in range(header_block.shape[0])]
                    parts = [p for p in parts if p and p.lower() != 'nan']
                    combined = " ".join(parts).strip()
                    if combined == "":
                        combined = f"col_{col}"
                    combined_header.append(combined)
                df = pd.read_excel(path, sheet_name=sheet_name, header=detected, dtype=str, engine='openpyxl')
                if len(combined_header) == df.shape[1]:
                    df.columns = combined_header
                else:
                    df.columns = [str(c).strip() for c in df.columns]
        elif ext == '.csv':
            preview = pd.read_csv(path, header=None, nrows=preview_rows, dtype=str, engine='python', encoding='utf-8')
            if header_mode == 'auto':
                detected = detect_header_row_from_df(preview)
            elif header_mode == 'manual' and header_row_index is not None:
                detected = int(header_row_index)
            else:
                detected = None
            if detected is None:
                df = pd.read_csv(path, dtype=str, engine='python', encoding='utf-8')
            else:
                df = pd.read_csv(path, header=detected, dtype=str, engine='python', encoding='utf-8')
            df.columns = [str(c).strip() for c in df.columns]
        else:
            df = pd.read_excel(path, dtype=str, engine='openpyxl')
            df.columns = [str(c).strip() for c in df.columns]
    except Exception:
        raise
    df.columns = [str(c).strip() for c in df.columns]
    return df

@app.context_processor
def inject_helpers():
    return dict(t=lambda k: TRANSLATIONS[get_lang_from_request()].get(k, k),
                get_lang=get_lang_from_request())

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html', old_id=None, new_id=None)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file_old' not in request.files or 'file_new' not in request.files:
        flash(t('no_files'))
        return redirect(url_for('index'))
    f_old = request.files['file_old']
    f_new = request.files['file_new']
    try:
        old_path = safe_save_upload(f_old, "old")
        new_path = safe_save_upload(f_new, "new")
    except Exception as e:
        flash(t('save_failed') + str(e))
        return redirect(url_for('index'))

    def get_sheets(path):
        ext = os.path.splitext(path)[1].lower()
        if ext in ('.xls', '.xlsx'):
            try:
                x = ExcelFile(path, engine='openpyxl')
                return x.sheet_names
            except Exception:
                return []
        else:
            return []

    sheets_old = get_sheets(old_path)
    sheets_new = get_sheets(new_path)

    old_id = os.path.basename(old_path)
    new_id = os.path.basename(new_path)

    resp = make_response(render_template('index.html', old_id=old_id, new_id=new_id,
                                  sheets_old=sheets_old, sheets_new=sheets_new,
                                  header_mode='auto', header_row_index=None))
    lang = request.args.get('lang')
    if lang and lang in TRANSLATIONS:
        resp.set_cookie('lang', lang, max_age=365*24*60*60, path='/')
    return resp

@app.route('/prepare_fields', methods=['POST'])
def prepare_fields():
    old_id = request.form.get('old_id')
    new_id = request.form.get('new_id')
    sheet_old = request.form.get('sheet_old') or 0
    sheet_new = request.form.get('sheet_new') or 0
    header_mode = request.form.get('header_mode', 'auto')
    header_row_index = request.form.get('header_row_index', None)

    if not all([old_id, new_id]):
        flash(t('no_files'))
        return redirect(url_for('index'))

    old_path = os.path.join(UPLOAD_DIR, old_id)
    new_path = os.path.join(UPLOAD_DIR, new_id)
    try:
        df_old = read_table(old_path, sheet_name=sheet_old, header_mode=header_mode, header_row_index=header_row_index)
        df_new = read_table(new_path, sheet_name=sheet_new, header_mode=header_mode, header_row_index=header_row_index)
    except Exception as e:
        flash(t('read_failed') + str(e))
        return redirect(url_for('index'))

    headers_common = [h for h in df_old.columns if h in df_new.columns]
    if not headers_common:
        flash(t('no_common_columns'))
        return redirect(url_for('index'))

    return render_template('index.html', headers=headers_common,
                                  old_id=old_id, new_id=new_id,
                                  sheet_old=sheet_old, sheet_new=sheet_new,
                                  header_mode=header_mode, header_row_index=header_row_index)

def sanitize_filename_component(s: str) -> str:
    if s is None:
        return ''
    s = str(s)
    s = s.strip()
    s = re.sub(r'\s+', '_', s)
    s = re.sub(r'[^\w\-.]', '', s)
    return s[:100]

@app.route('/compare', methods=['POST'])
def compare():
    key = request.form.get('key')
    dup_strategy = request.form.get('dup_strategy', 'last')
    old_id = request.form.get('old_id')
    new_id = request.form.get('new_id')
    sheet_old = request.form.get('sheet_old') or 0
    sheet_new = request.form.get('sheet_new') or 0
    header_mode = request.form.get('header_mode', 'auto')
    header_row_index = request.form.get('header_row_index', None)

    if not all([key, old_id, new_id]):
        flash(t('key_missing'))
        return redirect(url_for('index'))
    old_path = os.path.join(UPLOAD_DIR, old_id)
    new_path = os.path.join(UPLOAD_DIR, new_id)
    if not os.path.exists(old_path) or not os.path.exists(new_path):
        flash(t('temp_missing'))
        return redirect(url_for('index'))
    try:
        df_old = read_table(old_path, sheet_name=sheet_old, header_mode=header_mode, header_row_index=header_row_index)
        df_new = read_table(new_path, sheet_name=sheet_new, header_mode=header_mode, header_row_index=header_row_index)
    except Exception as e:
        flash(t('read_failed') + str(e))
        return redirect(url_for('index'))

    if key not in df_old.columns or key not in df_new.columns:
        flash(t('key_missing'))
        return redirect(url_for('index'))

    if dup_strategy == 'error':
        dup_old = df_old[df_old.duplicated(subset=[key], keep=False)]
        dup_new = df_new[df_new.duplicated(subset=[key], keep=False)]
        if not dup_old.empty or not dup_new.empty:
            flash(t('dup_found'))
            return redirect(url_for('index'))

    df_old[key] = df_old[key].astype(str).str.strip()
    df_new[key] = df_new[key].astype(str).str.strip()

    keep_choice = 'last' if dup_strategy == 'last' else 'first'
    df_old_u = df_old.drop_duplicates(subset=[key], keep=keep_choice).set_index(key, drop=False)
    df_new_u = df_new.drop_duplicates(subset=[key], keep=keep_choice).set_index(key, drop=False)

    idx_old = set(df_old_u.index)
    idx_new = set(df_new_u.index)
    common_idx = sorted(idx_old.intersection(idx_new))
    only_old_idx = sorted(idx_old - idx_new)
    only_new_idx = sorted(idx_new - idx_old)

    df_old_u2 = df_old_u.set_index(key)
    df_new_u2 = df_new_u.set_index(key)
    df_both = df_old_u2.loc[common_idx].join(df_new_u2.loc[common_idx], lsuffix='_old', rsuffix='_new', how='inner')
    df_only_old = df_old_u2.loc[only_old_idx]
    df_only_new = df_new_u2.loc[only_new_idx]

    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_key = sanitize_filename_component(key)
    n_both = len(common_idx)
    n_old = len(only_old_idx)
    n_new = len(only_new_idx)
    out_filename = f"compare_{now}_{safe_key}_both{n_both}_old{n_old}_new{n_new}.xlsx"
    out_path = os.path.join(UPLOAD_DIR, out_filename)
    try:
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            df_both.to_excel(writer, sheet_name='intersection')
            df_only_old.to_excel(writer, sheet_name='only_in_old')
            df_only_new.to_excel(writer, sheet_name='only_in_new')
    except Exception as e:
        flash(t('generate_failed') + str(e))
        return redirect(url_for('index'))

    download_link = url_for('download_file', filename=out_filename)
    resp = make_response(render_template('index.html', download_link=download_link, download_name=out_filename))
    lang = request.args.get('lang')
    if lang and lang in TRANSLATIONS:
        resp.set_cookie('lang', lang, max_age=365*24*60*60, path='/')
    return resp

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    path = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(path):
        flash(t('temp_missing'))
        return redirect(url_for('index'))
    return send_file(path, as_attachment=True, download_name=filename)

# --------------------
# 新增：shutdown 路由（用于关闭本地服务器）
# 前端点击“关闭程序”时会 POST 到此处以请求优雅结束进程
# --------------------
@app.route('/shutdown', methods=['POST'])
def shutdown():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        # 如果使用非 werkzeug 服务器，尝试强制退出（在 exe 场景中一般可行）
        try:
            os._exit(0)
        except Exception:
            return jsonify({"ok": False, "msg": "shutdown not supported"}), 500
    else:
        func()
    return jsonify({"ok": True, "msg": "shutting down"}), 200

if __name__ == '__main__':
    port = 5000
    url = f"http://127.0.0.1:{port}/"
    print("正在启动本地服务，稍后会在默认浏览器打开界面：", url)
    webbrowser.open(url)
    app.run(host='127.0.0.1', port=port, debug=False)
