import os
import re
import zipfile
import uuid
from copy import copy
from flask import Flask, request, send_file, render_template, jsonify, after_this_request
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill, GradientFill
app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_sheet_names(file_path):
    wb = load_workbook(file_path, data_only=True)
    sheets = wb.sheetnames
    wb.close()
    return sheets

def get_columns_from_sheet(file_path, sheet_name, header_row):
    """读取指定工作表的表头行，返回列名列表（支持合并单元格）"""
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]
    column_names = []
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(header_row, col)
        # 处理合并单元格
        if cell.coordinate in ws.merged_cells:
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    top_left = ws.cell(merged_range.min_row, merged_range.min_col)
                    value = top_left.value
                    break
            else:
                value = cell.value
        else:
            value = cell.value
        if value is None:
            value = f"列{col}"
        column_names.append(str(value))
    wb.close()
    return column_names

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_sheets_and_columns', methods=['POST'])
def get_sheets_and_columns():
    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': '不支持的文件类型'}), 400

    filename = secure_filename(file.filename)
    temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{uuid.uuid4()}_{filename}")
    file.save(temp_path)

    try:
        sheet_name = request.form.get('sheet_name')
        header_row = request.form.get('header_row')
        if sheet_name and header_row:
            header_row = int(header_row)
            columns = get_columns_from_sheet(temp_path, sheet_name, header_row)
            result = {'columns': columns}
        else:
            sheets = get_sheet_names(temp_path)
            result = {'sheets': sheets}
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.remove(temp_path)
    return jsonify(result)

@app.route('/get_split_values', methods=['POST'])
def get_split_values():
    """获取拆分列的所有不重复值（支持多列组合）"""
    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': '不支持的文件类型'}), 400

    filename = secure_filename(file.filename)
    temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{uuid.uuid4()}_{filename}")
    file.save(temp_path)

    try:
        header_row = int(request.form.get('header_row', 3))
        split_mode = request.form.get('split_mode', 'col_name')
        split_value = request.form.get('split_value', '').strip()
        split_sheets = request.form.getlist('split_sheets')
        use_multi = request.form.get('use_multi') == 'true'
        extra_split_columns = request.form.getlist('extra_split_columns')
        extra_split_indexes = request.form.get('extra_split_indexes', '').strip()

        if not split_sheets:
            return jsonify({'error': '请至少选择一个拆分工作表'}), 400

        wb = load_workbook(temp_path, data_only=True)
        all_values = set()
        first_sheet = split_sheets[0]
        ws = wb[first_sheet]

        col_numbers = []
        if split_mode == 'col_index':
            col_numbers.append(int(split_value))
            if use_multi and extra_split_indexes:
                for idx in extra_split_indexes.split(','):
                    col_numbers.append(int(idx.strip()))
        else:
            # 按列名：找到主拆分列号
            found_main = False
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(header_row, col)
                cell_value = cell.value
                if cell.coordinate in ws.merged_cells:
                    for merged_range in ws.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            top_left = ws.cell(merged_range.min_row, merged_range.min_col)
                            cell_value = top_left.value
                            break
                if cell_value == split_value:
                    col_numbers.append(col)
                    found_main = True
                    break
            if not found_main:
                return jsonify({'error': f'未找到主拆分列 "{split_value}"'}), 400
            # 额外列名
            if use_multi and extra_split_columns:
                for extra_col_name in extra_split_columns:
                    found = False
                    for col in range(1, ws.max_column + 1):
                        cell = ws.cell(header_row, col)
                        cell_value = cell.value
                        if cell.coordinate in ws.merged_cells:
                            for merged_range in ws.merged_cells.ranges:
                                if cell.coordinate in merged_range:
                                    top_left = ws.cell(merged_range.min_row, merged_range.min_col)
                                    cell_value = top_left.value
                                    break
                        if cell_value == extra_col_name:
                            col_numbers.append(col)
                            found = True
                            break
                    if not found:
                        return jsonify({'error': f'未找到额外列名 "{extra_col_name}"'}), 400

        col_numbers = list(set(col_numbers))

        for sheet_name in split_sheets:
            ws = wb[sheet_name]
            valid_cols = all(col_num <= ws.max_column for col_num in col_numbers)
            if not valid_cols:
                continue
            for row in range(header_row + 1, ws.max_row + 1):
                combo_parts = []
                valid = True
                for col_num in col_numbers:
                    val = ws.cell(row, col_num).value
                    if val is None or str(val).strip() == '':
                        valid = False
                        break
                    combo_parts.append(str(val))
                if valid:
                    combo = '_'.join(combo_parts)
                    all_values.add(combo)

        wb.close()
        os.remove(temp_path)
        return jsonify({'values': sorted(list(all_values))})
    except Exception as e:
        os.remove(temp_path)
        return jsonify({'error': str(e)}), 500

# ------------------------------------------------------------
# 核心样式复制函数（修复背景色丢失 + 样式不可变错误）
# ------------------------------------------------------------
from openpyxl.styles import PatternFill, GradientFill

def copy_cell_with_style(src_cell, dst_cell):
    """复制单元格值和所有样式（背景色、字体、边框等）- 增强版"""
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        # 复制字体、边框、对齐、数字格式
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format

        # 处理背景色填充 - 针对 PatternFill 进行显式重建
        src_fill = src_cell.fill
        if src_fill:
            try:
                if isinstance(src_fill, PatternFill):
                    # 新建 PatternFill 对象，复制所有属性
                    new_fill = PatternFill(
                        fill_type=src_fill.fill_type,
                        fgColor=copy(src_fill.fgColor) if src_fill.fgColor else None,
                        bgColor=copy(src_fill.bgColor) if src_fill.bgColor else None
                    )
                    dst_cell.fill = new_fill
                elif isinstance(src_fill, GradientFill):
                    # 渐变填充直接复制（通常很少用）
                    dst_cell.fill = copy(src_fill)
                else:
                    dst_cell.fill = copy(src_fill)
            except Exception:
                # 如果出错，退化为直接复制（但通常不会）
                dst_cell.fill = copy(src_fill)

def copy_sheet_dimensions(src_ws, dst_ws):
    for col in src_ws.column_dimensions:
        dst_ws.column_dimensions[col].width = src_ws.column_dimensions[col].width
    for row in src_ws.row_dimensions:
        dst_ws.row_dimensions[row].height = src_ws.row_dimensions[row].height

def copy_merged_cells(src_ws, dst_ws):
    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

def split_excel_advanced(original_file_path, output_dir, header_row, split_mode, split_value,
                         split_sheets, keep_sheets, filename_template, selected_values,
                         use_multi=False, extra_split_columns=None, extra_split_indexes=''):
    """
    增强拆分函数，支持值筛选和多列组合，完整保留样式
    """
    wb = load_workbook(original_file_path, data_only=False)
    base_name = os.path.splitext(os.path.basename(original_file_path))[0]

    # 确定所有拆分列号
    first_sheet = split_sheets[0]
    ws = wb[first_sheet]
    col_numbers = []
    if split_mode == 'col_index':
        col_numbers.append(int(split_value))
        if use_multi and extra_split_indexes:
            for idx in extra_split_indexes.split(','):
                col_numbers.append(int(idx.strip()))
    else:
        # 主拆分列
        found_main = False
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(header_row, col)
            cell_value = cell.value
            if cell.coordinate in ws.merged_cells:
                for merged_range in ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        top_left = ws.cell(merged_range.min_row, merged_range.min_col)
                        cell_value = top_left.value
                        break
            if cell_value == split_value:
                col_numbers.append(col)
                found_main = True
                break
        if not found_main:
            return False, f"未找到主拆分列 '{split_value}'"
        # 额外列名
        if use_multi and extra_split_columns:
            for extra_col_name in extra_split_columns:
                found = False
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(header_row, col)
                    cell_value = cell.value
                    if cell.coordinate in ws.merged_cells:
                        for merged_range in ws.merged_cells.ranges:
                            if cell.coordinate in merged_range:
                                top_left = ws.cell(merged_range.min_row, merged_range.min_col)
                                cell_value = top_left.value
                                break
                    if cell_value == extra_col_name:
                        col_numbers.append(col)
                        found = True
                        break
                if not found:
                    return False, f"未找到额外列名 '{extra_col_name}'"
    col_numbers = list(set(col_numbers))

    values_to_split = set(selected_values)
    if not values_to_split:
        return False, "没有选择任何拆分值"

    # 为每个拆分值生成新文件
    for value in values_to_split:
        safe_value = re.sub(r'[\\/*?:"<>|]', '_', str(value))
        filename = filename_template.replace('{value}', safe_value).replace('{base_name}', base_name)
        filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        new_filepath = os.path.join(output_dir, filename)

        new_wb = Workbook()
        new_wb.remove(new_wb.active)

        # 处理拆分工作表
        for sheet_name in split_sheets:
            old_ws = wb[sheet_name]
            new_ws = new_wb.create_sheet(title=sheet_name)

            copy_sheet_dimensions(old_ws, new_ws)

            # 复制前 header_row 行（标题及上方）
            for row in range(1, header_row + 1):
                for col in range(1, old_ws.max_column + 1):
                    copy_cell_with_style(old_ws.cell(row, col), new_ws.cell(row, col))
            for merged_range in old_ws.merged_cells.ranges:
                if merged_range.min_row <= header_row:
                    new_ws.merge_cells(str(merged_range))

            # 复制数据行：只保留匹配 value 的行
            target_row = header_row + 1
            for row in range(header_row + 1, old_ws.max_row + 1):
                combo_parts = []
                valid = True
                for col_num in col_numbers:
                    if col_num > old_ws.max_column:
                        valid = False
                        break
                    val = old_ws.cell(row, col_num).value
                    if val is None or str(val).strip() == '':
                        valid = False
                        break
                    combo_parts.append(str(val))
                if valid and '_'.join(combo_parts) == value:
                    for col in range(1, old_ws.max_column + 1):
                        copy_cell_with_style(old_ws.cell(row, col), new_ws.cell(target_row, col))
                    target_row += 1

        # 处理完整保留的工作表
        for sheet_name in keep_sheets:
            if sheet_name not in wb.sheetnames:
                continue
            old_ws = wb[sheet_name]
            new_ws = new_wb.create_sheet(title=sheet_name)
            copy_sheet_dimensions(old_ws, new_ws)
            for row in range(1, old_ws.max_row + 1):
                for col in range(1, old_ws.max_column + 1):
                    copy_cell_with_style(old_ws.cell(row, col), new_ws.cell(row, col))
            copy_merged_cells(old_ws, new_ws)

        new_wb.save(new_filepath)

    wb.close()
    return True, f"成功拆分为 {len(values_to_split)} 个文件"

@app.route('/split', methods=['POST'])
def split_file():
    if 'file' not in request.files:
        return "没有文件", 400
    file = request.files['file']
    if file.filename == '':
        return "文件名为空", 400
    if not allowed_file(file.filename):
        return "不支持的文件类型", 400

    header_row = int(request.form.get('header_row', 3))
    split_mode = request.form.get('split_mode', 'col_name')
    split_value = request.form.get('split_value', '').strip()
    split_sheets = request.form.getlist('split_sheets')
    keep_sheets = request.form.getlist('keep_sheets')
    filename_template = request.form.get('filename_template', '{value}_{base_name}').strip()
    selected_values = request.form.getlist('selected_values')
    use_multi = request.form.get('use_multi') == 'true'
    extra_split_columns = request.form.getlist('extra_split_columns')
    extra_split_indexes = request.form.get('extra_split_indexes', '').strip()

    if not split_value:
        return "请填写拆分列名或列号", 400
    if not split_sheets:
        return "请至少选择一个拆分工作表", 400
    if split_mode == 'col_index':
        try:
            int(split_value)
        except ValueError:
            return "主拆分列号必须为数字", 400
    if not selected_values:
        return "请至少选择一个要拆分的值", 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)

    unique_id = str(uuid.uuid4())
    output_dir = os.path.join(app.config['UPLOAD_FOLDER'], f"split_{unique_id}")
    os.makedirs(output_dir, exist_ok=True)

    success, message = split_excel_advanced(
        input_path, output_dir, header_row, split_mode, split_value,
        split_sheets, keep_sheets, filename_template, selected_values,
        use_multi, extra_split_columns, extra_split_indexes
    )

    os.remove(input_path)

    if not success:
        if os.path.exists(output_dir):
            os.rmdir(output_dir)
        return f"拆分失败：{message}", 400

    zip_name = f"{os.path.splitext(filename)[0]}_拆分结果.zip"
    zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_name)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(output_dir):
            for f in files:
                full_path = os.path.join(root, f)
                zipf.write(full_path, arcname=f)

    # 清理临时输出目录
    for root, _, files in os.walk(output_dir):
        for f in files:
            os.remove(os.path.join(root, f))
    os.rmdir(output_dir)

    # 发送 ZIP 文件后删除它
    @after_this_request
    def remove_zip(response):
        try:
            os.remove(zip_path)
        except Exception:
            pass
        return response

    return send_file(zip_path, as_attachment=True, download_name=zip_name)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)