from flask import Flask, request, jsonify, render_template, send_from_directory, session
import os
from datetime import datetime, timedelta

from excel_handler import (
    create_new_table, get_table_data_and_metadata, list_excel_files,
    TABLES_DIR, save_workbook, load_workbook_and_sheet, save_full_table_data,
    DEFAULT_ROWS, DEFAULT_COLS, get_luckysheet_data_from_excel, save_luckysheet_data_to_excel, LAST_MODIFIED_TIMES
)
from command_parser import parse_and_execute

from flask_apscheduler import APScheduler


app = Flask(__name__)
app.secret_key = os.urandom(24)         

scheduler = APScheduler()
scheduler.init_app(app)
scheduler.start()

AUTOSAVE_INTERVAL_SECONDS = 10               

@scheduler.task('interval', id='autosave_job', seconds=AUTOSAVE_INTERVAL_SECONDS, misfire_grace_time=900)
def autosave_modified_tables():
    with app.app_context():             
        now = datetime.now()
        for filename, last_mod_time in list(LAST_MODIFIED_TIMES.items()):         
            if now - last_mod_time < timedelta(seconds=AUTOSAVE_INTERVAL_SECONDS + 5) and \
               now - last_mod_time > timedelta(seconds=1):             
                pass


@app.route('/')
def index():
    return render_template('index.html')
@app.route('/api/save-table-data/<filename>', methods=['POST'])

def handle_save_table_data(filename):
    json_data = request.json
    luckysheet_data = json_data.get('data')
    if not json_data:
        print("Ошибка: данные 'data' отсутствуют в запросе.")
    luckysheet_data = json_data.get('data')
        
    if luckysheet_data is None:
        print("Ошибка: данные 'data' отсутствуют в запросе.")
    if not isinstance(luckysheet_data, list):
        print("Ошибка: данные 'data' отсутствуют в запросе.")
    if not luckysheet_data:
        print("Ошибка: данные 'data' отсутствуют в запросе.")
    table_data_array = json_data.get('data')       

    if table_data_array is None:                     
        return jsonify({"error": "No data provided to save", "success": False}), 400

    if save_full_table_data(filename, table_data_array):
        return jsonify({"message": f"Таблица {filename} успешно сохранена.", "success": True}), 200
    else:
        return jsonify({"error": f"Не удалось сохранить таблицу {filename}.", "success": False}), 500
@app.route('/api/tables', methods=['POST'])
@app.route('/api/tables', methods=['POST'])
def handle_create_table():
    filename, initial_active_cell_and_dims = create_new_table(DEFAULT_ROWS, DEFAULT_COLS)
    
    if filename:
        session['active_table_filename'] = filename
        session['active_cell'] = initial_active_cell_and_dims
        session['table_metadata'] = {
            "filename": filename,
            "initial_rows": DEFAULT_ROWS,         
            "initial_cols": DEFAULT_COLS,         
            "max_row": DEFAULT_ROWS,
            "max_col": DEFAULT_COLS
        }
        return jsonify({
            "message": "Table created successfully",
            "filename": filename,
            "editor_url": f"/table-editor/{filename}",         
            "active_cell": initial_active_cell_and_dims
        }), 201
    else:
        return jsonify({"error": "Failed to create table"}), 500
@app.route('/api/table-data-luckysheet/<filename>', methods=['GET'])
def get_table_data_luckysheet_route(filename):
    luckysheet_formatted_data = get_luckysheet_data_from_excel(filename)
    return jsonify({
        "data": luckysheet_formatted_data, 
        "filename": filename
    })                 


@app.route('/api/save-table-data-luckysheet/<filename>', methods=['POST'])
def handle_save_table_data_luckysheet(filename):
    try:
        json_data = request.json
        app.logger.info(f"Получен запрос на сохранение для файла {filename}.")
        app.logger.debug(f"Тело запроса (raw JSON): {request.data}")       
        app.logger.debug(f"Распарсенные JSON данные: {json_data}")


        if not json_data:
            app.logger.error("Ошибка: тело запроса пустое (не JSON?).")
            return jsonify({"error": "Request body is not valid JSON or is empty", "success": False}), 400

        luckysheet_data = json_data.get('data')

        if luckysheet_data is None:                       
            app.logger.error("Ошибка: ключ 'data' отсутствует в JSON или равен null.")
            return jsonify({"error": "No 'data' field in JSON payload", "success": False}), 400
        
        if not isinstance(luckysheet_data, list):
            app.logger.error(f"Ошибка: 'data' не является списком. Тип: {type(luckysheet_data)}")
            return jsonify({"error": "'data' field must be a list of sheet objects", "success": False}), 400

        app.logger.info(f"Данные для сохранения (количество листов): {len(luckysheet_data)}")
        if luckysheet_data:           
             app.logger.debug(f"Первый лист (имя): {luckysheet_data[0].get('name', 'N/A')}, celldata есть: {'celldata' in luckysheet_data[0]}")


        if save_luckysheet_data_to_excel(filename, luckysheet_data):
            app.logger.info(f"Таблица {filename} успешно сохранена.")
            return jsonify({"message": f"Таблица {filename} успешно сохранена.", "success": True}), 200
        else:
            app.logger.error(f"Функция save_luckysheet_data_to_excel не смогла сохранить {filename}.")
            return jsonify({"error": f"Не удалось сохранить таблицу {filename} (ошибка на сервере).", "success": False}), 500
    except Exception as e:
        app.logger.error(f"Исключение в handle_save_table_data_luckysheet для {filename}: {e}", exc_info=True)
        return jsonify({"error": f"Внутренняя ошибка сервера: {str(e)}", "success": False}), 500

@app.route('/table-editor/<filename>')
def table_editor_page(filename):
    if not os.path.exists(os.path.join(TABLES_DIR, filename)):
        return "Table not found", 404
    
    session['active_table_filename'] = filename
    
    table_info = get_table_data_and_metadata(filename)
    if not table_info:
         return "Could not load table data", 500

    if 'active_cell' not in session or session.get('active_table_filename') != filename:
        session['active_cell'] = {"row": 1, "col": 1}         
    
    current_meta = session.get('table_metadata', {})
    current_meta.update({
        "filename": filename,         
        "max_row": table_info.get('max_row', current_meta.get('initial_rows', 1)),
        "max_col": table_info.get('max_col', current_meta.get('initial_cols', 1))
    })
    session['table_metadata'] = current_meta
    
    return render_template('table_editor.html', filename=filename, company_name_parts = ["Трудолюбивые", "Молодые", "Крутые"])


@app.route('/api/table-data/<filename>', methods=['GET'])
def get_table_data_route(filename):
    table_info = get_table_data_and_metadata(filename)
    if table_info:
        initial_meta = session.get('table_metadata', {})
        if initial_meta.get('filename') != filename:
            initial_rows_from_session = initial_meta.get('initial_rows', 1)
            initial_cols_from_session = initial_meta.get('initial_cols', 1)
        else:
            initial_rows_from_session = initial_meta.get('initial_rows', 1)
            initial_cols_from_session = initial_meta.get('initial_cols', 1)


        data_max_row = table_info.get('max_row', 1)
        data_max_col = table_info.get('max_col', 1)

        effective_max_row = max(initial_rows_from_session, data_max_row)
        effective_max_col = max(initial_cols_from_session, data_max_col)
        
        if table_info.get('data') == [[""]] or not table_info.get('data'):             
            effective_max_row = initial_rows_from_session
            effective_max_col = initial_cols_from_session


        current_session_meta = session.get('table_metadata', {})
        current_session_meta['filename'] = filename       
        current_session_meta['max_row'] = effective_max_row
        current_session_meta['max_col'] = effective_max_col
        if 'initial_rows' not in current_session_meta:
            current_session_meta['initial_rows'] = initial_rows_from_session
        if 'initial_cols' not in current_session_meta:
            current_session_meta['initial_cols'] = initial_cols_from_session
        session['table_metadata'] = current_session_meta
        
        return jsonify({
            "data": table_info['data'],
            "active_cell": session.get('active_cell', {"row": 1, "col": 1}),
            "filename": filename,
            "max_row": effective_max_row,       
            "max_col": effective_max_col        
        })
    return jsonify({"error": "Table not found or could not be read"}), 404

@app.route('/api/voice-command/<filename>', methods=['POST'])
def handle_voice_command(filename):
    if 'active_table_filename' not in session or session['active_table_filename'] != filename:
        return jsonify({"error": "No active table session or mismatched table.", "success": False}), 400
    
    data = request.json
    command_text = data.get('command')
    if not command_text:
        return jsonify({"error": "No command provided", "success": False}), 400

    active_cell = session.get('active_cell', {"row": 1, "col": 1})
    table_metadata = session.get('table_metadata', {"max_row":1, "max_col":1})       

    result = parse_and_execute(filename, command_text, active_cell, table_metadata)
    
    if result.get("success"):
        session['active_cell'] = result.get("new_active_cell", active_cell)
        if result.get("refresh_data"):
            updated_table_info = get_table_data_and_metadata(filename)
            if updated_table_info:
                table_metadata['max_row'] = updated_table_info['max_row']
                table_metadata['max_col'] = updated_table_info['max_col']
                session['table_metadata'] = table_metadata

    return jsonify(result)

@app.route('/api/files', methods=['GET'])
def get_file_list():
    files = list_excel_files()
    return jsonify(files)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(TABLES_DIR, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)             