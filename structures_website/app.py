#!/usr/bin/env python
import os
import logging
from flask import Flask, request, render_template, redirect, url_for, flash, jsonify
from sqlalchemy import create_engine, text
import pandas as pd
from datetime import datetime
from sqlalchemy import text

app = Flask(__name__)
app.secret_key = 'cyhung^209_123' 

DATABASE_URL = "sqlite:///structures.db"
engine = create_engine(DATABASE_URL)

ALLOWED_SEARCH_FIELDS = ["FID", "構造物編碼", "權屬分署", "調查年分", "所屬工程", "所屬工程編號", "X_TWD97",
                         "Y_TWD97", "縣市", "功能評估", "子集水區"]


with engine.begin() as conn:
    conn.execute(text("""
        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """))

@app.route('/', methods=['GET', 'POST'])
def index():
    field = None
    keyword = None
    structure_data = []  # 體檢表資料
    inspection_data = []  # 巡查表資料
    inspection_years = []  # 有巡查的年份
    structure_fields = []  # 體檢表的欄位名稱
    last_update_time = "尚未更新"
    with engine.connect() as conn:
        result = conn.execute(text("SELECT value FROM meta WHERE key='last_update'")).fetchone()
        if result:
            last_update_time = result[0]

    # 獲取體檢表的欄位名稱
    with engine.connect() as conn:
        field_query = "PRAGMA table_info(structures)"
        fields_result = conn.execute(text(field_query))
        structure_fields = [row._mapping['name'] for row in fields_result]

    if request.method == 'POST':
        field = request.form.get('field')  # 使用者選擇的欄位
        keyword = request.form.get('keyword')  # 使用者輸入的關鍵字

        if field and keyword:
            # 檢查選擇的欄位是否有效
            if field not in ALLOWED_SEARCH_FIELDS:
                return "The selected field is not searchable", 400

            # 決定搜尋方式
            if field == "FID":  # 精確搜尋
                operator = "="
            else:  # 模糊搜尋
                operator = "LIKE"

            # 查詢體檢表 (structures)
            structure_query = f"SELECT * FROM structures WHERE {field} {operator} :keyword"
            # 查詢巡查表 (inspection_table)
            inspection_query = f"""
                SELECT * FROM inspection_table 
                WHERE FID IN (SELECT FID FROM structures WHERE {field} {operator} :keyword)
            """
            # 查詢有巡查的年份
            year_query = f"""
                SELECT DISTINCT substr("巡查年分", 1, 4) AS Year 
                FROM inspection_table 
                WHERE FID IN (SELECT FID FROM structures WHERE {field} {operator} :keyword)
            """

            # 為精確或模糊搜尋準備參數
            query_param = keyword if operator == "=" else f"%{keyword}%"

            # 執行查詢
            with engine.connect() as conn:
                structure_result = conn.execute(text(structure_query), {'keyword': query_param})
                inspection_result = conn.execute(text(inspection_query), {'keyword': query_param})
                year_result = conn.execute(text(year_query), {'keyword': query_param})

                # 使用 row._mapping 來安全地處理查詢結果
                structure_data = [row._mapping for row in structure_result]
                inspection_data = [row._mapping for row in inspection_result]
                inspection_data = sorted(inspection_data, key=lambda x: x['巡查年分'])

                
                # 如果有多條體檢表資料，清空巡查年份，否則正常處理
                if len(structure_data) > 1:
                    inspection_years = []
                else:
                    inspection_years = [row._mapping['Year'] for row in year_result]

    return render_template('index.html', 
                       structure_data=structure_data, 
                       inspection_data=inspection_data, 
                       inspection_years=inspection_years, 
                       structure_fields=ALLOWED_SEARCH_FIELDS, 
                       selected_field=field, 
                       search_value=keyword,
                       last_update_time=last_update_time)

@app.route('/fetch_inspection_data', methods=['GET'])
def fetch_inspection_data():
    fid = request.args.get('fid')  # Get the FID from request
    if not fid or not fid.isdigit():
        return {"error": "Invalid FID provided"}, 400

    try:
        inspection_query = "SELECT * FROM inspection_table WHERE FID = :fid"
        year_query = "SELECT DISTINCT substr(\"巡查年分\", 1, 4) AS Year FROM inspection_table WHERE FID = :fid"

        with engine.connect() as conn:

            column_query = "PRAGMA table_info(inspection_table)"
            columns_result = conn.execute(text(column_query)).mappings().all()
            columns = [row['name'] for row in columns_result]  # Get the column


            inspection_result = conn.execute(text(inspection_query), {'fid': fid}).mappings().all()
            year_result = conn.execute(text(year_query), {'fid': fid}).mappings().all()

            # Extract data and transform results
            inspection_data = [dict(row) for row in inspection_result]
            inspection_data = sorted(inspection_data, key=lambda x: x.get('巡查年分', ''))
            years = [row['Year'] for row in year_result]

        if not inspection_data:
            return jsonify({"inspectionData": [], "years": [], "fields": columns})

        return jsonify({"inspectionData": inspection_data, "years": years, "fields": columns})

    except Exception as e:
        logging.error(f"Error fetching data for FID {fid}: {e}")
        return {"error": "An unexpected error occurred", "details": str(e)}, 500


@app.route('/upload', methods=['GET', 'POST'])
def upload_excel():
    if request.method == 'POST':
        file1 = request.files.get('體檢表')
        file2 = request.files.get('巡查表')

        try:
            msg_list = []
            if file1:
                df1 = pd.read_excel(file1)
                df1.to_sql('structures', engine, if_exists='replace', index=False)
                msg_list.append("✅ 體檢表更新成功")

            if file2:
                df2 = pd.read_excel(file2)
                df2.to_sql('inspection_table', engine, if_exists='replace', index=False)
                msg_list.append("✅ 巡查表更新成功")

            if not msg_list:
                flash("⚠️ 請至少上傳一份 Excel 檔", "warning")
                return redirect(url_for('upload_excel'))

            flash("、".join(msg_list), "success")

            with engine.begin() as conn:
                conn.execute(text("REPLACE INTO meta (key, value) VALUES ('last_update', :val)"),
                            {"val": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            return redirect(url_for('index'))  # 跳回主頁

        except Exception as e:
            flash(f"❌ 上傳失敗：{str(e)}", "danger")
            return redirect(url_for('upload_excel'))

    return render_template('upload.html')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)









