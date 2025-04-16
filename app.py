# Vlookup Web App (Render 部署版)
from pywebio.input import file_upload, select
from pywebio.output import put_text, put_table, put_success, popup, put_buttons, put_link
from pywebio.platform.flask import webio_view
from flask import Flask, send_file, request
import pandas as pd
import io
import tempfile

app = Flask(__name__)

def export_result(df):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False)
        tmp.flush()
        put_success("导出成功")
        put_link("点击此处下载 Excel 文件", url=f"/download?path={tmp.name}", new_window=True)

def vlookup_web():
    put_text("寨富村办公可视化工具")

    f1 = file_upload("上传主表", accept=".csv,.xlsx")
    f2 = file_upload("上传查找表", accept=".csv,.xlsx")

    df1 = pd.read_excel(io.BytesIO(f1['content'])) if f1['filename'].endswith('xlsx') else pd.read_csv(io.BytesIO(f1['content']))
    df2 = pd.read_excel(io.BytesIO(f2['content'])) if f2['filename'].endswith('xlsx') else pd.read_csv(io.BytesIO(f2['content']))

    common_keys = list(set(df1.columns) & set(df2.columns))
    key = select("选择用于匹配的字段", options=common_keys)

    method = select("选择匹配方式", options=["交集", "左连接", "并集"])
    method_map = {"交集": "inner", "左连接": "left", "并集": "outer"}
    how = method_map[method]

    df_result = pd.merge(df1, df2, how=how, on=key, suffixes=("_主表", "_查找表"))

    put_success(f"匹配成功，共 {len(df_result)} 行")
    put_table([list(df_result.columns)] + df_result.head(10).astype(str).values.tolist())

    put_buttons(["联系方式", "导出结果"], onclick=[
        lambda: popup("帮助", [put_text("有疑问请联系：\n18291971545\n樊思恺")]),
        lambda: export_result(df_result)
    ])

@app.route('/')
def index():
    return webio_view(vlookup_web)()

@app.route('/download')
def download():
    file_path = request.args.get('path')
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run()
