import streamlit as st
import win32com.client as win32
import os
import pythoncom
import time
import pandas as pd
import subprocess
import re
import io
from openpyxl import load_workbook
from zhipuai import ZhipuAI

# --- 1. 基础配置 ---
API_KEY = st.secrets.get("ZHIPUAI_API_KEY") or os.getenv("ZHIPUAI_API_KEY")
client = ZhipuAI(api_key=API_KEY)

def force_cleanup():
    """清理后台进程，防止文件锁定"""
    subprocess.run("taskkill /F /IM excel.exe /T", shell=True, capture_output=True)
    subprocess.run("taskkill /F /IM winword.exe /T", shell=True, capture_output=True)

# --- 2. 深度数据提取函数 (完全保留您脚本#1的核心逻辑) ---
def get_column_letter_from_ref(ref_str):
    if not ref_str: return None
    clean_ref = ref_str.replace('$', '')
    match = re.search(r'([A-Z]+)\d+', clean_ref.split('!')[-1])
    return match.group(1) if match else None

def deep_extract_excel_data(file_path):
    """
    完全保留脚本#1的算法：利用迭代队列扫描提取所有图表数据
    返回格式：{(sheet_index, chart_index): DataFrame}
    """
    wb_data = load_workbook(file_path, data_only=True)
    results_map = {}

    for sheet_idx, ws in enumerate(wb_data.worksheets):
        if not hasattr(ws, '_charts') or not ws._charts:
            continue
        
        for chart_idx, chart_obj in enumerate(ws._charts):
            col_data_map = {}
            queue = [chart_obj]
            visited = set()
            found_refs = set()
            
            # --- 脚本#1 的核心扫描算法 ---
            while queue:
                curr = queue.pop(0)
                if id(curr) in visited: continue
                visited.add(id(curr))
                if hasattr(curr, 'f') and isinstance(curr.f, str) and '!' in curr.f:
                    found_refs.add(curr.f)
                if hasattr(curr, '__dict__'):
                    for attr in curr.__dict__:
                        val = getattr(curr, attr)
                        if val is None or isinstance(val, (int, str, float, bool, bytes)): continue
                        if isinstance(val, list): queue.extend(val)
                        else: queue.append(val)

            # 解析引用的单元格
            for ref in found_refs:
                try:
                    ref_ws_name, cells = ref.replace("'", "").split("!") if "!" in ref else (ws.title, ref)
                    if ref_ws_name in wb_data.sheetnames:
                        target_ws = wb_data[ref_ws_name]
                        # 脚本#1 的数据抓取方式
                        data_vals = [cell[0].value for cell in target_ws[cells]]
                        col_letter = get_column_letter_from_ref(cells)
                        if col_letter:
                            header = target_ws[f"{col_letter}1"].value or f"指标_{col_letter}"
                            col_data_map[col_letter] = {
                                "header": str(header).strip().replace('\n', ' '), 
                                "data": data_vals
                            }
                except: continue

            # 组装 DataFrame
            if col_data_map:
                sorted_keys = sorted(col_data_map.keys())
                max_l = max(len(v["data"]) for v in col_data_map.values())
                df_dict = {}
                for k in sorted_keys:
                    info = col_data_map[k]
                    df_dict[info["header"]] = info["data"] + [None] * (max_l - len(info["data"]))
                
                # 使用 (sheet索引, chart索引) 作为 key，确保与 win32 顺序严格一致
                results_map[(sheet_idx, chart_idx)] = pd.DataFrame(df_dict)
                
    return results_map

# --- 3. 自动化填充函数 ---
def run_full_process(excel_path, word_path, output_path, user_prompt):
    # A. 首先调用深度提取逻辑
    st.info("正在深度扫描 Excel 图表数据源...")
    data_inventory = deep_extract_excel_data(excel_path)
    
    # B. 启动 Office 自动化
    force_cleanup()
    pythoncom.CoInitialize()
    
    excel = None
    word = None
    try:
        excel = win32.DispatchEx('Excel.Application')
        word = win32.DispatchEx('Word.Application')
        excel.Visible = False
        word.Visible = False
        
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        doc = word.Documents.Open(os.path.abspath(word_path))

        global_counter = 0
        
        # 遍历所有工作表 (与 openpyxl 顺序一致)
        for s_idx in range(1, wb.Sheets.Count + 1):
            sheet = wb.Sheets(s_idx)
            charts = sheet.ChartObjects()
            
            for c_idx in range(1, charts.Count + 1):
                global_counter += 1
                chart_obj = charts.Item(c_idx)
                
                # 从我们提前准备好的 data_inventory 中取数据 (下标从0开始)
                # openpyxl 的索引和 win32 顺序是物理匹配的
                df = data_inventory.get((s_idx-1, c_idx-1))
                
                # AI 分析
                if df is not None and not df.empty:
                    with st.spinner(f"正在分析第 {global_counter} 个图表的数据..."):
                        context_data = df.to_string(index=False)
                        response = client.chat.completions.create(
                            model="glm-4-flash",
                            messages=[
                                {"role": "system", "content": "你是一个资深数据分析师。请分析提取出的数据趋势，对比指标，并给出专业结论。请直接回复报告内容。"},
                                {"role": "user", "content": f"数据如下：\n{context_data}\n指令：{user_prompt}"}
                            ]
                        )
                        analysis_text = response.choices[0].message.content
                else:
                    analysis_text = "未能在 Excel 中定位到该图表对应的底层 X/Y 轴数据区域。"

                # 填充 Word
                # 1. 复制图表
                chart_obj.Chart.ChartArea.Copy()
                time.sleep(1) # 给剪贴板一点时间
                
                # 2. 替换 [图表n]
                word.Selection.HomeKey(Unit=6)
                if word.Selection.Find.Execute(f"[图表{global_counter}]"):
                    word.Selection.Paste()
                
                # 3. 替换 [分析n]
                word.Selection.HomeKey(Unit=6)
                if word.Selection.Find.Execute(f"[分析{global_counter}]"):
                    word.Selection.Text = analysis_text

        doc.SaveAs(os.path.abspath(output_path))
        doc.Close()
        wb.Close(False)
        return global_counter

    finally:
        if excel: excel.Quit()
        if word: word.Quit()
        pythoncom.CoUninitialize()

# --- 4. Streamlit 界面 ---
st.set_page_config(page_title="AI 报表生成器", layout="wide")
st.title("📊 能源项目 AI 自动化分析报告")

with st.sidebar:
    st.header("文件上传")
    ex_file = st.file_uploader("1. 上传 Excel 数据", type=["xlsx"])
    wd_file = st.file_uploader("2. 上传 Word 模板", type=["docx"])
    st.divider()
    prompt = st.text_area("分析需求", value="请结合 X 轴数据分析随时间的变化趋势，并指出异常点。")
    out_name = st.text_input("输出文件名", value="分析报告_生成版.docx")

if st.button("🚀 开始一键生成报告", use_container_width=True):
    if ex_file and wd_file:
        t_ex, t_wd = "input_data.xlsx", "input_tpl.docx"
        with open(t_ex, "wb") as f: f.write(ex_file.getbuffer())
        with open(t_wd, "wb") as f: f.write(wd_file.getbuffer())
        
        target_path = os.path.join(os.path.expanduser("~"), "Desktop", out_name)
        
        try:
            total = run_full_process(t_ex, t_wd, target_path, prompt)
            st.success(f"处理完成！成功解析并填充了 {total} 组数据。报告已存至桌面。")
        except Exception as e:
            st.error(f"运行失败: {e}")
    else:
        st.warning("请上传完整的文件。")
