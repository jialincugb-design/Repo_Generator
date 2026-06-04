import streamlit as st
import os
import time
import pandas as pd
import subprocess
import re
import io
from openpyxl import load_workbook
from zhipuai import ZhipuAI
import sys

# 🔐 1. Windows 专属高级组件动态导入（防止 Linux 云端部署崩溃）
if sys.platform == "win32":
    import pythoncom
    import win32com.client as win32
    IS_WINDOWS = True
else:
    IS_WINDOWS = False

# 🔐 2. 本地环境变量安全加载
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# --- 3. 基础安全配置 ---
API_KEY = st.secrets.get("ZHIPUAI_API_KEY") or os.getenv("ZHIPUAI_API_KEY")

if not API_KEY:
    st.error("⚠️ 当前大模型引擎未激活：请在 Streamlit Secrets 或本地 .env 文件中配置 ZHIPUAI_API_KEY")
    st.stop()

client = ZhipuAI(api_key=API_KEY)

def force_cleanup():
    """清理后台进程，防止文件锁定"""
    if IS_WINDOWS:
        subprocess.run("taskkill /F /IM excel.exe /T", shell=True, capture_output=True)
        subprocess.run("taskkill /F /IM winword.exe /T", shell=True, capture_output=True)

# --- 4. 深度数据提取函数 (openpyxl 跨平台，保留其优良的核心逻辑) ---
def get_column_letter_from_ref(ref_str):
    if not ref_str: return None
    clean_ref = ref_str.replace('$', '')
    match = re.search(r'([A-Z]+)\d+', clean_ref.split('!')[-1])
    return match.group(1) if match else None

def deep_extract_excel_data(file_path):
    """
    利用迭代队列扫描提取所有图表数据 (纯 Python 实现，全平台通用)
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

            for ref in found_refs:
                try:
                    ref_ws_name, cells = ref.replace("'", "").split("!") if "!" in ref else (ws.title, ref)
                    if ref_ws_name in wb_data.sheetnames:
                        target_ws = wb_data[ref_ws_name]
                        data_vals = [cell[0].value for cell in target_ws[cells]]
                        col_letter = get_column_letter_from_ref(cells)
                        if col_letter:
                            header = target_ws[f"{col_letter}1"].value or f"指标_{col_letter}"
                            col_data_map[col_letter] = {
                                "header": str(header).strip().replace('\n', ' '), 
                                "data": data_vals
                            }
                except: continue

            if col_data_map:
                sorted_keys = sorted(col_data_map.keys())
                max_l = max(len(v["data"]) for v in col_data_map.values())
                df_dict = {}
                for k in sorted_keys:
                    info = col_data_map[k]
                    df_dict[info["header"]] = info["data"] + [None] * (max_l - len(info["data"]))
                
                results_map[(sheet_idx, chart_idx)] = pd.DataFrame(df_dict)
                
    return results_map

# --- 5. 自动化填充函数 ---
def run_full_process(excel_path, word_path, output_path, user_prompt):
    # A. 提取数据
    st.info("🔄 正在深度扫描 Excel 图表底层数据源...")
    data_inventory = deep_extract_excel_data(excel_path)
    
    # B. 安全校验：检查当前系统环境是否支持 win32 自动化
    if not IS_WINDOWS:
        st.warning("⚠️ 检测到当前非 Windows 环境。已成功提取并分析数据，但无法调用本地 Office 软件进行图表复制与粘贴。")
        # 云端只做 AI 分析演练
        for key, df in data_inventory.items():
            st.write(f"📊 图表 (Sheet:{key[0]+1}, Index:{key[1]+1}) 的结构化数据:")
            st.dataframe(df)
        return len(data_inventory)

    # C. 启动 Windows Office 自动化
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
        
        for s_idx in range(1, wb.Sheets.Count + 1):
            sheet = wb.Sheets(s_idx)
            charts = sheet.ChartObjects()
            
            for c_idx in range(1, charts.Count + 1):
                global_counter += 1
                chart_obj = charts.Item(c_idx)
                
                df = data_inventory.get((s_idx-1, c_idx-1))
                
                # AI 分析：升级使用 markdown 排版喂给大模型
                if df is not None and not df.empty:
                    with st.spinner(f"🤖 AI 正在深度分析第 {global_counter} 个图表的数据趋势..."):
                        try:
                            # 转换为大模型更好识别的 Markdown 表格
                            context_data = df.to_markdown(index=False)
                        except:
                            context_data = df.to_string(index=False)

                        response = client.chat.completions.create(
                            model="glm-4-flash",
                            messages=[
                                {"role": "system", "content": "你是一个严谨的油气能源行业资深数据分析师。请分析提取出的数据趋势，对比核心指标，并给出专业简明的结论。请直接回复报告文本，不需要说客套话。"},
                                {"role": "user", "content": f"提取出的图表数据如下：\n{context_data}\n分析指令：{user_prompt}"}
                            ]
                        )
                        analysis_text = response.choices[0].message.content
                else:
                    analysis_text = "未能在 Excel 中定位到该图表对应的底层 X/Y 轴数据源数据。"

                # 填充 Word 文档
                # 1. 复制 Excel 图表
                chart_obj.Chart.ChartArea.Copy()
                time.sleep(0.8)  # 给剪贴板系统反应时间
                
                # 2. 替换 [图表n] 占位符
                word.Selection.HomeKey(Unit=6)
                if word.Selection.Find.Execute(f"[图表{global_counter}]"):
                    word.Selection.Paste()
                    time.sleep(0.5)
                
                # 3. 替换 [分析n] 占位符
                word.Selection.HomeKey(Unit=6)
                if word.Selection.Find.Execute(f"[分析{global_counter}]"):
                    word.Selection.Text = analysis_text
                    time.sleep(0.2)

        doc.SaveAs(os.path.abspath(output_path))
        doc.Close()
        wb.Close(False)
        return global_counter

    finally:
        if excel: excel.Quit()
        if word: word.Quit()
        pythoncom.CoUninitialize()

# --- 6. Streamlit 现代感交互界面 ---
st.set_page_config(page_title="AI 智能报表自动生成系统", layout="wide")
st.title("📊 智能报表开发平台 (AI + 报告自动化)")

with st.sidebar:
    st.header("📂 核心数据源配置")
    ex_file = st.file_uploader("1. 上传数据源 Excel", type=["xlsx"])
    wd_file = st.file_uploader("2. 上传报告模板 Word", type=["docx"])
    st.divider()
    prompt = st.text_area("🧠 AI 深度分析控制指令", value="请结合数据分析随时间或深度的变化趋势，指出显著的极值、异常点，并给出生产建议。")
    out_name = st.text_input("💾 生成文件名", value="智能化评估报告_自动生成.docx")

# 运行控制区域
if st.button("🚀 开始一键激活引擎并生成报告", use_container_width=True):
    if ex_file and wd_file:
        t_ex, t_wd = "input_data.xlsx", "input_tpl.docx"
        with open(t_ex, "wb") as f: f.write(ex_file.getbuffer())
        with open(t_wd, "wb") as f: f.write(wd_file.getbuffer())
        
        # 默认保存到桌面
        target_path = os.path.join(os.path.expanduser("~"), "Desktop", out_name)
        
        try:
            total = run_full_process(t_ex, t_wd, target_path, prompt)
            if IS_WINDOWS:
                st.success(f"🎉 处理完成！成功解析并完美填充了 {total} 组图表与 AI 深度分析报告。报告已安全存至您的桌面。")
            else:
                st.info(f"💡 非 Windows 模拟运行完成。共解析出 {total} 组图表数据源。")
        except Exception as e:
            st.error(f"❌ 运行遭遇异常: {e}")
            st.info("提示：请确认当前没有在后台手动打开相同的 Excel 或 Word 模板文件造成冲突。")
    else:
        st.warning("⚠️ 关键配置缺失：请先在左侧侧边栏上传完整的 Excel 数据文件与 Word 模板文件。")
