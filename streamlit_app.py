import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
import re
import io
import pandas as pd

# ================= 1. 核心逻辑区 (保持不变) =================

DEFAULT_HEADERS = [
    "团队", "福田数量", "序号", "真实姓名", "推荐人", "居住地", 
    "职业", "出身年月日", "电话号码", "现在生活事业家庭情况", 
    "想收获什么梦想", "有无宗教信仰"
]

def normalize_birth_date(value):
    if not value: return ""
    nums = re.findall(r"\d+", value)
    if len(nums) >= 3:
        year, month, day = nums[:3]
        return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
    return value

def extract_person_info(text):
    text = text.replace("\r\n", "\n")
    field_alias = {
        "真实姓名": ["真实姓名", "姓名"],
        "推荐人": ["推荐人", "分享人"],
        "居住地": ["居住地", "地址"],
        "职业": ["职业"],
        "出身年月日": ["出身年月日", "出生年月日", "生日"],
        "电话号码": ["电话号码", "手机号码", "电话", "手机"],
        "现在生活事业家庭情况": ["现在生活事业家庭情况"],
        "想收获什么梦想": ["想收获什么梦想"],
        "有无宗教信仰": ["有无宗教信仰"]
    }
    reverse_map = {}
    for k, v in field_alias.items():
        for name in v: reverse_map[name] = k

    result = {k: "" for k in field_alias}
    current_field = None

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line: continue
        line = line.lstrip("0123456789. ")

        if "：" in line or ":" in line:
            parts = line.replace(":", "：").split("：", 1)
            if len(parts) > 1:
                key, val = parts
                key = key.split("（")[0].split("(")[0].strip().replace(" ", "")
                if key in reverse_map:
                    current_field = reverse_map[key]
                    if val.strip(): result[current_field] = val.strip()
                    continue

        if current_field:
            if result[current_field]: result[current_field] += "\n" + line
            else: result[current_field] = line

    result["出身年月日"] = normalize_birth_date(result["出身年月日"])
    return result

def create_blank_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(DEFAULT_HEADERS)
    widths = {"A": 10, "B": 10, "C": 6, "D": 12, "E": 12, "H": 15, "I": 15}
    for col_letter, width in widths.items():
         ws.column_dimensions[col_letter].width = width
    return wb

def append_data_to_workbook(wb, info_dict):
    sheet = wb.active
    header_map = {}
    for col_idx, cell in enumerate(sheet[1], 1):
        if cell.value: header_map[str(cell.value).strip()] = col_idx
            
    if not header_map: return False, "表格没有表头，无法识别列名"

    next_row = sheet.max_row + 1
    for field, value in info_dict.items():
        if field in header_map:
            col_index = header_map[field]
            cell = sheet.cell(row=next_row, column=col_index)
            cell.value = value
            cell.alignment = Alignment(wrap_text=True)

    if "序号" in header_map:
        seq_col = header_map["序号"]
        seq_num = next_row - 1 
        sheet.cell(row=next_row, column=seq_col).value = seq_num

    return True, f"成功添加：{info_dict.get('真实姓名', '未知')}"

def to_excel_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ================= 2. 界面交互逻辑 =================

st.set_page_config(page_title="智能填表助手", page_icon="📝", layout="wide")

# 初始化 Session State 变量
if 'workbook' not in st.session_state: st.session_state.workbook = None
if 'file_name' not in st.session_state: st.session_state.file_name = "团队统计表.xlsx"
if 'last_loaded_key' not in st.session_state: st.session_state.last_loaded_key = None
# 新增：用于存储操作反馈信息
if 'status_msg' not in st.session_state: st.session_state.status_msg = None
if 'last_extracted_info' not in st.session_state: st.session_state.last_extracted_info = None

# --- 回调函数：处理提交并清空 ---
def submit_data():
    """点击按钮时执行的函数"""
    text = st.session_state.user_input # 获取输入框的内容
    
    if not text.strip():
        st.session_state.status_msg = ("warning", "文本框是空的，请输入内容！")
        return

    if st.session_state.workbook is None:
        st.session_state.status_msg = ("error", "请先在左侧上传或新建表格！")
        return

    # 执行提取和追加
    info = extract_person_info(text)
    success, msg = append_data_to_workbook(st.session_state.workbook, info)

    if success:
        st.session_state.status_msg = ("success", msg)
        st.session_state.last_extracted_info = info
        # 关键步骤：清空输入框 (通过设置绑定的 key 为空字符串)
        st.session_state.user_input = "" 
    else:
        st.session_state.status_msg = ("error", msg)

# ================= 3. 页面布局 =================

st.title("📝 Excel 智能填表助手 (Web版)")

# --- Sidebar ---
with st.sidebar:
    st.header("1. 文件设置")
    upload_option = st.radio("模式:", ["📂 上传现有 Excel", "✨ 新建空白 Excel"])

    if upload_option == "📂 上传现有 Excel":
        uploaded_file = st.file_uploader("上传文件", type=["xlsx"])
        if uploaded_file:
            file_key = f"{uploaded_file.name}_{uploaded_file.size}"
            if st.session_state.last_loaded_key != file_key:
                st.session_state.workbook = openpyxl.load_workbook(uploaded_file)
                st.session_state.file_name = uploaded_file.name
                st.session_state.last_loaded_key = file_key
                st.success(f"已加载: {uploaded_file.name}")
    else:
        if st.button("初始化新表格"):
            st.session_state.workbook = create_blank_workbook()
            st.session_state.file_name = "新团队统计表.xlsx"
            st.session_state.last_loaded_key = "NEW_CREATED"
            st.success("已创建新表格模板！")
            
    st.markdown("---")
    st.info("💡 提示：追加数据后，请务必点击主界面下方的下载按钮保存文件。")

# --- Main Area ---

col_input, col_preview = st.columns([1, 1])

# 左侧：输入区
with col_input:
    st.subheader("2. 数据录入")
    # 绑定 key="user_input"，这样我们可以在回调函数里控制它
    st.text_area(
        "在此粘贴个人信息文本:", 
        height=300, 
        key="user_input",
        placeholder="粘贴格式如：\n姓名：张三\n电话：13800000000..."
    )
    
    # 按钮绑定 on_click 回调
    st.button("⚡ 提取并追加数据", type="primary", on_click=submit_data, use_container_width=True)

    # 显示操作反馈消息
    if st.session_state.status_msg:
        msg_type, msg_text = st.session_state.status_msg
        if msg_type == "success":
            st.success(msg_text)
            with st.expander("查看刚才提取的数据详情"):
                st.json(st.session_state.last_extracted_info)
        elif msg_type == "error":
            st.error(msg_text)
        elif msg_type == "warning":
            st.warning(msg_text)

# 右侧：全表预览区
with col_preview:
    st.subheader("3. 表格实时预览")
    
    if st.session_state.workbook:
        try:
            # 获取数据用于预览
            ws = st.session_state.workbook.active
            data = list(ws.values)
            if data:
                columns = data[0]
                rows = data[1:]
                # 转换为 DataFrame
                df = pd.DataFrame(rows, columns=columns)
                
                # 统计信息
                st.caption(f"当前共 **{len(df)}** 条数据")
                
                # 全表预览 (使用 dataframe 组件，支持滚动、排序、搜索)
                st.dataframe(df, use_container_width=True, height=300)
                # 全表预览 (改为静态表格，兼容旧版 iOS)
                # st.table(df)
                
                st.markdown("---")
                # 下载按钮放在这里更显眼
                excel_data = to_excel_bytes(st.session_state.workbook)
                st.download_button(
                    label="📥 下载最新 Excel 文件",
                    data=excel_data,
                    file_name=st.session_state.file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"预览生成失败: {e}")
    else:
        st.info("👈 请先在左侧加载 Excel 文件")


