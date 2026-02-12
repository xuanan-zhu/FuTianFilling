import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
import re
from datetime import datetime
import io
import pandas as pd

# ================= 1. 配置区 =================

DEFAULT_HEADERS = [
    "团队", 
    "福田数量", 
    "序号", 
    "真实姓名", 
    "推荐人", 
    "居住地", 
    "职业", 
    "出身年月日", 
    "电话号码", 
    "现在生活事业家庭情况", 
    "想收获什么梦想", 
    "有无宗教信仰"
]

# ================= 2. 核心逻辑区 (保持原逻辑) =================

def normalize_birth_date(value):
    """将各种格式的出生日期统一为：YYYY-MM-DD"""
    if not value:
        return ""
    nums = re.findall(r"\d+", value)
    if len(nums) >= 3:
        year, month, day = nums[:3]
        return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
    return value

def extract_person_info(text):
    """解析文本提取信息"""
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
        for name in v:
            reverse_map[name] = k

    result = {k: "" for k in field_alias}
    current_field = None

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        line = line.lstrip("0123456789. ")

        if "：" in line or ":" in line:
            parts = line.replace(":", "：").split("：", 1)
            if len(parts) > 1:
                key, val = parts
                key = key.split("（")[0].split("(")[0].strip()
                key = key.replace(" ", "").replace("　", "")
                
                if key in reverse_map:
                    current_field = reverse_map[key]
                    if val.strip():
                        result[current_field] = val.strip()
                    continue

        if current_field:
            if result[current_field]:
                result[current_field] += "\n" + line
            else:
                result[current_field] = line

    result["出身年月日"] = normalize_birth_date(result["出身年月日"])
    return result

def create_blank_workbook():
    """在内存中创建一个新的 Workbook"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(DEFAULT_HEADERS)
    
    # 设置列宽
    widths = {"A": 10, "B": 10, "C": 6, "D": 12, "E": 12, "H": 15, "I": 15}
    for col_letter, width in widths.items():
         ws.column_dimensions[col_letter].width = width
    for col in range(1, len(DEFAULT_HEADERS) + 1):
        letter = openpyxl.utils.get_column_letter(col)
        if letter not in widths:
            ws.column_dimensions[letter].width = 20
    return wb

def append_data_to_workbook(wb, info_dict):
    """将提取的数据追加到 workbook 对象中"""
    sheet = wb.active
    
    # 获取表头映射
    header_map = {}
    for col_idx, cell in enumerate(sheet[1], 1):
        if cell.value:
            header_map[str(cell.value).strip()] = col_idx
            
    if not header_map:
        return False, "表格没有表头，无法识别列名"

    next_row = sheet.max_row + 1
    
    # 1. 填入数据
    for field, value in info_dict.items():
        if field in header_map:
            col_index = header_map[field]
            cell = sheet.cell(row=next_row, column=col_index)
            cell.value = value
            # 自动换行
            cell.alignment = Alignment(wrap_text=True)

    # 2. 自动序号
    if "序号" in header_map:
        seq_col = header_map["序号"]
        # 假设第一行是表头，序号从1开始
        seq_num = next_row - 1 
        sheet.cell(row=next_row, column=seq_col).value = seq_num

    return True, f"成功添加：{info_dict.get('真实姓名', '未知')}"

def to_excel_bytes(wb):
    """将 workbook 转为二进制流以便下载"""
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ================= 3. Streamlit 界面区 =================

st.set_page_config(page_title="智能填表助手", page_icon="📝")

st.title("📝 Excel 智能填表助手 (Web版)")
st.markdown("---")

# --- Sidebar: 文件管理 ---
st.sidebar.header("1. 文件设置")
upload_option = st.sidebar.radio("选择模式:", ["📂 上传现有 Excel", "✨ 新建空白 Excel"])

# Session State 用于存储当前的 Workbook
if 'workbook' not in st.session_state:
    st.session_state.workbook = None
if 'file_name' not in st.session_state:
    st.session_state.file_name = "团队统计表.xlsx"

# 处理文件加载逻辑
if upload_option == "📂 上传现有 Excel":
    uploaded_file = st.sidebar.file_uploader("上传 .xlsx 文件", type=["xlsx"])
    if uploaded_file:
        try:
            # 只有当上传的文件改变时才重新加载
            if st.session_state.get('last_uploaded_id') != uploaded_file.id:
                st.session_state.workbook = openpyxl.load_workbook(uploaded_file)
                st.session_state.file_name = uploaded_file.name
                st.session_state.last_uploaded_id = uploaded_file.id
                st.sidebar.success("文件已加载！")
        except Exception as e:
            st.sidebar.error(f"文件读取失败: {e}")
else:
    if st.sidebar.button("初始化新表格"):
        st.session_state.workbook = create_blank_workbook()
        st.session_state.file_name = "新团队统计表.xlsx"
        st.sidebar.success("已创建新表格模板！")

# --- Main: 数据录入 ---
st.header("2. 数据录入")

if st.session_state.workbook is None:
    st.info("👈 请先在左侧侧边栏上传或新建 Excel 文件。")
else:
    # 文本输入
    input_text = st.text_area("在此粘贴个人信息文本:", height=200, placeholder="粘贴格式如：\n姓名：张三\n电话：13800000000...")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        process_btn = st.button("⚡ 提取并追加数据", type="primary")
    
    # 处理逻辑
    if process_btn and input_text:
        info = extract_person_info(input_text)
        success, msg = append_data_to_workbook(st.session_state.workbook, info)
        
        if success:
            st.success(msg)
            # 显示刚刚解析的数据预览
            st.markdown("**本次解析结果预览:**")
            st.json(info)
        else:
            st.error(msg)

    # --- Result: 下载区域 ---
    st.markdown("---")
    st.header("3. 下载结果")
    
    # 预览当前 Excel 的最后几行（可选功能，方便用户确认）
    try:
        # 将 openpyxl worksheet 转为 pandas dataframe 用于预览
        ws = st.session_state.workbook.active
        data = ws.values
        columns = next(data)
        df = pd.DataFrame(data, columns=columns)
        
        st.caption(f"当前表格共有 {len(df)} 条数据，预览最后 3 条：")
        st.dataframe(df.tail(3))
        
        # 下载按钮
        excel_data = to_excel_bytes(st.session_state.workbook)
        st.download_button(
            label="📥 下载更新后的 Excel 文件",
            data=excel_data,
            file_name=st.session_state.file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.warning("暂无数据或预览失败，但你可以继续添加。")