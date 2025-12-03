import streamlit as st
import pandas as pd
import io
import json
import re
from typing import Dict, List, Any
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

# 页面配置
st.set_page_config(
    page_title="医学编码数据预处理器",
    page_icon="📊",
    layout="wide"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main {
        background: linear-gradient(135deg, #e0f2fe 0%, #ddd6fe 100%);
    }
    .stButton>button {
        border-radius: 0.5rem;
        font-weight: 600;
        transition: all 0.3s;
    }
    h1 {
        color: #1f2937;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .subtitle {
        text-align: center;
        color: #6b7280;
        margin-bottom: 2rem;
    }
    .section-header {
        background: linear-gradient(90deg, #4f46e5 0%, #7c3aed 100%);
        color: white;
        padding: 0.75rem 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0 0.5rem 0;
        font-weight: 600;
    }
    .rule-card {
        background: white;
        border-left: 4px solid #4f46e5;
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 0.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .variable-header {
        background: #f3f4f6;
        padding: 0.75rem;
        border-radius: 0.5rem;
        margin: 1rem 0 0.5rem 0;
        font-weight: 600;
        color: #1f2937;
    }
    .rule-summary {
        background: #f9fafb;
        padding: 0.5rem;
        margin: 0.25rem 0;
        border-radius: 0.25rem;
        font-family: monospace;
        font-size: 0.9rem;
        color: #374151;
    }
</style>
""", unsafe_allow_html=True)

# 初始化session state
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None
if 'selected_sheets' not in st.session_state:
    st.session_state.selected_sheets = {}
if 'sheet_variables' not in st.session_state:
    st.session_state.sheet_variables = {}  # {sheet_name: {var_name: {...config}}}

# ==================== 配置管理功能 ====================

def load_all_configs():
    try:
        with open('excel_processor_configs.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def save_all_configs(all_configs):
    try:
        with open('excel_processor_configs.json', 'w', encoding='utf-8') as f:
            json.dump(all_configs, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"保存失败: {str(e)}")
        return False

def save_current_config(config_name):
    all_configs = load_all_configs()
    all_configs[config_name] = {
        'sheet_variables': st.session_state.sheet_variables,
        'saved_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    return save_all_configs(all_configs)

def load_config(config_name):
    all_configs = load_all_configs()
    if config_name in all_configs:
        st.session_state.sheet_variables = all_configs[config_name]['sheet_variables']
        return True
    return False

def delete_config(config_name):
    all_configs = load_all_configs()
    if config_name in all_configs:
        del all_configs[config_name]
        return save_all_configs(all_configs)
    return False

# ==================== 通用数据提取函数 ====================

def evaluate_condition(row_value, operator, compare_value):
    """评估条件是否满足"""
    # 处理空值
    if pd.isna(row_value):
        row_value = ""
    else:
        row_value = str(row_value)
    
    compare_value = str(compare_value) if compare_value is not None else ""
    
    if operator == "=":
        return row_value == compare_value
    elif operator == "<>":
        return row_value != compare_value
    elif operator == "包含":
        return compare_value in row_value
    elif operator == "不包含":
        return compare_value not in row_value
    elif operator == ">":
        try:
            return float(row_value) > float(compare_value)
        except:
            return False
    elif operator == "<":
        try:
            return float(row_value) < float(compare_value)
        except:
            return False
    elif operator == ">=":
        try:
            return float(row_value) >= float(compare_value)
        except:
            return False
    elif operator == "<=":
        try:
            return float(row_value) <= float(compare_value)
        except:
            return False
    return False

def extract_value(row, extract_type, extract_value_type, extract_value, regex_pattern=None, capture_group=1):
    """根据提取方式提取值"""
    # 处理提取值
    if extract_value_type == "固定文本":
        source_value = extract_value
    else:  # 从列提取
        if extract_value not in row.index:
            return []
        source_value = row[extract_value]
        if pd.isna(source_value):
            source_value = ""
        else:
            source_value = str(source_value)
    
    # 根据提取方式处理
    if extract_type == "直接提取":
        return [source_value] if source_value else []
    
    elif extract_type == "正则提取":
        if not regex_pattern or not source_value:
            return []
        
        results = []
        try:
            for match in re.finditer(regex_pattern, source_value):
                groups = match.groups()
                if len(groups) >= capture_group:
                    extracted = groups[capture_group - 1].strip()
                    if extracted:
                        results.append(extracted)
        except Exception as e:
            st.warning(f"正则表达式错误: {str(e)}")
        
        return results
    
    return []

def process_variable_rules(row, rules, separator):
    """处理一个变量的所有规则"""
    all_values = []
    
    for rule in rules:
        # 检查条件
        condition_column = rule.get('condition_column', '')
        condition_operator = rule.get('condition_operator', '=')
        condition_value = rule.get('condition_value', '')
        
        if not condition_column or condition_column not in row.index:
            continue
        
        # 评估条件
        if evaluate_condition(row[condition_column], condition_operator, condition_value):
            # 提取值
            extracted = extract_value(
                row,
                rule.get('extract_type', '直接取出'),
                rule.get('extract_value_type', '从列提取'),
                rule.get('extract_value', ''),
                rule.get('regex_pattern', ''),
                rule.get('capture_group', 1)
            )
            all_values.extend(extracted)
    
    # 拼接、拆分、去重、排序
    if not all_values:
        return ''
    
    combined = separator.join(all_values)
    split_values = [v.strip() for v in combined.split(separator) if v.strip()]
    unique_sorted = sorted(set(split_values))
    
    return separator.join(unique_sorted)

# ==================== 侧边栏：配置管理 ====================

with st.sidebar:
    st.markdown("### 💾 配置管理")
    
    # 保存配置
    with st.expander("保存当前配置", expanded=False):
        save_name = st.text_input(
            "配置名称",
            placeholder="例如: 默认配置",
            key="save_config_name"
        )
        if st.button("💾 保存", key="save_config_btn", use_container_width=True):
            if save_name:
                if save_current_config(save_name):
                    st.success(f"✅ 配置 '{save_name}' 已保存!")
                    st.rerun()
            else:
                st.warning("⚠️ 请输入配置名称")
    
    # 加载配置
    with st.expander("加载配置", expanded=False):
        all_configs = load_all_configs()
        if all_configs:
            config_options = list(all_configs.keys())
            selected_config = st.selectbox(
                "选择配置",
                options=config_options,
                key="load_config_select"
            )
            
            if selected_config:
                saved_time = all_configs[selected_config].get('saved_time', '未知')
                st.caption(f"⏰ {saved_time}")
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("📥 加载", key="load_btn", use_container_width=True):
                        if load_config(selected_config):
                            st.success(f"✅ 已加载 '{selected_config}'")
                            st.rerun()
                
                with col2:
                    if st.button("🗑️ 删除", key="delete_btn", use_container_width=True):
                        if delete_config(selected_config):
                            st.success(f"✅ 已删除 '{selected_config}'")
                            st.rerun()
        else:
            st.info("ℹ️ 暂无保存的配置")

# ==================== 主页面 ====================

st.markdown("<h1>📊 医学编码数据预处理器</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle'>导入、配置、导出 - 轻松处理您的数据</p>", unsafe_allow_html=True)

# ==================== 布局：上传区域（居中） ====================

col_left, col_center, col_right = st.columns([1, 2, 1])

with col_center:
    st.markdown("<div class='section-header'>📁 上传 Excel 文件</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "选择Excel文件",
        type=['xlsx', 'xls'],
        help="支持 .xlsx 和 .xls 格式",
        label_visibility="collapsed"
    )
    
    if uploaded_file is not None:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            st.session_state.uploaded_file = uploaded_file
            st.session_state.excel_data = excel_file
            
            if not st.session_state.selected_sheets:
                st.session_state.selected_sheets = {
                    sheet: True for sheet in excel_file.sheet_names
                }
            
            st.success(f"✅ 成功加载: {uploaded_file.name} ({len(excel_file.sheet_names)} 个工作表)")
            
        except Exception as e:
            st.error(f"❌ 文件读取失败: {str(e)}")

st.markdown("---")

# ==================== 主要区域：Sheet选择 + 配置 ====================

if st.session_state.excel_data is not None:
    
    # 布局：左侧Sheet选择，右侧配置
    col_sheets, col_config = st.columns([1, 3])
    
    # ========== 左侧：Sheet选择 ==========
    with col_sheets:
        st.markdown("<div class='section-header'>📋 选择工作表</div>", unsafe_allow_html=True)
        
        # 用于触发复选框重新渲染
        if 'sheet_select_trigger' not in st.session_state:
            st.session_state.sheet_select_trigger = 0
        
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("✅ 全选", use_container_width=True):
                for sheet in st.session_state.selected_sheets:
                    st.session_state.selected_sheets[sheet] = True
                st.session_state.sheet_select_trigger += 1
                st.rerun()
        with col_b:
            if st.button("❌ 全不选", use_container_width=True):
                for sheet in st.session_state.selected_sheets:
                    st.session_state.selected_sheets[sheet] = False
                st.session_state.sheet_select_trigger += 1
                st.rerun()
        
        st.markdown("---")
        
        for sheet_name in st.session_state.excel_data.sheet_names:
            checked = st.checkbox(
                f"📄 {sheet_name}",
                value=st.session_state.selected_sheets.get(sheet_name, True),
                key=f"sheet_{sheet_name}_{st.session_state.sheet_select_trigger}"
            )
            st.session_state.selected_sheets[sheet_name] = checked
    
    # ========== 右侧：变量配置 ==========
    with col_config:
        st.markdown("<div class='section-header'>⚙️ 变量配置</div>", unsafe_allow_html=True)
        
        selected_sheets = [name for name, sel in st.session_state.selected_sheets.items() if sel]
        
        if not selected_sheets:
            st.warning("⚠️ 请先选择至少一个工作表")
        else:
            # 为每个选中的sheet配置
            for sheet_name in selected_sheets:
                with st.expander(f"📊 {sheet_name}", expanded=True):
                    
                    # 初始化该sheet的变量配置
                    if sheet_name not in st.session_state.sheet_variables:
                        st.session_state.sheet_variables[sheet_name] = {}
                    
                    sheet_vars = st.session_state.sheet_variables[sheet_name]
                    
                    # 添加新变量按钮
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        new_var_name = st.text_input(
                            "新变量名",
                            placeholder="例如: ROUTE, INDICATION",
                            key=f"new_var_{sheet_name}"
                        )
                    with col2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("➕ 添加变量", key=f"add_var_{sheet_name}"):
                            if new_var_name and new_var_name not in sheet_vars:
                                sheet_vars[new_var_name] = {
                                    'separator': ';',
                                    'rules': []
                                }
                                st.rerun()
                            elif new_var_name in sheet_vars:
                                st.warning("⚠️ 变量名已存在")
                            else:
                                st.warning("⚠️ 请输入变量名")
                    
                    st.markdown("---")
                    
                    # 显示每个变量及其规则
                    for var_name in list(sheet_vars.keys()):
                        var_config = sheet_vars[var_name]
                        
                        st.markdown(f"<div class='variable-header'>📋 {var_name}</div>", unsafe_allow_html=True)
                        
                        # 分隔符和删除按钮
                        col1, col2, col3 = st.columns([2, 2, 1])
                        with col1:
                            var_config['separator'] = st.text_input(
                                "分隔符",
                                value=var_config.get('separator', ';'),
                                key=f"sep_{sheet_name}_{var_name}"
                            )
                        with col2:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button(f"➕ 添加规则", key=f"add_rule_{sheet_name}_{var_name}"):
                                var_config['rules'].append({
                                    'condition_column': '',
                                    'condition_operator': '=',
                                    'condition_value': '',
                                    'extract_type': '直接取出',
                                    'extract_value_type': '从列提取',
                                    'extract_value': '',
                                    'regex_pattern': '',
                                    'capture_group': 1
                                })
                                st.rerun()
                        with col3:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("🗑️", key=f"del_var_{sheet_name}_{var_name}"):
                                del sheet_vars[var_name]
                                st.rerun()
                        
                        # 显示规则预览
                        if var_config['rules']:
                            for idx, rule in enumerate(var_config['rules']):
                                cond_col = rule.get('condition_column', '')
                                cond_op = rule.get('condition_operator', '=')
                                cond_val = rule.get('condition_value', '')
                                ext_type = rule.get('extract_type', '直接取出')
                                ext_val_type = rule.get('extract_value_type', '从列提取')
                                ext_val = rule.get('extract_value', '')
                                
                                # 构建预览文本
                                rule_text = f"{'├─' if idx < len(var_config['rules'])-1 else '└─'} 规则{idx+1}: "
                                rule_text += f"当 {cond_col} {cond_op} "
                                rule_text += f'"{cond_val}"' if cond_val else '(空)'
                                rule_text += f" 时，{ext_type} "
                                
                                if ext_val_type == "固定文本":
                                    rule_text += f'"{ext_val}"'
                                else:
                                    rule_text += f'{ext_val}'
                                
                                if ext_type == "正则取出":
                                    regex = rule.get('regex_pattern', '')
                                    cap_grp = rule.get('capture_group', 1)
                                    rule_text += f" (模式: {regex}, 组{cap_grp})"
                                
                                st.markdown(f"<div class='rule-summary'>{rule_text}</div>", unsafe_allow_html=True)
                        
                        # 编辑每个规则
                        for idx, rule in enumerate(var_config['rules']):
                            # 使用折叠面板展示单条规则
                            with st.expander(f"🔧 规则 {idx + 1}", expanded=False):
                                
                                # 删除规则按钮
                                if st.button("🗑️ 删除此规则", key=f"del_rule_{sheet_name}_{var_name}_{idx}"):
                                    var_config['rules'].pop(idx)
                                    st.rerun()
                                
                                # 条件配置
                                st.markdown("**条件设置**")
                                col1, col2, col3 = st.columns(3)
                                
                                with col1:
                                    rule['condition_column'] = st.text_input(
                                        "判断变量(列名)",
                                        value=rule.get('condition_column', ''),
                                        placeholder="例如: CMROUTE",
                                        key=f"cond_col_{sheet_name}_{var_name}_{idx}"
                                    )
                                
                                with col2:
                                    operators = ["=", "<>", "包含", "不包含", ">", "<", ">=", "<="]
                                    current_op = rule.get('condition_operator', '=')
                                    rule['condition_operator'] = st.selectbox(
                                        "逻辑比较符",
                                        options=operators,
                                        index=operators.index(current_op) if current_op in operators else 0,
                                        key=f"cond_op_{sheet_name}_{var_name}_{idx}"
                                    )
                                
                                with col3:
                                    rule['condition_value'] = st.text_input(
                                        "判断值",
                                        value=rule.get('condition_value', ''),
                                        placeholder="留空表示空值",
                                        key=f"cond_val_{sheet_name}_{var_name}_{idx}"
                                    )
                                
                                st.markdown("---")
                                
                                # 提取配置
                                st.markdown("**提取设置**")
                                
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    extract_types = ["直接取出", "正则取出"]
                                    current_ext = rule.get('extract_type', '直接取出')
                                    rule['extract_type'] = st.selectbox(
                                        "提取方式",
                                        options=extract_types,
                                        index=extract_types.index(current_ext) if current_ext in extract_types else 0,
                                        key=f"ext_type_{sheet_name}_{var_name}_{idx}"
                                    )
                                
                                with col2:
                                    value_types = ["从列提取", "固定文本"]
                                    current_val_type = rule.get('extract_value_type', '从列提取')
                                    rule['extract_value_type'] = st.selectbox(
                                        "提取值类型",
                                        options=value_types,
                                        index=value_types.index(current_val_type) if current_val_type in value_types else 0,
                                        key=f"ext_val_type_{sheet_name}_{var_name}_{idx}"
                                    )
                                
                                # 提取值
                                if rule['extract_value_type'] == "从列提取":
                                    rule['extract_value'] = st.text_input(
                                        "提取值(列名)",
                                        value=rule.get('extract_value', ''),
                                        placeholder="例如: CMROUTE",
                                        key=f"ext_val_{sheet_name}_{var_name}_{idx}"
                                    )
                                else:
                                    rule['extract_value'] = st.text_input(
                                        "提取值(固定文本)",
                                        value=rule.get('extract_value', ''),
                                        placeholder="例如: 预防感冒",
                                        key=f"ext_val_{sheet_name}_{var_name}_{idx}"
                                    )
                                
                                # 正则配置
                                if rule['extract_type'] == "正则取出":
                                    col1, col2 = st.columns([3, 1])
                                    with col1:
                                        rule['regex_pattern'] = st.text_input(
                                            "正则表达式",
                                            value=rule.get('regex_pattern', ''),
                                            placeholder=r"例如: (\d+)#(.+?)[;,]",
                                            key=f"regex_{sheet_name}_{var_name}_{idx}",
                                            help="使用 .+? 进行非贪婪匹配"
                                        )
                                    with col2:
                                        rule['capture_group'] = st.number_input(
                                            "捕获组序号",
                                            value=rule.get('capture_group', 1),
                                            min_value=1,
                                            step=1,
                                            key=f"cap_grp_{sheet_name}_{var_name}_{idx}"
                                        )
                        
                        st.markdown("<br>", unsafe_allow_html=True)
    
    # ==================== 导出区域 ====================
    st.markdown("---")
    st.markdown("<div class='section-header'>📥 导出处理后的文件</div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col2:
        if st.button("🚀 处理并导出", type="primary", use_container_width=True):
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    
                    for sheet_name in selected_sheets:
                        # 读取数据
                        df = pd.read_excel(
                            st.session_state.uploaded_file,
                            sheet_name=sheet_name,
                            dtype=str
                        )
                        
                        # 处理该sheet的所有变量
                        if sheet_name in st.session_state.sheet_variables:
                            for var_name, var_config in st.session_state.sheet_variables[sheet_name].items():
                                separator = var_config.get('separator', ';')
                                rules = var_config.get('rules', [])
                                
                                if rules:
                                    df[var_name] = df.apply(
                                        lambda row: process_variable_rules(row, rules, separator),
                                        axis=1
                                    )
                        
                        # 写入Excel
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 获取工作表对象进行格式化
                        worksheet = writer.sheets[sheet_name]
                        
                        # 设置边框样式
                        thin_border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
                        
                        # 设置首行样式（淡蓝色背景）
                        header_fill = PatternFill(start_color='B4C7E7', end_color='B4C7E7', fill_type='solid')
                        header_font = Font(bold=True)
                        header_alignment = Alignment(horizontal='center', vertical='center')
                        
                        # 应用首行格式
                        for col_idx, col in enumerate(df.columns, 1):
                            cell = worksheet.cell(row=1, column=col_idx)
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = header_alignment
                            cell.border = thin_border
                        
                        # 为所有数据单元格添加边框
                        for row_idx in range(2, len(df) + 2):
                            for col_idx in range(1, len(df.columns) + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.border = thin_border
                        
                        # 冻结首行
                        worksheet.freeze_panes = 'A2'
                        
                        # 开启自动筛选
                        worksheet.auto_filter.ref = worksheet.dimensions
                
                output.seek(0)
                
                # 生成下载文件名
                original_name = st.session_state.uploaded_file.name
                if original_name.endswith('.xlsx'):
                    new_name = original_name.replace('.xlsx', '_processed.xlsx')
                elif original_name.endswith('.xls'):
                    new_name = original_name.replace('.xls', '_processed.xlsx')
                else:
                    new_name = original_name + '_processed.xlsx'
                
                st.download_button(
                    label="⬇️ 下载处理后的文件",
                    data=output,
                    file_name=new_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success("✅ 文件处理完成!")
                
            except Exception as e:
                st.error(f"❌ 处理失败: {str(e)}")
                st.exception(e)

# 页脚
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #6b7280;'>医学编码数据预处理器 v2.0 | Powered by Streamlit & Pandas</p>",
    unsafe_allow_html=True
)
