import streamlit as st
import pandas as pd
import io
import json
import re
from typing import Dict, List, Any
from datetime import datetime

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
        width: 100%;
        border-radius: 0.5rem;
        font-weight: 600;
        transition: all 0.3s;
    }
    .step-indicator {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 2rem 0;
    }
    .step-circle {
        width: 40px;
        height: 40px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        margin: 0 10px;
    }
    .step-active {
        background-color: #4f46e5;
        color: white;
    }
    .step-inactive {
        background-color: #d1d5db;
        color: #6b7280;
    }
    .step-line {
        width: 96px;
        height: 4px;
        background-color: #d1d5db;
    }
    .step-line-active {
        background-color: #4f46e5;
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
    .config-section {
        background: #f9fafb;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .success-message {
        padding: 1rem;
        background-color: #d1fae5;
        border-left: 4px solid #10b981;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# 初始化session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None
if 'selected_sheets' not in st.session_state:
    st.session_state.selected_sheets = {}
if 'sheet_configs' not in st.session_state:
    st.session_state.sheet_configs = {}
if 'select_all_trigger' not in st.session_state:
    st.session_state.select_all_trigger = 0

# ==================== 配置管理功能 ====================

# 加载所有保存的配置
def load_all_configs():
    try:
        with open('excel_processor_configs.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

# 保存所有配置
def save_all_configs(all_configs):
    try:
        with open('excel_processor_configs.json', 'w', encoding='utf-8') as f:
            json.dump(all_configs, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"保存失败: {str(e)}")
        return False

# 保存当前配置
def save_current_config(config_name):
    all_configs = load_all_configs()
    all_configs[config_name] = {
        'sheet_configs': st.session_state.sheet_configs,
        'saved_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    return save_all_configs(all_configs)

# 加载指定配置
def load_config(config_name):
    all_configs = load_all_configs()
    if config_name in all_configs:
        st.session_state.sheet_configs = all_configs[config_name]['sheet_configs']
        return True
    return False

# 删除指定配置
def delete_config(config_name):
    all_configs = load_all_configs()
    if config_name in all_configs:
        del all_configs[config_name]
        return save_all_configs(all_configs)
    return False

# 重命名配置
def rename_config(old_name, new_name):
    all_configs = load_all_configs()
    if old_name in all_configs and new_name not in all_configs:
        all_configs[new_name] = all_configs.pop(old_name)
        all_configs[new_name]['saved_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        return save_all_configs(all_configs)
    return False

# 步骤指示器
def render_step_indicator(current_step):
    steps_html = '<div class="step-indicator">'
    for i in range(1, 4):
        step_class = "step-active" if i <= current_step else "step-inactive"
        steps_html += f'<div class="step-circle {step_class}">{i}</div>'
        if i < 3:
            line_class = "step-line-active" if i < current_step else ""
            steps_html += f'<div class="step-line {line_class}"></div>'
    steps_html += '</div>'
    st.markdown(steps_html, unsafe_allow_html=True)

# 标题
st.markdown("<h1>📊 医学编码数据预处理器</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle'>导入、配置、导出 - 轻松处理您的数据</p>", unsafe_allow_html=True)

# 显示步骤指示器
render_step_indicator(st.session_state.step)

# ==================== 步骤1: 上传文件 ====================
if st.session_state.step == 1:
    st.markdown("### 📁 上传 Excel 文件")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        uploaded_file = st.file_uploader(
            "选择Excel文件",
            type=['xlsx', 'xls'],
            help="支持 .xlsx 和 .xls 格式"
        )
        
        if uploaded_file is not None:
            try:
                # 读取所有sheets
                excel_file = pd.ExcelFile(uploaded_file)
                st.session_state.uploaded_file = uploaded_file
                st.session_state.excel_data = excel_file
                
                # 初始化选中状态
                st.session_state.selected_sheets = {
                    sheet: True for sheet in excel_file.sheet_names
                }
                
                st.success(f"✅ 成功加载文件: {uploaded_file.name}")
                st.info(f"📄 发现 {len(excel_file.sheet_names)} 个工作表")
                
                if st.button("▶️ 下一步:选择工作表", type="primary"):
                    st.session_state.step = 2
                    st.rerun()
                    
            except Exception as e:
                st.error(f"❌ 文件读取失败: {str(e)}")

# ==================== 步骤2: 选择Sheet ====================
elif st.session_state.step == 2:
    st.markdown("### 📋 选择要保留的工作表")
    
    col1, col2 = st.columns([3, 1])
    with col1:
        if st.button("⬅️ 上一步"):
            st.session_state.step = 1
            st.rerun()
    with col2:
        pass
    
    # 全选/全不选按钮
    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("✅ 全选", key="select_all_btn"):
            for sheet in st.session_state.selected_sheets:
                st.session_state.selected_sheets[sheet] = True
            st.session_state.select_all_trigger += 1
            st.rerun()
    with col2:
        if st.button("❌ 全不选", key="deselect_all_btn"):
            for sheet in st.session_state.selected_sheets:
                st.session_state.selected_sheets[sheet] = False
            st.session_state.select_all_trigger += 1
            st.rerun()
    
    st.markdown("---")
    
    # 显示所有sheets的复选框 - 使用 session_state 直接控制
    for sheet_name in st.session_state.excel_data.sheet_names:
        # 使用唯一的 key,并通过 session_state 直接管理状态
        checkbox_key = f"sheet_select_{sheet_name}_{st.session_state.select_all_trigger}"
        selected = st.checkbox(
            f"📄 {sheet_name}",
            value=st.session_state.selected_sheets.get(sheet_name, True),
            key=checkbox_key
        )
        st.session_state.selected_sheets[sheet_name] = selected
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬅️ 上一步", key="back_from_step2"):
            st.session_state.step = 1
            st.rerun()
    with col2:
        if st.button("▶️ 下一步:配置列生成", type="primary"):
            st.session_state.step = 3
            st.rerun()

# ==================== 步骤3: 配置列生成 ====================
elif st.session_state.step == 3:
    st.markdown("### ⚙️ 配置列生成规则")
    
    # ==================== 配置管理区域 ====================
    st.markdown("#### 💾 配置管理")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # 保存配置
        with st.expander("💾 保存当前配置", expanded=False):
            save_name = st.text_input(
                "配置名称",
                placeholder="例如: 默认配置",
                key="save_config_name"
            )
            if st.button("保存", key="save_config_btn"):
                if save_name:
                    if save_current_config(save_name):
                        st.success(f"✅ 配置 '{save_name}' 已保存!")
                    else:
                        st.error("❌ 保存失败,请重试")
                else:
                    st.warning("⚠️ 请输入配置名称")
    
    with col2:
        # 加载配置
        with st.expander("📂 加载配置", expanded=False):
            all_configs = load_all_configs()
            if all_configs:
                config_options = list(all_configs.keys())
                selected_config = st.selectbox(
                    "选择配置",
                    options=config_options,
                    key="load_config_select"
                )
                
                if selected_config:
                    # 显示配置信息
                    saved_time = all_configs[selected_config].get('saved_time', '未知')
                    st.caption(f"保存时间: {saved_time}")
                    
                    col_a, col_b, col_c = st.columns(3)
                    
                    with col_a:
                        if st.button("📥 加载", key="load_config_btn"):
                            if load_config(selected_config):
                                st.success(f"✅ 已加载配置 '{selected_config}'")
                                st.rerun()
                    
                    with col_b:
                        if st.button("🗑️ 删除", key="delete_config_btn"):
                            if delete_config(selected_config):
                                st.success(f"✅ 已删除配置 '{selected_config}'")
                                st.rerun()
                            else:
                                st.error("❌ 删除失败")
                    
                    with col_c:
                        # 重命名功能
                        if st.button("✏️ 重命名", key="rename_config_btn"):
                            st.session_state.show_rename = True
                    
                    # 重命名输入框
                    if st.session_state.get('show_rename', False):
                        new_name = st.text_input(
                            "新名称",
                            value=selected_config,
                            key="rename_config_input"
                        )
                        col_x, col_y = st.columns(2)
                        with col_x:
                            if st.button("确认重命名", key="confirm_rename_btn"):
                                if new_name and new_name != selected_config:
                                    if rename_config(selected_config, new_name):
                                        st.success(f"✅ 已重命名为 '{new_name}'")
                                        st.session_state.show_rename = False
                                        st.rerun()
                                    else:
                                        st.error("❌ 重命名失败(可能名称已存在)")
                        with col_y:
                            if st.button("取消", key="cancel_rename_btn"):
                                st.session_state.show_rename = False
                                st.rerun()
            else:
                st.info("ℹ️ 暂无保存的配置")
    
    st.markdown("---")
    
    # ==================== Sheet配置区域 ====================
    # 为每个选中的sheet配置
    selected_sheet_names = [
        name for name, selected in st.session_state.selected_sheets.items() 
        if selected
    ]
    
    for sheet_name in selected_sheet_names:
        with st.expander(f"📊 {sheet_name}", expanded=True):
            
            # 初始化配置
            if sheet_name not in st.session_state.sheet_configs:
                st.session_state.sheet_configs[sheet_name] = {
                    'generate_route': False,
                    'route_config': {
                        'source_column_a': '',
                        'source_column_b': '',
                        'condition_value': '其他'
                    },
                    'generate_indication': False,
                    'indication_config': {
                        'separator': ';',
                        'join_separator': ';',
                        'columns': []
                    }
                }
            
            config = st.session_state.sheet_configs[sheet_name]
            
            # ROUTE列配置
            st.markdown("#### 🚗 ROUTE 列配置")
            config['generate_route'] = st.checkbox(
                "生成 ROUTE 列",
                value=config['generate_route'],
                key=f"route_enable_{sheet_name}"
            )
            
            if config['generate_route']:
                st.markdown('<div class="config-section">', unsafe_allow_html=True)
                config['route_config']['source_column_a'] = st.text_input(
                    "给药途径列",
                    value=config['route_config']['source_column_a'],
                    placeholder="例如: 给药途径",
                    key=f"route_cola_{sheet_name}"
                )
                config['route_config']['source_column_b'] = st.text_input(
                    "其他给药途径列",
                    value=config['route_config']['source_column_b'],
                    placeholder="例如: 若其他途径，请详述",
                    key=f"route_colb_{sheet_name}"
                )
                config['route_config']['condition_value'] = st.text_input(
                    "条件值 (当给药途径列等于此值时使用其他给药途径列的值)",
                    value=config['route_config']['condition_value'],
                    placeholder="例如: 其他",
                    key=f"route_cond_{sheet_name}"
                )
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown("---")
            
            # INDICATION列配置
            st.markdown("#### 🏷️ INDICATION 列配置")
            config['generate_indication'] = st.checkbox(
                "生成 INDICATION 列",
                value=config['generate_indication'],
                key=f"indication_enable_{sheet_name}"
            )
            
            if config['generate_indication']:
                st.info("💡 提示：在正则表达式中使用 `.+?` (非贪婪模式) 而不是 `.+` (贪婪模式)，可以避免跨越分隔符匹配过多内容。")
            
            if config['generate_indication']:
                st.markdown('<div class="config-section">', unsafe_allow_html=True)
                
                col_a, col_b = st.columns(2)
                with col_a:
                    config['indication_config']['separator'] = st.text_input(
                        "输出分隔符",
                        value=config['indication_config']['separator'],
                        placeholder="例如: ;",
                        key=f"indication_sep_{sheet_name}",
                        help="最终 INDICATION 列中各值之间的分隔符"
                    )
                with col_b:
                    config['indication_config']['join_separator'] = st.text_input(
                        "拼接分隔符",
                        value=config['indication_config'].get('join_separator', config['indication_config']['separator']),
                        placeholder="例如: ;",
                        key=f"indication_join_sep_{sheet_name}",
                        help="用于拼接所有提取值的分隔符，拼接后会按此分隔符拆分、去重、排序"
                    )
                
                st.markdown("**提取列配置**")
                
                # 添加列按钮
                if st.button(f"➕ 添加列", key=f"add_col_{sheet_name}"):
                    config['indication_config']['columns'].append({
                        'column_name': '',
                        'extract_type': 'direct',
                        'regex_pattern': '',
                        'capture_group': 2,
                        'conditional_column': '',
                        'conditional_value': '',
                        'mapping_column': ''
                    })
                    st.rerun()
                
                # 显示每个列配置
                for idx, col_config in enumerate(config['indication_config']['columns']):
                    st.markdown(f"**列 {idx + 1}**")
                    
                    col1, col2 = st.columns([5, 1])
                    with col1:
                        col_config['column_name'] = st.text_input(
                            "列名",
                            value=col_config['column_name'],
                            placeholder="例如: 不良事件描述、既往史描述、预防用药描述",
                            key=f"col_name_{sheet_name}_{idx}"
                        )
                    with col2:
                        if st.button("🗑️", key=f"del_col_{sheet_name}_{idx}"):
                            config['indication_config']['columns'].pop(idx)
                            st.rerun()
                    
                    col_config['extract_type'] = st.selectbox(
                        "提取方式",
                        options=['direct', 'regex', 'conditional'],
                        format_func=lambda x: {
                            'direct': '直接取值',
                            'regex': '正则提取',
                            'conditional': '条件映射'
                        }[x],
                        index=['direct', 'regex', 'conditional'].index(col_config['extract_type']),
                        key=f"extract_type_{sheet_name}_{idx}"
                    )
                    
                    if col_config['extract_type'] == 'regex':
                        col_config['regex_pattern'] = st.text_input(
                            "正则表达式",
                            value=col_config['regex_pattern'],
                            placeholder=r"例如: (\d+)#([^,;]+)",
                            key=f"regex_{sheet_name}_{idx}",
                            help="使用 .+? 进行非贪婪匹配，避免跨越分隔符"
                        )
                        col_config['capture_group'] = st.number_input(
                            "捕获组序号",
                            value=col_config['capture_group'],
                            min_value=1,
                            step=1,
                            key=f"capture_{sheet_name}_{idx}",
                            help="指定使用第几个括号捕获的内容"
                        )
                    
                    elif col_config['extract_type'] == 'conditional':
                        col_config['conditional_column'] = st.text_input(
                            "条件列名",
                            value=col_config['conditional_column'],
                            placeholder="例如: O",
                            key=f"cond_col_{sheet_name}_{idx}"
                        )
                        col_config['conditional_value'] = st.text_input(
                            "条件值",
                            value=col_config['conditional_value'],
                            placeholder="当条件列等于此值时",
                            key=f"cond_val_{sheet_name}_{idx}"
                        )
                        col_config['mapping_column'] = st.text_input(
                            "取值列名",
                            value=col_config['mapping_column'],
                            placeholder="例如: P",
                            key=f"map_col_{sheet_name}_{idx}"
                        )
                    
                    st.markdown("---")
                
                st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 处理和导出
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬅️ 上一步", key="back_from_step3"):
            st.session_state.step = 2
            st.rerun()
    with col2:
        if st.button("📥 导出处理后的文件", type="primary"):
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    
                    for sheet_name in selected_sheet_names:
                        # 读取sheet数据
                        df = pd.read_excel(
                            st.session_state.uploaded_file,
                            sheet_name=sheet_name,
                            dtype=str
                        )
                        
                        config = st.session_state.sheet_configs.get(sheet_name, {})
                        
                        # 生成ROUTE列
                        if config.get('generate_route', False):
                            route_cfg = config['route_config']
                            col_a = route_cfg['source_column_a']
                            col_b = route_cfg['source_column_b']
                            cond_val = route_cfg['condition_value']
                            
                            if col_a in df.columns and col_b in df.columns:
                                df['ROUTE'] = df.apply(
                                    lambda row: row[col_b] if str(row[col_a]) == cond_val else row[col_a],
                                    axis=1
                                )
                        
                        # 生成INDICATION列
                        if config.get('generate_indication', False):
                            indication_cfg = config['indication_config']
                            output_separator = indication_cfg['separator']
                            join_separator = indication_cfg.get('join_separator', output_separator)
                            
                            def extract_indication(row):
                                values = []
                                
                                for col_cfg in indication_cfg['columns']:
                                    col_name = col_cfg['column_name']
                                    if col_name not in df.columns:
                                        continue
                                    
                                    cell_value = str(row[col_name]) if pd.notna(row[col_name]) else ''
                                    if not cell_value:
                                        continue
                                    
                                    if col_cfg['extract_type'] == 'direct':
                                        values.append(cell_value)
                                    
                                    elif col_cfg['extract_type'] == 'regex':
                                        pattern = col_cfg['regex_pattern'] or r'(\d+)#([^,;]+)'
                                        capture_group = int(col_cfg['capture_group'])
                                        
                                        # 使用 finditer 进行全局匹配
                                        for match in re.finditer(pattern, cell_value):
                                            groups = match.groups()
                                            if len(groups) >= capture_group:
                                                extracted_value = groups[capture_group - 1].strip()
                                                if extracted_value:
                                                    values.append(extracted_value)
                                    
                                    elif col_cfg['extract_type'] == 'conditional':
                                        cond_col = col_cfg['conditional_column']
                                        cond_val = col_cfg['conditional_value']
                                        map_col = col_cfg['mapping_column']
                                        
                                        if cond_col in df.columns and map_col in df.columns:
                                            if str(row[cond_col]) == cond_val:
                                                map_value = str(row[map_col]) if pd.notna(row[map_col]) else ''
                                                if map_value:
                                                    values.append(map_value)
                                
                                # 处理流程：拼接 → 拆分 → 去重 → 排序 → 输出
                                if not values:
                                    return ''
                                
                                # 1. 用拼接分隔符拼接所有值
                                combined = join_separator.join(values)
                                
                                # 2. 按拼接分隔符拆分
                                split_values = [v.strip() for v in combined.split(join_separator) if v.strip()]
                                
                                # 3. 去重并排序
                                unique_sorted_values = sorted(set(split_values))
                                
                                # 4. 用输出分隔符拼接
                                return output_separator.join(unique_sorted_values)
                            
                            df['INDICATION'] = df.apply(extract_indication, axis=1)
                        
                        # 写入Excel
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                output.seek(0)
                
                # 提供下载
                original_name = st.session_state.uploaded_file.name
                # 正确处理文件后缀
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
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("✅ 文件处理完成!")
                
            except Exception as e:
                st.error(f"❌ 处理失败: {str(e)}")
                st.exception(e)

# 页脚
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #6b7280;'>Excel 数据处理器 | Powered by Streamlit & Pandas</p>",
    unsafe_allow_html=True
)
