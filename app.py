import streamlit as st
import pandas as pd
from collections import defaultdict
import sys
import re
import io
from datetime import datetime
import tempfile
import os

# Streamlit App Title and Description
st.title("批量广告上传模版-生成工具")
st.markdown("""
### 代码内容说明
新：此工具用于从上传的 Excel 文件（默认 sheet: '广告模版'）中提取全局设置、活动数据和关键词信息，生成广告 Header 文件。  
**主要功能：**  
- 支持（品牌旗舰店、商品集、商品详情页、SP-商品推广）主题的动态区域检测和数据提取。  
- 处理广告活动、广告组、视频/商品集广告、关键词、否定关键词、商品定向等行生成。  
- 自动填充默认值（如预算类型 '每日'、状态 '已启用'）。  
- 检测重复否定关键词并暂停生成（打印警告）。  
- 输出多Sheet工作簿：'品牌广告' Sheet (SB/SBV) 和 'SP-商品推广' Sheet (SP)，每个有独立列头。  

**使用步骤：**  
1. 上传 Excel 文件（文件名任意，需包含 '广告模版' sheet）。  
2. 点击 "生成 Header 文件" 按钮。  
3. 下载生成的 "header-YYYY-MM-DD HH:MM.xlsx" 文件。  

**注意：**  
- 文件需符合脚本预期结构（A 列主题行、B 列活动名称等）。  
- 如遇错误（如未找到主题），页面将显示日志。  
- 生成时间精确到分钟（基于当前时间）。  
""")

# File Uploader
uploaded_file = st.file_uploader("上传 Excel 文件", type=['xlsx', 'xls'])

# Function from the original script (copied and adapted)
def generate_header_for_sbv_brand_store(uploaded_bytes, sheet_name='广告模版'):
    # Create a temporary file from bytes
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_bytes)
        input_file = tmp.name
    
    try:
        # Read the entire file, header=0
        df_survey = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
        st.write(f"成功读取文件，数据形状：{df_survey.shape}")
        st.write(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        st.error(f"错误：未找到文件。请确保文件包含 '{sheet_name}' sheet。")
        os.unlink(input_file)
        return None
    except Exception as e:
        st.error(f"读取文件时出错：{e}")
        os.unlink(input_file)
        return None
    
    #Fill NaN with empty string
    df_survey = df_survey.fillna('')

    # ======== 【修改 1：建立外部显示区】 ========
    # 在 expander 外面建立一个容器，专门用来显示错误，这样不用点开折叠框也能看到
    error_area = st.container()
    # =========================================

    # 大 expander 包裹所有详细日志
    with st.expander("查看详细日志", expanded=False):
    
        # 新加：动态区域检测函数
        def find_region_start_end(df, target_theme):
            """扫描A列找到主题行，返回 (header_row, end_row) (0-based索引)"""
            theme_row = None
            next_theme_row = None
            for idx, val in enumerate(df.iloc[:, 0]):  # A列 (index 0)
                if pd.notna(val) and target_theme in str(val).strip():
                    theme_row = idx
                    break
            if theme_row is None:
                st.warning(f"错误：未找到主题 '{target_theme}' 在A列")
                return None, None
            
            # 找下一个主题（顺序：详情页 → 旗舰店 → 商品集 → SP）
            next_themes = ["SBV落地页：品牌旗舰店", "SB落地页：商品集", "SBV落地页：商品详情页", "SP-商品推广"]  # 从当前开始找下一个
            for idx in range(theme_row + 1, len(df)):
                val = str(df.iloc[idx, 0]).strip()
                if any(nt in val for nt in next_themes if nt != target_theme):
                    next_theme_row = idx
                    break
            end_row = next_theme_row - 1 if next_theme_row else len(df) - 1  # 到文件末尾
            header_row = theme_row + 1  # header在主题行下一行
            st.write(f"找到 '{target_theme}' 区域: 主题行 {theme_row+1}, header行 {header_row+1}, 数据到行 {end_row+1}")
            return header_row, end_row

        # 先找主题行，用于限全局设置范围（取第一个主题前）
        temp_result = find_region_start_end(df_survey, 'SBV落地页：品牌旗舰店')
        if temp_result[0] is None:
            temp_result = find_region_start_end(df_survey, 'SB落地页：商品集')
        if temp_result[0] is None:
            temp_result = find_region_start_end(df_survey, 'SBV落地页：商品详情页')
        if temp_result[0] is None:
            temp_result = find_region_start_end(df_survey, 'SP-商品推广')
        if temp_result[0] is None:
            st.error("未找到任何支持的主题区域")
            os.unlink(input_file)
            return None
        global_limit = temp_result[0]  # 用 [0] 是 header_row，即主题前

        # Extract global settings: from rows 0-20, column A (0) labels, B (1) values
        global_settings = {}
        for i in range(0, min(20, global_limit)):
            if i >= len(df_survey):
                break
            label = str(df_survey.iloc[i, 0]).strip() if pd.notna(df_survey.iloc[i, 0]) else ''
            value = str(df_survey.iloc[i, 1]).strip() if pd.notna(df_survey.iloc[i, 1]) and len(df_survey.columns) > 1 else ''
            st.write(f"Row {i+1}: label='{label}', value='{value}'")
            
            # Robust matching similar to test SB.py
            if '品牌实体编号' in label or 'ENTITY' in label.upper():
                global_settings['entity_id'] = value
            elif '品牌名称' in label:
                global_settings['brand_name'] = value
            elif '预算类型' in label:
                global_settings['budget_type'] = value if value else '每日'
            elif '创意素材标题' in label:
                global_settings['creative_title'] = value
            elif '落地页 URL' in label:
                global_settings['landing_url'] = value
        
        st.write(f"全局设置: {global_settings}")
        
        # ======== 【修改 1 强力调试版：精准锁定“广告活动名称”列】Start ========
        
        # 1. 定义哪些主题必须依赖全局设置
        strict_themes = ['SB落地页：商品集', 'SBV落地页：品牌旗舰店']
        
        # 2. 预扫描
        has_strict_activities = False
        
        for theme in strict_themes:
            h_row, e_row = find_region_start_end(df_survey, theme)
            
            if h_row is not None and e_row > h_row:
                # A. 找“广告活动名称”这一列的索引
                header_vals = df_survey.iloc[h_row]
                target_col_idx = -1
                for idx, val in enumerate(header_vals):
                    if '广告活动名称' in str(val).strip():
                        target_col_idx = idx
                        break
                
                # B. 检查该列内容
                if target_col_idx != -1:
                    # 提取数据 (从 header行+1 到 结束行)
                    col_data = df_survey.iloc[h_row + 1 : e_row + 1, target_col_idx]
                    
                    # --- 🕵️ 调试信息开始 (帮你找出哪里有隐藏字符) ---
                    # 转换成字符串并去空格
                    clean_series = col_data.astype(str).str.strip()
                    # 找出非空的值 (如果不为空，说明有数据)
                    non_empty_values = clean_series[clean_series != '']
                    
                    if not non_empty_values.empty:
                        has_strict_activities = True
                        st.error(f"⚠️ 在 '{theme}' 发现有效活动！")
                        st.write("👇 程序检测到的内容如下（请检查 Excel 对应行）：")
                        st.write(non_empty_values) # 把检测到的“幽灵数据”打印出来
                        st.write(f"对应行号（Excel显示行号需+2）：{non_empty_values.index}")
                        break 
                    else:
                        # 只有真的全是空的时候，才会显示这个
                        st.write(f"✅ '{theme}' 区域的“广告活动名称”列是空的（检测通过）。")
                    # --- 🕵️ 调试信息结束 ---

                else:
                    st.warning(f"⚠️ 在 '{theme}' 区域没找到“广告活动名称”这一列，跳过检测。")

        # 3. 执行报错
        if has_strict_activities:
            required_globals = ['creative_title', 'landing_url']
            missing_globals = [k for k in required_globals if not global_settings.get(k)]
            
            if missing_globals:
                error_area.error(f"❌ 【严重错误】全局设置缺失！")
                error_area.error(f"原因：上述调试信息显示，'{theme}' 区域里确实有数据。")
                error_area.error(f"请检查打印出来的那些单元格，可能包含空格或遗留字符。")
                # 暂停生成
                os.unlink(input_file)
                return None
        else:
            st.info("ℹ️ 全局设置检查通过：未检测到需要必填信息的活动。")
            
        # ======== 【修改 1 强力调试版：精准锁定“广告活动名称”列】End ========
        
        # Keyword columns: from header row (iloc[0]), but dynamic like test SB.py
        header_row_full = df_survey.iloc[0].tolist()
        keyword_columns = [col for col in header_row_full if isinstance(col, str) and ('精准词' in col or '广泛词' in col or '否' in col)]
        st.write(f"关键词相关列: {keyword_columns}")
        
        # Identify keyword categories like in test SB.py
        keyword_categories = set()
        for col in keyword_columns:
            col_lower = str(col).lower()
            if '/' in col_lower:
                parts = col_lower.split('/')
                if len(parts) > 0 and parts[0]:
                    keyword_categories.add(parts[0].strip())
                if len(parts) > 1 and parts[1]:
                    chinese_part = parts[1].split('-')[0].strip() if '-' in parts[1] else parts[1].strip()
                    keyword_categories.add(chinese_part)
            else:
                for suffix in ['精准词', '广泛词', '精准', '广泛']:
                    if col_lower.endswith(suffix):
                        prefix = col_lower[:-len(suffix)].strip()
                        if prefix:
                            keyword_categories.add(prefix)
                            break
        keyword_categories.update(['suzhu', '宿主', 'host', 'case', '包', '对手', 'tape'])
        st.write(f"识别到的关键词类别: {keyword_categories}")
        
        # Negative keywords extraction: map to specific columns like test SB.py
        # Col indices mapping
        col_indices = {
            'W': df_survey.columns.get_loc('宿主精准-否精准') if '宿主精准-否精准' in df_survey.columns else None,
            'X': df_survey.columns.get_loc('宿主精准-否词组') if '宿主精准-否词组' in df_survey.columns else None,
            'AA': df_survey.columns.get_loc('宿主广泛-否精准') if '宿主广泛-否精准' in df_survey.columns else None,
            'AB': df_survey.columns.get_loc('宿主广泛-否词组') if '宿主广泛-否词组' in df_survey.columns else None,
            'Y': df_survey.columns.get_loc('case精准-否精准') if 'case精准-否精准' in df_survey.columns else None,
            'Z': df_survey.columns.get_loc('case精准-否词组') if 'case精准-否词组' in df_survey.columns else None,
            'AC': df_survey.columns.get_loc('case广泛-否精准') if 'case广泛-否精准' in df_survey.columns else None,
            'AD': df_survey.columns.get_loc('case广泛-否词组') if 'case广泛-否词组' in df_survey.columns else None,
        }
        
        # Col names for logging
        col_names_dict = {
            'W': '宿主精准-否精准',
            'X': '宿主精准-否词组',
            'AA': '宿主广泛-否精准',
            'AB': '宿主广泛-否词组',
            'Y': 'case精准-否精准',
            'Z': 'case精准-否词组',
            'AC': 'case广泛-否精准',
            'AD': 'case广泛-否词组'
        }
        
        # Extract neg_asin and neg_brand from specific columns
        neg_asin = []
        neg_brand = []
        neg_asin_col = None
        neg_brand_col = None
        for col_idx, col_name in enumerate(df_survey.columns):
            if '否定asin' in str(col_name).lower():
                neg_asin_col = col_idx
            elif '否品牌' in str(col_name).lower():
                neg_brand_col = col_idx
        if neg_asin_col is not None:
            neg_asin = [str(x).strip() for x in df_survey.iloc[:, neg_asin_col].dropna() if str(x).strip()]
            neg_asin = list(dict.fromkeys(neg_asin))
        if neg_brand_col is not None:
            neg_brand = [str(int(x)).strip() for x in df_survey.iloc[:, neg_brand_col].dropna() if str(x).strip()]
            neg_brand = list(dict.fromkeys(neg_brand))
        st.write(f"否定ASIN: {neg_asin}")
        st.write(f"否品牌: {neg_brand}")
        
        # Output columns for Brand (SB/SBV) - original 27 columns
        output_columns_brand = [
            '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告编号', 
            '广告活动名称', '广告组名称', '广告名称', '状态', '品牌实体编号', 
            '预算类型', '预算', '商品位置', '竞价', '关键词文本', '匹配类型', '拓展商品投放编号', 
            '落地页 URL', '落地页类型', '品牌名称', '同意翻译', '品牌徽标素材编号', 
            '创意素材标题', '创意素材 ASIN', '视频素材编号', '自定义图片', '落地页 ASIN'
        ]
        
        # Output columns for SP - based on header-B_US (25 columns)
        output_columns_sp = [
            '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告组合编号', '广告编号', '关键词编号', 
            '商品投放 ID', '广告活动名称', '广告组名称', '开始日期', '结束日期', '投放类型', '状态', 
            '每日预算', 'SKU', '广告组默认竞价', '竞价', '关键词文本', '匹配类型', '竞价方案', 
            '广告位', '百分比', '拓展商品投放编号'
        ]
        
        product_brand = '品牌推广'
        product_sp = '商品推广'
        operation = 'Create'
        status = '已启用'
        
        # Separate rows for brand and SP
        brand_rows = []
        sp_rows = []
        
        default_bid = 0.6
        default_sp_budget = 12  # SP default budget from header-B_US

        # [修改 2] 初始化错误日志列表
        validation_errors = []
        
        # 支持的主题列表（原代码）
        targets = ['SBV落地页：品牌旗舰店', 'SB落地页：商品集', 'SBV落地页：商品详情页', 'SP-商品推广']
        
        # 支持的主题列表（添加SP）
        targets = ['SBV落地页：品牌旗舰店', 'SB落地页：商品集', 'SBV落地页：商品详情页', 'SP-商品推广']
        
        for target_theme in targets:
            header_row, end_row = find_region_start_end(df_survey, target_theme)
            if header_row is None:
                st.warning(f"跳过主题 '{target_theme}'：未找到区域")
                continue

            # 读取header行作为列名
            header_data = pd.read_excel(input_file, sheet_name=sheet_name, skiprows=header_row, nrows=1)
            col_names = header_data.iloc[0].tolist()  # 获取列名
            
            # 读取数据行 (从header下一行到end_row)
            activity_df = pd.DataFrame()
            if end_row > header_row:
                activity_df = pd.read_excel(input_file, sheet_name=sheet_name, skiprows=header_row + 1, nrows=end_row - header_row)
                activity_df.columns = col_names  # 设置列名
                st.write(f"活动数据形状 ({target_theme}): {activity_df.shape}")
                st.write(f"活动列名 ({target_theme}): {list(activity_df.columns)}")
            else:
                st.warning(f"无活动数据行 ({target_theme})")
                continue

            # 加填充 NaN
            activity_df = activity_df.fillna('')

            # 用activity_df构建activity_rows列表，根据主题不同提取不同
            activity_rows = []
            if 'SP-商品推广' in target_theme:
                # SP 逻辑：动态查找列名索引，类似SB但调整为SP字段
                for idx, row in activity_df.iterrows():
                    # 动态获取列索引 for SP
                    campaign_col = None
                    cpc_col = None
                    sku_col = None
                    budget_col = None
                    group_bid_col = None  # 新加：声明变量，找“广告组默认竞价”列
                    ad_position_col = None  # 新增：广告位列索引
                    percentage_col = None
                    for col_idx, col_name in enumerate(activity_df.columns):
                        col_str = str(col_name).strip().lower()
                        if '广告活动名称' in col_str:
                            campaign_col = col_idx
                        elif 'cpc' in col_str:
                            cpc_col = col_idx
                        elif 'sku' in col_str:
                            sku_col = col_idx
                        elif '预算' in col_str:
                            budget_col = col_idx
                        elif '广告组默认竞价' in col_str:
                            group_bid_col = col_idx
                        elif '广告位' in col_str:
                            ad_position_col = col_idx
                        elif '百分比' in col_str:
                            percentage_col = col_idx
                    
                    # 提取值
                    campaign_name = str(row.iloc[campaign_col]).strip() if campaign_col is not None else ''
                    cpc = str(row.iloc[cpc_col]).strip() if cpc_col is not None else ''
                    sku = str(row.iloc[sku_col]).strip() if sku_col is not None else ''
                    budget = str(row.iloc[budget_col]).strip() if budget_col is not None else ''
                    group_bid = str(row.iloc[group_bid_col]).strip() if group_bid_col is not None else ''
                    ad_position = str(row.iloc[ad_position_col]).strip() if ad_position_col is not None else ''
                    percentage = str(int(float(row.iloc[percentage_col]))) if percentage_col is not None and pd.notna(row.iloc[percentage_col]) and row.iloc[percentage_col] != '' else ''
                    
                    if campaign_name:
                        activity = {
                            'campaign_name': campaign_name,
                            'cpc': cpc,
                            'sku': sku,
                            'budget': budget,
                            'group_bid': group_bid,
                            'ad_position': ad_position,
                            'percentage': percentage
                        }
                        activity_rows.append(activity)
                        st.write(f"  SP 活动: {campaign_name}, CPC={cpc}, 预算={budget}, 广告位={ad_position}, 百分比={percentage}")
            
            else:
                # Brand 逻辑：动态查找列名索引
                for idx, row in activity_df.iterrows():
                    # 动态获取列索引 for Brand
                    campaign_col = None
                    cpc_col = None
                    budget_col = None # 确保预算列也在
                    asins_cols = [3, 4, 5]
                    video_media_col = None  # 新增：初始化视频媒体列索引
                    custom_image_col = None  # 新增：初始化自定义图片列索引
                    landing_type_col = None
                    for col_idx, col_name in enumerate(activity_df.columns):
                        col_str = str(col_name).strip().lower()
                        if '广告活动名称' in col_str:
                            campaign_col = col_idx
                        elif 'cpc' in col_str:
                            cpc_col = col_idx
                        elif '预算' in col_str:
                            budget_col = col_idx
                        elif '视频媒体' in col_str and '编号' in col_str:  # 新增：匹配“视频媒体编号”列
                            video_media_col = col_idx
                        elif '自定义图片' in col_str:  # 新增：匹配“自定义图片”列
                            custom_image_col = col_idx
                        elif '落地页类型' in col_str: # 【新增】如果列名包含落地页类型，记录它的位置
                            landing_type_col = col_idx
                    
                    # 提取值
                    campaign_name = str(row.iloc[campaign_col]).strip() if campaign_col is not None else ''
                    cpc = str(row.iloc[cpc_col]).strip() if cpc_col is not None else ''
                    
                    # 【新增】提取当前行的落地页类型。如果没填，则根据大区域自动补全
                    row_landing_type = str(row.iloc[landing_type_col]).strip() if landing_type_col is not None else ''

                    asins_list = []
                    for col in asins_cols:  # 用列表asins_cols
                        cell_val = str(row.iloc[col]).strip()
                        if cell_val:
                            asins_list.extend([asin.strip() for asin in cell_val.split(',')])  # split逗号扩展
                    unique_asins = list(dict.fromkeys(asins_list))  # 有序去重（保持D→E→F顺序）
                    asins_str = ', '.join(unique_asins) if unique_asins else ''
                    video_asset = str(row.iloc[video_media_col]).strip() if video_media_col is not None else ''  # 新增：提取视频素材值
                    custom_image = str(row.iloc[custom_image_col]).strip() if custom_image_col is not None else ''  # 新增：提取自定义图片值（覆盖原硬编码custom_image = ''）
                    print(f"  自定义图片: '{custom_image}' (col={custom_image_col})")
                    # 新增：为当前行提取品牌徽标素材编号
                    logo_asset = ''
                    logo_col_idx = None

                    # 优先：按列名查找（最稳健）
                    for col_idx, col_name in enumerate(activity_df.columns):
                        if '品牌徽标素材编号' in str(col_name):
                            logo_col_idx = col_idx
                            break

                    # Fallback 到固定 J 列 (index 9)
                    if logo_col_idx is None:
                        if len(activity_df.columns) > 9:
                            logo_col_idx = 9
                            st.write(f"  未找到‘品牌徽标素材编号’列名，使用固定J列（第10列） (活动: {campaign_name})")
                        else:
                            st.warning(f"  数据列不足10列，无法读取品牌徽标素材编号 (活动: {campaign_name})")

                    # 从当前行读取
                    if logo_col_idx is not None:
                        cell_value = row.iloc[logo_col_idx] if logo_col_idx < len(row) else ''
                        if pd.notna(cell_value):
                            logo_asset = str(cell_value).strip()

                    if campaign_name:
                        activity = {
                            'campaign_name': campaign_name,
                            'cpc': cpc,
                            'asins': asins_str,
                            'budget': str(row.iloc[budget_col]).strip() if budget_col is not None else '12',
                            'video_asset': video_asset,  # 新增：保存视频
                            'custom_image': custom_image,  # 新增：保存自定义图片
                            'logo_asset': logo_asset,
                            'landing_type': row_landing_type # 【新增】保存到活动信息里
                        }
                        activity_rows.append(activity)
                        st.write(f"  Brand 活动: {campaign_name}, CPC={cpc}")

            st.write(f"Found {len(activity_rows)} activity rows ({target_theme}): {[r['campaign_name'] for r in activity_rows]}")
            
            
            # Generate rows for this region
            for activity in activity_rows:
                campaign_name = activity['campaign_name']
                st.write(f"处理活动 ({target_theme}): {campaign_name}")

                # ======== 【第3处插入：开始】 ========
                # 2. 必填项与 ASIN 逻辑检查
                # A. 检查通用必填项
                if not str(activity.get('cpc', '')).strip():
                    validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 'CPC'")
                if not str(activity.get('budget', '')).strip():
                    validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 '预算'")

                # B. 根据类型检查特定字段
                if 'SP-商品推广' in target_theme:
                    if not str(activity.get('sku', '')).strip():
                        validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 'SKU'")
                    if not str(activity.get('group_bid', '')).strip():
                        validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 '广告组默认竞价'")
                else:
                    # Brand 检查
                    # 1. 视频检查：如果是视频广告，必须有视频 ID
                    if '品牌旗舰店' in target_theme or '商品详情页' in target_theme:
                         if not str(activity.get('video_asset', '')).strip():
                            validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 '视频媒体编号'")
                    
                    # 2. Logo 检查：【修改这里】排除 "商品详情页"，只有其他类型才查 Logo
                    if '商品详情页' not in target_theme:
                        if not str(activity.get('logo_asset', '')).strip():
                            validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 '品牌徽标素材编号'")
                    
                    # 3. 落地页类型检查
                    if not str(activity.get('landing_type', '')).strip():
                        validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 '落地页类型'")

                    # 4. 检查创意素材 ASIN (D、E、F 列)
                    # 只要是 旗舰店、详情页、商品集 这三类，ASIN 不能为空
                    check_asin_themes = ['品牌旗舰店', '商品详情页', '商品集']
                    if any(x in target_theme for x in check_asin_themes):
                        # activity['asins'] 是从 D/E/F 列提取并合并的字符串
                        if not str(activity.get('asins', '')).strip():
                            validation_errors.append(f"❌ 活动 [{campaign_name}]: 缺少 '创意素材 ASIN' (请检查 D、E、F 列是否填写)")

                # C. ASIN 定向智能检查
                temp_name_check = str(campaign_name).lower()
                is_asin_check = any(x in temp_name_check for x in ['asin'])
                
                if is_asin_check:
                    asin_found_check = False
                    for col in df_survey.columns:
                        if str(col).strip() == str(campaign_name):
                            # 检查该列是否有值
                            col_idx = df_survey.columns.get_loc(col)
                            vals = [x for x in df_survey.iloc[:, col_idx].dropna() if str(x).strip()]
                            if vals:
                                asin_found_check = True
                            break
                    
                    if not asin_found_check:
                         validation_errors.append(f"❌ 活动 [{campaign_name}]: 是 ASIN 投放，但在表头未找到对应列或列下无数据！")
                # ======== 【第3处插入：结束】 ========
                
                is_asin = False  # 初始化变量，避免 UnboundLocalError
                
                if 'SP-商品推广' in target_theme:
                    # SP-specific generation
                    cpc = float(activity['cpc']) if activity['cpc'] != '' else default_bid
                    budget = float(activity['budget']) if activity['budget'] != '' else default_sp_budget
                    sku = activity.get('sku', 'SKU-1')
                    group_bid = float(activity.get('group_bid', default_bid))
                    
                    campaign_name_normalized = str(campaign_name).lower()
                    
                    # Detect category and match type like test SB.py
                    matched_category = None
                    for cat in keyword_categories:
                        if cat in campaign_name_normalized:
                            matched_category = cat
                            break
                    
                    is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact', 'sp_exact'])
                    is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad', 'sp_broad'])
                    is_asin = any(x in campaign_name_normalized for x in ['asin', 'sp_asin'])  # 覆盖赋值
                    match_type = '精准' if is_exact else '广泛' if is_broad else '精准'  # Default exact/精准
                    
                    # Row1: 广告活动
                    row1 = [product_sp, '广告活动', operation, campaign_name, '', '', '', '', '', campaign_name, '', '', '', '手动', status, 
                            budget, '', '', '', '', '', '动态竞价 - 仅降低', '', '', '']
                    sp_rows.append(row1)
                    
                    # Row2: 广告组
                    row2 = [product_sp, '广告组', operation, campaign_name, campaign_name, '', '', '', '', campaign_name, campaign_name, '', '', '', status, 
                            '', '', group_bid, '', '', '', '', '', '', '']
                    sp_rows.append(row2)
                    
                    # Row3: 商品广告
                    row3 = [product_sp, '商品广告', operation, campaign_name, campaign_name, '', '', '', '', campaign_name, campaign_name, '', '', '', status, 
                            '', sku, '', '', '', '', '', '', '', '']
                    sp_rows.append(row3)
                    
                    if not is_asin:
                        # Keywords: dynamic column selection based on region rules (SP original)
                        keywords = []
                        keyword_col_idx = None
                        col_name = None  # For logging
                        
                        if match_type == '精准':
                            if matched_category in ['suzhu', '宿主', 'host']:
                                col_name = 'suzhu/宿主/host-精准词'
                            elif matched_category in ['case', '包']:
                                col_name = 'case/包-精准词'
                        elif match_type == '广泛':
                            # SP: original rules
                            if matched_category in ['suzhu', '宿主', 'host']:
                                col_name = 'suzhu/宿主/host-广泛词'  # M列
                            elif matched_category in ['case', '包']:
                                col_name = 'case/包-广泛词'  # P列
                        
                        if col_name and keyword_col_idx is None:
                            try:
                                keyword_col_idx = df_survey.columns.get_loc(col_name)
                            except KeyError:
                                st.warning(f"列 '{col_name}' 未找到，fallback到硬编码")
                                # Fallback for SP: original indices
                                if '精准' in match_type and matched_category in ['suzhu', '宿主', 'host']:
                                    keyword_col_idx = 11
                                elif '广泛' in match_type and matched_category in ['suzhu', '宿主', 'host']:
                                    keyword_col_idx = 12  # M
                                elif '精准' in match_type and matched_category in ['case', '包']:
                                    keyword_col_idx = 14
                                elif '广泛' in match_type and matched_category in ['case', '包']:
                                    keyword_col_idx = 15  # P
                        
                        if keyword_col_idx is not None and keyword_col_idx < len(df_survey.columns):
                            col_data = [str(kw).strip() for kw in df_survey.iloc[:, keyword_col_idx].dropna() if str(kw).strip()]
                            keywords = list(dict.fromkeys(col_data))
                            col_name = str(df_survey.columns[keyword_col_idx]) if col_name is None else col_name
                            st.write(f"  匹配的列: {col_name} (idx={keyword_col_idx})")
                            st.write(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
                        else:
                            keywords = []
                            st.warning(f"  无匹配列 for {matched_category} {match_type} in {target_theme}")
                    
                        if keywords:
                            for kw in keywords:
                                row_keyword = [product_sp, '关键词', operation, campaign_name, campaign_name, '', '', '', '', campaign_name, campaign_name, '', '', '', status, 
                                            '', '', '', cpc, kw, match_type, '', '', '', '']
                                sp_rows.append(row_keyword)
                        else:
                            st.warning(f"  无关键词数据，跳过生成关键词层级 (活动: {campaign_name})")
                        
                        # Negative keywords: dynamic like test SB.py, with specific column selection
                        if matched_category:
                            # Select columns based on category and type (SP similar to Brand)
                            selected_cols = []
                            if matched_category in ['suzhu', '宿主', 'host']:
                                if is_exact:
                                    selected_cols = ['W', 'X']
                                elif is_broad:
                                    selected_cols = ['AA', 'AB']
                            elif matched_category in ['case', '包']:
                                if is_exact:
                                    selected_cols = ['Y', 'Z']
                                elif is_broad:
                                    selected_cols = ['AC', 'AD']
                            
                            # Collect data, track sources for duplicates
                            neg_data_sources = {
                                '否定精准匹配': defaultdict(list),  # kw -> [col_keys]
                                '否定词组': defaultdict(list)
                            }
                            for col_key in selected_cols:
                                if col_indices.get(col_key) is not None:
                                    col_idx = col_indices[col_key]
                                    col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
                                    col_data = list(dict.fromkeys(col_data))  # column dedup
                                    m_type = '否定精准匹配' if col_key in ['W', 'AA', 'Y', 'AC'] else '否定词组'
                                    for kw in col_data:
                                        neg_data_sources[m_type][kw].append(col_key)
                            
                            # Check duplicates: kw with multiple sources
                            duplicates_detected = False
                            for m_type, kw_sources in neg_data_sources.items():
                                for kw, sources in kw_sources.items():
                                    if len(sources) > 1:
                                        duplicates_detected = True
                                        source_names = [col_names_dict.get(s, s) for s in sources]
                                        st.error(f"\n=== 检测到重复否定关键词 ===")
                                        st.error(f"活动: {campaign_name}")
                                        st.error(f"类型: {m_type}")
                                        st.error(f"重复关键词: '{kw}'")
                                        st.error(f"来源列: {', '.join(source_names)}")
                                        st.error(f"原因: 该关键词在多个否定列中出现，导致生成重复行。请检查 survey 文件的这些列并清理重复值。")
                                        st.error("暂停生成 header 表。")
                                        os.unlink(input_file)
                                        return None  # Pause generation
                            
                            st.write("\n=== 重复检测完成（无重复）===")
                            
                            # Generate rows: deduped kws
                            for m_type, kw_sources in neg_data_sources.items():
                                kws = list(kw_sources.keys())
                                if kws:
                                    st.write(f"  {m_type} 否定关键词数量: {len(kws)}")
                                for kw in kws:
                                    row_neg = [product_sp, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '', campaign_name, campaign_name, '', '', '', status, 
                                            '', '', '', '', kw, m_type, '', '', '', '']
                                    sp_rows.append(row_neg)
                    
                    # ASIN group: generate 商品定向 and 否定商品定向
                    if is_asin:
                        # 商品定向: exact column match to campaign_name
                        asin_targets = []
                        for col in df_survey.columns:
                            if str(col).strip() == str(campaign_name):
                                col_idx = df_survey.columns.get_loc(col)
                                if col_idx is not None:
                                    asin_targets = [str(asin).strip() for asin in df_survey.iloc[:, col_idx].dropna() if str(asin).strip()]
                                    asin_targets = list(dict.fromkeys(asin_targets))
                                    st.write(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
                                    break
                            
                        if asin_targets:
                            for asin in asin_targets:
                                row_product_target = [product_sp, '商品定向', operation, campaign_name, campaign_name, '', '', '', '', campaign_name, campaign_name, '', '', '', status, 
                                                    '', '', '', cpc, '', '', '', '', '', f'asin="{asin}"']
                                sp_rows.append(row_product_target)
                            
                        # 否定商品定向: from global neg_asin and neg_brand
                        for neg in neg_asin:
                            row_neg_product = [product_sp, '否定商品定向', operation, campaign_name, campaign_name, '', '', '', '', campaign_name, campaign_name, '', '', '', status, 
                                            '', '', '', '', '', '', '', '', '', f'asin="{neg}"']
                            sp_rows.append(row_neg_product)
                        
                        # 条件禁用: 否品牌循环
                        if False:  # 禁用 SP 否品牌生成 (改为 True 恢复)
                            for negb in neg_brand:
                                row_neg_brand = [product_sp, '否定商品定向', operation, campaign_name, campaign_name, '', '', '', '', campaign_name, campaign_name, '', '', '', status, 
                                                '', '', '', '', '', '', '', '', '', f'brand="{negb}"']
                                sp_rows.append(row_neg_brand)
                    
                    # 新增/修复：竞价调整层级（仅SP，为每个活动生成1行，如果条件满足）- 移到if is_asin外
                    row_bid_adjust = None  # 防护：初始化为空，避免UnboundLocalError
                    ad_position = activity.get('ad_position', '').strip()
                    percentage = activity.get('percentage', '').strip()
                    if ad_position and percentage:  # 只有两者都有值才生成
                        st.write(f"  生成竞价调整行 (活动: {campaign_name}, 广告位: {ad_position}, 百分比: {percentage})")
                        row_bid_adjust = [
                            product_sp, '竞价调整', operation,
                            campaign_name, '', '', '', '', '',
                            campaign_name, campaign_name, '', '',
                            '手动', status,
                            '', '', '', '', '', '',
                            '动态竞价 - 仅降低',
                            ad_position, percentage, ''
                        ]
                        sp_rows.append(row_bid_adjust)
                    else:
                        st.write(f"  跳过竞价调整行 (活动: {campaign_name})：广告位或百分比为空")
                
                else:
                    # Original Brand (SB/SBV) generation logic - with regional keyword rules
                    cpc = float(activity['cpc']) if activity['cpc'] != '' else default_bid
                    brand_budget = float(activity['budget']) if activity['budget'] != '' else 12
                    asins_str = activity.get('asins', '')
                    video_asset = activity.get('video_asset', '')  # 新增：从 activity 获取
                    custom_image = activity.get('custom_image', '')  # 新增：从 activity 获取
                    landing_url = global_settings.get('landing_url', '')
                    landing_type = activity.get('landing_type', '')
                    brand_name = global_settings.get('brand_name', '')
                    creative_title = global_settings.get('creative_title', '')
                    
                    # 直接从 activity 字典中获取之前保存好的 logo_asset
                    logo_asset = activity.get('logo_asset', '')
                    
                    campaign_name_normalized = str(campaign_name).lower()
                    
                    # Detect category and match type
                    matched_category = None
                    for cat in keyword_categories:
                        if cat in campaign_name_normalized:
                            matched_category = cat
                            break
                    
                    is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
                    is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
                    is_asin = any(x in campaign_name_normalized for x in ['asin'])  # 覆盖赋值
                    match_type = '精准' if is_exact else '广泛' if is_broad else '精准'
                    
                    # Row1: 广告活动
                    row1 = [product_brand, '广告活动', operation, campaign_name, '', '', campaign_name, '', '', status, 
                            global_settings.get('entity_id', ''), global_settings.get('budget_type', '每日'), brand_budget, '在亚马逊上出售', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
                    brand_rows.append(row1)
                    
                    # Row2: 广告组
                    row2 = [product_brand, '广告组', operation, campaign_name, campaign_name, '', campaign_name, campaign_name, '', status, 
                            '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
                    brand_rows.append(row2)
                    
                    # Row3: 广告实体层级（品牌视频广告 / 商品集广告 / 视频广告） - 按主题分开处理，避免共用逻辑
                    if 'SBV落地页：品牌旗舰店' in target_theme:
                        # 品牌旗舰店视频广告
                        row3 = [product_brand, '品牌视频广告', operation,
                                campaign_name, campaign_name, campaign_name, '', '', campaign_name, status,
                                '', '', '', '', '', '', '', '',
                                landing_url, landing_type, brand_name, 'False', logo_asset, creative_title,
                                asins_str, video_asset, custom_image, '']
                        brand_rows.append(row3)

                    elif 'SBV落地页：商品详情页' in target_theme:
                        row3 = [
                            product_brand, '视频广告', operation,
                            campaign_name, campaign_name, campaign_name, '', '', campaign_name, status,
                            '', '', '', '', '', '', '', '',
                            '', landing_type or '', '', 'False',
                            '', '',
                            asins_str, video_asset, '', ''
                        ]
                        brand_rows.append(row3)

                    elif 'SB落地页：商品集' in target_theme:
                        # 1. 默认设置（常规规则） 
                        final_landing_url = landing_url
                        final_creative_asin = asins_str
                        final_landing_asin = ''

                        # 2. 应用你的新规则：如果是“商品列表” 
                        if landing_type == '商品列表':
                            final_landing_url = ''      # 落地页 URL 为空
                            final_creative_asin = ''   # 创意素材 ASIN 为空
                            final_landing_asin = asins_str  # 将原来的 ASIN 填到“落地页 ASIN”列
                        
                        # 3. 生成第 28 列数据行 
                        row3 = [
                            product_brand, '商品集广告', operation,
                            campaign_name, campaign_name, campaign_name, '', '', campaign_name, status,
                            '', '', '', '', '', '', '', '',
                            final_landing_url,     # 对应第19列：落地页 URL
                            landing_type,          # 对应第20列：落地页类型
                            brand_name,            # 对应第21列：品牌名称
                            'False',               # 对应第22列：同意翻译
                            logo_asset,            # 对应第23列：品牌徽标素材编号
                            creative_title,        # 对应第24列：创意素材标题
                            final_creative_asin,   # 对应第25列：创意素材 ASIN
                            video_asset,           # 对应[cite: 5, 6]：视频素材编号
                            custom_image,          # 对应[cite: 7]：自定义图片
                            final_landing_asin     # 对应第28列：落地页 ASIN (新规则核心)
                        ]
                        brand_rows.append(row3)
                    
                    else:
                        st.warning(f"未识别的 Brand 主题：{target_theme}，跳过生成广告实体行")
                    
                    # Keywords: dynamic column selection based on regional rules (SB/SBV)
                    if not is_asin:
                        keywords = []
                        keyword_col_idx = None
                        col_name = None  # For logging
                        
                        if match_type == '精准':
                            # All regions: original precise rules
                            if matched_category in ['suzhu', '宿主', 'host']:
                                col_name = 'suzhu/宿主/host-精准词'  # L列，无空格
                            elif matched_category in ['case', '包']:
                                col_name = 'case/包-精准词'  # O列
                        elif match_type == '广泛':
                            # SB/SBV: regional rules - suzhu → N, case → Q
                            if matched_category in ['suzhu', '宿主', 'host']:
                                col_name = 'suzhu/宿主/host-广泛词带加号'  # N列，无空格
                            elif matched_category in ['case', '包']:
                                col_name = 'case/包-广泛词带加号'  # Q列
                        
                        if col_name and keyword_col_idx is None:  # Only if not already set
                            try:
                                keyword_col_idx = df_survey.columns.get_loc(col_name)
                            except KeyError:
                                st.warning(f"列 '{col_name}' 未找到，fallback到硬编码")
                                # Fallback: regional indices for SB/SBV
                                if '精准' in match_type and matched_category in ['suzhu', '宿主', 'host']:
                                    keyword_col_idx = 11  # L
                                elif '精准' in match_type and matched_category in ['case', '包']:
                                    keyword_col_idx = 14  # O
                                elif '广泛' in match_type and matched_category in ['suzhu', '宿主', 'host']:
                                    keyword_col_idx = 13  # N
                                elif '广泛' in match_type and matched_category in ['case', '包']:
                                    keyword_col_idx = 16  # Q
                        
                        if keyword_col_idx is not None and keyword_col_idx < len(df_survey.columns):
                            col_data = [str(kw).strip() for kw in df_survey.iloc[:, keyword_col_idx].dropna() if str(kw).strip()]
                            keywords = list(dict.fromkeys(col_data))
                            col_name = str(df_survey.columns[keyword_col_idx]) if col_name is None else col_name
                            st.write(f"  匹配的列: {col_name} (idx={keyword_col_idx})")
                            st.write(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
                        else:
                            keywords = []
                            st.warning(f"  无匹配列 for {matched_category} {match_type} in {target_theme}")
                
                        if keywords:
                            for kw in keywords:
                                row_keyword = [product_brand, '关键词', operation, campaign_name, campaign_name, '', '', '', '', status, 
                                            '', '', '', '', cpc, kw, match_type, '', '', '', '', '', '', '', '', '', '', '']
                                brand_rows.append(row_keyword)
                        else:
                            st.warning(f"  无关键词数据，跳过生成关键词层级 (活动: {campaign_name})")
                        
                        # Negative keywords: dynamic like test SB.py, with specific column selection
                        if matched_category:
                            # Select columns based on category and type
                            selected_cols = []
                            if matched_category in ['suzhu', '宿主', 'host']:
                                if is_exact:
                                    selected_cols = ['W', 'X']
                                elif is_broad:
                                    selected_cols = ['AA', 'AB']
                            elif matched_category in ['case', '包']:
                                if is_exact:
                                    selected_cols = ['Y', 'Z']
                                elif is_broad:
                                    selected_cols = ['AC', 'AD']
                            
                            # Collect data, track sources for duplicates
                            neg_data_sources = {
                                '否定精准匹配': defaultdict(list),  # kw -> [col_keys]
                                '否定词组': defaultdict(list)
                            }
                            for col_key in selected_cols:
                                if col_indices.get(col_key) is not None:
                                    col_idx = col_indices[col_key]
                                    col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
                                    col_data = list(dict.fromkeys(col_data))  # column dedup
                                    m_type = '否定精准匹配' if col_key in ['W', 'AA', 'Y', 'AC'] else '否定词组'
                                    for kw in col_data:
                                        neg_data_sources[m_type][kw].append(col_key)
                            
                            # Check duplicates: kw with multiple sources
                            duplicates_detected = False
                            for m_type, kw_sources in neg_data_sources.items():
                                for kw, sources in kw_sources.items():
                                    if len(sources) > 1:
                                        duplicates_detected = True
                                        source_names = [col_names_dict.get(s, s) for s in sources]
                                        st.error(f"\n=== 检测到重复否定关键词 ===")
                                        st.error(f"活动: {campaign_name}")
                                        st.error(f"类型: {m_type}")
                                        st.error(f"重复关键词: '{kw}'")
                                        st.error(f"来源列: {', '.join(source_names)}")
                                        st.error(f"原因: 该关键词在多个否定列中出现，导致生成重复行。请检查 survey 文件的这些列并清理重复值。")
                                        st.error("暂停生成 header 表。")
                                        os.unlink(input_file)
                                        return None  # Pause generation
                            
                            st.write("\n=== 重复检测完成（无重复）===")
                            
                            # Generate rows: deduped kws
                            for m_type, kw_sources in neg_data_sources.items():
                                kws = list(kw_sources.keys())
                                if kws:
                                    st.write(f"  {m_type} 否定关键词数量: {len(kws)}")
                                for kw in kws:
                                    row_neg = [product_brand, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '', status, 
                                            '', '', '', '', '', kw, m_type, '', '', '', '', '', '', '', '', '', '', '']
                                    brand_rows.append(row_neg)
                    
                    # ASIN group: generate 商品定向 and 否定商品定向
                    if is_asin:
                        # 商品定向: exact column match to campaign_name
                        asin_targets = []
                        for col in df_survey.columns:
                            if str(col).strip() == str(campaign_name):
                                col_idx = df_survey.columns.get_loc(col)
                                if col_idx is not None:
                                    asin_targets = [str(asin).strip() for asin in df_survey.iloc[:, col_idx].dropna() if str(asin).strip()]
                                    asin_targets = list(dict.fromkeys(asin_targets))
                                    st.write(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
                                    break
                        
                        if asin_targets:
                            for asin in asin_targets:
                                row_product_target = [product_brand, '商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                                    '', '', '', '', cpc, '', '', f'asin="{asin}"', '', '', '', '', '', '', '', '', '', '']
                                brand_rows.append(row_product_target)
                        
                        # 否定商品定向: from global neg_asin and neg_brand
                        for neg in neg_asin:
                            row_neg_product = [product_brand, '否定商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                            '', '', '', '', '', '', '', f'asin="{neg}"', '', '', '', '', '', '', '', '', '', '']
                            brand_rows.append(row_neg_product)
                        
                        for negb in neg_brand:
                            row_neg_brand = [product_brand, '否定商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                            '', '', '', '', '', '', '', f'brand="{negb}"', '', '', '', '', '', '', '', '', '', '']
                            brand_rows.append(row_neg_brand)
        
        # ======== 【修改 4 更新版】最终错误拦截 ========
        if validation_errors:
            # 使用 with error_area 确保这一堆错误都显示在外面
            with error_area:
                st.error("🚫 检测到 Excel 模版填写不完整，已停止生成！请修复以下问题：")
                for err in validation_errors:
                    st.error(err) # 这里每一条错误都会列在最显眼的地方
            
            os.unlink(input_file)
            return None
        # =============================================
        
        # Create DFs
        df_brand = pd.DataFrame(brand_rows, columns=output_columns_brand) if brand_rows else pd.DataFrame(columns=output_columns_brand)
        df_sp = pd.DataFrame(sp_rows, columns=output_columns_sp) if sp_rows else pd.DataFrame(columns=output_columns_sp)
        df_brand = df_brand.fillna('')
        df_sp = df_sp.fillna('')
        
        # Save to BytesIO for download - Multi-sheet
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            if not df_brand.empty:
                df_brand.to_excel(writer, index=False, sheet_name='品牌广告')
            if not df_sp.empty:
                df_sp.to_excel(writer, index=False, sheet_name='SP-商品推广')
        output_buffer.seek(0)
        
    st.success(f"生成完成！品牌行数：{len(brand_rows)}, SP行数：{len(sp_rows)}")
        
    # Cleanup temp file
    os.unlink(input_file)
        
    return output_buffer

# Generate Button
if uploaded_file is not None:
    if st.button("生成 Header 文件"):
        with st.spinner("正在处理文件..."):
            output_buffer = generate_header_for_sbv_brand_store(uploaded_file.read())
            if output_buffer is not None:
                # Generate filename with current time (precise to minute)
                now = datetime.now()
                timestamp = now.strftime("%Y-%m-%d %H:%M")
                filename = f"header-{timestamp}.xlsx"
                
                st.download_button(
                    label="下载生成的 Header 文件",
                    data=output_buffer.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("请上传 Excel 文件以开始生成。")