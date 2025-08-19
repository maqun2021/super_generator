

import pandas as pd
from docxtpl import DocxTemplate
from datetime import date
import streamlit as st
import io
import zipfile
import re
import calendar
from datetime import datetime

st.set_page_config(page_title="Enhanced KOC Contract Generator", layout="wide", page_icon="📝")

st.markdown("""
<style>
    .main {
        background-color: #f0f6fb;
        font-family: 'Segoe UI', sans-serif;
    }
    .title {
        font-size: 2.2em;
        font-weight: bold;
        color: #2563eb;
    }
    .subtitle {
        font-size: 1.1em;
        color: #3b82f6;
    }
    .stFileUploader > label {
        font-weight: 600;
        color: #2563eb;
    }
    .stButton>button {
        background-color: #2563eb;
        color: white;
        border-radius: 6px;
        border: none;
        padding: 0.5em 1.5em;
        font-size: 1.1em;
        font-weight: bold;
        margin-top: 1em;
    }
    .stButton>button:hover {
        background-color: #1d4ed8;
        color: #fff;
    }
    .summary-box {
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
    }
    .summary-title {
        font-weight: bold;
        color: #2563eb;
        margin-bottom: 10px;
    }
    .pair-info {
        background-color: #e8f4fd;
        border-left: 4px solid #2563eb;
        padding: 10px;
        margin: 10px 0;
        border-radius: 4px;
    }
    .form-section {
        background-color: #ffffff;
        border: 1px solid #e1e5e9;
        border-radius: 8px;
        padding: 20px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="title">📄 Enhanced KOC Contract Generator</div>', unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Generate Contracts + Summaries in Perfect Pairs 🚀</div>", unsafe_allow_html=True)

st.markdown("---")

# 定义函数
def validate_form_data(data):
    """验证表单数据"""
    errors = []
    
    # 必填字段验证
    if not data.get('party_b_name', '').strip():
        errors.append("乙方姓名为必填项")
    
    if not data.get('email', '').strip():
        errors.append("邮箱地址为必填项")
    elif not re.match(r"[^@]+@[^@]+\.[^@]+", data.get('email', '')):
        errors.append("邮箱格式不正确")
    
    if not data.get('platforms'):
        errors.append("请至少选择一个发布平台")
    
    if not data.get('video_rate') or data.get('video_rate') <= 0:
        errors.append("视频金额必须大于0")
    
    if not data.get('estimated_videos', '').strip():
        errors.append("预计视频数量为必填项")
    
    if not data.get('start_date'):
        errors.append("开始日期为必填项")
    
    if not data.get('payment_method'):
        errors.append("支付方式为必填项")
    
    # 日期逻辑验证
    if data.get('start_date') and data.get('end_date'):
        if data['start_date'] > data['end_date']:
            errors.append("开始日期不能晚于结束日期")
    
    # 实际上线视频数量验证（当状态为已履行完毕时）
    if data.get('statement') == "已履行完毕":
        if not data.get('actual_video_number', '').strip():
            errors.append("已履行完毕状态下，实际上线视频数量为必填项")
    
    return errors

def create_form_record(form_data):
    """从表单数据创建记录"""
    # 平台映射
    platform_map = {
        'TikTok': 'TT',
        'Instagram': 'IG', 
        'YouTube': 'YT',
        'Facebook': 'FB',
        'Kwai': 'kwai'
    }
    
    # 获取主平台昵称
    main_platform_nickname = ""
    if form_data.get('main_platform') and form_data.get('platform_usernames'):
        main_platform_nickname = form_data['platform_usernames'].get(form_data['main_platform'], '')
    
    # 创建记录
    record = {
        'Party B Name': form_data['party_b_name'],
        'Email': form_data['email'],
        'Contact': form_data['contact'] or '',
        'Address': form_data['address'] or '',
        'Video Rate': str(form_data['video_rate']),
        'Estimated Videos': form_data['estimated_videos'],
        'Start date': form_data['start_date'].strftime('%Y-%m-%d'),
        'end date': form_data['end_date'].strftime('%Y-%m-%d') if form_data['end_date'] else '',
        'Payment method': form_data['payment_method'],
        'Bonus': form_data['bonus_level'],
        'Main Platform nickname': main_platform_nickname,
        'Statement': form_data['statement'],
        'No. of Posted Videos': form_data['actual_video_number'] or '',
        'Payment Info': form_data['payment_info'] or ''
    }
    
    # 添加平台字段
    for platform in form_data['platforms']:
        if platform in platform_map:
            record[platform_map[platform]] = form_data['platform_usernames'].get(platform, '')
    
    # 添加平台显示信息
    platform_display = ' ＆ '.join(form_data['platforms'])
    record['platform_display'] = platform_display
    
    return record

# 输入方式选择
st.markdown("### 📋 选择输入方式")
input_mode = st.radio(
    "请选择数据输入方式：",
    ["📝 表单填写（推荐）", "📄 CSV文件上传"],
    index=0
)

# 表单填写界面
if input_mode == "📝 表单填写（推荐）":
    st.markdown("### 📝 表单填写模式")
    
    # 初始化session state
    if 'form_records' not in st.session_state:
        st.session_state.form_records = []
    
    # 表单填写区域
    with st.expander("📋 填写KOC信息", expanded=True):
        st.markdown('<div class="form-section">', unsafe_allow_html=True)
        
        # 基本信息
        st.markdown("#### 👤 基本信息")
        col1, col2 = st.columns(2)
        with col1:
            party_b_name = st.text_input("乙方姓名 *", key="name_input", placeholder="请输入KOC姓名")
            email = st.text_input("邮箱地址 *", key="email_input", placeholder="example@email.com")
        with col2:
            contact = st.text_input("联系电话", key="contact_input", placeholder="+1-234-567-8900")
            address = st.text_area("地址", key="address_input", placeholder="详细地址信息", height=80)
        
        # 平台信息
        st.markdown("#### 📱 平台信息")
        platforms = st.multiselect(
            "选择发布平台 *",
            ["TikTok", "Instagram", "YouTube", "Facebook", "Kwai"],
            default=["TikTok"],
            key="platforms_select"
        )
        
        # 动态生成平台用户名输入框
        platform_usernames = {}
        if platforms:
            cols = st.columns(len(platforms))
            for i, platform in enumerate(platforms):
                with cols[i]:
                    platform_usernames[platform] = st.text_input(
                        f"{platform} 用户名",
                        key=f"{platform}_username",
                        placeholder=f"请输入{platform}用户名"
                    )
            
            # 主平台选择
            st.markdown("#### 🎯 主平台选择")
            main_platform = st.selectbox(
                "选择主平台（用于显示昵称）",
                platforms,
                key="main_platform_select",
                help="选择哪个平台作为主要展示平台，其昵称将作为主平台昵称"
            )
        
        # 合作详情
        st.markdown("#### 📋 合作详情")
        col1, col2, col3 = st.columns(3)
        with col1:
            video_rate = st.number_input("单支视频金额 ($) *", min_value=0.0, value=10.0, key="video_rate_input")
            estimated_videos = st.text_input("预计视频数量 *", key="estimated_videos_input", placeholder="如：10-15")
        with col2:
            start_date = st.date_input("开始日期 *", value=date.today(), key="start_date_input")
            end_date = st.date_input("结束日期", value=None, key="end_date_input")
        with col3:
            payment_method = st.selectbox("支付方式 *", ["bank", "paypal"], key="payment_method_select")
            bonus_level = st.selectbox("奖励等级", ["none", "lower", "higher"], key="bonus_level_select")
        
        # 其他信息
        st.markdown("#### 📄 其他信息")
        col1, col2 = st.columns(2)
        with col1:
            statement = st.selectbox("履行状态", ["正在履行/未开始履行", "已履行完毕"], key="statement_select")
            payment_info = st.text_area("支付信息", key="payment_info_input", placeholder="银行账户信息等", height=80)
        with col2:
            # 根据履行状态显示实际上线视频数量
            if statement == "已履行完毕":
                actual_video_number = st.text_input("实际上线视频数量 *", key="actual_videos_input", placeholder="如：7")
            else:
                actual_video_number = ""
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 表单操作按钮
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            if st.button("➕ 添加记录", key="add_record_btn"):
                # 验证必填字段
                errors = validate_form_data({
                    'party_b_name': party_b_name,
                    'email': email,
                    'platforms': platforms,
                    'video_rate': video_rate,
                    'estimated_videos': estimated_videos,
                    'start_date': start_date,
                    'payment_method': payment_method,
                    'statement': statement,
                    'actual_video_number': actual_video_number
                })
                
                if errors:
                    for error in errors:
                        st.error(f"❌ {error}")
                else:
                    # 创建记录
                    record = create_form_record({
                        'party_b_name': party_b_name,
                        'email': email,
                        'contact': contact,
                        'address': address,
                        'platforms': platforms,
                        'platform_usernames': platform_usernames,
                        'main_platform': main_platform,
                        'video_rate': video_rate,
                        'estimated_videos': estimated_videos,
                        'start_date': start_date,
                        'end_date': end_date,
                        'payment_method': payment_method,
                        'bonus_level': bonus_level,
                        'statement': statement,
                        'actual_video_number': actual_video_number,
                        'payment_info': payment_info
                    })
                    
                    st.session_state.form_records.append(record)
                    st.success(f"✅ 已添加记录：{party_b_name}")
                    st.rerun()
        
        with col2:
            if st.button("🗑️ 清空表单", key="clear_form_btn"):
                st.rerun()
    
    # 显示已添加的记录
    if st.session_state.form_records:
        st.markdown("### 📋 已添加的记录")
        for i, record in enumerate(st.session_state.form_records):
            with st.expander(f"📄 {record['Party B Name']} - {record.get('Main Platform nickname', '')}", expanded=False):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**邮箱**: {record['Email']}")
                    st.write(f"**平台**: {record.get('platform_display', '')}")
                    st.write(f"**视频金额**: ${record['Video Rate']}")
                    st.write(f"**预计视频**: {record['Estimated Videos']}")
                with col2:
                    if st.button(f"❌ 删除", key=f"delete_{i}"):
                        st.session_state.form_records.pop(i)
                        st.rerun()
        
        st.success(f"📊 当前共有 {len(st.session_state.form_records)} 条记录")

# 模板上传（两种模式都需要）
st.markdown("### 📄 上传Word模板")
uploaded_template = st.file_uploader("📄 Upload Word Template (.docx)", type=["docx"])

# CSV上传界面（原有功能）
if input_mode == "📄 CSV文件上传":
    st.markdown("### 📄 CSV文件上传模式")
    uploaded_csv = st.file_uploader("📑 Upload CSV File", type=["csv"])

# 生成选项
st.markdown("### 生成选项")
generate_mode = st.radio(
    "请选择生成模式：",
    ["只批量生成合同（仅合同文件）", "合同及配对的基本内容（合同+概括配对输出）"],
    index=1
)

# 生成按钮
generate = st.button("🚀 Generate")

# 生成选项逻辑
if generate_mode == "只批量生成合同（仅合同文件）":
    generate_contracts = True
    generate_summaries = False
    output_mode = "配对输出"
else:
    generate_contracts = True
    generate_summaries = True
    output_mode = "配对输出"

def generate_contract_summary(row):
    """生成合同基本内容概括 - 使用中文字段"""
    try:
        kol_name = str(row.get('Party B Name', '')).strip()
        nickname = str(row.get('Main Platform nickname', '')).strip()
        video_rate = str(row.get('Video Rate', '')).strip()
        video_number = str(row.get('Estimated Videos', '')).strip()
        statement = str(row.get('Statement', '')).strip()
        actual_video_number = str(row.get('No. of Posted Videos', '')).strip()
        bonus = str(row.get('Bonus', '')).strip().lower()

        # 获取平台信息 - 从各个平台字段组合
        platform_fields = infer_platform_fields(row)
        platforms = platform_fields['platform']

        # 获取日期信息 - 使用infer_chinese_date_versions函数（用于中文摘要）
        date_fields = infer_chinese_date_versions(row)
        promotion_date = date_fields['promotion_date']

        # 获取支付信息
        payment_method = str(row.get('Payment method', '')).strip()
        if payment_method.lower() == 'bank':
            payment_text = "银行转账，手续费共同承担，视频上线后 Net 30 days/video。"
        elif payment_method.lower() == 'paypal':
            payment_text = "PayPal转账，手续费对方承担，预先扣除2%，视频上线后 Net 30 days/video。"
        else:
            payment_text = "视频上线后 Net 30 days/video。"

        try:
            video_rate_clean = "{:.0f}".format(float(video_rate)) if video_rate else "0"
        except:
            video_rate_clean = video_rate

        # 奖励机制：根据bonus字段判断
        bonus_text = ""
        if bonus in ['lower', 'higher']:
            bonus_text = "\n3. 有奖励机制，合同中有根据播放量制定的详细奖励机制，具体参见合同。"

        koc_display_name = nickname if nickname else kol_name

        # 根据statement判断履行状态，选择不同的时间表述
        if statement == "已履行完毕":
            line2 = f"2. 单支视频金额${video_rate_clean}，签约 {video_number} 期视频，实际上线视频数量为 {actual_video_number} 支，视频上线时间为 {promotion_date}。"
        else:
            line2 = f"2. 单支视频金额${video_rate_clean}，签约 {video_number} 期视频，视频预计上线时间为 {promotion_date}。"
        
        summary = f'''合作事项：\n1. 海外KOC（{koc_display_name}），发布平台{platforms}。\n{line2}{bonus_text}\n\n权利义务：(重点highlight)\n1. 未经甲方同意，乙方不得删除视频，内容永久保留，否则支付甲方50%的费用。\n2. 乙方发布未经批准/错误版本视频，甲方可以选择补偿方式（删除重发、另行协商补偿、终止合作拒绝付款）。\n\n付款条件：\n{payment_text}'''
        return summary
    except Exception as e:
        return f"生成概括时出错: {str(e)}"

def infer_platform_fields(row):
    platform_map = {
        'TT': 'TikTok',
        'IG': 'Instagram',
        'YT': 'YouTube',
        'FB': 'Facebook',
        'kwai': 'Kwai'
    }
    username_map = {
        'TT': 'Tiktok Video',
        'IG': 'Instagram Reels',
        'YT': 'YouTube Shorts',
        'FB': 'Facebook Reels',
        'kwai': 'Kwai Video'
    }
    link_map = {
        'TT': 'https://www.tiktok.com/@{}',
        'IG': 'https://www.instagram.com/{}',
        'YT': 'https://www.youtube.com/@{}',
        'FB': 'https://www.facebook.com/{}',
        'kwai': 'https://www.kwai.com/user/{}'
    }
    platforms = []
    usernames = []
    links = []
    for key in platform_map:
        uname = str(row.get(key, '')).strip()
        if uname:
            platforms.append(platform_map[key])
            usernames.append(f"{username_map[key]} - {uname}")
            links.append(link_map[key].format(uname))
    return {
        'platform': ' ＆ '.join(platforms),
        'platform_username': '\n'.join(usernames),
        'Influencer_links': '\n'.join(links)
    }

def infer_date_versions(row):
    start = row.get('Start date', '')
    end = row.get('end date', '')
    try:
        start_dt = datetime.strptime(start, '%Y-%m-%d')
        if end and str(end).strip():  # 检查end date是否为空
            end_dt = datetime.strptime(end, '%Y-%m-%d')
            if start_dt.year == end_dt.year and start_dt.month == end_dt.month:
                english = f"{calendar.month_name[start_dt.month]} {start_dt.year}"
            else:
                english = f"{calendar.month_name[start_dt.month]} {start_dt.year} - {calendar.month_name[end_dt.month]} {end_dt.year}"
        else:
            # 如果end date为空，只显示开始日期
            english = f"{calendar.month_name[start_dt.month]} {start_dt.year}"
    except Exception as e:
        print(f"Date parsing error: {e}, start: {start}, end: {end}")
        english = ""
    return {
        'promotion_date': english
    }

def infer_chinese_date_versions(row):
    start = row.get('Start date', '')
    end = row.get('end date', '')
    try:
        start_dt = datetime.strptime(start, '%Y-%m-%d')
        if end and str(end).strip():  # 检查end date是否为空
            end_dt = datetime.strptime(end, '%Y-%m-%d')
            if start_dt.year == end_dt.year and start_dt.month == end_dt.month:
                chinese = f"{start_dt.year}年{start_dt.month:02d}月"
            else:
                chinese = f"{start_dt.year}年{start_dt.month:02d}月 - {end_dt.year}年{end_dt.month:02d}月"
        else:
            # 如果end date为空，只显示开始日期
            chinese = f"{start_dt.year}年{start_dt.month:02d}月"
    except Exception as e:
        print(f"Date parsing error: {e}, start: {start}, end: {end}")
        chinese = ""
    return {
        'promotion_date': chinese
    }

def infer_bonus_info(row):
    bonus = str(row.get('Bonus', 'none')).strip().lower()
    # 处理不同的bonus值格式
    if bonus == 'lower':
        return {
            'bonus_info': """*The bonuses will be paid with basic video production fee.\n\nBonus Payment Policy:\n\n*One video will only get the bonus once in the limited time. Pay the bonus with the highest amount.\n\nBonus per video reaches 100k views in 3 days from the date of posting, USD[15.00].\nBonus per video reaches 200k views in 3 days from the date of posting, USD[20.00].\nBonus per video reaches 300k views in 3 days from the date of posting, USD[30.00].\nBonus per video reaches 500k views in 3 days from the date of posting, USD[45.00].\nBonus per video reaches 1M views in 3 days from the date of posting, USD[65.00].\n\nTotal Budget up to $1500."""
        }
    elif bonus == 'higher':
        return {
            'bonus_info': """*The bonuses will be paid with basic video production fee.\n\nBonus Payment Policy:\n\n*One video will only get the bonus once in the limited time. Pay the bonus with the highest amount.\n\nBonus per video reaches 100k views in 3 days from the date of posting, USD[30.00].\nBonus per video reaches 200k views in 3 days from the date of posting, USD[40.00].\nBonus per video reaches 300k views in 3 days from the date of posting, USD[60.00].\nBonus per video reaches 500k views in 3 days from the date of posting, USD[90.00].\nBonus per video reaches 1M views in 3 days from the date of posting, USD[110.00].\n\nTotal Budget up to $1500."""
        }
    else:
        return {'bonus_info': ''}

def infer_payment_fields(row):
    method = str(row.get('Payment method', '')).strip().lower()
    if method == 'bank':
        return {
            'payment_charges': "Payment charges shall be borne by each party independently (SHA)."
        }
    elif method == 'paypal':
        return {
            'payment_charges': "Party A processing fee with 2% of total payments shall be deducted in advance."
        }
    else:
        return {
            'payment_charges': ""
        }

def fix_multiline_csv(df):
    """修复包含多行字段的CSV数据"""
    clean_data = []
    current_record = None
    skip_count = 0  # 跳过前两行样例数据
    
    for index, row in df.iterrows():
        row_data = [str(cell) if pd.notna(cell) else '' for cell in row]
        party_b_name = row_data[0].strip()
        
        if party_b_name and party_b_name.lower() != 'nan' and party_b_name != '':
            # 跳过前两行样例数据
            if skip_count < 2:
                skip_count += 1
                continue
                
            # 保存上一条完整记录
            if current_record is not None:
                clean_data.append(current_record)
            # 开始新记录
            current_record = row_data.copy()
        else:
            # 这是续行，合并到当前记录
            if current_record is not None:
                for i, cell in enumerate(row_data):
                    if cell.strip() and cell.lower() != 'nan':
                        if i < len(current_record):
                            if current_record[i].strip():
                                current_record[i] += '\n' + cell
                            else:
                                current_record[i] = cell
    
    # 添加最后一条记录
    if current_record is not None:
        clean_data.append(current_record)
    
    # 创建新的DataFrame
    clean_df = pd.DataFrame(clean_data, columns=df.columns)
    clean_df.columns = clean_df.columns.str.strip()
    st.success(f"✅ CSV修复完成：跳过2行样例数据，找到 {len(clean_df)} 条有效记录")
    return clean_df

def process_data(df, uploaded_template, generate_contracts, generate_summaries, output_mode):
    today = date.today().isoformat()
    zip_buffer = io.BytesIO()
    summaries = []
    contract_files = []
    
    # 修复多行CSV数据
    df = fix_multiline_csv(df)
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            name_value = row['Party B Name'] if 'Party B Name' in row else ''
            if str(name_value).strip() == "":
                continue
            try:
                safe_name = str(row['Party B Name']).replace(" ", "_").replace("/", "-")
                if generate_contracts:
                    template = DocxTemplate(uploaded_template)
                    video_rate_value = row['Video Rate'] if 'Video Rate' in row else ''
                    is_video_rate_empty = False
                    try:
                        is_video_rate_empty = bool(pd.isna(video_rate_value)) or str(video_rate_value).strip() == ""
                    except Exception:
                        is_video_rate_empty = str(video_rate_value).strip() == ""
                    # 自动推断内容
                    platform_fields = infer_platform_fields(row)
                    date_fields = infer_date_versions(row)
                    bonus_info = infer_bonus_info(row)
                    payment_fields = infer_payment_fields(row)
                    context = {
                        'Influencer_name': row['Party B Name'],
                        'Influencer_email': row['Email'],
                        'Influencer_contact': "N/A" if pd.isna(row['Contact']) or str(row['Contact']).strip() == "" else row['Contact'],
                        'Influencer_address': row['Address'],
                        'platform': platform_fields['platform'],
                        'platform_username': platform_fields['platform_username'],
                        'Influencer_links': platform_fields['Influencer_links'],
                        'promotion_date': date_fields['promotion_date'],
                        'video_rate': "{:.2f}".format(float(video_rate_value)) if not is_video_rate_empty else "",
                        'video_number': row['Estimated Videos'],
                        'bonus_info': bonus_info['bonus_info'],
                        'payment_method': row['Payment method'],
                        'payment_information': row['Payment Info'],
                        'payment_charges': payment_fields['payment_charges']
                    }
                    template.render(context)
                    start_date = row.get('Start date', '')
                    start_datetime = datetime.strptime(str(start_date).strip(), '%Y-%m-%d')
                    contract_month = start_datetime.strftime('%Y-%m')
                    contract_filename = f'FW-ARETIS & {name_value}_{contract_month}.docx'
                    contract_files.append(contract_filename)
                    doc_stream = io.BytesIO()
                    template.save(doc_stream)
                    doc_stream.seek(0)
                    zip_file.writestr(contract_filename, doc_stream.read())
                if generate_summaries:
                    summary = generate_contract_summary(row)
                    summaries.append({
                        'name': str(row['Party B Name']).strip(),
                        'summary': summary,
                        'filename': f'Summary_{safe_name}_{today}.txt'
                    })
                    nickname = str(row.get('Main Platform nickname', '')).strip()
                    summary_filename = f'{name_value}_{nickname}_summary.txt'
                    if output_mode in ["配对输出", "单独文件"]:
                        zip_file.writestr(summary_filename, summary)
            except Exception as e:
                st.error(f"❌ Error processing {row.get('Party B Name', f'Row {index}')} (row {index}): {e}")
    if generate_summaries and output_mode == "合并文件" and summaries:
        combined_summary = ""
        for i, item in enumerate(summaries, 1):
            combined_summary += f"=== {item['name']} ===\n"
            combined_summary += item['summary']
            combined_summary += "\n\n"
        combined_filename = f'All_Summaries_{today}.txt'
        zip_file.writestr(combined_filename, combined_summary)
    return zip_buffer, summaries, contract_files

def process_form_data(form_records, uploaded_template, generate_contracts, generate_summaries, output_mode):
    """处理表单数据"""
    today = date.today().isoformat()
    zip_buffer = io.BytesIO()
    summaries = []
    contract_files = []
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
        for record in form_records:
            try:
                safe_name = str(record['Party B Name']).replace(" ", "_").replace("/", "-")
                
                if generate_contracts:
                    template = DocxTemplate(uploaded_template)
                    video_rate_value = record['Video Rate']
                    is_video_rate_empty = str(video_rate_value).strip() == ""
                    
                    # 自动推断内容
                    platform_fields = infer_platform_fields(record)
                    date_fields = infer_date_versions(record)
                    bonus_info = infer_bonus_info(record)
                    payment_fields = infer_payment_fields(record)
                    
                    context = {
                        'Influencer_name': record['Party B Name'],
                        'Influencer_email': record['Email'],
                        'Influencer_contact': "N/A" if str(record['Contact']).strip() == "" else record['Contact'],
                        'Influencer_address': record['Address'],
                        'platform': platform_fields['platform'],
                        'platform_username': platform_fields['platform_username'],
                        'Influencer_links': platform_fields['Influencer_links'],
                        'promotion_date': date_fields['promotion_date'],
                        'video_rate': "{:.2f}".format(float(video_rate_value)) if not is_video_rate_empty else "",
                        'video_number': record['Estimated Videos'],
                        'bonus_info': bonus_info['bonus_info'],
                        'payment_method': record['Payment method'],
                        'payment_information': record['Payment Info'],
                        'payment_charges': payment_fields['payment_charges']
                    }
                    
                    template.render(context)
                    start_date = record.get('Start date', '')
                    start_datetime = datetime.strptime(str(start_date).strip(), '%Y-%m-%d')
                    contract_month = start_datetime.strftime('%Y-%m')
                    contract_filename = f'FW-ARETIS & {record["Party B Name"]}_{contract_month}.docx'
                    contract_files.append(contract_filename)
                    
                    doc_stream = io.BytesIO()
                    template.save(doc_stream)
                    doc_stream.seek(0)
                    zip_file.writestr(contract_filename, doc_stream.read())
                
                if generate_summaries:
                    summary = generate_contract_summary(record)
                    summaries.append({
                        'name': str(record['Party B Name']).strip(),
                        'summary': summary,
                        'filename': f'Summary_{safe_name}_{today}.txt'
                    })
                    
                    nickname = str(record.get('Main Platform nickname', '')).strip()
                    summary_filename = f'{record["Party B Name"]}_{nickname}_summary.txt'
                    if output_mode in ["配对输出", "单独文件"]:
                        zip_file.writestr(summary_filename, summary)
                        
            except Exception as e:
                st.error(f"❌ Error processing {record.get('Party B Name', 'Unknown')}: {e}")
    
    if generate_summaries and output_mode == "合并文件" and summaries:
        combined_summary = ""
        for i, item in enumerate(summaries, 1):
            combined_summary += f"=== {item['name']} ===\n"
            combined_summary += item['summary']
            combined_summary += "\n\n"
        combined_filename = f'All_Summaries_{today}.txt'
        zip_file.writestr(combined_filename, combined_summary)
    
    return zip_buffer, summaries, contract_files

# Process files
if generate:
    if input_mode == "📝 表单填写（推荐）":
        # 表单模式处理
        if not st.session_state.form_records:
            st.warning("⚠️ 请先添加至少一条记录！")
        elif not uploaded_template:
            st.warning("⚠️ 请上传Word模板文件！")
        else:
            st.success("✅ 表单数据准备就绪！")
            
            # 显示处理进度
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                zip_buffer, summaries, contract_files = process_form_data(
                    st.session_state.form_records, 
                    uploaded_template, 
                    generate_contracts, 
                    generate_summaries, 
                    output_mode
                )
                progress_bar.progress(100)
                status_text.text("✅ 处理完成！")
                
                zip_buffer.seek(0)
                
                # 显示生成结果
                st.markdown("### ✅ Generation Complete!")
                
                if generate_contracts:
                    st.success(f"📄 Generated {len(contract_files)} contract files")
                
                if generate_summaries:
                    st.success(f"📝 Generated {len(summaries)} summary files")
                
                # 下载按钮
                download_filename = f"KOC_Form_Output_{date.today().isoformat()}.zip"
                st.download_button(
                    "📥 Download All Files", 
                    zip_buffer, 
                    file_name=download_filename, 
                    mime="application/zip"
                )
            
            except Exception as e:
                st.error(f"❌ Processing failed: {str(e)}")
    
    else:
        # CSV模式处理（原有功能）
        if not uploaded_csv or not uploaded_template:
            st.warning("⚠️ 请上传CSV文件和Word模板文件！")
        elif not generate_contracts and not generate_summaries:
            st.warning("请至少选择一个生成选项！")
        else:
            st.success("✅ Files uploaded successfully!")
            try:
                df = pd.read_csv(uploaded_csv, encoding='utf-8', keep_default_na=False, dtype=str)
            except UnicodeDecodeError:
                uploaded_csv.seek(0)
                df = pd.read_csv(uploaded_csv, encoding='gbk', keep_default_na=False, dtype=str)
            df.columns = df.columns.str.strip()
            
            # 显示处理进度
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                zip_buffer, summaries, contract_files = process_data(df, uploaded_template, generate_contracts, generate_summaries, output_mode)
                progress_bar.progress(100)
                status_text.text("✅ 处理完成！")
                
                zip_buffer.seek(0)
                
                # 显示生成结果
                st.markdown("### ✅ Generation Complete!")
                
                if generate_contracts:
                    st.success(f"📄 Generated {len(contract_files)} contract files")
                
                if generate_summaries:
                    st.success(f"📝 Generated {len(summaries)} summary files")
                
                # 下载按钮
                download_filename = f"KOC_Output_{date.today().isoformat()}.zip"
                st.download_button(
                    "📥 Download All Files", 
                    zip_buffer, 
                    file_name=download_filename, 
                    mime="application/zip"
                )
            
            except Exception as e:
                st.error(f"❌ Processing failed: {str(e)}")

else:
    if input_mode == "📝 表单填写（推荐）":
        st.info("📝 请填写表单信息，然后点击 Generate 按钮生成合同。")
    else:
        st.info("⬆️ Upload both files above to get started, then click Generate.")
