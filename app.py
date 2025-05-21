import streamlit as st
import pandas as pd
import tempfile
import os
import logging
from weekly_report_generator import WeeklyReportGenerator, generate_word_report

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

st.set_page_config(page_title="周报生成器", layout="wide")

# 注入自定义CSS，极简扁平化科技风
st.markdown('''
    <style>
    body, .stApp {background: #f7f9fb;}
    .stButton>button, .stDownloadButton>button {
        border-radius: 6px;
        background: linear-gradient(90deg, #3a8dde 0%, #5ad1e6 100%);
        color: white;
        border: none;
        padding: 0.6em 2em;
        font-size: 1.1em;
        font-weight: 600;
        margin: 0 0.2em 0 0;
        box-shadow: 0 2px 8px rgba(58,141,222,0.08);
        transition: background 0.1s, box-shadow 0.1s, transform 0.1s;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        filter: brightness(0.95);
        transform: translateY(2px) scale(0.98);
    }
    .stButton>button:active, .stDownloadButton>button:active {
        filter: brightness(0.90);
        transform: translateY(4px) scale(0.96);
    }
    .stFileUploader>div>div {border-radius: 6px; border: 1.5px solid #e3e8ee; background: #fff;}
    .stTextInput>div>div>input {border-radius: 6px; border: 1.5px solid #e3e8ee; background: #fff;}
    .stDataFrame {background: #fff; border-radius: 8px;}
    .stDownloadButton>button {
        border-radius: 6px;
        background: linear-gradient(90deg, #3a8dde 0%, #5ad1e6 100%);
        color: white;
        border: none;
        padding: 0.6em 2em;
        font-size: 1.1em;
        font-weight: 600;
        margin-right: 12px;
        margin-bottom: 0.5em;
        box-shadow: 0 2px 8px rgba(58,141,222,0.08);
        transition: background 0.1s, box-shadow 0.1s, transform 0.1s;
    }
    .stDownloadButton:last-child>button {margin-right: 0;}
    .stDownloadButton>button:hover {
        filter: brightness(0.95);
        transform: translateY(2px) scale(0.98);
    }
    .stDownloadButton>button:active {
        filter: brightness(0.90);
        transform: translateY(4px) scale(0.96);
    }
    </style>
''', unsafe_allow_html=True)

st.title("综合组周报生成器")

st.markdown("""
- 支持上传周报Excel，自动生成PDF/Word版本周报，上传Excel之后才会出现下载按钮
- 周报Excel严格按照要求填写哦，不然映射关系可能会乱，展示效果不好
- 如果PDF版本有格式或提取数据不正确的情况，可下载Word版本手动调整
""")

# 文件上传（中文提示）
uploaded_file = st.file_uploader("请上传周报Excel文件：", type=["xlsx", "xls"], help="仅支持Excel格式，直接从企微下载周报")

if uploaded_file:
    try:
        # 读取Excel数据
        df = pd.read_excel(uploaded_file)
        
        # 验证必需字段
        required_columns = [
            '姓名', '工作类型', '项目名称',
            '项目阶段', '上周三至本周二工作内容', '本周三至下周二工作计划',
            '问题反馈', '通过简历数量', '面试人员数量', '面试通过人员数量'
        ]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Excel文件缺少必需字段：{', '.join(missing_columns)}")
            st.stop()
            
        # 确保数值字段为数字类型
        numeric_columns = ['通过简历数量', '面试人员数量', '面试通过人员数量']
        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
        st.dataframe(df)
        
        with st.form("report_form"):
            col1, col2 = st.columns(2)
            with col1:
                issue = st.text_input("期数", value="", placeholder="如：1")
            with col2:
                date_str = st.text_input("日期", value="", placeholder="如：2025年5月20日")
            submitted = st.form_submit_button("生成周报")
            
        # 只要生成过一次，下载按钮就一直显示
        if (submitted and issue and date_str) or ("pdf" in st.session_state and "word" in st.session_state):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_xlsx:
                    df.to_excel(tmp_xlsx.name, index=False)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf, \
                         tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
                        try:
                            generator = WeeklyReportGenerator(tmp_xlsx.name, tmp_pdf.name, issue, date_str)
                            generator.run()
                            generate_word_report(tmp_xlsx.name, tmp_docx.name, issue, date_str)
                            
                            btn_col1, btn_col2 = st.columns([1, 1])
                            with btn_col1:
                                with open(tmp_pdf.name, "rb") as f:
                                    st.download_button(
                                        "下载PDF文件",
                                        f,
                                        file_name=f"产品研发部-综合业务组周报汇总-{date_str}.pdf",
                                        use_container_width=True,
                                        key="pdf"
                                    )
                            with btn_col2:
                                with open(tmp_docx.name, "rb") as f:
                                    st.download_button(
                                        "下载Word文件",
                                        f,
                                        file_name=f"产品研发部-综合业务组周报汇总-{date_str}.docx",
                                        use_container_width=True,
                                        key="word"
                                    )
                        except Exception as e:
                            logger.error(f"生成报告时出错: {str(e)}")
                            st.error(f"生成报告时出错: {str(e)}")
                        finally:
                            # 清理临时文件
                            try:
                                os.unlink(tmp_pdf.name)
                                os.unlink(tmp_docx.name)
                            except Exception as e:
                                logger.error(f"清理临时文件时出错: {str(e)}")
                os.unlink(tmp_xlsx.name)
            except Exception as e:
                logger.error(f"处理文件时出错: {str(e)}")
                st.error(f"处理文件时出错: {str(e)}")
        elif submitted:
            st.error("请填写期数和日期后再生成下载！")
    except Exception as e:
        logger.error(f"读取Excel文件时出错: {str(e)}")
        st.error(f"读取Excel文件时出错: {str(e)}") 