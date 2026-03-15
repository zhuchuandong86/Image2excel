# pages/04_🖼️_图片转Excel.py
import os
import io
import streamlit as st
from dotenv import load_dotenv
import core.paths
from modules.img2excel.core import process_image_to_df

# 加载全局环境变量
load_dotenv(core.paths.ENV_FILE)
API_KEY = os.getenv("INTERNAL_API_KEY")
API_BASE = os.getenv("INTERNAL_API_BASE")
MODEL_VISION = os.getenv("MODEL_VISION", "deepseek-v3-0324") # 如果环境变量没配，给个默认值

st.set_page_config(page_title="图片转Excel", page_icon="🖼️", layout="wide")
st.title("🖼️ 智能 OCR：图片提取转 Excel")
st.markdown("上传含有表格的截图或照片，AI 将精准提取数据并允许您下载为标准 Excel 文件。")

# 左侧上传，右侧预览
col1, col2 = st.columns([1, 1])

with col1:
    uploaded_file = st.file_uploader("📂 上传图片文件 (支持 JPG/PNG)", type=['png', 'jpg', 'jpeg'])

with col2:
    if uploaded_file is not None:
        st.image(uploaded_file, caption="原图预览", use_container_width=True)

st.divider()

# 执行逻辑
if uploaded_file is not None:
    if st.button("🚀 开始提取表格数据", type="primary", use_container_width=True):
        if not API_KEY:
            st.error("缺失 API KEY 配置，请检查根目录的 .env 文件！")
            st.stop()
            
        with st.spinner("🤖 视觉模型正在逐行扫描并提取数据，请稍候..."):
            try:
                # 传入文件的字节流 getvalue()
                df, raw_md = process_image_to_df(
                    image_bytes=uploaded_file.getvalue(), 
                    api_key=API_KEY, 
                    api_base=API_BASE, 
                    model_name=MODEL_VISION
                )
                
                st.success("✅ 提取成功！预览如下：")
                st.dataframe(df, use_container_width=True)
                
                # 在内存中生成 Excel 文件供下载，不产生磁盘垃圾文件
                excel_buffer = io.BytesIO()
                df.to_excel(excel_buffer, index=False)
                excel_data = excel_buffer.getvalue()
                
                col_btn, _ = st.columns([1, 3])
                with col_btn:
                    st.download_button(
                        label="📥 一键下载 Excel 文件",
                        data=excel_data,
                        file_name=f"表格提取结果_{uploaded_file.name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                with st.expander("🛠️ 查看大模型原始 Markdown 返回值 (供排查)"):
                    st.code(raw_md)
                    
            except Exception as e:
                st.error(f"提取失败: {e}")
