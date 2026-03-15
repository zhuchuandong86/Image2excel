# pages/04_🖼️_图片转Excel.py
import os
import io
import streamlit as st
import pandas as pd
from dotenv import load_dotenv
import core.paths
from modules.img2excel.core import process_image_to_df

# 加载全局环境变量
load_dotenv(core.paths.ENV_FILE)
API_KEY = os.getenv("INTERNAL_API_KEY")
API_BASE = os.getenv("INTERNAL_API_BASE")
DEFAULT_MODEL = os.getenv("MODEL_VISION", "deepseek-v3-0324")

st.set_page_config(page_title="图片转Excel", page_icon="🖼️", layout="wide")
st.title("🖼️ 智能 OCR：图片提取转 Excel")
st.markdown("上传含有表格的截图或照片，支持**多张图片**和**Ctrl+V直接粘贴**，AI 将提取并允许您下载为标准 Excel 文件。")

# 侧边栏：配置多模型 A+B -> C 架构
with st.sidebar:
    st.header("⚙️ 多模型校验配置")
    st.markdown("支持填入多个提取模型（以逗号分隔）。如果填入多个，会并发调用它们，然后交给审阅模型整合。")
    extract_input = st.text_input("📝 提取模型 (A, B...)", value=DEFAULT_MODEL, help="用英文逗号分隔，如: qwen-vl-plus, deepseek-vl")
    reviewer_input = st.text_input("🕵️ 审阅模型 (C)", value="", help="可选。用于审阅和整合提取模型的结果。留空则只用单模型或默认模型合并。")
    
    extract_models = [m.strip() for m in extract_input.split(",") if m.strip()]
    reviewer_model = reviewer_input.strip() if reviewer_input.strip() else None

# 左侧上传，右侧预览
col1, col2 = st.columns([1, 1])

with col1:
    # 开启 accept_multiple_files=True 支持多图
    # Streamlit 原生支持鼠标点击该区域后直接 Ctrl+V / Cmd+V 粘贴截图
    uploaded_files = st.file_uploader(
        "📂 上传/粘贴图片 (支持多选及快捷键粘贴)", 
        type=['png', 'jpg', 'jpeg'], 
        accept_multiple_files=True
    )

with col2:
    if uploaded_files:
        st.write(f"已加载 {len(uploaded_files)} 张图片：")
        # 如果图片太多，可以用 columns 或 tabs 并排显示预览
        preview_cols = st.columns(min(len(uploaded_files), 3))
        for i, file in enumerate(uploaded_files):
            with preview_cols[i % 3]:
                st.image(file, caption=file.name, use_container_width=True)

st.divider()

# 执行逻辑
if uploaded_files:
    if st.button("🚀 开始批量提取表格数据", type="primary", use_container_width=True):
        if not API_KEY:
            st.error("缺失 API KEY 配置，请检查根目录的 .env 文件！")
            st.stop()
            
        if not extract_models:
            st.error("请至少在左侧配置一个提取模型！")
            st.stop()
            
        all_dfs = []
        all_mds = []
        
        # 添加总体进度提示
        progress_text = "🤖 视觉模型正在逐图扫描并提取数据，请稍候..."
        progress_bar = st.progress(0, text=progress_text)
        
        try:
            for idx, uploaded_file in enumerate(uploaded_files):
                st.toast(f"正在处理第 {idx+1}/{len(uploaded_files)} 张图片: {uploaded_file.name}")
                
                # 传入多模型配置进行处理
                df, raw_md = process_image_to_df(
                    image_bytes=uploaded_file.getvalue(), 
                    api_key=API_KEY, 
                    api_base=API_BASE, 
                    extract_models=extract_models,
                    reviewer_model=reviewer_model
                )
                
                all_dfs.append(df)
                all_mds.append(f"### {uploaded_file.name} 提取结果 ###\n{raw_md}")
                
                # 更新进度条
                progress_bar.progress((idx + 1) / len(uploaded_files), text=progress_text)
            
            st.success("✅ 全部提取成功！")
            
            # 将多张图提取出的表格垂直合并拼接 (如果多图是同一个大表的不同部分，这非常有用)
            final_df = pd.concat(all_dfs, ignore_index=True)
            st.dataframe(final_df, use_container_width=True)
            
            # 在内存中生成 Excel 文件
            excel_buffer = io.BytesIO()
            final_df.to_excel(excel_buffer, index=False)
            excel_data = excel_buffer.getvalue()
            
            col_btn, _ = st.columns([1, 3])
            with col_btn:
                st.download_button(
                    label="📥 一键下载合并后的 Excel",
                    data=excel_data,
                    file_name="批量表格提取结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
            with st.expander("🛠️ 查看大模型原始 Markdown 返回值 (供排查)"):
                st.code("\n\n".join(all_mds))
                
        except Exception as e:
            st.error(f"提取过程中发生错误: {e}")
