# modules/img2excel/core.py
import base64
import pandas as pd
from openai import OpenAI
import os



PROMPT_TEXT = """
# 角色定义
你是一个高精度的文档数字处理专家，擅长从复杂的拍照图像中提取结构化表格数据。

# 核心任务
将提供的图片中的核心数据表格提取出来，并转换为标准的 Markdown 表格。

# 约束与具体指令
1. ### 忽略背景与干扰 ###
   - 忽略纸张以外的桌面、纹理或背景。
2. ### 适度读取行标和列标 ###
   - 如图片中有 Excel 界面自带的最顶端列标(A, B...)和最左侧行号(1, 2...)，请将其分别作为 Markdown 表格的“表头”和“第一列”。如果没有，则只提取内部实际数据。
3. ### 应对畸变与模糊 ###
   - 图片存在拍照畸变，请按行列逻辑对齐数据。对于模糊不清的单元格，填入 `[模糊]`。
4. ### 完整性要求（绝对禁止省略） ###
   - 你必须从图片表格的第一行数据开始，逐行提取直到最后一行！
   - 绝对不允许“偷懒”，绝对不允许使用“...”、省略号或“等”字样来跳过中间或尾部的数据。
   - **必须原原本本地输出每一个单元格，不要因为重复空值就停止！**
5. ### 严格输出格式 ###
   - 必须且仅输出标准的 Markdown 表格文本，不要有任何多余的解释文字或代码块符号。
"""

def process_image_to_df(image_bytes: bytes, api_key: str, api_base: str, model_name: str) -> pd.DataFrame:
    """接收图片字节流，调用大模型，返回 Pandas DataFrame"""
    
    # 1. 直接将内存中的图片字节流转为 Base64
    base64_img = base64.b64encode(image_bytes).decode('utf-8')
    
    # 2. 调用视觉大模型
    client = OpenAI(api_key=api_key, base_url=api_base)
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": PROMPT_TEXT},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}", "detail": "high"}}
                ]
            }],
            temperature=0.1, 
            max_tokens=4096
        )
        md_text = response.choices[0].message.content
    except Exception as e:
        raise Exception(f"大模型 API 调用失败: {e}")

    # 3. 将 Markdown 解析为 DataFrame (复用你的优秀逻辑)
    if not md_text:
        raise Exception("模型返回结果为空")
        
    lines = [line.strip() for line in md_text.strip().split('\n') if line.strip()]
    data_lines = [line for line in lines if not set(line.replace('|', '').replace('-', '').replace(':', '').strip()).issubset(set(' '))]
    
    if not data_lines:
        raise Exception("未能在返回结果中找到有效的表格数据")
        
    table_data = []
    for line in data_lines:
        if line.startswith('|'): line = line[1:]
        if line.endswith('|'): line = line[:-1]
        row = [cell.strip() for cell in line.split('|')]
        table_data.append(row)
        
    if len(table_data) > 1:
        df = pd.DataFrame(table_data[1:], columns=table_data[0])
    else:
        df = pd.DataFrame(table_data)
        
    return df, md_text
