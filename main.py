# modules/img2excel/core.py
import base64
import pandas as pd
from openai import OpenAI
import concurrent.futures
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

REVIEWER_PROMPT = """
# 角色定义
你是一个高精度的数据核对专家与文档处理大师。

# 背景
我们使用了多个AI视觉模型对同一张图片进行了表格数据提取，以下是它们各自的提取结果：
{extracted_results}

# 核心任务
请你结合原图，仔细对比上述不同模型的提取结果，找出并修正其中的错漏、对齐问题或省略的部分，整合出一份最准确、最完整的最终数据表格。

# 严格约束
1. 必须原原本本地输出每一个单元格，应对畸变与模糊。绝对不允许省略数据。
2. 必须且仅输出标准的 Markdown 表格文本，不要有任何多余的解释文字或代码块符号。
"""

def _call_vision_model(client: OpenAI, image_base64: str, model_name: str, prompt_text: str) -> str:
    """底层的单次模型调用函数"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt_text},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_base64}", "detail": "high"}}
            ]
        }],
        temperature=0.1, 
        max_tokens=4096
    )
    return response.choices[0].message.content

def parse_markdown_to_df(md_text: str) -> pd.DataFrame:
    """将 Markdown 解析为 DataFrame"""
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
        return pd.DataFrame(table_data[1:], columns=table_data[0])
    else:
        return pd.DataFrame(table_data)

def process_image_to_df(image_bytes: bytes, api_key: str, api_base: str, extract_models: list, reviewer_model: str = None) -> tuple:
    """接收图片字节流，可调用多个大模型提取并使用审阅模型合并，返回 Pandas DataFrame 和 MD 文本"""
    base64_img = base64.b64encode(image_bytes).decode('utf-8')
    client = OpenAI(api_key=api_key, base_url=api_base)
    
    # 1. 只有单个提取模型，且没有配置审阅模型时（单模型模式）
    if len(extract_models) == 1 and not reviewer_model:
        md_text = _call_vision_model(client, base64_img, extract_models[0], PROMPT_TEXT)
    
    # 2. 多模型并发提取 + 审阅机制
    else:
        results = []
        # 并发请求所有的提取模型
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future_to_model = {
                executor.submit(_call_vision_model, client, base64_img, model, PROMPT_TEXT): model 
                for model in extract_models
            }
            for i, future in enumerate(concurrent.futures.as_completed(future_to_model)):
                model_name = future_to_model[future]
                try:
                    res = future.result()
                    results.append(f"### 提取结果 (来自模型 {model_name}) ###\n{res}\n")
                except Exception as e:
                    print(f"模型 {model_name} 提取失败: {e}")
        
        if not results:
            raise Exception("所有前置提取模型均调用失败")

        # 将结果合并提交给审阅模型（如果未指定审阅模型，就用第一个模型代替）
        combined_text = "\n".join(results)
        final_prompt = REVIEWER_PROMPT.format(extracted_results=combined_text)
        final_model = reviewer_model if reviewer_model else extract_models[0]
        
        md_text = _call_vision_model(client, base64_img, final_model, final_prompt)

    # 3. 将最终的 Markdown 解析为 DataFrame
    df = parse_markdown_to_df(md_text)
        
    return df, md_text
