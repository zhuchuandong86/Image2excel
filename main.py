import os
import base64
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd
from config import API_KEY,API_URL,MODEL_VISION

# ---------------------------------------------------------
# 1. 配置内网 API 参数 (请替换为你们实际的地址、Key 和模型名)
# ---------------------------------------------------------
API_BASE_URL = API_URL  # 注意通常以 /v1 结尾
API_KEY = API_KEY
MODEL_NAME = MODEL_VISION # 必须与内网注册的模型名一致


# 文件夹配置
INPUT_FOLDER = "07_img2excel/input"   # 存放原图的文件夹
OUTPUT_FOLDER = "07_img2excel/output" # 存放生成的 Excel 文件夹

# 初始化客户端
client = OpenAI(api_key=API_KEY, base_url=API_BASE_URL)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 核心 Prompt (已使用优化后的版本)
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

# ---------------------------------------------------------
# 2. 核心处理函数
# ---------------------------------------------------------
def local_image_to_base64(file_path):
    with open(file_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def call_vlm_for_table(base64_img):
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
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
        return response.choices[0].message.content
    except Exception as e:
        print(f"  [!] 大模型 API 调用失败: {e}")
        return None

def markdown_to_excel(md_text, output_filepath):
    """
    将 Markdown 表格字符串解析并保存为 Excel 文件
    """
    if not md_text:
        return False
        
    # 1. 按行分割，过滤掉空行
    lines = [line.strip() for line in md_text.strip().split('\n') if line.strip()]
    
    # 2. 过滤掉 Markdown 表格中的分隔行 (例如 |---|---|---|)
    # 逻辑：如果一行全由 '-', '|', ':' 和空格组成，说明它是分隔行
    data_lines = [line for line in lines if not set(line.replace('|', '').replace('-', '').replace(':', '').strip()).issubset(set(' '))]
    
    if not data_lines:
        print("  [!] 未能在返回结果中找到有效的表格数据")
        return False
        
    # 3. 解析每一行的数据
    table_data = []
    for line in data_lines:
        # 去除开头和结尾的 '|'
        if line.startswith('|'): line = line[1:]
        if line.endswith('|'): line = line[:-1]
        
        # 按 '|' 分割并去除每个单元格两端的空格
        row = [cell.strip() for cell in line.split('|')]
        table_data.append(row)
        
    # 4. 转换为 Pandas DataFrame 并导出
    try:
        if len(table_data) > 1:
            # 第一行作为表头 (列名)
            df = pd.DataFrame(table_data[1:], columns=table_data[0])
        else:
            # 只有一行数据的情况
            df = pd.DataFrame(table_data)
            
        # 导出为 Excel (如果你想导出 CSV，可以直接把下面这行改成 df.to_csv(output_filepath.replace('.xlsx', '.csv'), index=False, encoding='utf-8-sig'))
        df.to_excel(output_filepath, index=False)
        return True
    except Exception as e:
        print(f"  [!] 转换为 Excel 时出错: {e}")
        return False

# ---------------------------------------------------------
# 3. 主干逻辑
# ---------------------------------------------------------
def process_folder():
    valid_image_exts = ('.jpg', '.jpeg', '.png')
    
    for filename in os.listdir(INPUT_FOLDER):
        filepath = os.path.join(INPUT_FOLDER, filename)
        if os.path.isdir(filepath): continue
            
        ext = os.path.splitext(filename)[1].lower()
        base_name = os.path.splitext(filename)[0]

        if ext in valid_image_exts:
            print(f"正在处理图片: {filename} ...")
            
            # 1. 提图转 Base64
            b64_img = local_image_to_base64(filepath)
            
            # 2. 调用模型获取 Markdown
            md_result = call_vlm_for_table(b64_img)
            
            # 3. 转换为 Excel 并保存
            if md_result:
                out_excel_path = os.path.join(OUTPUT_FOLDER, f"{base_name}.xlsx")
                success = markdown_to_excel(md_result, out_excel_path)
                if success:
                    print(f"  [+] 成功导出 Excel: {out_excel_path}")
        else:
            print(f"跳过非图片文件: {filename}")

if __name__ == "__main__":
    print(f"开始扫描文件夹: {INPUT_FOLDER}")
    process_folder()
    print("全部任务处理完成！")
