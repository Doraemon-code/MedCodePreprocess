import os
import time
from openai import OpenAI
import pandas as pd

df = pd.read_excel(r"C:\Users\WeiQ\Desktop\test.xlsx", sheet_name="WHODRrugResult")
list = df["药物通用名"].tolist()
print(list)
API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-97c1523049f1427cb9379b082069e97a")
BASE_URL = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
MODEL = os.environ.get("DEEPSEEK_MODEL", "deepseek-chat")

def ai_extract(values):
    if not API_KEY:
        raise RuntimeError("DEEPSEEK_API_KEY 未设置")
    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    instr = (
        """你是一个具备多年医药行业经验的医药专家。任务：从每行药物名称中提取【核心成分】或【通用名】，以方便后续的医学编码的匹配工作，主要不要擅自添加信息。
        你需要严格遵守以下规则：
        1. 严格保持输出行数与输入行数一致；
        2. 如果无法提取或当前行为空，必须输出原始结果；
        3. 不要输出解释，只输出json结果；
        4. 只输出提取后的结果以及对应的行号，示例只是为了方便理解所有输入和输出同时给出。
        5. 大部分药物名称可能会包含一定程度的剂型信息,剂量信息或者给药途径的信息，你需要根据上下文理解并提取核心成分或通用名。

        以下是提取示例：
        裸花紫珠片 -- 裸花紫珠
        康复新液 -- 康复新
        头孢呋辛片 -- 头孢呋辛
        膏药 -- 膏药
        注射液用核黄素磷酸钠 -- 核黄素磷酸钠
        吸入用布地奈德混悬液 -- 布地奈德
        吸入用乙酰半胱氨酸溶液 -- 乙酰半胱氨酸
        地塞米松磷酸钠涂剂 -- 地塞米松磷酸钠
        精蛋白锌重组赖脯胰岛素混合注射液 -- 精蛋白锌重组赖脯胰岛素
        非那雄胺片 -- 非那雄胺
        坦索罗辛缓释胶囊 -- 坦索罗辛
        碳酸钙D3颗粒（Ⅱ） -- 碳酸钙D3
        维生素D滴剂（胶囊型） -- 维生素D
        左氨氯地平片 -- 左氨氯地平
        缬沙坦胶囊 -- 缬沙坦
        0.9%氯化钠注射液 -- 氯化钠
        艾瑞昔布 -- 艾瑞昔布
        阿司匹林 -- 阿司匹林
        中药 -- 中药
        """
    )
    batch = 25
    chunks = [values[i:i+batch] for i in range(0, len(values), batch)]
    out = []
    for chunk in chunks:
        orig = [str(v) if v is not None else "" for v in chunk]
        proc = [v if v.strip() else "N/A" for v in orig]
        user_content = "请提取以下数据的成分，严格按行对应输出：\n" + "\n".join(proc)
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": instr},
                {"role": "user", "content": user_content},
            ],
            response_format={"type": "json_object"},
            stream=False,
            temperature=0
        )
        content = resp.choices[0].message.content if resp and resp.choices else ""
        lines = [str(l).strip() for l in str(content).splitlines()]
        if len(lines) < len(orig):
            lines.extend([""] * (len(orig) - len(lines)))
        norm = []
        for i, x in enumerate(lines[:len(orig)]):
            if x == "N/A" or not x:
                norm.append(orig[i])
            else:
                norm.append(x)
        out.extend(norm)
        time.sleep(0.5)
    return out

if __name__ == "__main__":
    values = list
    results = ai_extract(values)
    for r in results:
        print(r)
