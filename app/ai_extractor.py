import json
import logging
import re
import time
from typing import Dict, List, Optional

from openai import OpenAI

from .settings import AI_CONFIG

logger = logging.getLogger(__name__)

AI_SYSTEM_PROMPT = """你是一个具备多年医药行业经验的医药专家。任务：从每行药物名称中提取【核心成分】或【通用名】，以方便后续的医学编码的匹配工作，主要不要擅自添加信息。
你需要严格遵守以下规则：
1. 只输出 JSON，不要输出任何解释、代码块或多余文本；
2. 输入为 JSON，包含 items 列表，每个元素含 id 和 text；
3. 输出为 JSON 对象：{"results":[{"id":0,"value":"..."}] }；
4. id 必须原样返回且一一对应；
5. 如果无法提取、无法确定或输入为空，value 必须与输入 text 完全一致（不确定就原样返回）；
6. 不要省略盐基成分，例如"硫酸氨基葡萄糖片"应输出“硫酸氨基葡萄糖”、"复方氨酚烷胶囊"应输出"复方氨酚烷",这里不应该移除复方；
7. 只输出提取后的结果，不要输出示例或解释。

以下是提取示例（理解规则用）：
苯磺酸左氨氯地平片 -- 苯磺酸左氨氯地平
硫酸氨基葡萄糖片 -- 硫酸氨基葡萄糖
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

UNCERTAIN_TOKENS = {"n/a", "na", "null", "none"}


def _strip_code_fences(text: str) -> str:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3].strip()
    return cleaned


def _extract_json(content: str) -> Optional[object]:
    cleaned = _strip_code_fences(content)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    match = re.search(r"(\{.*\}|\[.*\])", cleaned, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            return None
    return None


def _normalize_result(raw_value: object, original: str) -> str:
    if raw_value is None:
        return original

    text = str(raw_value).strip()
    if not text:
        return original

    lower = text.lower()
    if lower in UNCERTAIN_TOKENS:
        return original

    if "不确定" in text or "无法确定" in text or "无法判断" in text:
        return original

    return text


def ai_extract_batch(values: List[object], column_name: str = "未知列", cache: Optional[Dict[str, str]] = None) -> List[str]:
    """使用AI提取药物成分（单批次），带缓存和JSON协议"""
    logger.info("AI提取批次 - 列名: %s, 数据量: %s", column_name, len(values))

    if cache is None:
        cache = {}

    if not AI_CONFIG["API_KEY"]:
        logger.error("AI提取失败: API_KEY未设置")
        raise RuntimeError("DEEPSEEK_API_KEY 未设置")

    orig = [str(v) if v is not None else "" for v in values]
    results: List[Optional[str]] = [None] * len(orig)
    pending_map: Dict[str, List[int]] = {}
    cache_hits = 0
    empty_count = 0

    for idx, text in enumerate(orig):
        if text in cache:
            results[idx] = cache[text]
            cache_hits += 1
            continue
        if not text.strip():
            results[idx] = text
            empty_count += 1
            continue
        pending_map.setdefault(text, []).append(idx)

    pending_count = sum(len(indices) for indices in pending_map.values())
    unique_pending = len(pending_map)
    logger.info(
        "缓存命中: %s/%s, 空值: %s, 待请求: %s (去重后 %s)",
        cache_hits,
        len(orig),
        empty_count,
        pending_count,
        unique_pending,
    )

    if not pending_map:
        logger.info("批次处理完成, 全部命中缓存")
        return [r if r is not None else "" for r in results]

    items = []
    id_to_text: Dict[int, str] = {}
    for item_id, text in enumerate(pending_map.keys()):
        items.append({"id": item_id, "text": text})
        id_to_text[item_id] = text

    user_content = json.dumps({"items": items}, ensure_ascii=False)

    try:
        client = OpenAI(api_key=AI_CONFIG["API_KEY"], base_url=AI_CONFIG["BASE_URL"])
        logger.info("OpenAI客户端初始化成功")
    except Exception as e:
        logger.error("OpenAI客户端初始化失败: %s", str(e))
        raise

    try:
        logger.info("开始调用AI API")
        start_time = time.time()

        resp = client.chat.completions.create(
            model=AI_CONFIG["MODEL"],
            messages=[
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
            response_format={"type": "json_object"},
            stream=False,
            temperature=AI_CONFIG["TEMPERATURE"],
        )

        elapsed_time = time.time() - start_time
        logger.info("AI API调用成功, 耗时: %.2f秒", elapsed_time)

        content = resp.choices[0].message.content if resp and resp.choices else ""
        data = _extract_json(content) if content else None

        if data is None:
            raise ValueError("AI返回不是有效JSON")

        if isinstance(data, dict):
            results_list = data.get("results")
        else:
            results_list = data

        id_to_value: Dict[int, object] = {}
        if isinstance(results_list, list):
            for item in results_list:
                if not isinstance(item, dict):
                    continue
                item_id = item.get("id")
                if isinstance(item_id, str) and item_id.isdigit():
                    item_id = int(item_id)
                if isinstance(item_id, int):
                    id_to_value[item_id] = item.get("value")
        else:
            raise ValueError("AI返回JSON结构不正确")

        for item_id, text in id_to_text.items():
            raw_value = id_to_value.get(item_id)
            normalized = _normalize_result(raw_value, text)
            cache[text] = normalized
            for idx in pending_map[text]:
                results[idx] = normalized

        logger.info("批次处理完成, 结果数: %s", len(orig))
        return [r if r is not None else "" for r in results]

    except Exception as e:
        logger.error("AI API调用失败: %s", str(e), exc_info=True)
        logger.info("使用原始数据作为后备")
        for text, indices in pending_map.items():
            for idx in indices:
                results[idx] = text
        return [r if r is not None else "" for r in results]
