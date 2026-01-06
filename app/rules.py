import logging
import re

import pandas as pd
import streamlit as st

logger = logging.getLogger(__name__)


def evaluate_condition(row_value, operator, compare_value):
    """评估条件是否满足"""
    if pd.isna(row_value):
        row_value = ""
    else:
        row_value = str(row_value)

    compare_value = str(compare_value) if compare_value is not None else ""

    if operator == "=":
        return row_value == compare_value
    if operator == "<>":
        return row_value != compare_value
    if operator == "包含":
        return compare_value in row_value
    if operator == "不包含":
        return compare_value not in row_value
    if operator == ">":
        try:
            return float(row_value) > float(compare_value)
        except Exception:
            return False
    if operator == "<":
        try:
            return float(row_value) < float(compare_value)
        except Exception:
            return False
    if operator == ">=":
        try:
            return float(row_value) >= float(compare_value)
        except Exception:
            return False
    if operator == "<=":
        try:
            return float(row_value) <= float(compare_value)
        except Exception:
            return False
    return False


def extract_value(row, extract_type, extract_value_type, extract_value, regex_pattern=None, capture_group=1):
    """根据提取方式提取值"""
    if extract_value_type == "固定文本":
        source_value = extract_value
    else:
        if extract_value not in row.index:
            return []
        source_value = row[extract_value]
        if pd.isna(source_value):
            source_value = ""
        else:
            source_value = str(source_value)

    if extract_type == "直接提取":
        return [source_value] if source_value else []

    if extract_type == "正则提取":
        if not regex_pattern or not source_value:
            return []

        results = []
        try:
            for match in re.finditer(regex_pattern, source_value):
                groups = match.groups()
                if len(groups) >= capture_group:
                    extracted = groups[capture_group - 1].strip()
                    if extracted:
                        results.append(extracted)
        except Exception as e:
            logger.error("正则表达式错误: %s", str(e))
            st.warning(f"正则表达式错误: {str(e)}")

        return results

    if extract_type == "AI提取":
        return [source_value] if source_value else []

    return []


def process_variable_rules(row, rules, separator):
    """处理一个变量的所有规则（非AI提取）"""
    all_values = []

    for rule in rules:
        condition_column = rule.get("condition_column", "")
        condition_operator = rule.get("condition_operator", "=")
        condition_value = rule.get("condition_value", "")

        if not condition_column or condition_column not in row.index:
            continue

        if evaluate_condition(row[condition_column], condition_operator, condition_value):
            extracted = extract_value(
                row,
                rule.get("extract_type", "直接提取"),
                rule.get("extract_value_type", "从列提取"),
                rule.get("extract_value", ""),
                rule.get("regex_pattern", ""),
                rule.get("capture_group", 1),
            )
            all_values.extend(extracted)

    if not all_values:
        return ""

    combined = separator.join(all_values)
    split_values = [v.strip() for v in combined.split(separator) if v.strip()]
    unique_sorted = sorted(set(split_values))

    return separator.join(unique_sorted)
