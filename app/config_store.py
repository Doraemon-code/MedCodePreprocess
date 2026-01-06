import json
import logging
from datetime import datetime

import streamlit as st

logger = logging.getLogger(__name__)


def load_all_configs():
    logger.info("尝试加载所有配置")
    try:
        with open("excel_processor_configs.json", "r", encoding="utf-8") as f:
            configs = json.load(f)
            logger.info("配置加载成功: 共 %s 个配置", len(configs))
            return configs
    except FileNotFoundError:
        logger.warning("配置文件不存在，返回空配置")
        return {}
    except Exception as e:
        logger.error("配置加载失败: %s", str(e), exc_info=True)
        return {}


def save_all_configs(all_configs):
    logger.info("尝试保存配置: 共 %s 个", len(all_configs))
    try:
        with open("excel_processor_configs.json", "w", encoding="utf-8") as f:
            json.dump(all_configs, f, ensure_ascii=False, indent=2)
        logger.info("配置保存成功")
        return True
    except Exception as e:
        logger.error("配置保存失败: %s", str(e), exc_info=True)
        st.error(f"保存失败: {str(e)}")
        return False


def save_current_config(config_name):
    logger.info("保存当前配置: %s", config_name)
    all_configs = load_all_configs()
    all_configs[config_name] = {
        "sheet_variables": st.session_state.sheet_variables,
        "saved_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    result = save_all_configs(all_configs)
    if result:
        logger.info("配置 '%s' 保存成功", config_name)
    return result


def load_config(config_name):
    logger.info("加载配置: %s", config_name)
    all_configs = load_all_configs()
    if config_name in all_configs:
        st.session_state.sheet_variables = all_configs[config_name]["sheet_variables"]
        logger.info("配置 '%s' 加载成功", config_name)
        return True
    logger.warning("配置 '%s' 不存在", config_name)
    return False


def delete_config(config_name):
    logger.info("删除配置: %s", config_name)
    all_configs = load_all_configs()
    if config_name in all_configs:
        del all_configs[config_name]
        result = save_all_configs(all_configs)
        if result:
            logger.info("配置 '%s' 删除成功", config_name)
        return result
    logger.warning("配置 '%s' 不存在，无需删除", config_name)
    return False
