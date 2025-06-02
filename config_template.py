# -*- coding: utf-8 -*-
"""
File cấu hình cho nghiên cứu thị trường với OpenAI GPT API
Template file for deployment - rename to config.py and add your API key
"""

import os
import streamlit as st

# OpenAI API Configuration
# Try to get API key from multiple sources (deployment-friendly)
def get_openai_api_key():
    # 1. Try Streamlit secrets (for Streamlit Cloud deployment)
    try:
        return st.secrets["OPENAI_API_KEY"]
    except:
        pass
    
    # 2. Try environment variable
    env_key = os.getenv("OPENAI_API_KEY")
    if env_key:
        return env_key
    
    # 3. Fallback for local development
    # REPLACE WITH YOUR ACTUAL API KEY
    return "your-openai-api-key-here"

OPENAI_API_KEY = get_openai_api_key()

MODEL = 'gpt-3.5-turbo'

# Cấu hình nghiên cứu thị trường
RESEARCH_CONFIG = {
    "industry": "Công nghiệp ô tô",  # Ngành nghiên cứu
    "market": "Việt Nam",      # Thị trường nghiên cứu
    "input_file": "output/market_research_structured.json",  # File JSON input
    "delay_seconds": 3,          # Delay between API calls
    "max_retries": 3,           # Số lần thử lại khi API lỗi
}

# Cấu hình output
OUTPUT_CONFIG = {
    "save_intermediate": True,   # Lưu kết quả trung gian
    "output_format": "json",     # Format output: json, csv, excel
    "include_prompts": True,     # Có lưu prompts được tạo ra không
}

# Cấu hình logging
LOGGING_CONFIG = {
    "show_progress": True,       # Hiển thị progress bar
    "log_level": "INFO",         # DEBUG, INFO, WARNING, ERROR
    "save_logs": True,           # Lưu logs ra file
}

# Các thông số khác
MISC_CONFIG = {
    "backup_results": True,      # Tạo backup kết quả
    "compress_output": False,    # Nén file output
    "validate_input": True,      # Kiểm tra tính hợp lệ của input
}

# Rate Limiting Configuration
RATE_LIMIT_CONFIG = {
    "base_delay": 3,            # Base delay between requests (seconds)
    "max_retries": 3,          # Maximum retry attempts for rate limit
    "backoff_factor": 2,       # Exponential backoff multiplier  
    "max_backoff": 60,         # Maximum wait time (seconds)
} 