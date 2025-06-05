#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🔬 Market Research Automation System - Web UI
Beautiful Streamlit interface for marketing teams
"""

import streamlit as st
import pandas as pd
import json
import os
import time
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from streamlit_option_menu import option_menu
import re

# Get OpenAI API key from multiple sources (deployment-friendly)
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
    
    # 3. Try importing from config (local development)
    try:
        from config import OPENAI_API_KEY
        return OPENAI_API_KEY
    except ImportError:
        pass
    
    # 4. Default fallback
    return "your-openai-api-key-here"

OPENAI_API_KEY = get_openai_api_key()

# Import our core modules
from excel_to_structured_json import ExcelToStructuredJSON
from openai_market_research import OpenAIMarketResearch
from export_comprehensive_report import create_comprehensive_word_report, find_latest_research_file

# World countries list for Target Market selection
WORLD_COUNTRIES = [
    "🇻🇳 Việt Nam", "🇺🇸 United States", "🇨🇳 China", "🇯🇵 Japan", "🇰🇷 South Korea",
    "🇹🇭 Thailand", "🇸🇬 Singapore", "🇲🇾 Malaysia", "🇮🇩 Indonesia", "🇵🇭 Philippines",
    "🇬🇧 United Kingdom", "🇩🇪 Germany", "🇫🇷 France", "🇮🇹 Italy", "🇪🇸 Spain",
    "🇨🇦 Canada", "🇦🇺 Australia", "🇳🇿 New Zealand", "🇧🇷 Brazil", "🇲🇽 Mexico",
    "🇮🇳 India", "🇷🇺 Russia", "🇿🇦 South Africa", "🇪🇬 Egypt", "🇦🇪 UAE",
    "🇸🇦 Saudi Arabia", "🇹🇷 Turkey", "🇳🇱 Netherlands", "🇸🇪 Sweden", "🇳🇴 Norway",
    "🇩🇰 Denmark", "🇫🇮 Finland", "🇨🇭 Switzerland", "🇦🇹 Austria", "🇧🇪 Belgium",
    "🇵🇱 Poland", "🇨🇿 Czech Republic", "🇭🇺 Hungary", "🇬🇷 Greece", "🇵🇹 Portugal",
    "🇮🇪 Ireland", "🇮🇱 Israel", "🇭🇰 Hong Kong", "🇹🇼 Taiwan", "🇦🇷 Argentina",
    "🇨🇱 Chile", "🇨🇴 Colombia", "🇵🇪 Peru", "🇻🇪 Venezuela", "🇪🇨 Ecuador",
    "🇺🇾 Uruguay", "🇧🇴 Bolivia", "🇵🇾 Paraguay", "🇳🇬 Nigeria", "🇰🇪 Kenya",
    "🇬🇭 Ghana", "🇪🇹 Ethiopia", "🇺🇬 Uganda", "🇹🇿 Tanzania", "🇿🇼 Zimbabwe",
    "🌏 Southeast Asia", "🌍 Asia-Pacific", "🌎 Global Market"
]

# Page config
st.set_page_config(
    page_title="🔬 Market Research AI",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .feature-box {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin: 1rem 0;
    }
    .success-box {
        background: #d4edda;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    .warning-box {
        background: #fff3cd;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #ffc107;
        margin: 1rem 0;
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
    }
    .download-section {
        background: #e8f5e8;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main Streamlit application"""
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🔬 Market Research Automation System</h1>
        <p>AI-powered market research reports for marketing teams</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Simplified navigation - only essential pages
    with st.sidebar:
        selected = option_menu(
            menu_title="Navigation",
            options=["🚀 New Research", "⚙️ Settings"],
            icons=["rocket", "gear"],
            menu_icon="cast",
            default_index=0,
        )
    
    # Main content based on selection
    if selected == "🚀 New Research":
        show_research_page()
    elif selected == "⚙️ Settings":
        show_settings_page()

def validate_excel_template(file_path):
    """
    Validate uploaded Excel file format
    Returns: (is_valid, error_message, main_question_count, sub_question_count, total_estimated_cost)
    """
    try:
        # Try to read the Excel file
        excel_file = pd.ExcelFile(file_path)
        
        # Check if 'template' sheet exists
        if 'template' not in excel_file.sheet_names:
            return False, "❌ Sheet 'template' not found! Please use the correct template format or rename your sheet to 'template'.", 0, 0, 0
        
        # Read template sheet
        df = pd.read_excel(file_path, sheet_name="template")
        
        # Check if dataframe is not empty
        if df.empty:
            return False, "❌ Sheet 'template' is empty! Please add data to your template.", 0, 0, 0
        
        # Check for basic required structure
        if df.shape[1] < 3:  # At least 3 columns
            return False, "❌ Template needs at least 3 columns! Please use the standard template format.", 0, 0, 0
        
        # Look for purpose row
        purpose_found = False
        for idx, row in df.iterrows():
            if any("Mục đích" in str(cell) for cell in row if pd.notna(cell)):
                purpose_found = True
                break
        
        if not purpose_found:
            return False, "❌ 'Market Research Purpose' row not found! Please use the standard template format.", 0, 0, 0
        
        # Count questions for cost estimation
        question_count, sub_question_count = count_questions_in_excel(df)
        
        # Calculate total cost: Layer 3 + Layer 4 enhancements
        layer3_cost = estimate_research_cost(question_count)
        layer4_cost = sub_question_count * 0.05  # $0.05 per Layer 4 enhancement
        total_estimated_cost = layer3_cost + layer4_cost
        
        return True, "✅ Template valid!", question_count, sub_question_count, total_estimated_cost
        
    except Exception as e:
        return False, f"❌ Error reading Excel file: {str(e)}. Please check if the file is corrupted.", 0, 0, 0

def count_questions_in_excel(df):
    """
    Count total number of questions (Layer 3) and sub-questions (Layer 4+) in Excel template
    Supports dynamic number of layers (Layer 5, 6, etc.)
    Returns: (main_question_count, sub_question_count)
    """
    main_question_count = 0
    sub_question_count = 0
    
    # Find layer header row
    layer_header_row = None
    layer_columns = []
    
    for idx, row in df.iterrows():
        if any("Layer" in str(cell) for cell in row if pd.notna(cell)):
            layer_header_row = idx
            # Detect all layer columns dynamically
            for col_idx, col_val in enumerate(row):
                if pd.notna(col_val) and "Layer" in str(col_val):
                    layer_columns.append(col_idx)
            break
    
    if layer_header_row is None or len(layer_columns) < 3:
        return 0, 0
    
    # Debug information
    print(f"🔍 DEBUG: Found {len(layer_columns)} layers at columns {layer_columns}")
    
    # Map layer positions for counting
    layer3_col = layer_columns[2] if len(layer_columns) > 2 else None  # Layer 3 (main questions)
    layer4_plus_cols = layer_columns[3:] if len(layer_columns) > 3 else []  # Layer 4+ (sub-questions)
    
    print(f"🔍 DEBUG: Layer 3 column: {layer3_col}, Layer 4+ columns: {layer4_plus_cols}")
    
    # Count questions with more detailed logging
    question_details = []
    try:
        for idx in range(layer_header_row + 1, len(df)):
            row = df.iloc[idx]
            
            # Skip empty rows
            if row.isna().all():
                continue
            
            # Check Layer 3 column (main questions)
            if layer3_col is not None and layer3_col < len(row):
                layer3_val = row.iloc[layer3_col] if pd.notna(row.iloc[layer3_col]) else None
                if layer3_val and str(layer3_val).strip() and str(layer3_val) != 'None':
                    main_question_count += 1
                    
                    # Check Layer 4+ columns for sub-questions
                    has_sub_questions = False
                    total_individual_sub_questions = 0
                    sub_question_texts = []
                    for layer_col in layer4_plus_cols:
                        if layer_col < len(row):
                            layer_val = row.iloc[layer_col] if pd.notna(row.iloc[layer_col]) else None
                            if layer_val and str(layer_val).strip() and str(layer_val) != 'None':
                                has_sub_questions = True
                                
                                # Count individual sub-questions for debugging
                                sub_q_text = str(layer_val)
                                # Split by common patterns: newlines, semicolons, bullet points
                                import re
                                sub_questions_list = re.split(r'[\n;•\-]\s*', sub_q_text)
                                # Filter out empty strings and count actual sub-questions
                                valid_sub_questions = [sq.strip() for sq in sub_questions_list if sq.strip()]
                                individual_count = len(valid_sub_questions)
                                total_individual_sub_questions += individual_count
                                sub_question_texts.append(f"{individual_count} individual sub-questions")
                    
                    if has_sub_questions:
                        sub_question_count += 1  # Count as 1 enhancement per main question with sub-questions
                    
                    # Store for debugging
                    question_details.append({
                        'row': idx,
                        'main_question': str(layer3_val)[:50] + "...",
                        'has_sub_questions': has_sub_questions,
                        'individual_sub_question_count': total_individual_sub_questions,
                        'sub_question_preview': sub_question_texts[:2]  # First 2 sub-questions preview
                    })
                        
    except Exception as e:
        print(f"Error counting questions: {e}")
        return 0, 0
    
    # Debug output
    total_individual_sub_questions = sum(detail['individual_sub_question_count'] for detail in question_details)
    print(f"🔍 DEBUG: Final counts - Main questions: {main_question_count}, Layer 4 enhancements: {sub_question_count}")
    print(f"🔍 DEBUG: Total individual sub-questions across all main questions: {total_individual_sub_questions}")
    print(f"🔍 DEBUG: Sample questions:")
    for i, detail in enumerate(question_details[:5]):  # Show first 5
        print(f"  {i+1}. Row {detail['row']}: {detail['main_question']} | Has sub-questions: {detail['has_sub_questions']}")
        if detail['sub_question_preview']:
            print(f"     Sub-questions preview: {detail['sub_question_preview']} | Individual count: {detail['individual_sub_question_count']}")
    
    return main_question_count, sub_question_count

def calculate_dynamic_cost(question_count, sub_question_count, analysis_level="Layer 4 Analysis"):
    """
    Calculate cost dynamically based on analysis level choice
    
    Args:
        question_count: Number of main questions (Layer 3)
        sub_question_count: Number of sub-questions (Layer 4)
        analysis_level: "Layer 3 Analysis" or "Layer 4 Analysis"
    
    Returns:
        tuple: (total_cost, layer3_cost, layer4_cost, category_count_estimate)
    """
    
    if analysis_level == "Layer 3 Analysis":
        # Layer 3 mode: Cost based on categories, not individual questions
        # Estimate categories (typically 4-8 categories per framework)
        category_count_estimate = max(2, min(8, question_count // 3))  # Rough estimate
        
        # Cost per category analysis (more expensive per unit but fewer units)
        if category_count_estimate <= 3:
            category_cost = 0.08  # $0.08 per category
        elif category_count_estimate <= 6:
            category_cost = 0.06  # $0.06 per category  
        else:
            category_cost = 0.05  # $0.05 per category
        
        layer3_cost = round(category_count_estimate * category_cost, 2)
        layer4_cost = 0  # No Layer 4 in Layer 3 mode
        total_cost = layer3_cost
        
        return total_cost, layer3_cost, layer4_cost, category_count_estimate
    
    else:
        # Layer 4 mode: Original calculation method
        layer3_cost = estimate_research_cost(question_count)
        layer4_cost = sub_question_count * 0.05  # $0.05 per Layer 4 enhancement
        total_cost = layer3_cost + layer4_cost
        
        return total_cost, layer3_cost, layer4_cost, 0

def estimate_research_cost(question_count):
    """
    Estimate research cost based on number of questions
    Returns: float - estimated cost in USD
    """
    if question_count <= 5:
        return round(question_count * 0.03, 2)
    elif question_count <= 25:
        return round(question_count * 0.02, 2)
    elif question_count <= 50:
        return round(question_count * 0.025, 2)
    elif question_count <= 100:
        return round(question_count * 0.03, 2)
    else:
        # For very large templates, increase cost per question
        return round(question_count * 0.04, 2)

def display_cost_warning(question_count, total_estimated_cost, analysis_level="Layer 4 Analysis", layer3_cost=0, layer4_cost=0, category_count=0):
    """Display cost warning with analysis level breakdown"""
    
    if total_estimated_cost <= 0.15:
        st.info(f"💰 Total Cost: ${total_estimated_cost:.2f} ({question_count} questions)")
    elif total_estimated_cost <= 0.50:
        st.info(f"💰 Total Cost: ${total_estimated_cost:.2f} ({question_count} questions) - Standard Level")
    elif total_estimated_cost <= 1.25:
        st.warning(f"⚠️ Total Cost: ${total_estimated_cost:.2f} ({question_count} questions) - High Level")
    elif total_estimated_cost <= 3.0:
        st.error(f"🚨 Total Cost: ${total_estimated_cost:.2f} ({question_count} questions) - Very High!")
    else:
        st.error(f"🔴 WARNING: Total Cost ${total_estimated_cost:.2f} ({question_count} questions) - Extremely Expensive!")
        st.error("💡 Recommendation: Split into multiple smaller research projects")
    
    # Show breakdown based on analysis level
    if analysis_level == "Layer 3 Analysis":
        if category_count > 0:
            st.caption(f"📊 Breakdown: {category_count} categories × ${layer3_cost/category_count:.2f} = ${total_estimated_cost:.2f}")
            st.caption(f"ℹ️ Layer 3 Mode: Comprehensive analysis per category (no sub-questions)")
    else:
        # Layer 4 breakdown
        if layer4_cost > 0:
            st.caption(f"📊 Breakdown: Layer 3 (${layer3_cost:.2f}) + Layer 4 (${layer4_cost:.2f}) = ${total_estimated_cost:.2f}")
            st.caption(f"ℹ️ Detected {int(layer4_cost / 0.05)} questions with sub-questions → auto Layer 4 reports")
        else:
            st.caption(f"📊 Layer 4 Mode: ${total_estimated_cost:.2f} for {question_count} individual question analysis")

def proceed_with_research():
    """Start research immediately for low-cost cases"""
    st.session_state['start_research'] = True
    st.rerun()

def show_research_page():
    """New research creation page - simplified and streamlined"""
    st.header("🚀 Create Market Research Report")
    
    # Check API key
    if not OPENAI_API_KEY or OPENAI_API_KEY == "your-openai-api-key-here":
        st.error("⚠️ Please configure your OpenAI API key in Settings first!")
        return
    
    # Simple template download
    template_path = 'input/market research template.xlsx'
    if os.path.exists(template_path):
        with open(template_path, 'rb') as template_file:
            st.download_button(
                label="📥 Download Excel Template",
                data=template_file.read(),
                file_name="market_research_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # Excel file upload with IMMEDIATE validation
    uploaded_file = st.file_uploader(
        "📁 Upload custom Excel framework (optional)",
        type=['xlsx'],
        help="Upload customized template or leave empty to use default"
    )
    
    # IMMEDIATE validation when file is uploaded
    if uploaded_file:
        st.markdown("#### 🔍 File Validation Results")
        
        # Save temp file for validation
        temp_path = os.path.join('output', 'temp_uploaded.xlsx')
        os.makedirs('output', exist_ok=True)
        
        try:
            with open(temp_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            
            # Validate file
            is_valid, message, question_count, sub_question_count, total_estimated_cost = validate_excel_template(temp_path)
            
            if is_valid:
                st.success(message)
                st.info("💡 **Cost will be calculated dynamically based on your Analysis Level choice below**")
                
                # Store template data in session state for cost preview
                st.session_state['template_data'] = {
                    'question_count': question_count,
                    'sub_question_count': sub_question_count,
                    'file_type': 'uploaded'
                }
                
                # Show basic breakdown details  
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("📊 Main Questions", question_count)
                with col2:
                    st.metric("🎯 Layer 4 Enhancements", sub_question_count, help="Number of main questions that have sub-questions")
                with col3:
                    st.metric("📋 Template Valid", "✅")
                    
            else:
                st.error(message)
                # Clear any previous template data on validation failure
                if 'template_data' in st.session_state:
                    del st.session_state['template_data']
                
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
    else:
        # Use default template - simple and clean
        default_template_path = 'input/market research template.xlsx'
        
        if os.path.exists(default_template_path):
            try:
                # Validate default template
                is_valid, message, question_count, sub_question_count, total_estimated_cost = validate_excel_template(default_template_path)
                
                if is_valid:
                    # Store default template data in session state
                    st.session_state['template_data'] = {
                        'question_count': question_count,
                        'sub_question_count': sub_question_count,
                        'file_type': 'default'
                    }
                else:
                    st.warning("⚠️ Default template has issues. Please upload a custom file.")
                    
            except Exception as e:
                # Fallback to estimated values and store in session state
                estimated_questions = 22
                estimated_sub_questions = 5
                st.session_state['template_data'] = {
                    'question_count': estimated_questions,
                    'sub_question_count': estimated_sub_questions,
                    'file_type': 'fallback'
                }
        else:
            st.warning("⚠️ Default template not found. Please upload a custom Excel file.")
            # Store fallback data in session state
            st.session_state['template_data'] = {
                'question_count': 22,
                'sub_question_count': 5,
                'file_type': 'fallback'
            }
    
    # Research form
    with st.form("research_form"):
        st.subheader("📝 Research Configuration")
        
        # Topic input
        research_topic = st.text_input(
            "🎯 Research Topic",
            placeholder="e.g., Electric vehicles, Food delivery, E-commerce",
            help="Enter the industry or market you want to research"
        )
        
        # Market selection in 2 columns
        col1, col2 = st.columns(2)
        
        with col1:
            # Market selection
            market = st.selectbox(
                "🌍 Target Market",
                WORLD_COUNTRIES,
                index=0,  # Default to Vietnam
                help="Select your target market from countries worldwide"
            )
        
        with col2:
            # Research mode
            research_mode = st.radio(
                "🔬 Research Mode",
                ["Complete Analysis", "Quick Test (5 questions)"],
                help="Complete: Full research report | Quick: Test with 5 questions only"
            )
        
        # Analysis Level Selection - NEW FEATURE
        st.markdown("---")
        analysis_level = st.radio(
            "📊 Analysis Level",
            ["Layer 3 Analysis", "Layer 4 Analysis"],
            index=1,  # Default to Layer 4 (current behavior)
            help="Layer 3: Comprehensive analysis per category (Political, Economic...) | Layer 4: Detailed analysis per main question with sub-questions"
        )
        
        # Explanation for each level
        if analysis_level == "Layer 3 Analysis":
            st.info("🎯 **Layer 3 Mode**: Groups all main questions by category (Political, Economic...) to create comprehensive analysis for entire categories. Sub-questions are not used.")
        else:
            st.info("🔍 **Layer 4 Mode**: Detailed analysis of each main question with its corresponding sub-questions.")
        
        # Submit button
        submitted = st.form_submit_button("🚀 Start Research", use_container_width=True)
    
    if submitted:
        if not research_topic:
            st.error("Please enter a research topic!")
            return
        
        # Validate uploaded file again before processing with dynamic cost calculation
        final_question_count = 0
        final_sub_question_count = 0
        final_estimated_cost = 0
        
        if uploaded_file:
            temp_path = os.path.join('output', 'temp_validation.xlsx')
            os.makedirs('output', exist_ok=True)
            
            try:
                with open(temp_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
                
                is_valid, error_msg, question_count, sub_question_count, _ = validate_excel_template(temp_path)
                
                if not is_valid:
                    st.error(error_msg)
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                    return
                
                # Use dynamic cost calculation based on analysis level
                final_estimated_cost, layer3_cost, layer4_cost, category_count = calculate_dynamic_cost(
                    question_count, sub_question_count, analysis_level
                )
                final_question_count = question_count
                final_sub_question_count = sub_question_count
                
                # Clean up validation file
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    
            except Exception as e:
                st.error(f"❌ File validation error: {str(e)}")
                return
        else:
            # Use default template question count - get actual values from default template
            default_template_path = 'input/market research template.xlsx'
            if os.path.exists(default_template_path):
                try:
                    is_valid, _, question_count, sub_question_count, _ = validate_excel_template(default_template_path)
                    if is_valid:
                        # Use dynamic cost calculation for default template too
                        final_estimated_cost, layer3_cost, layer4_cost, category_count = calculate_dynamic_cost(
                            question_count, sub_question_count, analysis_level
                        )
                        final_question_count = question_count
                        final_sub_question_count = sub_question_count
                    else:
                        # Fallback to estimated values
                        final_question_count = 22
                        final_sub_question_count = 5
                        final_estimated_cost, layer3_cost, layer4_cost, category_count = calculate_dynamic_cost(
                            final_question_count, final_sub_question_count, analysis_level
                        )
                except Exception:
                    # Fallback to estimated values
                    final_question_count = 22
                    final_sub_question_count = 5
                    final_estimated_cost, layer3_cost, layer4_cost, category_count = calculate_dynamic_cost(
                        final_question_count, final_sub_question_count, analysis_level
                    )
            else:
                # Fallback to estimated values
                final_question_count = 22
                final_sub_question_count = 5
                final_estimated_cost, layer3_cost, layer4_cost, category_count = calculate_dynamic_cost(
                    final_question_count, final_sub_question_count, analysis_level
                )
        
        # Handle file processing
        if uploaded_file:
            # Save uploaded file
            excel_path = os.path.join('output', 'uploaded_framework.xlsx')
            with open(excel_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
        else:
            # Use default
            excel_path = 'input/market research template.xlsx'
        
        # Convert Excel to structured data
        try:
            converter = ExcelToStructuredJSON()
            json_path = 'output/market_research_structured.json'
            success = converter.convert_excel_to_json(excel_path, json_path)
            
            if success:
                with open(json_path, 'r', encoding='utf-8') as f:
                    structured_data = json.load(f)
            else:
                st.error("❌ Failed to convert Excel file!")
                return
                
        except Exception as e:
            st.error(f"❌ Excel conversion error: {str(e)}")
            return
        
        # Store research params in session state
        st.session_state['pending_research'] = {
            'structured_data': structured_data,
            'industry': research_topic,
            'market': market,
            'testing_mode': research_mode == "Quick Test (5 questions)",
            'analysis_level': analysis_level,
            'api_key': OPENAI_API_KEY,
            'question_count': final_question_count,
            'estimated_cost': final_estimated_cost
        }
        
        # Cost confirmation for expensive research
        # For Layer 3: also warn if high question count even with low cost
        should_warn = (
            final_estimated_cost > 1.0 or 
            (analysis_level == "Layer 3 Analysis" and final_question_count > 50)
        )
        
        if should_warn:
            # Show cost warning dialog
            st.session_state['show_cost_dialog'] = True
            st.rerun()
        else:
            # Proceed directly for low-cost research
            proceed_with_research()
    
    # Show completed research results if available - NOW AT BOTTOM
    if st.session_state.get('research_completed', False):
        st.success("🎉 Research completed successfully!")
        
        # Word download button
        word_file = st.session_state.get('word_file')
        if word_file and os.path.exists(word_file):
            with open(word_file, "rb") as file:
                st.download_button(
                    label="📄 Download Word Report",
                    data=file.read(),
                    file_name=os.path.basename(word_file),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    
    # Check if research should start (after dialog confirmation)
    if st.session_state.get('start_research', False):
        st.session_state['start_research'] = False
        # Clean up dialog state
        if 'dialog_processed' in st.session_state:
            del st.session_state['dialog_processed']
        
        research_params = st.session_state.get('pending_research', {})
        if research_params:
            structured_data = research_params['structured_data']
            industry = research_params['industry']
            market = research_params['market']
            testing_mode = research_params['testing_mode']
            analysis_level = research_params['analysis_level']
            api_key = research_params['api_key']
            
            # Progress tracking - consistent with default file experience
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Execute research with clean progress tracking
                status_text.text("🤖 Initializing AI research engine...")
                progress_bar.progress(10)
                
                researcher = OpenAIMarketResearch(api_key=api_key, industry=industry, market=market)
                
                status_text.text("🔍 Starting market research analysis...")
                progress_bar.progress(20)
                
                results = researcher.run_layer3_research(
                    structured_data=structured_data,
                    topic=industry,
                    testing_mode=testing_mode,
                    analysis_level=analysis_level
                )
                
                status_text.text("💾 Saving research results...")
                progress_bar.progress(80)
                
                # Add API key to results for export function
                results['api_key'] = api_key
                
                # Save results
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                # Determine market suffix for filename
                market_suffix = ""
                if market and market.lower() != "global market":
                    if "việt nam" in market.lower() or "vietnam" in market.lower():
                        market_suffix = "_VN"
                    else:
                        # Clean market name for filename
                        clean_market = re.sub(r'[^\w\s-]', '', market).strip()
                        clean_market = re.sub(r'\s+', '_', clean_market)
                        market_suffix = f"_{clean_market}"
                
                industry_clean = re.sub(r'[^\w\s-]', '', industry)
                industry_clean = re.sub(r'\s+', '_', industry_clean)
                
                filename = f"layer3_research_{industry_clean}{market_suffix}_{timestamp}.json"
                output_path = os.path.join("output", filename)
                
                # Ensure output directory exists
                os.makedirs("output", exist_ok=True)
                
                # Save with proper encoding
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(results, f, ensure_ascii=False, indent=2)
                
                progress_bar.progress(90)
                status_text.text("📄 Generating comprehensive Word report...")
                
                # Generate Word report automatically with Vietnamese filename
                word_file = create_comprehensive_word_report(output_path, use_vietnamese_filename=True)
                
                progress_bar.progress(100)
                status_text.text("✅ Research completed successfully!")
                
                st.session_state['research_completed'] = True
                st.session_state['research_results'] = results
                st.session_state['results_file'] = output_path
                st.session_state['word_file'] = word_file
                
                # Clear pending research
                for key in ['pending_research', 'start_research']:
                    if key in st.session_state:
                        del st.session_state[key]
                
                st.success("✅ Research completed! Download files are ready.")
                st.rerun()
                
            except Exception as e:
                progress_bar.progress(0)
                status_text.text("❌ Research failed")
                st.error(f"❌ Research failed: {str(e)}")
                
                # Clear session state on error
                for key in ['pending_research', 'start_research']:
                    if key in st.session_state:
                        del st.session_state[key]

# Cost confirmation dialog
if st.session_state.get('show_cost_dialog', False) and not st.session_state.get('dialog_processed', False):
    # Get pending research data
    pending = st.session_state.get('pending_research', {})
    cost = pending.get('estimated_cost', 0)
    question_count = pending.get('question_count', 0)
    
    @st.dialog("⚠️ High Cost Warning")
    def show_cost_confirmation():
        # Get analysis level from pending research
        analysis_level = pending.get('analysis_level', 'Layer 4 Analysis')
        
        if analysis_level == "Layer 3 Analysis" and cost <= 1.0:
            # Special case for Layer 3 with high question count but low cost
            st.warning("⚠️ **HIGH QUESTION COUNT DETECTED**")
            st.markdown(f"""
            ### 📊 Analysis Details: **{question_count} questions** 
            
            **💰 Estimated Cost: ${cost:.2f}** (Low cost due to Layer 3 efficiency)
            
            **ℹ️ Layer 3 Analysis Info:**
            - Groups questions by category for comprehensive analysis
            - Much more cost-effective than Layer 4 individual analysis
            - {question_count} questions → ~8 category analyses
            
            **💡 Alternative Options:**
            - 🔄 Choose "Quick Test (5 questions)" → Only $0.15
            - ✂️ Split template into smaller parts  
            - ✅ **Proceed** - Layer 3 is efficient for large templates!
            """)
        else:
            # Original high cost warning
            st.error("🚨 **HIGH RESEARCH COST DETECTED**")
            st.markdown(f"""
            ### 💰 Estimated Cost: **${cost:.2f}** 
            
            **📊 Details:**
            - Total Questions: {question_count}
            - Layer 3 + Layer 4 Fees: ${cost:.2f}
            
            **⚠️ This is higher than normal cost!**
            
            **💡 Alternative Options:**
            - 🔄 Choose "Quick Test (5 questions)" → Only $0.15
            - ✂️ Split template into smaller parts  
            - 📝 Reduce questions in Excel template
            """)
        
        st.markdown("---")
        
        # Confirmation buttons
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("❌ **Cancel**", use_container_width=True):
                # Clear everything and reset
                for key in ['show_cost_dialog', 'pending_research', 'dialog_processed']:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
        
        with col2:
            if st.button(f"✅ **Proceed - ${cost:.2f}**", use_container_width=True, type="primary"):
                # Mark dialog as processed and start research
                st.session_state['dialog_processed'] = True
                st.session_state['show_cost_dialog'] = False
                st.session_state['start_research'] = True
                st.rerun()
    
    # Show the dialog
    show_cost_confirmation()

def show_settings_page():
    """Settings and configuration page"""
    st.header("⚙️ Settings & Configuration")
    
    # API Configuration
    st.subheader("🔑 API Configuration")
    
    # Check current API key status
    current_key = OPENAI_API_KEY if OPENAI_API_KEY != "your-openai-api-key-here" else ""
    key_status = "✅ Configured" if current_key else "❌ Not configured"
    
    st.info(f"API Key Status: {key_status}")
    
    # For deployment, show instructions for Streamlit Cloud
    if not os.path.exists('config.py'):
        st.warning("🌐 **Deployment Mode Detected**")
        st.markdown("""
        **For Streamlit Cloud deployment:**
        1. Go to your app dashboard
        2. Click **Settings** → **Secrets**
        3. Add your API key:
        ```toml
        OPENAI_API_KEY = "your-actual-openai-api-key-here"
        ```
        4. Save and restart your app
        """)
    else:
        # Local development mode
        with st.form("api_settings"):
            api_key = st.text_input(
                "OpenAI API Key",
                value=current_key,
                type="password",
                help="Enter your OpenAI API key from https://platform.openai.com/api-keys"
            )
            
            model_choice = st.selectbox(
                "AI Model",
                ["gpt-3.5-turbo", "gpt-4"],
                help="gpt-3.5-turbo is recommended for cost-effectiveness"
            )
            
            if st.form_submit_button("💾 Save API Settings"):
                if api_key:
                    # Update config file (local development only)
                    try:
                        with open('config.py', 'r', encoding='utf-8') as f:
                            content = f.read()
                        
                        # Replace API key
                        import re
                        new_content = re.sub(
                            r'return ".*"  # Fallback',
                            f'return "{api_key}"  # Fallback',
                            content
                        )
                        
                        with open('config.py', 'w', encoding='utf-8') as f:
                            f.write(new_content)
                        
                        st.success("✅ API key saved successfully!")
                        st.info("Please restart the application to apply changes.")
                        
                    except Exception as e:
                        st.error(f"Error saving API key: {str(e)}")
                else:
                    st.error("Please enter a valid API key!")
    
if __name__ == "__main__":
    main() 