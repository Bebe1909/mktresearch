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
            options=["🚀 New Research", "📊 Results & Export", "⚙️ Settings"],
            icons=["rocket", "bar-chart-fill", "gear"],
            menu_icon="cast",
            default_index=0,
        )
    
    # Main content based on selection
    if selected == "🚀 New Research":
        show_research_page()
    elif selected == "📊 Results & Export":
        show_results_and_export_page()
    elif selected == "⚙️ Settings":
        show_settings_page()

def validate_excel_template(file_path):
    """
    Validate uploaded Excel file format
    Returns: (is_valid, error_message, question_count, estimated_cost)
    """
    try:
        # Try to read the Excel file
        excel_file = pd.ExcelFile(file_path)
        
        # Check if 'template' sheet exists
        if 'template' not in excel_file.sheet_names:
            return False, "❌ Sheet 'template' không tìm thấy! Vui lòng sử dụng template đúng format hoặc đổi tên sheet thành 'template'.", 0, 0
        
        # Read template sheet
        df = pd.read_excel(file_path, sheet_name="template")
        
        # Check if dataframe is not empty
        if df.empty:
            return False, "❌ Sheet 'template' trống! Vui lòng thêm dữ liệu vào template.", 0, 0
        
        # Check for basic required structure
        if df.shape[1] < 3:  # At least 3 columns
            return False, "❌ Template cần ít nhất 3 cột! Vui lòng sử dụng template chuẩn.", 0, 0
        
        # Look for purpose row
        purpose_found = False
        for idx, row in df.iterrows():
            if any("Mục đích" in str(cell) for cell in row if pd.notna(cell)):
                purpose_found = True
                break
        
        if not purpose_found:
            return False, "❌ Không tìm thấy dòng 'Mục đích của Market Research'! Vui lòng sử dụng template chuẩn.", 0, 0
        
        # Count questions for cost estimation
        question_count = count_questions_in_excel(df)
        estimated_cost = estimate_research_cost(question_count)
        
        return True, "✅ Template hợp lệ!", question_count, estimated_cost
        
    except Exception as e:
        return False, f"❌ Lỗi đọc file Excel: {str(e)}. Vui lòng kiểm tra file có bị hỏng không.", 0, 0

def count_questions_in_excel(df):
    """
    Count total number of questions (Layer 3) in Excel template
    Returns: int - number of questions
    """
    question_count = 0
    
    # Find layer header row
    layer_header_row = None
    for idx, row in df.iterrows():
        if any("Layer 1" in str(cell) for cell in row if pd.notna(cell)):
            layer_header_row = idx
            break
    
    if layer_header_row is None:
        return 0
    
    # Count questions in Layer 3 column (usually column 3)
    try:
        for idx in range(layer_header_row + 1, len(df)):
            row = df.iloc[idx]
            
            # Skip empty rows
            if row.isna().all():
                continue
            
            # Check Layer 3 column (index 3, assuming 0-based indexing)
            if len(row) > 3:
                layer3_val = row.iloc[3] if pd.notna(row.iloc[3]) else None
                if layer3_val and str(layer3_val).strip() and str(layer3_val) != 'None':
                    question_count += 1
    except Exception as e:
        print(f"Error counting questions: {e}")
        return 0
    
    return question_count

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

def display_cost_warning(question_count, estimated_cost):
    """Display cost warning based on question count"""
    if question_count <= 5:
        st.info(f"💰 Chi phí ước tính: ${estimated_cost} ({question_count} câu hỏi)")
    elif question_count <= 25:
        st.info(f"💰 Chi phí ước tính: ${estimated_cost} ({question_count} câu hỏi) - Mức tiêu chuẩn")
    elif question_count <= 50:
        st.warning(f"⚠️ Chi phí ước tính: ${estimated_cost} ({question_count} câu hỏi) - Mức cao")
    elif question_count <= 100:
        st.error(f"🚨 Chi phí ước tính: ${estimated_cost} ({question_count} câu hỏi) - Mức rất cao!")
    else:
        st.error(f"🔴 CẢNH BÁO: Chi phí ước tính ${estimated_cost} ({question_count} câu hỏi) - Rất tốn kém!")
        st.error("💡 Khuyến nghị: Chia nhỏ thành nhiều nghiên cứu riêng biệt")

def show_research_page():
    """New research creation page - simplified and streamlined"""
    st.header("🚀 Create Market Research Report")
    
    # Check API key
    if not OPENAI_API_KEY or OPENAI_API_KEY == "your-openai-api-key-here":
        st.error("⚠️ Please configure your OpenAI API key in Settings first!")
        return
    
    # Template download section
    st.markdown("""
    <div class="download-section">
        <h4>📁 Excel Template</h4>
        <p>Download our template or upload your customized version</p>
    </div>
    """, unsafe_allow_html=True)
    
    template_path = 'input/market research template.xlsx'
    if os.path.exists(template_path):
        col1, col2 = st.columns([1, 3])
        with col1:
            with open(template_path, 'rb') as template_file:
                st.download_button(
                    label="📥 Download Template",
                    data=template_file.read(),
                    file_name="market_research_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        with col2:
            st.info("💡 **Template Structure:** Download → Customize → Upload back here")
    
    st.markdown("---")
    
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
        
        # Excel file upload with validation
        st.markdown("##### 📁 Excel Framework")
        uploaded_file = st.file_uploader(
            "Upload your research framework (optional)",
            type=['xlsx'],
            help="Upload customized template or leave empty to use default"
        )
        
        # Show file validation status
        file_validation_container = st.container()
        
        # Advanced options (collapsed by default)
        with st.expander("🔧 Advanced Options"):
            custom_purpose = st.text_area(
                "Custom Research Purpose",
                placeholder="Leave empty to use default purpose from template",
                help="Override default research objectives"
            )
        
        # Submit button
        submitted = st.form_submit_button("🚀 Start Research", use_container_width=True)
    
    # Validate uploaded file when selected
    if uploaded_file:
        with file_validation_container:
            # Save temp file for validation
            temp_path = os.path.join('output', 'temp_uploaded.xlsx')
            os.makedirs('output', exist_ok=True)
            
            try:
                with open(temp_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
                
                # Validate file
                is_valid, message, question_count, estimated_cost = validate_excel_template(temp_path)
                
                if is_valid:
                    st.success(message)
                    display_cost_warning(question_count, estimated_cost)
                else:
                    st.error(message)
                    
                # Clean up temp file
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    
            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
    
    if submitted:
        if not research_topic:
            st.error("Please enter a research topic!")
            return
        
        # Validate uploaded file again before processing
        final_question_count = 0
        final_estimated_cost = 0
        
        if uploaded_file:
            temp_path = os.path.join('output', 'temp_validation.xlsx')
            os.makedirs('output', exist_ok=True)
            
            try:
                with open(temp_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
                
                is_valid, error_msg, question_count, estimated_cost = validate_excel_template(temp_path)
                
                if not is_valid:
                    st.error(error_msg)
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                    return
                
                final_question_count = question_count
                final_estimated_cost = estimated_cost
                
                # Clean up validation file
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    
            except Exception as e:
                st.error(f"❌ File validation error: {str(e)}")
                return
        else:
            # Use default template question count
            # Estimate default template cost (around 20-25 questions)
            final_question_count = 22  # Default template estimate
            final_estimated_cost = estimate_research_cost(final_question_count)
        
        # Cost confirmation for expensive research
        if final_estimated_cost > 1.0:
            st.warning(f"🚨 CẢNH BÁO CHI PHÍ: ${final_estimated_cost} ({final_question_count} câu hỏi)")
            
            # Show confirmation message
            st.markdown("""
            **⚠️ Nghiên cứu này có chi phí cao!**
            
            💡 **Các lựa chọn khác:**
            - Chọn "Quick Test (5 questions)" để test với $0.15
            - Chia nhỏ template thành nhiều phần
            - Giảm số câu hỏi trong Excel template
            """)
            
            # Require explicit confirmation
            confirm_expensive = st.checkbox(f"✅ Tôi xác nhận tiếp tục với chi phí ${final_estimated_cost}")
            
            if not confirm_expensive:
                st.stop()  # Stop execution until user confirms
        
        # Clean market name (remove emoji and country code)
        clean_market = market.split(' ', 1)[-1] if ' ' in market else market
        
        # Debug info
        st.info(f"🔍 Processing: Topic='{research_topic}', Original Market='{market}', Clean Market='{clean_market}'")
        
        # Run research
        run_research(research_topic, clean_market, uploaded_file, research_mode == "Quick Test (5 questions)", custom_purpose)

def run_research(topic, market, uploaded_file, is_test_mode, custom_purpose):
    """Execute the research process with enhanced error handling"""
    
    # Create output directory
    os.makedirs('output', exist_ok=True)
    
    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Handle Excel file with validation
        status_text.text("📁 Processing Excel framework...")
        progress_bar.progress(10)
        
        if uploaded_file:
            # Save uploaded file
            excel_path = os.path.join('output', 'uploaded_framework.xlsx')
            with open(excel_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            
            # Double-check validation
            is_valid, error_msg, question_count, estimated_cost = validate_excel_template(excel_path)
            if not is_valid:
                st.error(f"File validation failed: {error_msg}")
                return
                
        else:
            # Use default
            excel_path = 'input/market research template.xlsx'
            if not os.path.exists(excel_path):
                st.error("❌ Default template not found! Please upload an Excel file or contact support.")
                return
        
        # Step 2: Convert Excel to JSON with error handling
        status_text.text("🔄 Converting Excel to structured data...")
        progress_bar.progress(20)
        
        try:
            converter = ExcelToStructuredJSON()
            json_path = 'output/market_research_structured.json'
            success = converter.convert_excel_to_json(excel_path, json_path, custom_purpose)
            
            if not success:
                st.error("❌ Failed to convert Excel file! Please check if your template follows the correct format.")
                return
                
        except Exception as e:
            st.error(f"❌ Excel conversion error: {str(e)}")
            st.info("💡 Try downloading a fresh template and make sure your file structure matches.")
            return
        
        # Load structured data
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                structured_data = json.load(f)
        except Exception as e:
            st.error(f"❌ Error loading converted data: {str(e)}")
            return
        
        # Step 3: Initialize research
        status_text.text("🤖 Initializing AI research engine...")
        progress_bar.progress(30)
        
        try:
            researcher = OpenAIMarketResearch(
                api_key=OPENAI_API_KEY,
                industry=topic,
                market=market
            )
        except Exception as e:
            st.error(f"❌ AI initialization error: {str(e)}")
            st.info("💡 Please check your API key in Settings.")
            return
        
        # Step 4: Run research
        status_text.text("🔍 Conducting market research analysis...")
        progress_bar.progress(40)
        
        try:
            research_results = researcher.run_layer3_research(
                structured_data=structured_data,
                topic=topic,
                testing_mode=is_test_mode
            )
        except Exception as e:
            st.error(f"❌ Research execution error: {str(e)}")
            st.info("💡 This might be an API quota or network issue. Please try again.")
            return
        
        # Step 5: Auto Layer 4 enhancement
        status_text.text("⚡ Creating comprehensive Layer 4 reports...")
        progress_bar.progress(60)
        
        # Save results
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        results_filename = f"layer3_research_{timestamp}.json"
        results_path = os.path.join('output', results_filename)
        
        with open(results_path, 'w', encoding='utf-8') as f:
            json.dump(research_results, f, ensure_ascii=False, indent=2)
        
        # Auto-enhance questions with sub-questions
        enhanced_count = 0
        total_main_questions = 0
        
        try:
            for layer in research_results.get('research_results', []):
                for category in layer.get('categories', []):
                    for question in category.get('questions', []):
                        total_main_questions += 1
                        sub_questions = question.get('sub_questions', [])
                        
                        if sub_questions and len(sub_questions) > 0:
                            # Has sub-questions -> create comprehensive Layer 4
                            progress_text = f"Creating comprehensive report {enhanced_count + 1}..."
                            status_text.text(f"🎯 {progress_text}")
                            
                            comprehensive_content = researcher.enhance_to_layer4_comprehensive(
                                research_results,
                                layer.get('layer_name'),
                                category.get('category_name'),
                                question.get('main_question')
                            )
                            
                            # Add to results
                            question['layer4_comprehensive_report'] = {
                                "comprehensive_content": comprehensive_content,
                                "enhancement_timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                                "sub_questions_integrated": sub_questions
                            }
                            
                            enhanced_count += 1
        except Exception as e:
            st.warning(f"⚠️ Layer 4 enhancement partially failed: {str(e)}")
            st.info("🔄 Continuing with available results...")
        
        # Save enhanced results
        progress_bar.progress(80)
        status_text.text("💾 Saving enhanced research results...")
        
        with open(results_path, 'w', encoding='utf-8') as f:
            json.dump(research_results, f, ensure_ascii=False, indent=2)
        
        # Step 6: Create Word report
        status_text.text("📄 Generating Word report...")
        progress_bar.progress(90)
        
        try:
            word_filename = f"Market_Research_{topic.replace(' ', '_')}_{timestamp}.docx"
            word_path = os.path.join('output', word_filename)
            
            create_comprehensive_word_report(results_path, word_path)
        except Exception as e:
            st.warning(f"⚠️ Word report generation failed: {str(e)}")
            st.info("📊 JSON results are still available for download.")
            word_path = None
        
        # Complete
        progress_bar.progress(100)
        status_text.text("✅ Research completed successfully!")
        
        # Show results
        st.success("🎉 Research completed successfully!")
        
        # Results summary
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("📊 Total Questions", total_main_questions)
        with col2:
            st.metric("🎯 Enhanced Reports", enhanced_count)
        with col3:
            estimated_cost = total_main_questions * 0.02 + enhanced_count * 0.05
            st.metric("💰 Estimated Cost", f"${estimated_cost:.2f}")
        
        # Download section
        st.markdown("### 📥 Download Results")
        
        col1, col2 = st.columns(2)
        
        # JSON download
        with col1:
            with open(results_path, 'r', encoding='utf-8') as f:
                st.download_button(
                    label="📊 Download JSON Data",
                    data=f.read(),
                    file_name=results_filename,
                    mime="application/json",
                    use_container_width=True
                )
        
        # Word download (if available)
        with col2:
            if word_path and os.path.exists(word_path):
                with open(word_path, 'rb') as f:
                    st.download_button(
                        label="📄 Download Word Report",
                        data=f.read(),
                        file_name=word_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            else:
                st.info("📄 Word report not available")
        
        # Store in session state for results page
        st.session_state['last_research'] = {
            'topic': topic,
            'market': market,
            'timestamp': timestamp,
            'total_questions': total_main_questions,
            'enhanced_reports': enhanced_count,
            'cost': estimated_cost,
            'results_path': results_path,
            'word_path': word_path
        }
        
    except Exception as e:
        st.error(f"❌ Unexpected error: {str(e)}")
        st.exception(e)
        st.info("💡 Please try again or contact support if the problem persists.")

def show_results_and_export_page():
    """Combined results analytics and export page"""
    st.header("📊 Research Results & Export")
    
    # Check for recent research
    if 'last_research' not in st.session_state:
        st.info("🔍 No recent research data. Create a new research first!")
        
        # Show available files for export
        st.subheader("📄 Available Research Files")
        show_available_files()
        return
    
    research_data = st.session_state['last_research']
    
    # Recent research overview
    st.subheader("🎯 Latest Research Overview")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📋 Topic", research_data['topic'])
    with col2:
        st.metric("🌍 Market", research_data['market'])
    with col3:
        st.metric("📊 Questions", research_data['total_questions'])
    with col4:
        st.metric("💰 Cost", f"${research_data['cost']:.2f}")
    
    # Download latest results
    st.subheader("📥 Download Latest Results")
    
    col1, col2 = st.columns(2)
    
    # JSON download
    with col1:
        if os.path.exists(research_data['results_path']):
            with open(research_data['results_path'], 'r', encoding='utf-8') as f:
                st.download_button(
                    label="📊 Download JSON Data",
                    data=f.read(),
                    file_name=f"research_{research_data['timestamp']}.json",
                    mime="application/json",
                    use_container_width=True
                )
    
    # Word download
    with col2:
        if research_data.get('word_path') and os.path.exists(research_data['word_path']):
            with open(research_data['word_path'], 'rb') as f:
                st.download_button(
                    label="📄 Download Word Report",
                    data=f.read(),
                    file_name=f"report_{research_data['timestamp']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.info("📄 Word report not available")
    
    st.markdown("---")
    
    # All available files
    st.subheader("📄 All Research Files")
    show_available_files()

def show_available_files():
    """Show all available research files for download"""
    output_dir = 'output'
    if os.path.exists(output_dir):
        research_files = [f for f in os.listdir(output_dir) if f.startswith('layer3_research_') and f.endswith('.json')]
        word_files = [f for f in os.listdir(output_dir) if f.endswith('.docx')]
    else:
        research_files = []
        word_files = []
    
    if not research_files and not word_files:
        st.info("📁 No research files found. Create your first research!")
        return
    
    # Show files in a clean format
    col1, col2 = st.columns(2)
    
    with col1:
        if research_files:
            st.markdown("#### 📊 JSON Data Files")
            for file in sorted(research_files, reverse=True)[:5]:
                file_path = os.path.join(output_dir, file)
                timestamp = file.replace('layer3_research_', '').replace('.json', '')
                
                with open(file_path, 'r', encoding='utf-8') as f:
                    st.download_button(
                        label=f"📊 {timestamp}",
                        data=f.read(),
                        file_name=file,
                        mime="application/json",
                        key=f"json_{file}"
                    )
    
    with col2:
        if word_files:
            st.markdown("#### 📄 Word Reports")
            for file in sorted(word_files, reverse=True)[:5]:
                file_path = os.path.join(output_dir, file)
                display_name = file[:30] + "..." if len(file) > 30 else file
                
                with open(file_path, 'rb') as f:
                    st.download_button(
                        label=f"📄 {display_name}",
                        data=f.read(),
                        file_name=file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"word_{file}"
                    )

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
    
    # System Information
    st.markdown("---")
    st.subheader("📊 System Information")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("🐍 Python Environment")
        st.text(f"Streamlit: {st.__version__}")
        
        # Deployment detection
        is_cloud = not os.path.exists('config.py')
        deployment_type = "☁️ Streamlit Cloud" if is_cloud else "💻 Local Development"
        st.text(f"Environment: {deployment_type}")
        
        # Check if output directory exists
        output_exists = os.path.exists('output')
        st.text(f"Output Directory: {'✅' if output_exists else '❌'}")
        
        # Check default framework
        framework_exists = os.path.exists('input/market research template.xlsx')
        st.text(f"Default Framework: {'✅' if framework_exists else '❌'}")
    
    with col2:
        st.info("💾 Storage")
        if os.path.exists('output'):
            files = os.listdir('output')
            json_files = len([f for f in files if f.endswith('.json')])
            word_files = len([f for f in files if f.endswith('.docx')])
            st.text(f"Research Files: {json_files}")
            st.text(f"Word Reports: {word_files}")
        else:
            st.text("No output directory found")
    
    # Clear data
    st.markdown("---")
    st.subheader("🗑️ Data Management")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🧹 Clear Session Data", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.success("Session data cleared!")
    
    with col2:
        if st.button("📁 Create Output Directory", use_container_width=True):
            os.makedirs('output', exist_ok=True)
            st.success("Output directory created!")

if __name__ == "__main__":
    main() 