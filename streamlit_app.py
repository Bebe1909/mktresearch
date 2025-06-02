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

# Import our core modules
from config import OPENAI_API_KEY
from excel_to_structured_json import ExcelToStructuredJSON
from openai_market_research import OpenAIMarketResearch
from export_comprehensive_report import create_comprehensive_word_report, find_latest_research_file

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
    
    # Sidebar navigation
    with st.sidebar:
        selected = option_menu(
            menu_title="Navigation",
            options=["🏠 Home", "🚀 New Research", "📊 Analytics", "📄 Export Reports", "⚙️ Settings"],
            icons=["house", "rocket", "bar-chart", "file-earmark-text", "gear"],
            menu_icon="cast",
            default_index=0,
        )
    
    # Main content based on selection
    if selected == "🏠 Home":
        show_home_page()
    elif selected == "🚀 New Research":
        show_research_page()
    elif selected == "📊 Analytics":
        show_analytics_page()
    elif selected == "📄 Export Reports":
        show_export_page()
    elif selected == "⚙️ Settings":
        show_settings_page()

def show_home_page():
    """Home dashboard"""
    st.header("📊 Dashboard Overview")
    
    # Quick stats
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div class="metric-card">
            <h3>🎯</h3>
            <h2>Fast</h2>
            <p>15-25 minutes</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="metric-card">
            <h3>💰</h3>
            <h2>Cost-effective</h2>
            <p>$0.15-0.25 per report</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="metric-card">
            <h3>🤖</h3>
            <h2>AI-powered</h2>
            <p>OpenAI GPT-3.5</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown("""
        <div class="metric-card">
            <h3>📄</h3>
            <h2>Professional</h2>
            <p>Word reports</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Recent activity
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("🚀 Quick Start")
        st.markdown("""
        <div class="feature-box">
            <h4>1. 📁 Upload Excel Framework</h4>
            <p>Upload your research framework Excel file or use our default template</p>
        </div>
        <div class="feature-box">
            <h4>2. 🎯 Enter Research Topic</h4>
            <p>Specify your industry and market focus (e.g., "Electric vehicles in Vietnam")</p>
        </div>
        <div class="feature-box">
            <h4>3. ⚡ Generate Research</h4>
            <p>AI automatically creates comprehensive analysis with Layer 3 & 4 insights</p>
        </div>
        <div class="feature-box">
            <h4>4. 📄 Export Report</h4>
            <p>Download professional Word document ready for presentation</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.subheader("📈 Benefits")
        st.info("✅ Save 10-20 hours of manual research")
        st.info("✅ Consistent, professional format")
        st.info("✅ Data-driven insights")
        st.info("✅ Ready for stakeholder presentations")
        st.info("✅ No technical skills required")

def show_research_page():
    """New research creation page"""
    st.header("🚀 Create New Market Research")
    
    # Check API key
    if not OPENAI_API_KEY or OPENAI_API_KEY == "your-openai-api-key-here":
        st.error("⚠️ Please configure your OpenAI API key in Settings first!")
        return
    
    # Research form
    with st.form("research_form"):
        st.subheader("📝 Research Configuration")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Topic input
            research_topic = st.text_input(
                "🎯 Research Topic",
                placeholder="e.g., Electric vehicles, Food delivery, E-commerce",
                help="Enter the industry or market you want to research"
            )
            
            # Market selection
            market = st.selectbox(
                "🌍 Target Market",
                ["Việt Nam", "Southeast Asia", "Asia-Pacific"],
                help="Select your target market"
            )
        
        with col2:
            # Excel file upload
            uploaded_file = st.file_uploader(
                "📁 Upload Excel Framework",
                type=['xlsx'],
                help="Upload your research framework or leave empty to use default template"
            )
            
            # Research mode
            research_mode = st.radio(
                "🔬 Research Mode",
                ["Complete Analysis", "Quick Test (5 questions)"],
                help="Complete: Full research report | Quick: Test with 5 questions only"
            )
        
        # Advanced options
        with st.expander("🔧 Advanced Options"):
            custom_purpose = st.text_area(
                "Custom Research Purpose",
                placeholder="Leave empty to use default purpose",
                help="Override default research objectives"
            )
        
        # Submit button
        submitted = st.form_submit_button("🚀 Start Research", use_container_width=True)
    
    if submitted:
        if not research_topic:
            st.error("Please enter a research topic!")
            return
        
        # Run research
        run_research(research_topic, market, uploaded_file, research_mode == "Quick Test (5 questions)", custom_purpose)

def run_research(topic, market, uploaded_file, is_test_mode, custom_purpose):
    """Execute the research process"""
    
    # Create output directory
    os.makedirs('output', exist_ok=True)
    
    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Handle Excel file
        status_text.text("📁 Processing Excel framework...")
        progress_bar.progress(10)
        
        if uploaded_file:
            # Save uploaded file
            excel_path = os.path.join('output', 'uploaded_framework.xlsx')
            with open(excel_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
        else:
            # Use default
            excel_path = 'input/Research Framework.xlsx'
            if not os.path.exists(excel_path):
                st.error("Default framework not found! Please upload an Excel file.")
                return
        
        # Step 2: Convert Excel to JSON
        status_text.text("🔄 Converting Excel to structured data...")
        progress_bar.progress(20)
        
        converter = ExcelToStructuredJSON()
        json_path = 'output/market_research_structured.json'
        success = converter.convert_excel_to_json(excel_path, json_path, custom_purpose)
        
        if not success:
            st.error("Failed to convert Excel file!")
            return
        
        # Load structured data
        with open(json_path, 'r', encoding='utf-8') as f:
            structured_data = json.load(f)
        
        # Step 3: Initialize research
        status_text.text("🤖 Initializing AI research engine...")
        progress_bar.progress(30)
        
        researcher = OpenAIMarketResearch(
            api_key=OPENAI_API_KEY,
            industry=topic,
            market=market
        )
        
        # Step 4: Run research
        status_text.text("🔍 Conducting market research analysis...")
        progress_bar.progress(40)
        
        research_results = researcher.run_layer3_research(
            structured_data=structured_data,
            topic=topic,
            testing_mode=is_test_mode
        )
        
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
        
        # Save enhanced results
        progress_bar.progress(80)
        status_text.text("💾 Saving enhanced research results...")
        
        with open(results_path, 'w', encoding='utf-8') as f:
            json.dump(research_results, f, ensure_ascii=False, indent=2)
        
        # Step 6: Create Word report
        status_text.text("📄 Generating Word report...")
        progress_bar.progress(90)
        
        word_filename = f"Báo_cáo_nghiên_cứu_thị_trường_{topic.replace(' ', '_')}_{timestamp}.docx"
        word_path = os.path.join('output', word_filename)
        
        create_comprehensive_word_report(results_path, word_path)
        
        # Complete
        progress_bar.progress(100)
        status_text.text("✅ Research completed successfully!")
        
        # Show results
        st.success("🎉 Research completed successfully!")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("📊 Total Questions", total_main_questions)
        with col2:
            st.metric("🎯 Enhanced Reports", enhanced_count)
        with col3:
            estimated_cost = total_main_questions * 0.02 + enhanced_count * 0.05
            st.metric("💰 Estimated Cost", f"${estimated_cost:.2f}")
        
        # Download buttons
        st.markdown("### 📥 Download Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            with open(word_path, 'rb') as f:
                st.download_button(
                    label="📄 Download Word Report",
                    data=f.read(),
                    file_name=word_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        
        with col2:
            with open(results_path, 'r', encoding='utf-8') as f:
                st.download_button(
                    label="📊 Download JSON Data",
                    data=f.read(),
                    file_name=results_filename,
                    mime="application/json",
                    use_container_width=True
                )
        
        # Store in session state for analytics
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
        st.error(f"❌ Error during research: {str(e)}")
        st.exception(e)

def show_analytics_page():
    """Analytics and insights page"""
    st.header("📊 Research Analytics")
    
    if 'last_research' not in st.session_state:
        st.info("No recent research data available. Create a new research first!")
        return
    
    research_data = st.session_state['last_research']
    
    # Load research results for analysis
    try:
        with open(research_data['results_path'], 'r', encoding='utf-8') as f:
            results = json.load(f)
        
        # Research overview
        st.subheader("🎯 Research Overview")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📋 Topic", research_data['topic'])
        with col2:
            st.metric("🌍 Market", research_data['market'])
        with col3:
            st.metric("📅 Date", research_data['timestamp'][:8])
        with col4:
            st.metric("💰 Cost", f"${research_data['cost']:.2f}")
        
        # Analysis breakdown
        st.subheader("📈 Analysis Breakdown")
        
        # Collect data for visualization
        layers_data = []
        categories_data = []
        
        for layer in results.get('research_results', []):
            layer_name = layer.get('layer_name', '')
            layer_questions = 0
            layer_enhanced = 0
            
            for category in layer.get('categories', []):
                category_name = category.get('category_name', '')
                questions = category.get('questions', [])
                category_questions = len(questions)
                category_enhanced = sum(1 for q in questions if q.get('layer4_comprehensive_report'))
                
                layer_questions += category_questions
                layer_enhanced += category_enhanced
                
                categories_data.append({
                    'Layer': layer_name,
                    'Category': category_name,
                    'Questions': category_questions,
                    'Enhanced': category_enhanced
                })
            
            layers_data.append({
                'Layer': layer_name,
                'Total Questions': layer_questions,
                'Enhanced Reports': layer_enhanced,
                'Enhancement Rate': f"{(layer_enhanced/layer_questions*100):.1f}%" if layer_questions > 0 else "0%"
            })
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            # Layer breakdown chart
            if layers_data:
                df_layers = pd.DataFrame(layers_data)
                fig = px.bar(
                    df_layers, 
                    x='Layer', 
                    y=['Total Questions', 'Enhanced Reports'],
                    title="Questions & Enhanced Reports by Layer",
                    barmode='group'
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Enhancement rate pie chart
            if layers_data:
                total_questions = sum(item['Total Questions'] for item in layers_data)
                total_enhanced = sum(item['Enhanced Reports'] for item in layers_data)
                
                fig = go.Figure(data=[go.Pie(
                    labels=['Enhanced Reports', 'Standard Reports'],
                    values=[total_enhanced, total_questions - total_enhanced],
                    title="Enhancement Coverage"
                )])
                st.plotly_chart(fig, use_container_width=True)
        
        # Detailed breakdown
        st.subheader("📋 Detailed Breakdown")
        
        if categories_data:
            df_categories = pd.DataFrame(categories_data)
            st.dataframe(df_categories, use_container_width=True)
        
    except Exception as e:
        st.error(f"Error loading analytics: {str(e)}")

def show_export_page():
    """Export and download page"""
    st.header("📄 Export Reports")
    
    # Find available research files
    output_dir = 'output'
    if os.path.exists(output_dir):
        research_files = [f for f in os.listdir(output_dir) if f.startswith('layer3_research_') and f.endswith('.json')]
        word_files = [f for f in os.listdir(output_dir) if f.endswith('.docx')]
    else:
        research_files = []
        word_files = []
    
    if not research_files:
        st.info("No research results found. Create a new research first!")
        return
    
    # Recent reports
    st.subheader("📊 Available Reports")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📋 Research Data Files")
        for file in sorted(research_files, reverse=True)[:5]:
            file_path = os.path.join(output_dir, file)
            file_size = os.path.getsize(file_path) / 1024  # KB
            timestamp = file.replace('layer3_research_', '').replace('.json', '')
            
            col_a, col_b = st.columns([3, 1])
            with col_a:
                st.text(f"📊 {timestamp}")
                st.caption(f"Size: {file_size:.1f} KB")
            with col_b:
                with open(file_path, 'r', encoding='utf-8') as f:
                    st.download_button(
                        "⬇️",
                        data=f.read(),
                        file_name=file,
                        mime="application/json",
                        key=f"json_{file}"
                    )
    
    with col2:
        st.markdown("#### 📄 Word Reports")
        for file in sorted(word_files, reverse=True)[:5]:
            file_path = os.path.join(output_dir, file)
            file_size = os.path.getsize(file_path) / 1024  # KB
            
            col_a, col_b = st.columns([3, 1])
            with col_a:
                st.text(f"📄 {file[:30]}...")
                st.caption(f"Size: {file_size:.1f} KB")
            with col_b:
                with open(file_path, 'rb') as f:
                    st.download_button(
                        "⬇️",
                        data=f.read(),
                        file_name=file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"word_{file}"
                    )
    
    # Export latest research
    st.markdown("---")
    st.subheader("🚀 Quick Export Latest Research")
    
    if st.button("📄 Export Latest Research to Word", use_container_width=True):
        latest_file = find_latest_research_file()
        if latest_file:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            word_filename = f"Latest_Market_Research_Report_{timestamp}.docx"
            word_path = os.path.join(output_dir, word_filename)
            
            with st.spinner("Creating Word report..."):
                create_comprehensive_word_report(latest_file, word_path)
            
            st.success("✅ Word report created!")
            
            with open(word_path, 'rb') as f:
                st.download_button(
                    label="📥 Download Word Report",
                    data=f.read(),
                    file_name=word_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.error("No research files found!")

def show_settings_page():
    """Settings and configuration page"""
    st.header("⚙️ Settings & Configuration")
    
    # API Configuration
    st.subheader("🔑 API Configuration")
    
    current_key = OPENAI_API_KEY if OPENAI_API_KEY != "your-openai-api-key-here" else ""
    
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
                # Update config file
                try:
                    with open('config.py', 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Replace API key
                    import re
                    new_content = re.sub(
                        r'OPENAI_API_KEY\s*=\s*["\'][^"\']*["\']',
                        f'OPENAI_API_KEY = "{api_key}"',
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
        
        # Check if output directory exists
        output_exists = os.path.exists('output')
        st.text(f"Output Directory: {'✅' if output_exists else '❌'}")
        
        # Check default framework
        framework_exists = os.path.exists('input/Research Framework.xlsx')
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