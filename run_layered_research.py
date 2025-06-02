#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Layered Market Research Tool
Hỗ trợ nghiên cứu thị trường với cấp độ Layer 3 và enhancement lên Layer 4
Updated to use OpenAI GPT-3.5-turbo
"""

import json
import os
from typing import List
from openai_market_research import OpenAIMarketResearch
from config import OPENAI_API_KEY, RESEARCH_CONFIG, MODEL
from excel_to_structured_json import convert_market_research_to_json

def run_complete_research_workflow():
    """Chạy workflow hoàn chỉnh: Excel → JSON → Research → Ready for Export"""
    print("\n🚀 COMPLETE RESEARCH WORKFLOW")
    print("="*60)
    
    # Step 1: Get Excel file path
    print("📁 BƯỚC 1: CHỌN FILE EXCEL")
    print("-" * 30)
    
    default_excel = "input/Research Framework.xlsx"
    print(f"💡 File mặc định: {default_excel}")
    
    excel_path = input("\n📝 Nhập đường dẫn file Excel (Enter để dùng file mặc định): ").strip()
    
    if not excel_path:
        excel_path = default_excel
        print(f"✅ Sử dụng file mặc định: {excel_path}")
    else:
        print(f"✅ Sử dụng file: {excel_path}")
    
    # Check if Excel file exists
    if not os.path.exists(excel_path):
        print(f"❌ Không tìm thấy file: {excel_path}")
        if excel_path == default_excel:
            print("💡 Hãy đảm bảo file 'Research Framework.xlsx' nằm trong folder 'input/'")
        return False
    
    # Step 2: Convert Excel to JSON
    print(f"\n📊 BƯỚC 2: CONVERT EXCEL → JSON")
    print("-" * 30)
    
    # Temporarily update the excel path in the converter
    import excel_to_structured_json
    original_convert = excel_to_structured_json.convert_market_research_to_json
    
    def convert_with_custom_path():
        # Temporarily modify the excel file path
        temp_excel_file = excel_to_structured_json.convert_market_research_to_json.__code__.co_filename
        # Read the file content and replace the path
        with open("excel_to_structured_json.py", 'r') as f:
            content = f.read()
        
        # For simplicity, we'll just call the original function but inform user about path
        print(f"🔄 Converting: {excel_path}")
        
        if excel_path != default_excel:
            print("⚠️ Đang sử dụng file tùy chỉnh. Hãy đảm bảo file có cấu trúc giống file mặc định.")
            
        # If using custom path, temporarily copy to expected location
        if excel_path != default_excel:
            import shutil
            os.makedirs("input", exist_ok=True)
            shutil.copy2(excel_path, default_excel)
            print(f"📋 Đã copy file vào: {default_excel}")
        
        return original_convert()
    
    result = convert_with_custom_path()
    
    if not result:
        print("❌ Convert Excel thất bại!")
        return False
    
    print(f"✅ Convert thành công!")
    print(f"📊 Đã tạo {len(result.get('layers', []))} layers")
    total_questions = 0
    for layer in result.get('layers', []):
        for category in layer.get('categories', []):
            total_questions += len(category.get('questions', []))
    print(f"❓ Tổng {total_questions} questions")
    
    # Step 3: Get research topic
    print(f"\n🎯 BƯỚC 3: THIẾT LẬP NGHIÊN CỨU")
    print("-" * 30)
    
    topic = input("📝 Nhập chủ đề nghiên cứu (ví dụ: 'Công nghiệp ô tô Việt Nam'): ").strip()
    if not topic:
        print("❌ Cần nhập chủ đề!")
        return False
    
    print(f"\n🎯 Topic: {topic}")
    print(f"🌍 Market: Việt Nam")
    print(f"💰 Estimated cost: ~$0.15-0.25")
    print(f"🤖 Auto Layer 4: Questions có sub-questions sẽ tự động được enhance")
    print(f"⏱️ Estimated time: ~15-25 minutes")
    
    if input("\n▶️ Bắt đầu nghiên cứu? (Y/n): ").strip().lower() not in ['', 'y', 'yes']:
        return False
    
    # Step 4: Run Research
    print(f"\n🔍 BƯỚC 4: CHẠY NGHIÊN CỨU")
    print("-" * 30)
    
    # Ensure output directory exists
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    
    # Load structured data
    structured_file = os.path.join(output_dir, 'market_research_structured.json')
    with open(structured_file, 'r', encoding='utf-8') as f:
        structured_data = json.load(f)

    # Process research (always full mode)
    researcher = OpenAIMarketResearch(
        api_key=OPENAI_API_KEY,
        industry=topic,
        market="Việt Nam",
        model=MODEL
    )
    
    print("🔄 Đang chạy Layer 3 Research...")
    result = researcher.run_layer3_research(structured_data, topic, testing_mode=False)
    
    if not result:
        print("❌ Nghiên cứu thất bại!")
        return False
    
    # Save research results
    output_file = os.path.join(output_dir, f"layer3_research_{topic.replace(' ', '_')}_openai.json")
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"✅ Hoàn thành Layer 3!")
    
    # Auto-generate comprehensive Layer 4 reports
    print(f"\n🚀 TỰ ĐỘNG TẠO BÁO CÁO LAYER 4...")
    
    comprehensive_count = 0
    total_questions = 0
    
    for layer in result.get('research_results', []):
        for category in layer.get('categories', []):
            for question in category.get('questions', []):
                total_questions += 1
                sub_questions = question.get('sub_questions', [])
                
                if sub_questions:  # Có sub-questions → tạo comprehensive report
                    print(f"  🔄 Tạo comprehensive report: {question.get('main_question', '')[:50]}...")
                    
                    try:
                        researcher.add_layer4_comprehensive_enhancement(
                            results_file=output_file,
                            layer_name=layer.get('layer_name'),
                            category_name=category.get('category_name'),
                            main_question=question.get('main_question')
                        )
                        comprehensive_count += 1
                        print(f"    ✅ Hoàn thành!")
                    except Exception as e:
                        print(f"    ❌ Lỗi: {e}")
    
    print(f"\n🎉 HOÀN THÀNH TẤT CẢ!")
    print(f"📊 Tổng questions: {total_questions}")
    print(f"🔍 Questions chỉ Layer 3: {total_questions - comprehensive_count}")
    print(f"🔍 Comprehensive reports (Layer 4): {comprehensive_count}")
    print(f"📁 File kết quả: {output_file}")
    print(f"\n💡 Sẵn sàng export Word document!")
    
    return True

def export_word_document():
    """Export Word document từ research results"""
    print("\n📄 EXPORT WORD DOCUMENT")
    print("="*50)
    
    # Find latest research file
    output_dir = 'output'
    files = []
    
    if os.path.exists(output_dir):
        for file in os.listdir(output_dir):
            if file.startswith('layer3_research_') and file.endswith('.json'):
                files.append(os.path.join(output_dir, file))
    
    if not files:
        print("❌ Không tìm thấy file research results!")
        print("💡 Hãy chạy Complete Research Workflow trước (Option 1).")
        return False
    
    # Get the latest file
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    latest_file = files[0]
    
    print(f"📁 Sử dụng file: {latest_file}")
    
    # Show basic info
    with open(latest_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"🎯 Topic: {data.get('industry', 'N/A')}")
    print(f"🌍 Market: {data.get('market', 'N/A')}")
    
    total_questions = 0
    comprehensive_reports = 0
    
    for layer in data.get('research_results', []):
        for category in layer.get('categories', []):
            for question in category.get('questions', []):
                total_questions += 1
                if question.get('layer4_comprehensive_report'):
                    comprehensive_reports += 1
    
    print(f"📋 Total questions: {total_questions}")
    print(f"🔍 Comprehensive reports: {comprehensive_reports}")
    
    if input("\n▶️ Tiếp tục export? (Y/n): ").strip().lower() not in ['', 'y', 'yes']:
        return False
    
    print(f"\n🔄 Đang export Word document...")
    
    # Run export script
    import subprocess
    try:
        result = subprocess.run(["python", "export_comprehensive_report.py"], 
                              capture_output=True, text=True, check=True)
        print("✅ Export thành công!")
        print(f"📄 File Word đã được tạo trong folder output/")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Lỗi export: {e}")
        print(f"Output: {e.stdout}")
        print(f"Error: {e.stderr}")
        return False

def main():
    """Hàm main với menu đơn giản 3 options"""
    
    print("="*80)
    print("🔬 MARKET RESEARCH AUTOMATION SYSTEM")
    print("="*80)
    print(f"🤖 API: OpenAI {MODEL}")
    print(f"🚀 Mode: Complete Workflow (Excel → Research → Word)")
    print("="*80)
    
    while True:
        print("\n📋 MENU:")
        print("1. 🚀 Complete Research Workflow (Excel → JSON → Research → Ready)")
        print("2. 📄 Export Word Document")
        print("3. 👋 Thoát")
        
        choice = input("\nLựa chọn của bạn (1-3): ").strip()
        
        if choice == "1":
            run_complete_research_workflow()
            
        elif choice == "2":
            export_word_document()
            
        elif choice == "3":
            print("👋 Cảm ơn bạn đã sử dụng Market Research Tool!")
            break
            
        else:
            print("❌ Lựa chọn không hợp lệ. Vui lòng chọn 1-3.")

if __name__ == "__main__":
    main()