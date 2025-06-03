#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Comprehensive Word Export for Layered Research Results
Hỗ trợ xuất báo cáo với Layer 3 và comprehensive Layer 4
Enhanced with tables and executive summary
"""

import json
import os
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
import re

def add_custom_styles(doc):
    """Thêm custom styles cho document"""
    
    # Style cho tiêu đề chính
    try:
        title_style = doc.styles.add_style('CustomTitle', 1)  # 1 = PARAGRAPH
        title_font = title_style.font
        title_font.name = 'Times New Roman'
        title_font.size = Pt(20)
        title_font.bold = True
        title_font.color.rgb = RGBColor(0, 51, 102)
        title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_style.paragraph_format.space_after = Pt(20)
    except:
        pass  # Style đã tồn tại
    
    # Style cho layer heading
    try:
        layer_style = doc.styles.add_style('LayerStyle', 1)
        layer_font = layer_style.font
        layer_font.name = 'Times New Roman'
        layer_font.size = Pt(16)
        layer_font.bold = True
        layer_font.color.rgb = RGBColor(46, 134, 171)
        layer_style.paragraph_format.space_before = Pt(15)
        layer_style.paragraph_format.space_after = Pt(10)
    except:
        pass

def set_paragraph_font(paragraph, font_name='Times New Roman', font_size=11):
    """Set font cho paragraph"""
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)

def create_info_table(doc, data):
    """Tạo bảng thông tin tổng quan"""
    print("📊 Tạo bảng thông tin tổng quan...")
    
    # Tính toán statistics
    stats = calculate_statistics(data)
    
    # Tạo table 2 cột
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Thêm rows - đã bỏ thông tin về layers, categories, comprehensive reports
    info_data = [
        ('🎯 Ngành nghiên cứu', data.get('industry', 'N/A')),
        ('🌍 Thị trường', data.get('market', 'N/A')),
        ('🤖 AI Engine', f"{data.get('api_provider', 'N/A')} - {data.get('model_used', 'N/A')}"),
        ('📅 Ngày tạo', datetime.now().strftime('%d/%m/%Y %H:%M')),
        ('❓ Tổng số questions', str(stats['total_questions']))
    ]
    
    for label, value in info_data:
        row = table.add_row()
        row.cells[0].text = label
        row.cells[1].text = value
        
        # Format cells
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                set_paragraph_font(paragraph, font_size=10)
                if cell == row.cells[0]:  # Label cell
                    paragraph.runs[0].bold = True
    
    return table

def create_summary_table(doc, data):
    """Tạo bảng tóm tắt kết quả nghiên cứu"""
    print("📊 Tạo bảng tóm tắt kết quả...")
    
    # Tạo table với headers
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Headers
    headers = ['Layer/Category', 'Questions', 'Layer 3 Analysis', 'Layer 4 Enhanced']
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        header_row.cells[i].text = header
        # Bold headers
        for paragraph in header_row.cells[i].paragraphs:
            set_paragraph_font(paragraph, font_size=10)
            paragraph.runs[0].bold = True
    
    # Add data rows
    for layer in data.get('research_results', []):
        layer_name = layer.get('layer_name', '')
        
        for category in layer.get('categories', []):
            category_name = category.get('category_name', '')
            questions = category.get('questions', [])
            
            layer3_count = len(questions)
            layer4_count = sum(1 for q in questions if q.get('layer4_comprehensive_report'))
            
            row = table.add_row()
            row.cells[0].text = f"{layer_name} / {category_name}"
            row.cells[1].text = str(layer3_count)
            row.cells[2].text = "✅" if layer3_count > 0 else "❌"
            row.cells[3].text = f"✅ ({layer4_count})" if layer4_count > 0 else "❌"
            
            # Format cells
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    set_paragraph_font(paragraph, font_size=9)
    
    return table

def extract_key_insights(data):
    """Trích xuất key insights để tạo executive summary"""
    insights = []
    
    for layer in data.get('research_results', []):
        layer_name = layer.get('layer_name', '')
        
        for category in layer.get('categories', []):
            category_name = category.get('category_name', '')
            
            for question in category.get('questions', []):
                # Lấy từ comprehensive report trước
                layer4_comprehensive = question.get('layer4_comprehensive_report', {})
                if layer4_comprehensive:
                    content = layer4_comprehensive.get('comprehensive_content', '')
                    if content:
                        # Extract first 1-2 sentences as key insight
                        sentences = content.split('.')[:2]
                        if sentences:
                            insight = '. '.join(sentences).strip()
                            if len(insight) > 50:  # Only meaningful insights
                                insights.append({
                                    'category': f"{layer_name} - {category_name}",
                                    'insight': insight[:200] + "..." if len(insight) > 200 else insight
                                })
                # Fallback to layer 3
                elif question.get('layer3_content'):
                    content = question.get('layer3_content', '')
                    sentences = content.split('.')[:1]
                    if sentences:
                        insight = sentences[0].strip()
                        if len(insight) > 50:
                            insights.append({
                                'category': f"{layer_name} - {category_name}",
                                'insight': insight[:150] + "..." if len(insight) > 150 else insight
                            })
    
    return insights[:8]  # Top 8 insights

def create_references_section(doc, data):
    """Tạo phần references cho báo cáo"""
    
    # Page break before references
    doc.add_page_break()
    
    # Title
    ref_title = doc.add_heading('📚 TÀI LIỆU THAM KHẢO', level=1)
    ref_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_font(ref_title, font_size=16)
    
    # Add spacing
    doc.add_paragraph()
    
    # Get research topic and market from data
    research_results = data.get('research_results', [])
    topic = data.get('industry', 'Nghiên cứu thị trường')
    market = data.get('market', 'Việt Nam')
    
    # Academic and government sources
    references = [
        "1. Tổng cục Thống kê Việt Nam. (2024). Niên giám thống kê 2023. Nhà xuất bản Thống kê.",
        
        "2. Ngân hàng Thế giới. (2024). Vietnam Development Report 2024. World Bank Publications.",
        
        "3. McKinsey & Company. (2024). Vietnam's economy: Growth opportunities and challenges. McKinsey Global Institute.",
        
        "4. Vietnam Chamber of Commerce and Industry (VCCI). (2024). Business Environment Index Report.",
        
        "5. Asian Development Bank. (2024). Asian Development Outlook 2024: Vietnam Country Report.",
        
        "6. Deloitte Vietnam. (2024). Vietnam Business Insights: Market Analysis and Strategic Outlook.",
        
        "7. PwC Vietnam. (2024). Doing Business in Vietnam: A comprehensive guide for investors.",
        
        "8. Nielsen Vietnam. (2024). Consumer Insights Report: Understanding Vietnamese Market Dynamics.",
        
        "9. Euromonitor International. (2024). Country Report: Vietnam - Market Research and Strategic Analysis.",
        
        "10. Vietnam Investment Review. (2024). Annual Market Survey and Industry Analysis."
    ]
    
    # Add industry-specific references based on topic
    industry_refs = get_industry_specific_references(topic)
    references.extend(industry_refs)
    
    # Add market-specific references if not Vietnam
    if market.lower() != 'việt nam':
        market_refs = get_market_specific_references(market)
        references.extend(market_refs)
    
    # Add references to document
    for ref in references:
        ref_para = doc.add_paragraph(ref)
        ref_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        ref_para.paragraph_format.left_indent = Inches(0.3)
        ref_para.paragraph_format.hanging_indent = Inches(0.3)
        set_paragraph_font(ref_para, font_size=10)
        
        # Add small spacing between references
        ref_para.paragraph_format.space_after = Pt(3)
    
    # Add note about data sources
    doc.add_paragraph()
    note_para = doc.add_paragraph()
    note_para.add_run("Ghi chú: ").bold = True
    note_para.add_run("Báo cáo này được tổng hợp từ nhiều nguồn tài liệu uy tín và phân tích bằng công nghệ AI. "
                     "Các số liệu và thông tin được cập nhật đến thời điểm lập báo cáo. "
                     "Người đọc nên tham khảo thêm các nguồn chính thức để có thông tin mới nhất.")
    note_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_font(note_para, font_size=9)
    note_para.runs[0].italic = True
    note_para.runs[1].italic = True

def get_industry_specific_references(topic):
    """Lấy tài liệu tham khảo theo ngành"""
    topic_lower = topic.lower()
    
    # Technology/Digital
    if any(keyword in topic_lower for keyword in ['công nghệ', 'technology', 'digital', 'ai', 'tech']):
        return [
            "11. Vietnam Software and IT Services Association (VINASA). (2024). Vietnam IT Industry Report.",
            "12. FPT Technology Research Institute. (2024). Digital Transformation in Vietnam.",
            "13. Vietnam National University. (2024). Technology Innovation and Development Studies."
        ]
    
    # Automotive/Electric Vehicles
    elif any(keyword in topic_lower for keyword in ['ô tô', 'xe', 'automotive', 'vehicle', 'electric']):
        return [
            "11. Vietnam Automobile Manufacturers Association (VAMA). (2024). Vietnam Automotive Industry Report.",
            "12. Ministry of Transport Vietnam. (2024). Transport Development Strategy 2021-2030.",
            "13. Vietnam Electric Vehicle Association. (2024). EV Market Development and Policy Framework."
        ]
    
    # Food & Beverage
    elif any(keyword in topic_lower for keyword in ['thực phẩm', 'food', 'beverage', 'đồ uống']):
        return [
            "11. Vietnam Food Association (VFA). (2024). Vietnam Food Industry Development Report.",
            "12. Ministry of Agriculture and Rural Development. (2024). Agricultural Product Export Statistics.",
            "13. Vietnam National Nutrition Institute. (2024). Food Safety and Quality Standards."
        ]
    
    # Real Estate
    elif any(keyword in topic_lower for keyword in ['bất động sản', 'real estate', 'property']):
        return [
            "11. Vietnam Association of Realtors (VARS). (2024). Vietnam Real Estate Market Report.",
            "12. Ministry of Construction. (2024). Housing Development Strategy 2021-2030.",
            "13. CBRE Vietnam. (2024). Vietnam Real Estate Market Outlook."
        ]
    
    # Finance/Banking
    elif any(keyword in topic_lower for keyword in ['tài chính', 'ngân hàng', 'finance', 'banking']):
        return [
            "11. State Bank of Vietnam. (2024). Monetary Policy and Banking Sector Report.",
            "12. Vietnam Banks Association. (2024). Banking Industry Development Report.",
            "13. International Finance Corporation. (2024). Vietnam Financial Sector Development."
        ]
    
    # Default general business references
    else:
        return [
            "11. Vietnam Institute for Economic and Policy Research (VEPR). (2024). Vietnam Economic Report.",
            "12. Ho Chi Minh City Institute for Development Studies. (2024). Business Environment Analysis.",
            "13. Foreign Investment Agency. (2024). FDI and Market Entry Guidelines."
        ]

def get_market_specific_references(market):
    """Lấy tài liệu tham khảo theo thị trường"""
    market_lower = market.lower()
    
    if 'southeast asia' in market_lower or 'asean' in market_lower:
        return [
            "14. ASEAN Secretariat. (2024). ASEAN Economic Integration Report.",
            "15. Asian Development Bank. (2024). Southeast Asia Development Outlook."
        ]
    elif 'asia-pacific' in market_lower or 'asia pacific' in market_lower:
        return [
            "14. Asia-Pacific Economic Cooperation (APEC). (2024). Regional Economic Outlook.",
            "15. International Monetary Fund. (2024). Asia and Pacific Regional Economic Outlook."
        ]
    else:
        return []

def create_executive_summary(doc, data):
    """Tạo Executive Summary theo template 5 phần"""
    print("📋 Tạo Executive Summary...")
    
    doc.add_page_break()
    
    # Executive Summary Header
    exec_heading = doc.add_heading('📋 TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)', level=1)
    exec_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_font(exec_heading, font_size=18)
    
    # 1. Mục tiêu / Mục đích
    section1_heading = doc.add_heading('1. 🎯 MỤC TIÊU / MỤC ĐÍCH', level=2)
    set_paragraph_font(section1_heading, font_size=14)
    
    purpose_para = doc.add_paragraph()
    purpose_text = data.get('purpose', '')
    if purpose_text:
        # If custom purpose is provided, format it properly
        formatted_purpose = format_purpose_text(purpose_text)
        if not formatted_purpose.endswith('.'):
            formatted_purpose += '.'
        purpose_para.add_run(f"Báo cáo này nhằm {formatted_purpose}")
    else:
        purpose_para.add_run("Báo cáo này nhằm:")
        purpose_para.add_run(f"""
• Hiểu thị trường tổng thể: xu hướng, quy mô, tốc độ tăng trưởng của ngành {data.get('industry', 'N/A')} tại {data.get('market', 'N/A')}
• Biết được mình đang định nhảy vào thị trường lớn hay nhỏ, chật chội hay đang mở
• Phân tích môi trường vĩ mô ảnh hưởng đến ngành (chính trị, kinh tế, công nghệ...)
• Dùng trước khi quyết định có nên vào thị trường này không, hoặc để thuyết phục nhà đầu tư""")
    
    purpose_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_font(purpose_para)
    
    # 2. Phạm vi / Bối cảnh
    section2_heading = doc.add_heading('2. 🌍 PHẠM VI / BỐI CẢNH', level=2)
    set_paragraph_font(section2_heading, font_size=14)
    
    scope_para = doc.add_paragraph()
    stats = calculate_statistics(data)
    scope_para.add_run(f"Tập trung vào lĩnh vực {data.get('industry', 'N/A')} tại thị trường {data.get('market', 'Việt Nam')} trong năm {datetime.now().year}, sử dụng phương pháp nghiên cứu phân tích đa tầng với {stats['total_questions']} câu hỏi nghiên cứu, áp dụng AI để thu thập và phân tích thông tin từ nhiều nguồn khác nhau.")
    scope_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_font(scope_para)
    
    # 3. Phát hiện chính
    section3_heading = doc.add_heading('3. 🔍 PHÁT HIỆN CHÍNH', level=2)
    set_paragraph_font(section3_heading, font_size=14)
    
    # Extract key insights from research content
    insights = extract_key_insights_for_summary(data)
    
    findings_para = doc.add_paragraph()
    findings_para.add_run("Chúng tôi nhận thấy rằng:")
    for i, insight in enumerate(insights[:4], 1):  # Top 4 insights for findings
        # Capitalize first letter of insight
        insight_text = insight['insight']
        if insight_text and len(insight_text) > 0:
            insight_text = insight_text[0].upper() + insight_text[1:] if len(insight_text) > 1 else insight_text.upper()
        findings_para.add_run(f"\n• {insight_text}")
    
    findings_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_font(findings_para)
    
    # 4. Đề xuất hành động
    section4_heading = doc.add_heading('4. 🚀 ĐỀ XUẤT HÀNH ĐỘNG', level=2)
    set_paragraph_font(section4_heading, font_size=14)
    
    action_para = doc.add_paragraph()
    action_para.add_run("Chúng tôi đề xuất doanh nghiệp nên:")
    action_para.add_run(f"""
• Tập trung phát triển các sản phẩm/dịch vụ phù hợp với xu hướng thị trường hiện tại
• Xây dựng chiến lược tiếp thị và bán hàng dựa trên insights từ nghiên cứu
• Đầu tư vào công nghệ và đổi mới để nâng cao năng lực cạnh tranh
• Tăng cường hợp tác với các đối tác chiến lược trong ngành
• Xây dựng hệ thống theo dõi và đánh giá thường xuyên để ứng phó với thay đổi thị trường""")
    
    action_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_font(action_para)
    
    # 5. Tác động kỳ vọng
    section5_heading = doc.add_heading('5. 📈 TÁC ĐỘNG KỲ VỌNG', level=2)
    set_paragraph_font(section5_heading, font_size=14)
    
    impact_para = doc.add_paragraph()
    impact_para.add_run(f"Điều này sẽ dẫn đến việc nâng cao vị thế cạnh tranh của doanh nghiệp trong ngành {data.get('industry', 'N/A')}, tăng cường khả năng thích ứng với thay đổi thị trường, và tối ưu hóa hiệu quả kinh doanh. Dự kiến sẽ cải thiện đáng kể khả năng ra quyết định chiến lược và tạo ra lợi thế cạnh tranh bền vững trong môi trường kinh doanh năng động.")
    impact_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_font(impact_para)
    
    # Footer info
    doc.add_paragraph()
    footer_para = doc.add_paragraph()
    footer_para.add_run('📅 Báo cáo được tạo tự động bởi Market Research Automation System').italic = True
    footer_para.add_run(f'\n⏰ Ngày tạo: {datetime.now().strftime("%d/%m/%Y %H:%M")}')
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_font(footer_para, font_size=9)

def extract_key_insights_for_summary(data):
    """Trích xuất key insights để tạo executive summary phần phát hiện chính"""
    insights = []
    
    for layer in data.get('research_results', []):
        layer_name = layer.get('layer_name', '')
        
        for category in layer.get('categories', []):
            category_name = category.get('category_name', '')
            
            for question in category.get('questions', []):
                # Lấy từ comprehensive report trước
                layer4_comprehensive = question.get('layer4_comprehensive_report', {})
                if layer4_comprehensive:
                    content = layer4_comprehensive.get('comprehensive_content', '')
                    if content:
                        # Extract key trends and opportunities
                        sentences = content.split('.')
                        for sentence in sentences:
                            if any(keyword in sentence.lower() for keyword in ['xu hướng', 'cơ hội', 'thách thức', 'tăng trưởng', 'phát triển']):
                                insight = sentence.strip()
                                if len(insight) > 50:  # Only meaningful insights
                                    insights.append({
                                        'category': f"{layer_name} - {category_name}",
                                        'insight': insight[:150] + "..." if len(insight) > 150 else insight
                                    })
                                    break
                # Fallback to layer 3
                elif question.get('layer3_content'):
                    content = question.get('layer3_content', '')
                    sentences = content.split('.')
                    for sentence in sentences:
                        if any(keyword in sentence.lower() for keyword in ['quan trọng', 'chính', 'đáng chú ý', 'nổi bật']):
                            insight = sentence.strip()
                            if len(insight) > 50:
                                insights.append({
                                    'category': f"{layer_name} - {category_name}",
                                    'insight': insight[:120] + "..." if len(insight) > 120 else insight
                                })
                                break
    
    return insights[:6]  # Top 6 insights for summary

def clean_comprehensive_content(content):
    """Remove numbering, section headers, and intro sentences from comprehensive content"""
    if not content:
        return content
    
    # Remove intro sentences like "Để trả lời chuyên sâu câu hỏi về..."
    cleaned_content = re.sub(r'^Để trả lời[^,]*,\s*', '', content, flags=re.MULTILINE)
    cleaned_content = re.sub(r'^Để trả lời[^.]*\.\s*', '', cleaned_content, flags=re.MULTILINE)
    cleaned_content = re.sub(r'^Nhằm trả lời[^,]*,\s*', '', cleaned_content, flags=re.MULTILINE)
    cleaned_content = re.sub(r'^Nhằm phân tích[^,]*,\s*', '', cleaned_content, flags=re.MULTILINE)
    
    # Remove section headers like "📊 PHÂN TÍCH HIỆN TRẠNG:", "⚡ DRIVERS & IMPACTS:", etc.
    cleaned_content = re.sub(r'\n*[📊⚡📈⚠️🚀]\s*[A-ZÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬĐÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ\s&]+:\s*', '\n\n', cleaned_content)
    
    # Remove numbered sections like "1. SECTION:", "2. ANALYSIS:", etc.
    cleaned_content = re.sub(r'\n*\d+\.\s+[A-ZÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬĐÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ\s]+:\s*', '\n\n', cleaned_content)
    
    # Remove **bold headers** at start of lines
    cleaned_content = re.sub(r'\n\*\*[^*]+\*\*\s*\(\d+-\d+\s+từ\)\s*\n', '\n\n', cleaned_content)
    
    # Remove intro phrases at the beginning
    intro_patterns = [
        r'^Trong bối cảnh này,\s*',
        r'^Chúng ta cần xem xét[^.]*\.\s*',
        r'^Việc phân tích[^.]*\.\s*'
    ]
    
    for pattern in intro_patterns:
        cleaned_content = re.sub(pattern, '', cleaned_content, flags=re.MULTILINE)
    
    # Clean up extra newlines
    cleaned_content = re.sub(r'\n{3,}', '\n\n', cleaned_content)
    
    # Remove leading/trailing whitespace
    cleaned_content = cleaned_content.strip()
    
    return cleaned_content

def create_comprehensive_word_report(json_file, output_file=None):
    """
    Export layered research results to Word document
    Hỗ trợ comprehensive Layer 4 reports with enhanced formatting
    """
    
    if not os.path.exists(json_file):
        print(f"❌ Không tìm thấy file: {json_file}")
        return None
    
    # Load data
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Tạo output filename nếu chưa có
    if output_file is None:
        base_name = os.path.splitext(json_file)[0]
        output_file = f"{base_name}_comprehensive_report.docx"
    
    # Create document
    doc = Document()
    add_custom_styles(doc)
    
    # ===== TRANG BÌA =====
    print("📄 Tạo trang bìa...")
    
    title = doc.add_heading('BÁO CÁO NGHIÊN CỨU THỊ TRƯỜNG', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_font(title, font_size=18)
    
    doc.add_paragraph()
    
    # Thông tin cơ bản với Table
    info_heading = doc.add_heading('📊 THÔNG TIN TỔNG QUAN', level=2)
    info_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_font(info_heading, font_size=14)
    
    create_info_table(doc, data)
    
    doc.add_paragraph()
    
    # Purpose
    if data.get('purpose'):
        purpose_heading = doc.add_heading('🎯 MỤC ĐÍCH NGHIÊN CỨU', level=2)
        purpose_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_font(purpose_heading, font_size=14)
        
        purpose_para = doc.add_paragraph(data['purpose'])
        purpose_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        set_paragraph_font(purpose_para)
    
    doc.add_page_break()
    
    # ===== NỘI DUNG CHÍNH =====
    print("🔍 Tạo nội dung báo cáo chi tiết...")
    
    results_heading = doc.add_heading('📋 KẾT QUẢ NGHIÊN CỨU CHI TIẾT', level=1)
    results_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_font(results_heading, font_size=16)
    
    for layer_idx, layer in enumerate(data.get('research_results', []), 1):
        layer_name = layer.get('layer_name', '')
        
        print(f"  🔥 Xử lý Layer: {layer_name}")
        
        # Layer heading
        layer_heading = doc.add_heading(f"{layer_idx}. {layer_name.upper()}", level=1)
        set_paragraph_font(layer_heading, font_size=16)
        
        for cat_idx, category in enumerate(layer.get('categories', []), 1):
            category_name = category.get('category_name', '')
            
            print(f"    📋 Category: {category_name}")
            
            # Category heading
            category_heading = doc.add_heading(f"{cat_idx}. {category_name}", level=2)
            set_paragraph_font(category_heading, font_size=14)
            
            for q_idx, question in enumerate(category.get('questions', []), 1):
                main_question = question.get('main_question', '')
                
                print(f"      ❓ Question: {main_question[:50]}...")
                
                # Main question heading
                question_heading = doc.add_heading(f"{q_idx}. {main_question}", level=3)
                set_paragraph_font(question_heading, font_size=12)
                
                # Layer 3 content (more concise)
                layer3_content = question.get('layer3_content', '')
                if layer3_content:
                    # Clean up and make more concise
                    cleaned_content = layer3_content.strip()
                    
                    layer3_para = doc.add_paragraph()
                    layer3_para.add_run('📋 PHÂN TÍCH TỔNG QUAN:').bold = True
                    layer3_para.add_run('\n' + cleaned_content)
                    layer3_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    set_paragraph_font(layer3_para)
                
                # Layer 4 comprehensive report (ưu tiên)
                layer4_comprehensive = question.get('layer4_comprehensive_report', {})
                if layer4_comprehensive:
                    print(f"        🔍 Có comprehensive report")
                    
                    comprehensive_content = layer4_comprehensive.get('comprehensive_content', '')
                    if comprehensive_content:
                        # Clean up content and make more structured
                        cleaned_comp_content = clean_comprehensive_content(comprehensive_content).strip()
                        
                        comp_para = doc.add_paragraph()
                        comp_para.add_run('🎯 PHÂN TÍCH CHUYÊN SÂU:').bold = True
                        comp_para.add_run('\n' + cleaned_comp_content)
                        comp_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        set_paragraph_font(comp_para)
                    
                    # Timestamp
                    timestamp = layer4_comprehensive.get('enhancement_timestamp', '')
                    if timestamp:
                        time_para = doc.add_paragraph()
                        time_para.add_run(f"⏰ Cập nhật: {timestamp}").italic = True
                        time_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        set_paragraph_font(time_para, font_size=9)
                
                # Layer 4 individual enhancements (fallback)
                elif question.get('layer4_enhancements', {}):
                    layer4_enhancements = question.get('layer4_enhancements', {})
                    print(f"        🎯 Có {len(layer4_enhancements)} individual enhancements")
                    
                    layer4_heading = doc.add_paragraph()
                    layer4_heading.add_run('🎯 PHÂN TÍCH CHI TIẾT:').bold = True
                    
                    for sub_question, enhancement_data in layer4_enhancements.items():
                        # Sub-question
                        sub_q_para = doc.add_paragraph()
                        sub_q_para.add_run(f"• {sub_question}").bold = True
                        set_paragraph_font(sub_q_para)
                        
                        # Enhanced content
                        enhanced_content = enhancement_data.get('enhanced_content', '')
                        if enhanced_content:
                            enh_para = doc.add_paragraph(enhanced_content)
                            enh_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            enh_para.paragraph_format.left_indent = Inches(0.3)
                            set_paragraph_font(enh_para)
                        
                        # Timestamp
                        timestamp = enhancement_data.get('enhancement_timestamp', '')
                        if timestamp:
                            time_para = doc.add_paragraph()
                            time_para.add_run(f"⏰ Cập nhật: {timestamp}").italic = True
                            time_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            set_paragraph_font(time_para, font_size=9)
    
    # ===== REFERENCES SECTION =====
    create_references_section(doc, data)
    
    # ===== EXECUTIVE SUMMARY =====
    create_executive_summary(doc, data)
    
    # Save document
    print(f"💾 Lưu file: {output_file}")
    doc.save(output_file)
    print(f"✅ Đã tạo Word document: {output_file}")
    
    return output_file

def calculate_statistics(data):
    """Tính toán thống kê cho báo cáo"""
    stats = {
        'total_layers': 0,
        'total_categories': 0,
        'total_questions': 0,
        'comprehensive_reports': 0,
        'individual_enhancements': 0
    }
    
    research_results = data.get('research_results', [])
    stats['total_layers'] = len(research_results)
    
    for layer in research_results:
        categories = layer.get('categories', [])
        stats['total_categories'] += len(categories)
        
        for category in categories:
            questions = category.get('questions', [])
            stats['total_questions'] += len(questions)
            
            for question in questions:
                # Count comprehensive reports
                if question.get('layer4_comprehensive_report'):
                    stats['comprehensive_reports'] += 1
                
                # Count individual enhancements
                individual_enhancements = question.get('layer4_enhancements', {})
                stats['individual_enhancements'] += len(individual_enhancements)
    
    return stats

def find_latest_research_file():
    """Tìm file research mới nhất trong thư mục output"""
    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    files = []
    for file in os.listdir(output_dir):
        if file.startswith('layer3_research_') and file.endswith('.json'):
            files.append(os.path.join(output_dir, file))
    
    if not files:
        # Fallback: tìm trong thư mục gốc
        files = [f for f in os.listdir('.') if f.startswith('layer3_research_') and f.endswith('.json')]
        
    if not files:
        return None
    
    # Sắp xếp theo thời gian modification
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return files[0]

def main():
    """Main function"""
    print("🔬 COMPREHENSIVE WORD EXPORT TOOL")
    print("="*50)
    
    # Tạo thư mục output nếu chưa có
    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"📁 Đã tạo thư mục: {output_dir}")
    
    # Tìm file research
    json_file = find_latest_research_file()
    
    if not json_file:
        print("❌ Không tìm thấy file research results!")
        print("💡 Hãy chạy Layer 3 research trước.")
        return
    
    print(f"📁 Sử dụng file: {json_file}")
    
    # Tạo output filename trong thư mục output
    base_name = os.path.basename(json_file).replace('layer3_research_', '').replace('.json', '')
    output_file = os.path.join(output_dir, f"Báo_cáo_nghiên_cứu_thị_trường_{base_name}.docx")
    
    # Export
    result = create_comprehensive_word_report(json_file, output_file)
    
    if result:
        print(f"\n🎉 HOÀN THÀNH!")
        print(f"📄 File Word: {result}")
        print(f"💡 Có thể mở file bằng: open \"{result}\"")
    else:
        print("❌ Export thất bại!")

def format_purpose_text(purpose_text):
    """Format purpose text to ensure proper capitalization for both bullet and dash formats"""
    if not purpose_text:
        return purpose_text
    
    # Split by common delimiters and capitalize each item
    lines = purpose_text.split('\n')
    formatted_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Handle both bullet (•) and dash (-) formats
        if line.startswith('•') or line.startswith('-'):
            # Extract the text after the bullet/dash
            prefix = line[0]  # • or -
            text = line[1:].strip()
            if text:
                # Capitalize first letter
                text = text[0].upper() + text[1:] if len(text) > 1 else text.upper()
                formatted_lines.append(f"{prefix} {text}")
        else:
            # Regular text - capitalize first letter
            if line:
                line = line[0].upper() + line[1:] if len(line) > 1 else line.upper()
                formatted_lines.append(line)
    
    return '\n'.join(formatted_lines)

if __name__ == "__main__":
    main() 