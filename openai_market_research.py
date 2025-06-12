#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script gửi request tới OpenAI GPT API để nghiên cứu thị trường
Tác giả: AI Assistant
Ngày: 2024
"""

import json
import requests
import time
from typing import Dict, List, Any
import openai

class OpenAIMarketResearch:
    def __init__(self, api_key: str, industry: str = "Technology", market: str = "Việt Nam", model: str = "gpt-3.5-turbo"):
        """
        Khởi tạo class nghiên cứu thị trường với OpenAI GPT
        
        Args:
            api_key (str): API key của OpenAI
            industry (str): Ngành công nghiệp nghiên cứu
            market (str): Thị trường nghiên cứu
            model (str): Model GPT sử dụng
        """
        self.client = openai.OpenAI(api_key=api_key)
        self.industry = industry
        self.market = market
        self.model = model
        self.base_url = "https://api.openai.com/v1/chat/completions"
        self.delay_seconds = 3  # Increased delay from 1 to 3 seconds
        
        # Reference tracking system
        self.reference_tracker = {}
        self.tracked_sources = set()
        
        print(f"🤖 Initialized OpenAI Market Research for {industry} in {market}")
        print(f"📊 API Provider: OpenAI")
        print(f"🔧 Model: {model}")
        
    def call_openai_api(self, prompt: str, max_retries: int = 3) -> str:
        """Gửi request tới OpenAI API với retry logic và track references"""
        
        for attempt in range(max_retries):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    max_tokens=2000,
                    temperature=0.7
                )
                
                if response.choices and response.choices[0].message:
                    content = response.choices[0].message.content
                    
                    # Track references from the response
                    self.track_references_from_response(content)
                    
                    return content
                else:
                    return "Không có nội dung trong phản hồi API"
                    
            except Exception as e:
                if "429" in str(e) or "rate_limit" in str(e).lower():  # Rate limit error
                    wait_time = (2 ** attempt) * 5  # Exponential backoff: 5, 10, 20 seconds
                    print(f"⚠️ Rate limit hit - Waiting {wait_time}s before retry {attempt + 1}/{max_retries}")
                    time.sleep(wait_time)
                    if attempt == max_retries - 1:
                        print(f"❌ Max retries reached. Error: {e}")
                        return f"Rate limit error after {max_retries} retries: {e}"
                else:
                    print(f"❌ API Error: {e}")
                    return f"API Error: {e}"
        
        return "Failed after all retries"
    
    def track_references_from_response(self, content: str):
        """Extract and track references from AI response"""
        import re
        
        # Patterns to detect references and sources in AI responses
        reference_patterns = [
            # Organizations and institutions
            r'\b(?:Tổng cục Thống kê|General Statistics Office|GSO)\b',
            r'\b(?:Ngân hàng Thế giới|World Bank)\b',
            r'\b(?:IMF|International Monetary Fund)\b',
            r'\b(?:ADB|Asian Development Bank)\b',
            r'\b(?:McKinsey|Deloitte|PwC|KPMG|BCG)\b',
            r'\b(?:Nielsen|Euromonitor|Statista)\b',
            r'\b(?:VCCI|Vietnam Chamber of Commerce)\b',
            r'\b(?:Bộ (?:Kế hoạch|Tài chính|Công Thương|Y tế|Giáo dục))\b',
            
            # Vietnam specific
            r'\b(?:VAMA|Vietnam Automobile)\b',
            r'\b(?:VINASA|Vietnam Software)\b',
            r'\b(?:VFA|Vietnam Food Association)\b',
            r'\b(?:FPT|Viettel|VNPT)\b',
            r'\b(?:Ngân hàng Nhà nước|State Bank of Vietnam)\b',
            
            # Global sources
            r'\b(?:Bloomberg|Reuters|Financial Times)\b',
            r'\b(?:Forbes|Harvard Business Review|MIT)\b',
            r'\b(?:Gartner|IDC|Forrester)\b',
            
            # Government and regulatory
            r'\b(?:Ministry of|Bộ)\s+[A-Za-zÀ-ỹ\s]+\b',
            r'\b(?:Government of|Chính phủ)\s+[A-Za-zÀ-ỹ\s]+\b',
        ]
        
        # Extract references
        for pattern in reference_patterns:
            matches = re.finditer(pattern, content, re.IGNORECASE)
            for match in matches:
                source = match.group().strip()
                if source and len(source) > 3:  # Avoid very short matches
                    # Normalize source name
                    normalized_source = self.normalize_source_name(source)
                    
                    # Track frequency
                    if normalized_source in self.reference_tracker:
                        self.reference_tracker[normalized_source] += 1
                    else:
                        self.reference_tracker[normalized_source] = 1
                    
                    self.tracked_sources.add(normalized_source)
    
    def normalize_source_name(self, source: str) -> str:
        """Normalize source names for consistent tracking"""
        source = source.strip()
        
        # Mapping for common variations
        mappings = {
            'Tổng cục Thống kê': 'General Statistics Office (GSO)',
            'General Statistics Office': 'General Statistics Office (GSO)',
            'GSO': 'General Statistics Office (GSO)',
            'Ngân hàng Thế giới': 'World Bank',
            'World Bank': 'World Bank',
            'IMF': 'International Monetary Fund (IMF)',
            'International Monetary Fund': 'International Monetary Fund (IMF)',
            'McKinsey': 'McKinsey & Company',
            'Deloitte': 'Deloitte Consulting',
            'PwC': 'PricewaterhouseCoopers (PwC)',
            'KPMG': 'KPMG International',
            'Nielsen': 'Nielsen Holdings',
            'Euromonitor': 'Euromonitor International',
            'Statista': 'Statista GmbH',
            'VCCI': 'Vietnam Chamber of Commerce and Industry (VCCI)',
            'Vietnam Chamber of Commerce': 'Vietnam Chamber of Commerce and Industry (VCCI)',
        }
        
        return mappings.get(source, source)
    
    def get_top_references(self, limit: int = 10) -> list:
        """Get top references sorted by frequency"""
        sorted_refs = sorted(
            self.reference_tracker.items(), 
            key=lambda x: x[1], 
            reverse=True
        )
        return sorted_refs[:limit]

    def create_layer3_prompt_request(self, layer1: str, layer2: str, main_question: str, purpose: str) -> str:
        """Tạo request để lấy prompt Layer 3 (main question level) - DIRECT ANSWER FOCUSED"""
        template = f'''Bạn là chuyên gia phân tích thị trường cho ngành "{self.industry}" tại thị trường "{self.market}".

MỤC ĐÍCH NGHIÊN CỨU: {purpose}

NGỮ CẢNH PHÂN TÍCH:
- Chủ đề chính: {layer1}
- Lĩnh vực: {layer2}

CÂU HỎI CẦN TRẢ LỜI: "{main_question}"

🎯 **YÊU CẦU BẮT BUỘC:**
1. **TRẢ LỜI TRỰC TIẾP** câu hỏi ngay từ câu đầu tiên
2. **BẮT ĐẦU** bằng: "Để trả lời câu hỏi về [tóm tắt ngắn câu hỏi]..."
3. **FOCUS 100%** vào nội dung mà câu hỏi đang hỏi - không drift sang chủ đề khác
4. **SỬ DỤNG** số liệu và ví dụ cụ thể từ thị trường {self.market}
5. **KẾT THÚC** bằng conclusion trả lời rõ ràng câu hỏi

**CẤU TRÚC:**
- Đoạn 1: Trả lời trực tiếp + evidence chính (3-4 câu)
- Đoạn 2: Phân tích deeper với data/examples (3-4 câu)  
- Đoạn 3: Impact/implications và conclusion (2-3 câu)

**TRÁNH:**
- Nói chung chung hoặc lạc đề
- Đặt câu hỏi thêm
- Phân tích những gì không được hỏi

**CHỈ TẬP TRUNG:** Trả lời chính xác và đầy đủ câu hỏi "{main_question}"'''
        
        return template

    def create_layer3_comprehensive_category_prompt(self, layer1: str, layer2: str, all_questions: list, purpose: str) -> str:
        """Tạo prompt để phân tích comprehensive cho toàn bộ category (Layer 3 mode)"""
        
        # Get Vietnamese market name for Vietnamese reports
        vietnamese_market = self.get_vietnamese_market_name(self.market)
        
        questions_text = "\n".join([f"- {q}" for q in all_questions])
        
        template = f'''Bạn là chuyên gia phân tích thị trường cho ngành "{self.industry}" tại thị trường "{vietnamese_market}".

NGỮ CẢNH: {layer1} > {layer2}
MỤC ĐÍCH NGHIÊN CỨU: "{purpose}"

CÁC CÂU HỎI CẦN TRẢ LỜI TOÀN DIỆN:
{questions_text}

🎯 **NHIỆM VỤ:** Viết phân tích comprehensive cho toàn bộ category "{layer2}" bằng cách trả lời tích hợp TẤT CẢ câu hỏi trên.

**BẮT ĐẦU NGAY VỚI PHÂN TÍCH** - KHÔNG có câu giới thiệu hay mở đầu, đi thẳng vào tình hình hiện tại.

**CÁCH VIẾT - FLOW TỰ NHIÊN:**

Viết một phân tích dạng văn xuôi, liền mạch theo logic:
1. **Tình hình hiện tại** (150-200 từ): Bắt đầu ngay với overview về {layer2} trong ngành {self.industry}
2. **Phân tích chi tiết** (200-250 từ): Deep dive các factors chính, data cụ thể 
3. **Tác động và xu hướng** (150-200 từ): Impacts, trends, opportunities 
4. **Khuyến nghị chiến lược** (100-150 từ): Actionable insights và recommendations

**YÊU CẦU:**
- BẮT ĐẦU NGAY với phân tích (VD: "Tình hình {layer2} hiện tại...", "Trong bối cảnh {layer2}...")
- TRẢ LỜI HẾT TẤT CẢ các câu hỏi được liệt kê
- VIẾT liền mạch như một bài phân tích chuyên nghiệp
- SỬ DỤNG data và examples cụ thể từ {vietnamese_market}
- TRÁNH section headers, viết dạng essay flow tự nhiên
- KẾT HỢP insights từ tất cả câu hỏi thành một comprehensive view

**PHONG CÁCH:**
- Professional analysis writing
- Smooth transitions giữa các ý
- Evidence-based với concrete examples
- Forward-looking perspective

**VÍ DỤ BẮT ĐẦU TỐT:**
"Tình hình {layer2} trong ngành {self.industry} tại {vietnamese_market} hiện đang..."
"Bối cảnh {layer2} cho thấy..."
"Môi trường {layer2} đang trải qua..."

Tạo một phân tích comprehensive trả lời hết các câu hỏi trên để hiểu toàn diện về {layer2} impact lên ngành {self.industry}.'''
        
        return template

    def create_layer4_enhancement_prompt(self, layer1: str, layer2: str, main_question: str, sub_question: str, layer3_content: str, purpose: str) -> str:
        """Tạo prompt để enhance specific section từ Layer 3 lên Layer 4"""
        template = f'''As a master prompt engineer, I need to enhance a specific section of an existing market research report from Layer 3 to Layer 4 standard for: "{self.industry}" in market: "{self.market}". 

The research purpose is: "{purpose}"
Structure: layer 1: {layer1} | layer 2: {layer2} | layer 3: {main_question}

EXISTING LAYER 3 CONTENT TO ENHANCE:
{layer3_content}

SPECIFIC ENHANCEMENT REQUEST (Layer 4):
{sub_question}

Create a prompt for GPT to provide deep, detailed analysis specifically for the enhancement request above, while building upon the existing Layer 3 content. The result should be much more detailed, with specific data, examples, and actionable insights. Vietnamese answer only.'''
        
        return template

    def create_layer4_comprehensive_report_prompt(self, layer1: str, layer2: str, main_question: str, sub_questions: list, layer3_content: str, purpose: str) -> str:
        """Tạo prompt để tạo báo cáo Layer 4 tổng hợp - DIRECT CONTENT, NO INTRO"""
        
        # Get Vietnamese market name for Vietnamese reports
        vietnamese_market = self.get_vietnamese_market_name(self.market)
        
        sub_questions_text = "\n".join([f"- {sq}" for sq in sub_questions])
        
        template = f'''Bạn là chuyên gia phân tích thị trường cho ngành "{self.industry}" tại thị trường "{vietnamese_market}".

NGỮ CẢNH: {layer1} > {layer2}
CÂU HỎI CHÍNH CẦN TRẢ LỜI: "{main_question}"

PHÂN TÍCH SẴN CÓ (Layer 3):
{layer3_content}

CÁC KHÍA CẠNH CHI TIẾT CẦN PHÂN TÍCH:
{sub_questions_text}

🎯 **NHIỆM VỤ:** Viết phân tích chuyên sâu trả lời câu hỏi chính "{main_question}" bằng cách tích hợp tất cả khía cạnh chi tiết.

**BẮT ĐẦU NGAY VỚI NỘI DUNG PHÂN TÍCH** - KHÔNG có câu giới thiệu, không có "Để trả lời câu hỏi...", đi thẳng vào tình hình hiện tại.

**CÁCH VIẾT - FLOW TỰ NHIÊN:**

Viết một phân tích dạng văn xuôi, liền mạch theo logic:
1. **Tình hình hiện tại** (120-150 từ): Bắt đầu ngay với phân tích tình trạng hiện tại, data cụ thể
2. **Động lực và tác động** (120-150 từ): Drivers chính, impacts lên players và consumers
3. **Cơ hội và xu hướng** (120-150 từ): Opportunities, growth areas, success cases
4. **Thách thức và rủi ro** (80-120 từ): Barriers, risks cần monitor
5. **Khuyến nghị chiến lược** (100-120 từ): Actionable steps cụ thể

**YÊU CẦU CRITICAL:**
- BẮT ĐẦU NGAY bằng câu về tình hình thực tế (VD: "Hiện tại thị trường...", "Trong bối cảnh...", "Tình trạng hiện tại...")
- KHÔNG sử dụng section headers hay bullet points
- KHÔNG có câu giới thiệu hay mở đầu
- VIẾT liền mạch như một bài phân tích chuyên nghiệp
- Use real {vietnamese_market} market data và case studies
- BE SPECIFIC - tránh generalities
- Total: 550-700 từ

**PHONG CÁCH:**
- Văn xuôi professional, mạch lạc
- Transition tự nhiên giữa các ý
- Bắt đầu ngay với facts và analysis
- Flow như một essay analysis, không intro

**VÍ DỤ BẮT ĐẦU TỐT:**
"Hiện tại ngành [X] đang trải qua..."
"Trong bối cảnh thị trường [Y]..."
"Tình trạng hiện tại cho thấy..."
"Thị trường [Z] đang chứng kiến..."

**KẾT THÚC** với conclusion trả lời hoàn chỉnh câu hỏi chính.'''
        
        return template

    def process_layer3_research(self, json_file: str = "market_research_structured.json", limit: int = None) -> Dict[str, Any]:
        """
        Xử lý nghiên cứu thị trường ở Layer 3 standard (main questions only)
        
        Args:
            json_file (str): Đường dẫn file JSON chứa dữ liệu
            limit (int): Giới hạn số questions để test (None = unlimited)
            
        Returns:
            Dict: Kết quả nghiên cứu thị trường Layer 3
        """
        
        # Đọc file JSON
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        purpose = data.get('purpose', '')
        layers = data.get('layers', [])
        
        results = {
            "industry": self.industry,
            "market": self.market,
            "purpose": purpose,
            "research_standard": "Layer 3",
            "api_provider": "openai",
            "model_used": self.model,
            "research_results": []
        }
        
        total_questions = 0
        processed_questions = 0
        
        # Đếm tổng số main questions (Layer 3)
        for layer in layers:
            for category in layer.get('categories', []):
                total_questions += len(category.get('questions', []))
        
        # Apply limit if specified
        if limit:
            total_questions = min(total_questions, limit)
            print(f"🚧 Testing mode: Giới hạn {limit} questions")
        
        print(f"🎯 Bắt đầu nghiên cứu thị trường Layer 3: {self.industry}")
        print(f"📊 Thị trường: {self.market}")
        print(f"🤖 API: OpenAI {self.model}")
        print(f"❓ Tổng số main questions (Layer 3): {total_questions}")
        print("="*60)
        
        for layer in layers:
            layer_name = layer.get('name', '')
            layer_result = {
                "layer_name": layer_name,
                "categories": []
            }
            
            print(f"\n🔥 Đang xử lý Layer: {layer_name}")
            
            for category in layer.get('categories', []):
                category_name = category.get('name', '')
                category_result = {
                    "category_name": category_name,
                    "questions": []
                }
                
                print(f"📋 Category: {category_name}")
                
                for question in category.get('questions', []):
                    # Check limit
                    if limit and processed_questions >= limit:
                        print(f"🚧 Đã đạt giới hạn {limit} questions, dừng lại.")
                        break
                        
                    main_question = question.get('main_question', '')
                    processed_questions += 1
                    progress = (processed_questions / total_questions) * 100
                    
                    print(f"  ❓ [{processed_questions}/{total_questions}] ({progress:.1f}%) Processing: {main_question[:50]}...")
                    
                    # Bước 1: Tạo Layer 3 prompt request
                    prompt_request = self.create_layer3_prompt_request(
                        layer_name, category_name, main_question, purpose
                    )
                    
                    # Bước 2: Gửi request để lấy prompt
                    print(f"    🔄 Tạo Layer 3 prompt...")
                    generated_prompt = self.call_openai_api(prompt_request)
                    
                    # Bước 3: Sử dụng prompt để lấy kết quả Layer 3
                    print(f"    🔍 Nghiên cứu Layer 3...")
                    research_result = self.call_openai_api(generated_prompt)
                    
                    # Lưu kết quả với cấu trúc mới cho Layer 3
                    question_result = {
                        "main_question": main_question,
                        "research_standard": "Layer 3",
                        "generated_prompt": generated_prompt,
                        "layer3_content": research_result,
                        "sub_questions": question.get('sub_questions', []),  # Lưu để có thể enhance sau
                        "layer4_enhancements": {}  # Dict để lưu các enhancement
                    }
                    
                    category_result["questions"].append(question_result)
                    
                    # Delay để tránh rate limit
                    time.sleep(self.delay_seconds)
                    
                    print(f"    ✅ Hoàn thành Layer 3!")
                
                if category_result["questions"]:  # Only add if has questions
                    layer_result["categories"].append(category_result)
                    
                # Check limit
                if limit and processed_questions >= limit:
                    break
            
            if layer_result["categories"]:  # Only add if has categories
                results["research_results"].append(layer_result)
                
            # Check limit
            if limit and processed_questions >= limit:
                break
        
        print("\n" + "="*60)
        print(f"🎉 Hoàn thành nghiên cứu thị trường Layer 3!")
        print(f"📊 Đã xử lý: {processed_questions} main questions")
        
        return results

    def enhance_to_layer4(self, layer3_results: Dict[str, Any], layer_name: str, category_name: str, main_question: str, sub_question: str) -> str:
        """
        Enhance một section cụ thể từ Layer 3 lên Layer 4
        """
        
        # Tìm layer3_content tương ứng
        layer3_content = ""
        purpose = layer3_results.get('purpose', '')
        
        for layer in layer3_results.get('research_results', []):
            if layer.get('layer_name') == layer_name:
                for category in layer.get('categories', []):
                    if category.get('category_name') == category_name:
                        for question in category.get('questions', []):
                            if question.get('main_question') == main_question:
                                layer3_content = question.get('layer3_content', '')
                                break
        
        if not layer3_content:
            return "Không tìm thấy nội dung Layer 3 để enhance"
        
        print(f"🔄 Enhancing to Layer 4: {sub_question[:50]}...")
        
        # Bước 1: Tạo enhancement prompt
        enhancement_prompt_request = self.create_layer4_enhancement_prompt(
            layer_name, category_name, main_question, sub_question, layer3_content, purpose
        )
        
        # Bước 2: Lấy enhancement prompt
        print(f"    🔄 Tạo Layer 4 enhancement prompt...")
        generated_prompt = self.call_openai_api(enhancement_prompt_request)
        
        # Bước 3: Chạy enhancement
        print(f"    🔍 Thực hiện Layer 4 enhancement...")
        enhancement_result = self.call_openai_api(generated_prompt)
        
        print(f"    ✅ Hoàn thành Layer 4 enhancement!")
        
        return enhancement_result

    def enhance_to_layer4_comprehensive(self, layer3_results: Dict[str, Any], layer_name: str, category_name: str, main_question: str) -> str:
        """
        Tạo báo cáo Layer 4 tổng hợp cho toàn bộ main question (tất cả sub-questions)
        """
        
        # Tìm layer3_content và sub_questions tương ứng
        layer3_content = ""
        sub_questions = []
        purpose = layer3_results.get('purpose', '')
        
        for layer in layer3_results.get('research_results', []):
            if layer.get('layer_name') == layer_name:
                for category in layer.get('categories', []):
                    if category.get('category_name') == category_name:
                        for question in category.get('questions', []):
                            if question.get('main_question') == main_question:
                                layer3_content = question.get('layer3_content', '')
                                sub_questions = question.get('sub_questions', [])
                                break
        
        if not layer3_content:
            return "Không tìm thấy nội dung Layer 3 để enhance"
        
        if not sub_questions:
            return "Không tìm thấy sub-questions để tạo báo cáo tổng hợp"
        
        print(f"🔄 Tạo báo cáo Layer 4 tổng hợp với {len(sub_questions)} sub-questions...")
        
        # Tạo comprehensive report prompt
        comprehensive_prompt = self.create_layer4_comprehensive_report_prompt(
            layer_name, category_name, main_question, sub_questions, layer3_content, purpose
        )
        
        # Gọi API để tạo comprehensive report
        print(f"    🔍 Thực hiện Layer 4 comprehensive analysis...")
        comprehensive_report = self.call_openai_api(comprehensive_prompt)
        
        print(f"    ✅ Hoàn thành Layer 4 comprehensive report!")
        
        return comprehensive_report

    def add_layer4_enhancement(self, results_file: str, layer_name: str, category_name: str, main_question: str, sub_question: str, output_file: str = None) -> str:
        """
        Thêm Layer 4 enhancement vào kết quả hiện có và lưu file
        """
        
        # Đọc kết quả hiện có
        with open(results_file, 'r', encoding='utf-8') as f:
            results = json.load(f)
        
        # Thực hiện enhancement
        enhancement_content = self.enhance_to_layer4(results, layer_name, category_name, main_question, sub_question)
        
        # Cập nhật kết quả
        for layer in results.get('research_results', []):
            if layer.get('layer_name') == layer_name:
                for category in layer.get('categories', []):
                    if category.get('category_name') == category_name:
                        for question in category.get('questions', []):
                            if question.get('main_question') == main_question:
                                # Thêm enhancement vào dict
                                question['layer4_enhancements'][sub_question] = {
                                    "enhanced_content": enhancement_content,
                                    "enhancement_timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
                                }
                                break
        
        # Lưu file
        if output_file is None:
            output_file = results_file  # Ghi đè file gốc
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        print(f"💾 Đã cập nhật Layer 4 enhancement vào: {output_file}")
        return output_file

    def add_layer4_comprehensive_enhancement(self, results_file: str, layer_name: str, category_name: str, main_question: str, output_file: str = None) -> str:
        """
        Thêm Layer 4 comprehensive enhancement vào kết quả hiện có và lưu file
        """
        
        # Đọc kết quả hiện có
        with open(results_file, 'r', encoding='utf-8') as f:
            results = json.load(f)
        
        # Thực hiện comprehensive enhancement
        comprehensive_content = self.enhance_to_layer4_comprehensive(results, layer_name, category_name, main_question)
        
        # Cập nhật kết quả với comprehensive report
        for layer in results.get('research_results', []):
            if layer.get('layer_name') == layer_name:
                for category in layer.get('categories', []):
                    if category.get('category_name') == category_name:
                        for question in category.get('questions', []):
                            if question.get('main_question') == main_question:
                                # Thêm comprehensive enhancement
                                question['layer4_comprehensive_report'] = {
                                    "comprehensive_content": comprehensive_content,
                                    "enhancement_timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                                    "sub_questions_integrated": question.get('sub_questions', [])
                                }
                                break
        
        # Lưu file
        if output_file is None:
            output_file = results_file  # Ghi đè file gốc
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        print(f"💾 Đã cập nhật Layer 4 comprehensive report vào: {output_file}")
        return output_file

    def run_layer3_research(self, structured_data: dict, topic: str, testing_mode: bool = False, analysis_level: str = "Layer 4 Analysis") -> dict:
        """Main research execution with comprehensive error handling and reference tracking"""
        
        print(f"🎯 Bắt đầu nghiên cứu thị trường Layer 3: {topic}")
        print(f"📊 Thị trường: {self.market}")
        print(f"🤖 API: OpenAI {self.model}")
        print(f"🔬 Analysis Level: {analysis_level}")
        
        # Reset reference tracking for new research
        self.reference_tracker = {}
        self.tracked_sources = set()
        
        # Get purpose from structured data
        purpose = structured_data.get('purpose', 'Nghiên cứu thị trường và phân tích cơ hội kinh doanh')
        
        # Create result structure
        result = {
            'research_metadata': {
                'industry': topic,
                'market': self.market,
                'model_used': self.model,
                'api_provider': 'OpenAI',
                'purpose': purpose,
                'analysis_level': analysis_level,
                'research_timestamp': time.strftime("%Y-%m-%d %H:%M:%S"),
                'testing_mode': testing_mode
            },
            'research_results': [],
            'research_statistics': {},
            'tracked_references': []  # Will be populated at the end
        }
        
        # Count total questions for progress tracking
        total_questions = 0
        total_questions_available = 0
        
        # First pass: count all available questions
        for layer in structured_data.get('layers', []):
            for category in layer.get('categories', []):
                questions = category.get('questions', [])
                total_questions_available += len(questions)
        
        if testing_mode:
            total_questions = min(5, total_questions_available)  # Limit to exactly 5 questions
            print(f"🧪 Testing mode: Limiting to {total_questions} questions out of {total_questions_available} available")
        else:
            total_questions = total_questions_available
        
        print(f"❓ Tổng số main questions (Layer 3): {total_questions}")
        print("=" * 60)
        
        processed_questions = 0
        
        # Branch logic based on analysis level
        if analysis_level == "Layer 3 Analysis":
            return self._run_layer3_category_analysis(structured_data, purpose, result, testing_mode)
        else:
            return self._run_layer4_detailed_analysis(structured_data, purpose, result, testing_mode)
    
    def _run_layer3_category_analysis(self, structured_data: dict, purpose: str, result: dict, testing_mode: bool) -> dict:
        """Layer 3 Analysis Mode: Comprehensive analysis per category"""
        
        print("🎯 Chế độ Layer 3: Phân tích comprehensive theo category")
        processed_categories = 0
        total_questions_processed = 0
        
        for layer in structured_data.get('layers', []):
            layer_name = layer.get('name', '')
            layer_result = {
                'layer_name': layer_name,
                'categories': []
            }
            
            print(f"🔥 Đang xử lý Layer: {layer_name}")
            
            for category in layer.get('categories', []):
                category_name = category.get('name', '')
                processed_categories += 1
                
                print(f"📋 Category [{processed_categories}]: {category_name}")
                
                questions = category.get('questions', [])
                
                # In testing mode, limit questions for this category
                if testing_mode:
                    questions_remaining = 5 - total_questions_processed
                    if questions_remaining <= 0:
                        print(f"    🧪 Testing mode: Reached 5 question limit, skipping remaining categories")
                        break
                    questions = questions[:min(len(questions), questions_remaining)]
                
                # Collect all main questions for this category
                all_main_questions = [q.get('main_question', '') for q in questions]
                total_questions_processed += len(all_main_questions)
                
                print(f"    🔄 Gom {len(all_main_questions)} questions cho comprehensive analysis...")
                
                # Create Layer 3 comprehensive analysis for entire category
                comprehensive_prompt = self.create_layer3_comprehensive_category_prompt(
                    layer_name, category_name, all_main_questions, purpose
                )
                
                print(f"    🔍 Phân tích comprehensive cho category {category_name}...")
                comprehensive_content = self.call_openai_api(comprehensive_prompt)
                
                # Store individual questions for reference (minimal processing)
                category_questions = []
                for question_data in questions:
                    question_result = {
                        'main_question': question_data.get('main_question', ''),
                        'sub_questions': question_data.get('sub_questions', []),
                        'layer3_content': None  # Not processed individually in Layer 3 mode
                    }
                    category_questions.append(question_result)
                
                # Create category result with comprehensive analysis
                category_result = {
                    'category_name': category_name,
                    'questions': category_questions,
                    'layer3_comprehensive_category': {
                        'comprehensive_content': comprehensive_content,
                        'analysis_timestamp': time.strftime("%Y-%m-%d %H:%M:%S"),
                        'questions_analyzed': all_main_questions,
                        'analysis_mode': 'Layer 3 Category Comprehensive'
                    }
                }
                
                layer_result['categories'].append(category_result)
                print(f"    ✅ Hoàn thành comprehensive analysis cho {category_name}!")
                
                # Add delay to avoid rate limiting
                time.sleep(self.delay_seconds)
                
                # Break if we've reached the testing limit
                if testing_mode and total_questions_processed >= 5:
                    print(f"    🧪 Testing mode: Reached 5 question limit")
                    break
                
            result['research_results'].append(layer_result)
            
            # Break if we've reached the testing limit
            if testing_mode and total_questions_processed >= 5:
                break
        
        # Final statistics
        print("=" * 60)
        print("🎉 Hoàn thành Layer 3 Category Analysis!")
        print(f"📊 Đã phân tích: {processed_categories} categories")
        if testing_mode:
            print(f"🧪 Testing mode: Processed {total_questions_processed} questions")
        
        # Add tracked references to result
        top_references = self.get_top_references(10)
        result['tracked_references'] = top_references
        
        print(f"📚 Tracked {len(self.tracked_sources)} unique sources")
        if top_references:
            print("🔝 Top references:")
            for source, count in top_references[:5]:
                print(f"   • {source} ({count}x)")
        
        # Add research statistics
        result['research_statistics'] = {
            'total_categories_processed': processed_categories,
            'total_questions_processed': total_questions_processed,
            'total_sources_tracked': len(self.tracked_sources),
            'total_api_calls': processed_categories,  # One call per category
            'processing_time_estimate': f"{processed_categories * self.delay_seconds / 60:.1f} minutes",
            'analysis_mode': 'Layer 3 Category Comprehensive'
        }
        
        return result
    
    def _run_layer4_detailed_analysis(self, structured_data: dict, purpose: str, result: dict, testing_mode: bool) -> dict:
        """Layer 4 Analysis Mode: Detailed analysis per main question (current behavior)"""
        
        print("🎯 Chế độ Layer 4: Phân tích chi tiết theo main question")
        processed_questions = 0
        total_questions = 0
        total_questions_available = 0
        
        # Count total questions available
        for layer in structured_data.get('layers', []):
            for category in layer.get('categories', []):
                questions = category.get('questions', [])
                total_questions_available += len(questions)
        
        if testing_mode:
            total_questions = min(5, total_questions_available)  # Limit to exactly 5 questions
            print(f"🧪 Testing mode: Limiting to {total_questions} questions out of {total_questions_available} available")
        else:
            total_questions = total_questions_available
        
        questions_remaining = total_questions if testing_mode else float('inf')
        
        for layer in structured_data.get('layers', []):
            layer_name = layer.get('name', '')
            layer_result = {
                'layer_name': layer_name,
                'categories': []
            }
            
            print(f"🔥 Đang xử lý Layer: {layer_name}")
            
            for category in layer.get('categories', []):
                category_name = category.get('name', '')
                category_result = {
                    'category_name': category_name,
                    'questions': []
                }
                
                print(f"📋 Category: {category_name}")
                
                questions = category.get('questions', [])
                
                # In testing mode, limit questions to remaining count
                if testing_mode:
                    questions_to_process = min(len(questions), questions_remaining)
                    questions = questions[:questions_to_process]
                    questions_remaining -= questions_to_process
                
                for question_data in questions:
                    if testing_mode and processed_questions >= total_questions:
                        break  # Stop if we've reached the limit
                        
                    processed_questions += 1
                    progress_percent = (processed_questions / total_questions) * 100
                    
                    main_question = question_data.get('main_question', '')
                    sub_questions = question_data.get('sub_questions', [])
                    
                    print(f"  ❓ [{processed_questions}/{total_questions}] ({progress_percent:.1f}%) Processing: {main_question[:50]}...")
                    
                    # Create Layer 3 analysis
                    print("    🔄 Tạo Layer 3 prompt...")
                    layer3_prompt = self.create_layer3_prompt_request(
                        layer_name, category_name, main_question, purpose
                    )
                    
                    print("    🔍 Nghiên cứu Layer 3...")
                    layer3_content = self.call_openai_api(layer3_prompt)
                    
                    # Create question result
                    question_result = {
                        'main_question': main_question,
                        'sub_questions': sub_questions,
                        'layer3_content': layer3_content
                    }
                    
                    # Auto Layer 4 comprehensive if has sub-questions
                    if sub_questions and len(sub_questions) > 0:
                        print(f"🔄 Tạo báo cáo Layer 4 tổng hợp với {len(sub_questions)} sub-questions...")
                        
                        # Create temporary structure with current question data for Layer 4 enhancement
                        temp_structure = {
                            'research_results': [{
                                'layer_name': layer_name,
                                'categories': [{
                                    'category_name': category_name,
                                    'questions': [question_result]  # Include current question with layer3_content
                                }]
                            }],
                            'purpose': purpose
                        }
                        
                        comprehensive_content = self.enhance_to_layer4_comprehensive(
                            temp_structure, 
                            layer_name, 
                            category_name, 
                            main_question
                        )
                        
                        question_result['layer4_comprehensive_report'] = {
                            "comprehensive_content": comprehensive_content,
                            "enhancement_timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                            "sub_questions_integrated": sub_questions
                        }
                    
                    category_result['questions'].append(question_result)
                    print("    ✅ Hoàn thành Layer 3!")
                    
                    # Add delay to avoid rate limiting
                    time.sleep(self.delay_seconds)
                
                if category_result['questions']:  # Only add category if it has questions
                    layer_result['categories'].append(category_result)
                
                # Break early if we've reached the testing limit
                if testing_mode and processed_questions >= total_questions:
                    break
                    
            if layer_result['categories']:  # Only add layer if it has categories
                result['research_results'].append(layer_result)
                
            # Break early if we've reached the testing limit
            if testing_mode and processed_questions >= total_questions:
                break
        
        print("=" * 60)
        print("🎉 Hoàn thành nghiên cứu thị trường Layer 4!")
        print(f"📊 Đã xử lý: {processed_questions} main questions")
        
        # Add tracked references to result
        top_references = self.get_top_references(10)
        result['tracked_references'] = top_references
        
        print(f"📚 Tracked {len(self.tracked_sources)} unique sources")
        if top_references:
            print("🔝 Top references:")
            for source, count in top_references[:5]:
                print(f"   • {source} ({count}x)")
        
        # Add research statistics
        result['research_statistics'] = {
            'total_questions_processed': processed_questions,
            'total_sources_tracked': len(self.tracked_sources),
            'total_api_calls': processed_questions * 2,  # Estimate including Layer 4
            'processing_time_estimate': f"{processed_questions * self.delay_seconds / 60:.1f} minutes",
            'analysis_mode': 'Layer 4 Detailed per Question'
        }
        
        return result

    def get_vietnamese_market_name(self, market: str) -> str:
        """Translate market name to Vietnamese for consistent Vietnamese reports"""
        market_translations = {
            "🇺🇸 United States": "🇺🇸 Hoa Kỳ",
            "🇨🇳 China": "🇨🇳 Trung Quốc", 
            "🇯🇵 Japan": "🇯🇵 Nhật Bản",
            "🇰🇷 South Korea": "🇰🇷 Hàn Quốc",
            "🇹🇭 Thailand": "🇹🇭 Thái Lan",
            "🇸🇬 Singapore": "🇸🇬 Singapore",
            "🇲🇾 Malaysia": "🇲🇾 Malaysia",
            "🇮🇩 Indonesia": "🇮🇩 Indonesia",
            "🇵🇭 Philippines": "🇵🇭 Philippines",
            "🇬🇧 United Kingdom": "🇬🇧 Vương quốc Anh",
            "🇩🇪 Germany": "🇩🇪 Đức",
            "🇫🇷 France": "🇫🇷 Pháp",
            "🇮🇹 Italy": "🇮🇹 Ý",
            "🇪🇸 Spain": "🇪🇸 Tây Ban Nha",
            "🇨🇦 Canada": "🇨🇦 Canada",
            "🇦🇺 Australia": "🇦🇺 Úc",
            "🇳🇿 New Zealand": "🇳🇿 New Zealand",
            "🇧🇷 Brazil": "🇧🇷 Brazil",
            "🇲🇽 Mexico": "🇲🇽 Mexico",
            "🇮🇳 India": "🇮🇳 Ấn Độ",
            "🇷🇺 Russia": "🇷🇺 Nga",
            "🇿🇦 South Africa": "🇿🇦 Nam Phi",
            "🇪🇬 Egypt": "🇪🇬 Ai Cập",
            "🇦🇪 UAE": "🇦🇪 UAE",
            "🇸🇦 Saudi Arabia": "🇸🇦 Ả Rập Saudi",
            "🇹🇷 Turkey": "🇹🇷 Thổ Nhĩ Kỳ",
            "🇳🇱 Netherlands": "🇳🇱 Hà Lan",
            "🇸🇪 Sweden": "🇸🇪 Thụy Điển",
            "🇳🇴 Norway": "🇳🇴 Na Uy",
            "🇩🇰 Denmark": "🇩🇰 Đan Mạch",
            "🇫🇮 Finland": "🇫🇮 Phần Lan",
            "🇨🇭 Switzerland": "🇨🇭 Thụy Sĩ",
            "🇦🇹 Austria": "🇦🇹 Áo",
            "🇧🇪 Belgium": "🇧🇪 Bỉ",
            "🇵🇱 Poland": "🇵🇱 Ba Lan",
            "🇨🇿 Czech Republic": "🇨🇿 Cộng hòa Séc",
            "🇭🇺 Hungary": "🇭🇺 Hungary",
            "🇬🇷 Greece": "🇬🇷 Hy Lạp",
            "🇵🇹 Portugal": "🇵🇹 Bồ Đào Nha",
            "🇮🇪 Ireland": "🇮🇪 Ireland",
            "🇮🇱 Israel": "🇮🇱 Israel",
            "🇭🇰 Hong Kong": "🇭🇰 Hồng Kông",
            "🇹🇼 Taiwan": "🇹🇼 Đài Loan",
            "🇦🇷 Argentina": "🇦🇷 Argentina",
            "🇨🇱 Chile": "🇨🇱 Chile",
            "🇨🇴 Colombia": "🇨🇴 Colombia",
            "🇵🇪 Peru": "🇵🇪 Peru",
            "🇻🇪 Venezuela": "🇻🇪 Venezuela",
            "🇪🇨 Ecuador": "🇪🇨 Ecuador",
            "🇺🇾 Uruguay": "🇺🇾 Uruguay",
            "🇧🇴 Bolivia": "🇧🇴 Bolivia",
            "🇵🇾 Paraguay": "🇵🇾 Paraguay",
            "🇳🇬 Nigeria": "🇳🇬 Nigeria",
            "🇰🇪 Kenya": "🇰🇪 Kenya",
            "🇬🇭 Ghana": "🇬🇭 Ghana",
            "🇪🇹 Ethiopia": "🇪🇹 Ethiopia",
            "🇺🇬 Uganda": "🇺🇬 Uganda",
            "🇹🇿 Tanzania": "🇹🇿 Tanzania",
            "🇿🇼 Zimbabwe": "🇿🇼 Zimbabwe",
            "🌏 Southeast Asia": "🌏 Đông Nam Á",
            "🌍 Asia-Pacific": "🌍 Châu Á - Thái Bình Dương",
            "🌎 Global Market": "🌎 Thị trường Toàn cầu"
        }
        
        return market_translations.get(market, market) 