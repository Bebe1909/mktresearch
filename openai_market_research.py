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

class OpenAIMarketResearch:
    def __init__(self, api_key: str, industry: str, market: str, model: str = "gpt-3.5-turbo"):
        """
        Khởi tạo class nghiên cứu thị trường với OpenAI GPT
        
        Args:
            api_key (str): API key của OpenAI
            industry (str): Ngành công nghiệp nghiên cứu
            market (str): Thị trường nghiên cứu
            model (str): Model GPT sử dụng
        """
        self.api_key = api_key
        self.industry = industry
        self.market = market
        self.model = model
        self.base_url = "https://api.openai.com/v1/chat/completions"
        self.delay_seconds = 3  # Increased delay from 1 to 3 seconds
        
    def call_openai_api(self, prompt: str, max_retries: int = 3) -> str:
        """Gửi request tới OpenAI API với retry logic"""
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        
        data = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "max_tokens": 2000,
            "temperature": 0.7
        }
        
        for attempt in range(max_retries):
            try:
                response = requests.post(self.base_url, headers=headers, json=data)
                response.raise_for_status()
                
                result = response.json()
                if 'choices' in result and len(result['choices']) > 0:
                    if 'message' in result['choices'][0] and 'content' in result['choices'][0]['message']:
                        return result['choices'][0]['message']['content']
                    else:
                        return "Không có nội dung trong phản hồi API"
                else:
                    return "Không có phản hồi từ API"
                    
            except requests.exceptions.RequestException as e:
                if "429" in str(e):  # Rate limit error
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
        
        sub_questions_text = "\n".join([f"- {sq}" for sq in sub_questions])
        
        template = f'''Bạn là chuyên gia phân tích thị trường cho ngành "{self.industry}" tại thị trường "{self.market}".

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
- Use real {self.market} market data và case studies
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

    def run_layer3_research(self, structured_data: dict, topic: str, testing_mode: bool = False) -> dict:
        """
        Chạy nghiên cứu Layer 3 với topic cụ thể
        
        Args:
            structured_data (dict): Dữ liệu cấu trúc từ JSON
            topic (str): Chủ đề nghiên cứu
            testing_mode (bool): Chế độ testing (5 câu hỏi)
            
        Returns:
            dict: Kết quả nghiên cứu
        """
        
        # Set industry from topic, but keep the market that was passed to __init__
        self.industry = topic
        
        # Set limit based on testing mode
        limit = 5 if testing_mode else None
        
        # Create temporary JSON file
        import tempfile
        import os
        
        temp_file = os.path.join('output', 'temp_structured_data.json')
        with open(temp_file, 'w', encoding='utf-8') as f:
            json.dump(structured_data, f, ensure_ascii=False, indent=2)
        
        try:
            # Process research
            result = self.process_layer3_research(temp_file, limit)
            
            # Clean up temp file
            if os.path.exists(temp_file):
                os.remove(temp_file)
                
            return result
            
        except Exception as e:
            # Clean up temp file on error
            if os.path.exists(temp_file):
                os.remove(temp_file)
            raise e 