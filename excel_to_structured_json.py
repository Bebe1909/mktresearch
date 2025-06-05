import pandas as pd
import json
import os

class ExcelToStructuredJSON:
    """
    Lớp chuyển đổi Excel framework thành JSON có cấu trúc
    Hỗ trợ cho cả web UI và CLI
    """
    
    def __init__(self):
        """Khởi tạo converter"""
        pass
    
    def convert_excel_to_json(self, excel_path, json_path, custom_purpose=None):
        """
        Chuyển đổi Excel framework thành JSON có cấu trúc
        
        Args:
            excel_path (str): Đường dẫn tới file Excel
            json_path (str): Đường dẫn output file JSON
            custom_purpose (str): Mục đích custom (nếu có)
            
        Returns:
            bool: True nếu thành công, False nếu lỗi
        """
        try:
            # Đọc sheet template (updated from "Market Research")
            df = pd.read_excel(excel_path, sheet_name="template")
            
            print("✅ Đã đọc file Excel thành công!")
            print(f"📊 Kích thước dữ liệu: {df.shape}")
            
            # Tìm dòng "Mục đích của Market Research"
            purpose_row = None
            purpose_content = None
            
            for idx, row in df.iterrows():
                if any("Mục đích của Market Research" in str(cell) for cell in row if pd.notna(cell)):
                    purpose_row = idx
                    # Tìm nội dung purpose trong các cột
                    for col in df.columns:
                        if pd.notna(row[col]) and "Mục đích của Market Research" not in str(row[col]):
                            if str(row[col]).strip():
                                purpose_content = str(row[col])
                                break
                    break
            
            # Sử dụng custom purpose nếu có, ngược lại dùng purpose từ Excel
            final_purpose = custom_purpose if custom_purpose else (purpose_content if purpose_content else "Nghiên cứu và phân tích thị trường")
            
            # Khởi tạo cấu trúc JSON kết quả
            result = {
                "purpose": final_purpose,
                "layers": []
            }
            
            # Phân tích các layer (bắt đầu từ sau dòng purpose)
            if purpose_row is not None:
                # Tìm header của layers (Layer 1, Layer 2, etc.)
                layer_header_row = None
                for idx in range(purpose_row + 1, len(df)):
                    row = df.iloc[idx]
                    if any("Layer" in str(cell) for cell in row if pd.notna(cell)):
                        layer_header_row = idx
                        break
                
                if layer_header_row is not None:
                    print(f"📍 Tìm thấy header layers ở dòng {layer_header_row}")
                    
                    # Get header row to determine number of layers dynamically
                    header_row = df.iloc[layer_header_row]
                    layer_columns = []
                    
                    # Find all Layer columns dynamically
                    for col_idx, col_val in enumerate(header_row):
                        if pd.notna(col_val) and "Layer" in str(col_val):
                            layer_columns.append(col_idx)
                    
                    print(f"🔍 Detected {len(layer_columns)} layer columns: {[f'Layer {i+1}' for i in range(len(layer_columns))]}")
                    
                    if len(layer_columns) < 3:
                        print("⚠️ Warning: Need at least 3 layers (Layer 1, Layer 2, Layer 3) for proper structure")
                        return False
                    
                    # Dynamic layer tracking
                    layer_stack = [None] * len(layer_columns)  # Track current value at each layer level
                    
                    # Phân tích dữ liệu từ dòng tiếp theo
                    for idx in range(layer_header_row + 1, len(df)):
                        row = df.iloc[idx]
                        
                        # Bỏ qua dòng trống
                        if row.isna().all():
                            continue
                        
                        # Get values for all layer columns
                        layer_values = []
                        for col_idx in layer_columns:
                            if col_idx < len(row):
                                val = row.iloc[col_idx]
                                layer_values.append(str(val) if pd.notna(val) and str(val) != 'None' else None)
                            else:
                                layer_values.append(None)
                        
                        # Process layers dynamically
                        for level, value in enumerate(layer_values):
                            if value:  # Non-empty value at this level
                                # Reset deeper levels
                                for deeper_level in range(level + 1, len(layer_stack)):
                                    layer_stack[deeper_level] = None
                                
                                if level == 0:  # Layer 1 - Top level framework
                                    current_layer1 = {
                                        "name": value,
                                        "categories": []
                                    }
                                    result["layers"].append(current_layer1)
                                    layer_stack[0] = current_layer1
                                    
                                elif level == 1 and layer_stack[0]:  # Layer 2 - Categories
                                    current_layer2 = {
                                        "name": value,
                                        "questions": []
                                    }
                                    layer_stack[0]["categories"].append(current_layer2)
                                    layer_stack[1] = current_layer2
                                    
                                elif level == 2 and layer_stack[1]:  # Layer 3 - Main questions
                                    question = {
                                        "main_question": value,
                                        "sub_questions": []
                                    }
                                    layer_stack[1]["questions"].append(question)
                                    layer_stack[2] = question
                                    
                                elif level >= 3 and layer_stack[2]:  # Layer 4+ - Sub-questions or deeper analysis
                                    # Add as sub-question to current main question
                                    layer_stack[2]["sub_questions"].append(value)
                                
                                break  # Process only the first non-empty value in this row
            
            # Đảm bảo folder output tồn tại
            os.makedirs(os.path.dirname(json_path), exist_ok=True)
            
            # Lưu kết quả ra file JSON
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            
            print(f"✅ Đã chuyển đổi thành công!")
            print(f"📁 File output: {json_path}")
            print(f"📊 Số lượng layers: {len(result['layers'])}")
            
            return True
            
        except FileNotFoundError:
            print(f"❌ Không tìm thấy file: {excel_path}")
            return False
        except Exception as e:
            print(f"❌ Lỗi: {e}")
            return False

def convert_market_research_to_json():
    """
    Hàm legacy để tương thích với CLI
    Chuyển đổi sheet 'template' từ market research template.xlsx 
    thành JSON có cấu trúc với purpose và layers
    """
    
    # Sử dụng class mới - updated to use new file name
    converter = ExcelToStructuredJSON()
    excel_file = "input/market research template.xlsx"
    json_file = "output/market_research_structured.json"
    
    success = converter.convert_excel_to_json(excel_file, json_file)
    
    if success:
        # Load và return kết quả để tương thích với code cũ
        with open(json_file, 'r', encoding='utf-8') as f:
            result = json.load(f)
        return result
    else:
        return None

if __name__ == "__main__":
    result = convert_market_research_to_json()
    
    if result:
        print("\n--- Preview JSON structure ---")
        print(f"Purpose: {result['purpose'][:100]}...")
        print(f"Number of layers: {len(result['layers'])}")
        if result['layers']:
            print(f"First layer: {result['layers'][0]['name']}")
            if result['layers'][0]['categories']:
                print(f"First category: {result['layers'][0]['categories'][0]['name']}") 