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
                    
                    # Phân tích dữ liệu từ dòng tiếp theo
                    current_layer1 = None
                    current_layer2 = None
                    current_layer3 = None
                    
                    for idx in range(layer_header_row + 1, len(df)):
                        row = df.iloc[idx]
                        
                        # Bỏ qua dòng trống
                        if row.isna().all():
                            continue
                        
                        # Lấy giá trị các cột
                        col_values = []
                        for col in df.columns:
                            val = row[col]
                            col_values.append(str(val) if pd.notna(val) else None)
                        
                        # Xử lý layer 1 (cột 1)
                        if col_values[1] and col_values[1] != 'None':  # Layer 1
                            current_layer1 = {
                                "name": col_values[1],
                                "categories": []
                            }
                            result["layers"].append(current_layer1)
                            current_layer2 = None
                            current_layer3 = None
                        
                        # Xử lý layer 2 (cột 2)
                        if col_values[2] and col_values[2] != 'None' and current_layer1:  # Layer 2
                            current_layer2 = {
                                "name": col_values[2],
                                "questions": []
                            }
                            current_layer1["categories"].append(current_layer2)
                            current_layer3 = None
                        
                        # Xử lý layer 3 và 4 (câu hỏi và chi tiết)
                        if current_layer2:
                            layer3_val = col_values[3] if len(col_values) > 3 else None
                            layer4_val = col_values[4] if len(col_values) > 4 else None
                            
                            if layer3_val and layer3_val != 'None':
                                question = {
                                    "main_question": layer3_val,
                                    "sub_questions": []
                                }
                                if layer4_val and layer4_val != 'None':
                                    question["sub_questions"].append(layer4_val)
                                current_layer2["questions"].append(question)
                                current_layer3 = question
                            elif layer4_val and layer4_val != 'None' and current_layer3:
                                # Thêm sub-question vào câu hỏi hiện tại
                                current_layer3["sub_questions"].append(layer4_val)
            
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