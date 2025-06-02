# 🔬 Market Research Automation System

Hệ thống tự động hóa nghiên cứu thị trường sử dụng AI, tự động tạo báo cáo đầy đủ từ framework Excel đến tài liệu Word chuyên nghiệp.

## 🎯 **Dành cho ai?**

### 👥 **Marketing Teams - Web UI** (Recommended)
- ✅ **No-code solution** - Giao diện web thân thiện
- ✅ **Drag & drop** Excel upload
- ✅ **Real-time progress tracking**
- ✅ **Analytics dashboard** 
- ✅ **One-click export**

### 💻 **Technical Users - Command Line**
- ✅ **Script automation**
- ✅ **Batch processing**
- ✅ **Custom integration**

---

## 🚀 **Quick Start - Web UI (Cho Marketing Teams)**

### **Bước 1: Setup & Launch**

```bash
# Clone repository
git clone <repository_url>
cd mktresearch

# Cài đặt dependencies
pip install -r requirements.txt

# Launch Web UI
python launch_ui.py
```

**🌐 Web interface sẽ mở tại: http://localhost:8501**

### **Bước 2: Configure API Key**
1. Vào tab **⚙️ Settings**
2. Nhập OpenAI API key
3. Click **💾 Save API Settings**

### **Bước 3: Create Research**
1. Vào tab **🚀 New Research**
2. Enter research topic (VD: "Electric vehicles")
3. Upload Excel framework hoặc dùng default
4. Click **🚀 Start Research**
5. Theo dõi progress real-time
6. **📄 Download Word report** khi hoàn thành!

### **Bước 4: View Analytics**
- Tab **📊 Analytics**: Xem breakdown research
- Tab **📄 Export Reports**: Quản lý và download reports

---

## 💻 **Command Line - Workflow Truyền Thống**

```bash
# Chạy hệ thống CLI
python run_layered_research.py
```

**Menu CLI:**
1. **🚀 Complete Research Workflow** (Excel → JSON → Research → Ready)
2. **📄 Export Word Document** 
3. **👋 Thoát**

---

## 🏗️ **Kiến trúc hệ thống**

### **Smart Layered Research:**
- **Layer 1**: Framework categories (PESTEL, Porter's Five Forces)
- **Layer 2**: Sub-categories (Political, Economic, etc.)
- **Layer 3**: Main questions → **Overview Analysis** (tự động tạo cho tất cả)
- **Layer 4**: Sub-questions → **Comprehensive Reports** (tự động tạo nếu có sub-questions)

### **Auto-Enhancement Logic:**
- ✅ **Questions có sub-questions**: Tự động tạo comprehensive Layer 4 reports
- ✅ **Questions không có sub-questions**: Chỉ tạo Layer 3 analysis
- 🤖 **Workflow đơn giản**: 1 lệnh → báo cáo hoàn chỉnh

### **API Support:**
- ✅ **OpenAI GPT-3.5-turbo** (recommended, cost-effective)

---

## 🌟 **Web UI Features**

### **🏠 Home Dashboard**
- Overview metrics và system status
- Quick start guide
- Benefits highlight

### **🚀 New Research**
- **Form-based input**: Topic, market, framework upload
- **Research modes**: Complete analysis hoặc Quick test (5 questions)
- **Real-time progress tracking** với progress bar
- **Advanced options**: Custom research purpose
- **Instant download** Word + JSON results

### **📊 Analytics Dashboard**
- **Visual charts**: Layer breakdown, enhancement coverage
- **Metrics overview**: Questions, costs, timing
- **Data tables**: Detailed category analysis

### **📄 Export Management**
- **File browser**: Tất cả research files và Word reports
- **Quick actions**: One-click download
- **Export latest**: Instant Word generation from latest research

### **⚙️ Settings**
- **API configuration**: OpenAI key management
- **System info**: Environment status, file counts
- **Data management**: Clear session, create directories

---

## 📁 **Cấu trúc thư mục**

```
mktresearch/
├── 🌐 streamlit_app.py                   # Web UI main app
├── 🚀 launch_ui.py                       # UI launcher script
├── 📄 README.md                          # Hướng dẫn này
├── ⚙️ config.py                          # Cấu hình API keys
├── 🔄 excel_to_structured_json.py       # Excel → JSON converter  
├── 🤖 openai_market_research.py         # OpenAI research engine
├── 🎮 run_layered_research.py           # CLI system (traditional)
├── 📝 export_comprehensive_report.py    # Word export tool
├── 📥 input/                            # Input files (optional)
│   └── Research Framework.xlsx           # Framework mặc định
├── 🗂️ output/                           # All generated files
│   ├── market_research_structured.json   # Framework data (converted)
│   ├── layer3_research_*.json           # Research results
│   └── Báo_cáo_*.docx                   # Final reports
└── 🛠️ Supporting modules...
```

---

## 💰 **Chi phí và hiệu suất**

### **OpenAI GPT-3.5-turbo**:
- **Complete Workflow**: ~$0.15-0.25 
- **Time**: ~15-25 minutes (Excel → Word)

### **Performance:**
- **Web UI**: Point-and-click, no technical skills needed
- **CLI**: 1 lệnh cho toàn bộ workflow
- **Smart input**: Tự động hoặc custom Excel path
- **Auto Layer 4**: Chỉ tạo khi cần thiết

---

## 🎯 **Use Cases**

### **🏢 Marketing Teams**
```bash
# Super simple - Launch web UI
python launch_ui.py
# → Click upload Excel → Enter topic → Download report!
```

### **🔧 Technical Integration**
```bash
# Automated workflow
python run_layered_research.py
# → CLI workflow for batch processing
```

### **📊 Business Users**
- Upload research framework Excel files
- Get professional Word reports in 15-25 minutes
- No coding or AI knowledge required
- Real-time progress tracking

---

## 🔧 **Advanced Features**

### **Web UI Benefits:**
- **No technical knowledge required**
- **Visual progress tracking**
- **Analytics dashboard với charts**
- **File management interface**
- **API key management**
- **Error handling với user-friendly messages**

### **Auto File Management:**
- Tất cả output tự động lưu trong `output/` folder
- Tự động detect file research mới nhất
- Clean project structure

### **Smart Content Generation:**
- **Layer 3**: Overview analysis với context Việt Nam
- **Layer 4**: Comprehensive reports tích hợp tất cả sub-questions
- **Clean Export**: Simplified headings, professional format

### **Error Handling:**
- Graceful handling cho API errors
- Automatic retry logic
- Clear error messages

---

*🚀 Hệ thống này giúp bạn tạo ra báo cáo nghiên cứu thị trường chuyên nghiệp chỉ với vài clicks!* 