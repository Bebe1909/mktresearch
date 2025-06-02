# 🚀 **Deployment Guide - Market Research Automation**

## **Quick Deploy to Streamlit Cloud (Free)**

### **Step 1: Setup GitHub Repository**

```bash
# Initialize git repository
git init
git add .
git commit -m "Initial commit: Market Research Automation System"

# Create GitHub repository (go to github.com and create new repo)
# Then connect local repo to GitHub:
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git branch -M main
git push -u origin main
```

### **Step 2: Deploy to Streamlit Cloud**

1. **Go to [share.streamlit.io](https://share.streamlit.io)**
2. **Sign in with GitHub**
3. **Click "New app"**
4. **Select your repository**
5. **Set these values:**
   - **Branch:** `main`
   - **Main file path:** `streamlit_app.py`
   - **App URL:** `your-app-name` (optional custom URL)

### **Step 3: Add API Key Secret**

1. **In Streamlit Cloud dashboard**, click your app
2. **Go to "Settings" → "Secrets"**
3. **Add this content:**
```toml
OPENAI_API_KEY = "your-actual-openai-api-key-here"
```
4. **Save and restart app**

### **Step 4: Done! 🎉**

Your app will be live at: `https://your-app-name.streamlit.app`

---

## **Alternative: Railway Deployment**

If you prefer Railway:

1. **Go to [railway.app](https://railway.app)**
2. **Connect GitHub repository**
3. **Add environment variable:**
   - `OPENAI_API_KEY` = your API key
4. **Railway auto-detects Streamlit and deploys**

---

## **Local Development**

```bash
# Install dependencies
pip install -r requirements.txt

# Run locally
python launch_ui.py
# or
streamlit run streamlit_app.py
```

---

## **Features Ready for Production:**

✅ **Responsive web interface**  
✅ **File uploads (Excel frameworks)**  
✅ **Real-time progress tracking**  
✅ **Professional Word report generation**  
✅ **Analytics dashboard**  
✅ **API key management**  
✅ **Error handling**  
✅ **Vietnamese market research focus**

---

## **Cost Estimation:**

- **Hosting:** Free (Streamlit Cloud)
- **OpenAI API:** ~$0.15-0.25 per complete research report
- **Total:** Essentially free for moderate usage

---

*🔬 Perfect for marketing teams who need professional market research reports without technical complexity!* 