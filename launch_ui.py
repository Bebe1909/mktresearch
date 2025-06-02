#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🚀 Market Research UI Launcher
Quick launcher for the Streamlit web interface
"""

import subprocess
import sys
import os

def main():
    """Launch the Streamlit web UI"""
    
    print("🔬 Starting Market Research Automation UI...")
    print("="*50)
    
    # Check if streamlit is installed
    try:
        import streamlit
        print("✅ Streamlit found")
    except ImportError:
        print("❌ Streamlit not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "streamlit", "streamlit-option-menu", "plotly"])
        print("✅ Streamlit installed")
    
    # Check output directory
    if not os.path.exists('output'):
        os.makedirs('output')
        print("✅ Output directory created")
    
    # Launch Streamlit
    print("\n🚀 Launching web interface...")
    print("📱 The app will open in your default browser")
    print("🔗 URL: http://localhost:8501")
    print("\n💡 Press Ctrl+C to stop the server")
    print("="*50)
    
    try:
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", "streamlit_app.py",
            "--server.port=8501",
            "--server.address=localhost",
            "--server.headless=false"
        ])
    except KeyboardInterrupt:
        print("\n👋 Shutting down server...")
    except Exception as e:
        print(f"❌ Error launching UI: {e}")
        print("\n💡 Try running manually: streamlit run streamlit_app.py")

if __name__ == "__main__":
    main() 