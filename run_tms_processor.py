#!/usr/bin/env python3
"""
TMS Data Processor Pro - Launcher
Professional launcher for the modern TMS data processing application
"""
import sys
import os

# Add current directory to path to ensure imports work
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def check_dependencies():
    """Check if all required packages are installed"""
    required_packages = ['pandas', 'openpyxl', 'numpy']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    return missing_packages

def install_dependencies():
    """Install missing dependencies"""
    print("Installing missing dependencies...")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("Dependencies installed successfully!")
        return True
    except Exception as e:
        print(f"Failed to install dependencies: {e}")
        return False

def main():
    print("🚛 TMS Data Processor Pro")
    print("=" * 40)
    
    # Check dependencies
    missing_packages = check_dependencies()
    
    if missing_packages:
        print(f"Missing packages: {', '.join(missing_packages)}")
        print("Attempting to install dependencies...")
        
        if not install_dependencies():
            print("\nPlease install the required packages manually:")
            print("pip install pandas openpyxl numpy")
            input("Press Enter to exit...")
            return
    
    try:
        from tms_processor import main as start_app
        print("Starting TMS Data Processor Pro...")
        start_app()
        
    except ImportError as e:
        print(f"Error importing required modules: {e}")
        print("Please ensure all required packages are installed:")
        print("pip install -r requirements.txt")
        input("Press Enter to exit...")
        
    except Exception as e:
        print(f"Error starting application: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()