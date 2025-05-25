#!/usr/bin/env python3
"""
Font Installer for Document Text Extractor
This script installs the required Arabic and Urdu fonts for the Document Text Extractor.
Run this script once before using the main script to ensure all fonts are installed.
"""

import os
import sys
import subprocess
import tempfile

def install_fonts():
    """Install required Arabic/Urdu fonts"""
    print("Installing Arabic and Urdu fonts for Document Text Extractor...")
    
    # Define font paths
    user_fonts_dir = os.path.expanduser("~/Library/Fonts")
    
    # Create fonts directory if it doesn't exist
    os.makedirs(user_fonts_dir, exist_ok=True)
    
    # Define required fonts and their direct download URLs
    required_fonts = {
        "Amiri-Regular.ttf": "https://github.com/alif-type/amiri/raw/master/amiri-regular.ttf",
        "NotoSansArabic-Regular.ttf": "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansArabic/NotoSansArabic-Regular.ttf",
        "NotoNastaliqUrdu-Regular.ttf": "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoNastaliqUrdu/NotoNastaliqUrdu-Regular.ttf"
    }
    
    # Check which fonts are missing
    missing_fonts = {}
    for font_file, url in required_fonts.items():
        if not os.path.exists(os.path.join(user_fonts_dir, font_file)):
            missing_fonts[font_file] = url
    
    if not missing_fonts:
        print("All required Arabic and Urdu fonts are already installed.")
        return True
    
    print(f"Installing {len(missing_fonts)} missing Arabic/Urdu fonts...")
    
    # Try to import requests
    try:
        import requests
        has_requests = True
    except ImportError:
        has_requests = False
        print("Note: 'requests' module not found. Using curl for downloads.")
    
    # Download missing fonts
    success_count = 0
    for font_file, url in missing_fonts.items():
        font_path = os.path.join(user_fonts_dir, font_file)
        print(f"Downloading {font_file} from {url}...")
        
        try:
            # Method 1: Using requests if available
            if has_requests:
                try:
                    response = requests.get(url, timeout=30)
                    response.raise_for_status()
                    
                    with open(font_path, 'wb') as f:
                        f.write(response.content)
                        
                    print(f"Successfully installed {font_file}")
                    success_count += 1
                    continue
                except Exception as e:
                    print(f"Requests download failed for {font_file}, trying curl: {e}")
            
            # Method 2: Using curl as fallback
            result = subprocess.run(
                ["curl", "-L", url, "-o", font_path, "--connect-timeout", "10", "--max-time", "30"],
                check=True, capture_output=True, text=True
            )
            print(f"Successfully installed {font_file} using curl")
            success_count += 1
            
        except subprocess.CalledProcessError as e:
            print(f"Error installing {font_file}: {e}")
            print(f"stdout: {e.stdout}")
            print(f"stderr: {e.stderr}")
        except Exception as e:
            print(f"Unexpected error installing {font_file}: {e}")
    
    # Verify installation
    if success_count == len(missing_fonts):
        print("All required Arabic and Urdu fonts have been successfully installed.")
        
        # On macOS, clear font cache
        try:
            subprocess.run(["atsutil", "databases", "-removeUser"], 
                          check=False, capture_output=True)
            print("Font cache cleared.")
        except Exception as e:
            print(f"Note: Could not clear font cache: {e}")
        
        return True
    else:
        print(f"Installed {success_count}/{len(missing_fonts)} required fonts.")
        print("Some fonts may need to be installed manually.")
        return False

if __name__ == "__main__":
    print("Font Installer for Document Text Extractor")
    print("=========================================")
    
    success = install_fonts()
    
    if success:
        print("\nAll fonts installed successfully!")
        print("You can now run the Document Text Extractor script.")
        sys.exit(0)
    else:
        print("\nSome fonts could not be installed automatically.")
        print("You may need to install them manually:")
        print("1. Amiri Regular: https://github.com/alif-type/amiri/raw/master/amiri-regular.ttf")
        print("2. Noto Sans Arabic: https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansArabic/NotoSansArabic-Regular.ttf")
        print("3. Noto Nastaliq Urdu: https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoNastaliqUrdu/NotoNastaliqUrdu-Regular.ttf")
        print("\nDownload these files and place them in your ~/Library/Fonts directory.")
        sys.exit(1)
