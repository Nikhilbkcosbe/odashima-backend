import os
import sys
from pathlib import Path

# Add the parent directory to Python path
sys.path.append(str(Path(__file__).parent.parent))

import pytesseract
from PIL import Image

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def test_tesseract_setup():
    """Test if Tesseract OCR is properly set up."""
    try:
        # Check if Tesseract is installed and accessible
        version = pytesseract.get_tesseract_version()
        print(f"Tesseract version: {version}")
        
        # Create a simple test image with text
        img = Image.new('RGB', (200, 50), color='white')
        img.save('test.png')
        
        # Try to perform OCR on the test image
        text = pytesseract.image_to_string('test.png')
        print("OCR test successful!")
        
        # Clean up test file
        os.remove('test.png')
        
        return True
    except Exception as e:
        print(f"Error testing Tesseract setup: {str(e)}")
        return False

if __name__ == "__main__":
    test_tesseract_setup() 