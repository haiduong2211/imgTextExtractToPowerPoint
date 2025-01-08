# **OCR Text Extraction and PowerPoint Integration**

## 📚 **Overview**  

This Python script extracts text from images using **Tesseract OCR** and integrates the extracted text into a **PowerPoint presentation**. It processes images from a specified folder, performs OCR on each image, and appends the recognized text to slides in a PowerPoint file.  

---

## 🚀 **Features**  

- Extracts text from `.jpeg` and `.jpg` images using **Tesseract OCR**.  
- Converts images to **grayscale** for better OCR accuracy.  
- Organizes extracted text into a **PowerPoint presentation**.  
- Supports adding text to an **existing presentation** or creating a **new one using a template**.  
- Processes images in **numerical order** for consistency.  

---

## 🛠️ **Dependencies**  

Ensure the following dependencies are installed:  

- **Python 3.x**  
- **OpenCV** (`opencv-python`)  
- **Tesseract OCR** (`pytesseract`)  
- **python-pptx** (`python-pptx`)  

### **Install Required Libraries:**  
```bash
pip install opencv-python pytesseract python-pptx
```

### Install Tesseract OCR:
	•	macOS: brew install tesseract
	•	Ubuntu: sudo apt install tesseract-ocr
	•	Windows: Download Tesseract OCR (https://github.com/tesseract-ocr/tesseract)


## 📂 Folder Structure

/ImageToPowerpoint
├── Input/
│   ├── Image1/
│   │   ├── image1.jpg
│   │   ├── image2.jpg
│   │   └── ...
├── Sample.pptx          # Template PowerPoint file
├── image1_extract.pptx  # Output file
├── extract_text.py      # The script
└── README.md            # Documentation


## ⚙️ Configuration

Update paths in the script:
```python
date = "AddonJan08"
image_folder_path = f"/path_to_your_images/{date}"
pptx_path = f"/path_to_output/{date}_extract.pptx"
template_path = "/path_to_template/Sample.pptx"
```


## ▶️ How to Run the Script
1.	Place your images in the Image/<date> folder.
2.	Ensure Sample.pptx exists as the template.
3.	Run the script:
```py
python extract_text.py
```

## 🐞 Troubleshooting
•	TesseractNotFoundError: Ensure Tesseract OCR is installed and added to the system path.
•	Image Read Errors: Verify image paths and formats.
•	PowerPoint Template Missing: Ensure Sample.pptx exists.

## 📧 Contact
For any questions or feedback, please reach out via email or open an issue.