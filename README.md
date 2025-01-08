OCR Text Extraction and PowerPoint Integration

📚 Overview

This Python script extracts text from images using Tesseract OCR and integrates the extracted text into a PowerPoint presentation. It processes images from a specified folder, performs OCR on each image, and appends the recognized text to slides in a PowerPoint file.

🚀 Features
	•	Extracts text from .jpeg and .jpg images using Tesseract OCR.
	•	Converts images to grayscale for better OCR accuracy.
	•	Organizes extracted text into a PowerPoint presentation.
	•	Supports adding text to an existing presentation or creating a new one using a template.
	•	Processes images in numerical order for consistency.

🛠️ Dependencies

Ensure the following dependencies are installed:
	•	Python 3.x
	•	OpenCV (opencv-python)
	•	Tesseract OCR (pytesseract)
	•	python-pptx (python-pptx)

Install Required Libraries:

pip install opencv-python pytesseract python-pptx

Install Tesseract OCR:
	•	macOS: brew install tesseract
	•	Ubuntu: sudo apt install tesseract-ocr
	•	Windows: Download Tesseract OCR

📂 Folder Structure

/UOB_OD1
│
├── Image/
│   ├── AddonJan08/
│   │   ├── image1.jpg
│   │   ├── image2.jpg
│   │   └── ...
│
├── Sample.pptx      # Template PowerPoint file
├── AddonJan08_extract.pptx  # Output file
├── main.py          # The script
└── README.md        # Documentation

⚙️ Configuration

Update Paths in the Script:

Modify the main() function:

date = "AddonJan08"
image_folder_path = f"/path_to_your_images/{date}"
pptx_path = f"/path_to_output/{date}_extract.pptx"
template_path = "/path_to_template/Sample.pptx"

▶️ How to Run the Script
	1.	Place your images in the Image/<date> folder.
	2.	Ensure Sample.pptx exists as the template.
	3.	Run the script:

python main.py

	4.	Check the output PowerPoint file (AddonJan08_extract.pptx) for extracted text.

📊 Functions Explained

1. ocr_extract_text_pysseract(image_path)
	•	Converts the image to grayscale.
	•	Extracts text using pytesseract.
	•	Returns extracted text.

2. add_extracted_text_to_powerpoint(pptx_path, template_path, text)
	•	Opens an existing PowerPoint file or uses a template.
	•	Adds extracted text to a new slide.
	•	Saves the updated presentation.

3. numerical_sort(value)
	•	Ensures files are processed in numerical order.

4. main()
	•	Iterates through images in the folder.
	•	Calls OCR and PowerPoint functions sequentially.

🐞 Troubleshooting
	•	TesseractNotFoundError: Ensure Tesseract OCR is installed and added to the system path.
	•	Image Read Errors: Verify image paths and formats.
	•	PowerPoint Template Missing: Ensure Sample.pptx exists.

🤝 Contributing

Feel free to fork the repository and submit pull requests with improvements.

📜 License

This project is licensed under the MIT License.

📧 Contact

For any questions or feedback, please reach out via email or open an issue.