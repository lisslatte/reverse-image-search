# I-Check (Image Similarity Search)

I-Check is a tool designed to find similar images by uploading an image (JPG, PNG) or document (PDF, Word, PPTX). The tool utilizes Structural Similarity Index (SSIM) and color histogram similarity metrics to identify and display similar images along with their metadata.

## Features

- Upload Images and Documents: Supports JPG, PNG, PDF, Word, and PPTX files.
- Similarity Search: Finds and displays similar images using SSIM and color histogram similarity metrics.
- Image Extraction: Extracts images from PDF, Word, and PPTX documents.
- Metadata Display: Shows detailed metadata for each image.
- Encryption and Hashing: Includes functionality for file encryption and SHA-256 hashing.
- Logging: Logs errors to `app.log` for troubleshooting.
- GUI: Built with Tkinter for a user-friendly interface.

## Requirements

- Python 3.6+
- Packages: 
  - numpy
  - opencv-python
  - scikit-image
  - scikit-learn
  - tkinter
  - pillow
  - python-docx
  - PyPDF2
  - python-pptx
  - pymupdf
  - cryptography
  - sympy

## How to Run

- Open the folder in your chosen IDE such as Visual Studio Code. Then run the FYP (I-CHECK)_TP062285.py file.
- Wait until it installs all the required libraries and import. Once installed you can start to use the app.