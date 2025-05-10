# Exampeer Blog Automation 🤖📝

**Exampeer Blog Automation** is a Python-based tool designed to automate the process of scraping data from a DOCX file, parsing it, converting it into an Excel file, and uploading the data to the Exampeer website on a daily basis. This tool helps streamline the content posting process for Exampeer's blog.

---

## 🧰 Features

- 📄 Scrapes data from DOCX files and converts it into an **Excel sheet**  
- 📊 Automatically uploads the parsed data to the **Exampeer website**  
- 🕒 Automates daily operations for blog content management  
- 💻 Handles error notifications with clear images and icons  
- 📝 Generates detailed reports in `report.xlsx`

---

## 📁 Project Structure

ExampeerBlogAutomation/  
│  
├── PostBlog.py            # Main script to scrape and upload blog data  
├── analysed.xlsx          # Excel file with parsed blog data  
├── autoAnalyser.txt       # Text file for analyzing the data  
├── autoDetector.txt       # File used for detecting relevant data  
├── autoUploader.txt       # Contains credentials and settings for auto-upload  
├── bg.png                 # Background image for notifications  
├── error.ico              # Error icon for notifications  
├── error.jpg              # Error image for notifications  
├── icon.ico               # App icon for UI or notifications  
├── notification.jpg       # Notification image shown during processes  
├── report.xlsx            # Excel report generated after parsing and upload  
├── rough.py               # Script for testing or debugging  
├── rough2.py              # Another testing script  
├── sample.xlsx            # Sample input Excel file for testing  
├── showBrowser.txt        # Configuration for showing browser during upload  
├── temp.txt               # Temporary file for intermediate data  
└── README.md              # Project documentation (this file)

---

## 🛠️ Requirements

- Python 3.7+  
- python-docx  
- pandas  
- requests  
- selenium (if using browser automation)  

Install dependencies:
``` bash
  pip install python-docx pandas requests selenium
```
## 🖥️ How to Use

1. Place your **DOCX file** with blog content in the working directory.  
2. Configure any required settings in `autoUploader.txt` (e.g., credentials or upload settings).  
3. Run the script:
``` bash
  python PostBlog.py
```

4. The script will:
   - Scrape and parse the DOCX file into an Excel sheet.
   - Automatically upload the parsed data to the website.
   - Generate a detailed report in `report.xlsx`.

5. You can check the `report.xlsx` for any errors or issues with the parsing process.

---

## 🔍 How It Works

- **PostBlog.py** is the main script that:
  - Reads the DOCX file and extracts relevant content.
  - Converts the content into an Excel file using **pandas**.
  - Automatically uploads the Excel data to the **Exampeer website** using configured credentials and settings.
  
- If there are errors, notification images (like `error.ico` and `error.jpg`) will be shown.

---

## ⚠️ Disclaimer

This tool is intended for **internal business automation purposes** only and may not be suitable for general public use without modifications.  
Make sure to comply with your website's terms of service before automating any data scraping and uploading processes.

---

## 📄 License

This project is licensed under the MIT License.

---

## 🙋‍♂️ Author

Developed by @princemehta-git  
https://github.com/princemehta-git


