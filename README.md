# 📊 T-Mobile Stock Data PowerPoint Generator

This Python project automatically generates a stylish PowerPoint presentation that summarizes **T-Mobile (TMUS)** stock data. It reads data from an Excel file, formats it into a visually appealing table, and creates a clean 3-slide `.pptx` file using `python-pptx`.

---

## 🖼️ Slides Overview

1. **Welcome Slide**  
   - Displays the T-Mobile logo and a brief company overview.  
   - Custom fonts and brand colors used for professional styling.

2. **Stock Analysis Slide**  
   - Reads stock data from `tmobileStock.xlsx`.  
   - Displays the data in a formatted table with alternating row colors.  
   - Uses `pandas` for data handling and `python-pptx` for rendering.

3. **Thank You Slide**  
   - Full-screen magenta background.  
   - Center-aligned thank-you message in clean white text.

---

## 📁 Files Needed

Make sure the following files are in the same directory as the script:

- `tmobileStock.xlsx` — Excel file containing the stock data  
- `tMobile.jpg` — T-Mobile logo image for the welcome slide  

---

## 📦 Requirements

Install the required Python libraries with:

```bash
pip install pandas python-pptx openpyxl
