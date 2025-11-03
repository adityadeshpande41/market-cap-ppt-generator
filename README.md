# Market Capitalization PPT Generator

This tool allows users to upload multiple Excel financial models and a PowerPoint template. It automatically extracts market capitalization data from the Excel files and populates a table in the provided PowerPoint presentation.

---

## Features
- Upload multiple Excel files and a PPTX template.
- Extracts the "Market Cap" and "Ticker" from Excel sheets.
- Fills the provided PowerPoint table with extracted data.
- Skips files that do not have a "Market Cap" column.
- Generates a downloadable, updated presentation.

---

## How to Use

1. **Install the requirements:**

```bash
pip install -r requirements.txt


2. Run the Streamlit App:
streamlit run app.py



3. On the Web App:

Upload all your Excel models (.xlsx).

Upload the PowerPoint template (.pptx).

Click on Generate Presentation.

Download the updated .pptx file.



4. Input File Requirements
Excel Files: Must contain a column named Ticker and Market Cap.

PowerPoint File: Should have a table where the extracted data will be inserted.

⚠️ Files missing required columns will be skipped with a warning.


5. Output
Updated_Presentation.pptx: A PowerPoint presentation with the extracted market cap data populated.

Author
Aditya Deshpande
