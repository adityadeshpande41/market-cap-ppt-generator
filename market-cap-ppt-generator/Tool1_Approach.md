Steps:

1)Extract Input Data:

2) Load all Excel model files.

3) Parse each file to locate the Market Cap field.

4) Standardize the data into a common format (ticker, market cap in USD).

5) Prepare the PowerPoint Template:

6) Load the provided PowerPoint presentation (.pptx).

7) Identify the target slide and locate the table object.

8) Populate the Table:

For each company, add a new row:

Insert ticker name into the first cell.

Insert market cap (formatted in USD) into the second cell.

9) Save the Output:

Save the updated presentation as a new .pptx file.

Ensure proper formatting is retained (currency, alignment, font).

10) Error Handling:

If market cap is missing for a company, skip or notify the user.

Validate that all Excel files have the necessary fields.

11) User Interface:

Build a simple Streamlit interface to upload:

Multiple Excel files

One PowerPoint template

Provide a Download button to get the final PPT.

12) Optional Enhancements:

Allow editing the title or footnotes dynamically.

Support other currencies (e.g., EUR, GBP).

