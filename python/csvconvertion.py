import pandas as pd
from docx import Document
from docx.shared import Pt

def csv_to_docx(csv_file, docx_file):
    # Read the CSV file
    df = pd.read_csv(csv_file)
    
    # Create a new Document
    doc = Document()
    doc.add_heading('CSV Data', 0)

    # Add table to document
    table = doc.add_table(rows=1, cols=len(df.columns))

    # Add the header row
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = column
    
    # Add the rows from the dataframe
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value)
    
    # Optional: Change font size for the table
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)

    # Save the document
    doc.save(docx_file)
    print(f"CSV data has been converted and saved to {docx_file}")

# Example usage
csv_file_path = 'example.csv'
docx_file_path = 'output.docx'
csv_to_docx(csv_file_path, docx_file_path)
