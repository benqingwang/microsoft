import docx
import pandas as pd

def read_table_from_docx(filepath):
    # Load the Word Document
    doc = docx.Document(filepath)

    # Initialize a list to hold all rows of data
    data = []

    # Assume the first table is the one we need
    table = doc.tables[0]

    # Read each row in table
    for row in table.rows:
        row_data = [cell.text for cell in row.cells]
        data.append(row_data)

    # Convert list to DataFrame
    df = pd.DataFrame(data)

    # If you want to use the first row as a header
    new_header = df.iloc[0]  # Grab the first row for the header
    df = df[1:]  # Take the data less the header row
    df.columns = new_header  # Set the header row as the df header

    return df

# Example usage
filepath = 'path_to_your_document.docx'
df = read_table_from_docx(filepath)
print(df)
