import camelot
import pandas as pd

# Extract tables from your PDF
tables = camelot.read_pdf(
    "Statement.pdf",
    pages="all",
    flavor="stream",
    strip_text="\n"
)

def clean_and_merge_rows(df):
    cleaned_rows = []
    current_row = None
    
    for _, row in df.iterrows():
        # If this is a continuation row (starts with R-0-0-0)
        if str(row[1]).startswith('R-0-0-0'):
            if current_row is not None:
                # Merge the item code and description
                current_row[1] = current_row[1] + " " + str(row[1])
                current_row[2] = current_row[2] + " " + str(row[2])
        else:
            # If we have a stored row, add it to our results
            if current_row is not None:
                cleaned_rows.append(current_row)
            current_row = row.tolist()
    
    # Don't forget to add the last row
    if current_row is not None:
        cleaned_rows.append(current_row)
    
    return pd.DataFrame(cleaned_rows, columns=df.columns)

with pd.ExcelWriter("output.xlsx", engine='openpyxl') as writer:
    if len(tables) == 0:
        pd.DataFrame({"Info": ["No data extracted"]}).to_excel(writer, sheet_name="Sheet1", index=False)
    else:
        # Combine all tables
        combined_df = pd.concat([table.df for table in tables])
        
        # Clean and merge rows
        cleaned_df = clean_and_merge_rows(combined_df)
        
        # Write to Excel
        cleaned_df.to_excel(writer, sheet_name="Farmer_Statement", index=False)

print("Conversion complete! File saved as output.xlsx")