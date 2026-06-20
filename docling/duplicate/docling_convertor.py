from docling.document_converter import DocumentConverter
import json
import pandas as pd
from pathlib import Path

# Install: pip install docling

converter = DocumentConverter()
source = "PdfDownload_new.pdf"
result = converter.convert(source)

result_dict = result.document.export_to_dict()
print(json.dumps(result_dict, indent=2))

# Extract tables
for table_idx, table in enumerate(result.document.tables):
    df = table.export_to_dataframe()
    print(f"Table {table_idx}:\n", df)
    df.to_csv(f"{table_idx}.csv")
    with open(f"{table_idx}.html", "w") as fp:
        fp.write(table.export_to_html())
