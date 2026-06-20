from docling.document_converter import DocumentConverter

#source = "https://arxiv.org/pdf/2408.09869"  # local path or URL
source = "Chennai South - 2019 LS Votes.xlsx"
converter = DocumentConverter()
result = converter.convert(source)
doc = result.document

print(doc.export_to_markdown())  # outputs a Markdown-formatted version of the document
