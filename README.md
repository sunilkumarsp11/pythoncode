# XML & XLSX Migration to MongoDB

This tool migrates `.xml` and `.xlsx` files from a directory structure into MongoDB collections.

## Features

- Supports both **XML** and **XLSX** files.
- Directory → Collection mapping:
  - **Top-level folder** = `category` = MongoDB collection name.
  - **Subfolder** = `orderLine` value.
  - **File name (without extension)** = `apiName`.
- XML:
  - Stored raw in the field `Content`.
- XLSX:
  - If [openpyxl](https://openpyxl.readthedocs.io/) is installed → parsed into JSON-like data per sheet (`ParsedContent`).
  - Otherwise → stored as Base64 in `Content_base64`.
- Prevents duplicates:
  - Creates a **unique compound index** on `(category, orderLine, apiName)`.
  - If a file with the same keys exists, it is skipped (no overwrite).
- Verbose output and **dry-run mode** (`DO_INSERT=False`).

## Example Document (XML)

```json
{
  "category": "inputXML",
  "orderLine": "1",
  "Content": "<Order ...>...</Order>",
  "ParsedContent": null,
  "apiName": "createOrder",
  "fileType": "xml"
}
