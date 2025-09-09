# XML & XLSX Migration to MongoDB

This tool migrates `.xml` and `.xlsx` files from a directory structure into MongoDB collections.

---

## 📌 Features

- Supports both **XML** and **XLSX** files.
- Directory → Collection mapping:
  - **Top-level folder** → `category` → MongoDB **collection name**.
  - **Subfolder** → `orderLine` value.
  - **File name (without extension)** → `apiName`.
- XML:
  - Stored raw in the field `Content`.
- XLSX:
  - If [`openpyxl`](https://pypi.org/project/openpyxl/) is installed → parsed into JSON-like structure in `ParsedContent`.
  - Otherwise → stored as Base64 string in `Content_base64`.
- Prevents duplicates:
  - Creates a **unique compound index** on `(category, orderLine, apiName)`.
  - Skips inserting if a matching document already exists.
- Verbose output and **dry-run mode** (`DO_INSERT = False` in the script).

---

## 📦 Requirements

- **Python 3.9+**
- **MongoDB** running locally or remotely
- Required Python packages:
  ```bash
  pip install pymongo
  pip install openpyxl   # optional, but recommended for XLSX parsing


testData/
│
├── inputXML/                # category = "inputXML" -> collection "inputXML"
│   ├── 1/                   # orderLine = "1"
│   │   └── createOrder.xml  # apiName = "createOrder"
│   └── 2/
│       └── updateOrder.xml
│
├── responseXlsx/            # category = "responseXlsx"
│   └── 1/
│       └── orders.xlsx

How to Run

BASE_DIR = Path(r"D:/migratetomongo/AutomationTool/src/test/resources/testData")
MONGO_URI = "mongodb://localhost:27017/"
DB_NAME = "te"
DO_INSERT = True   # Set False to dry-run (no writes)

python xmljsontomongo.py


