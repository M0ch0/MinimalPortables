# XLSX Sanitizer
![Screenshot](https://github.com/user-attachments/assets/66eccef7-dda3-452e-9f77-5f5c51c7f89c)
This project provides a simple tool to sanitize XLSX (Microsoft Excel) files by removing potentially harmful content such as:

* **Formulas:** All formulas are removed, leaving only the resulting values.
* **Macros:** VBA macros are removed to prevent malicious code execution.
* **External Links:** Links to external files or websites are removed.
* **Embedded Objects:** Embedded objects like images or OLE objects are removed.
* **Metadata:** Author, last author, creation date, and modification date are removed from the file's metadata.

## Installation
* **Option 1: Open index.html.**
        Open index.html in your default browser, even if it's on your phone, or without internet connection, this will works.
* **Option 2: Place this folder on your server.**
        Also works.

## How to Use

1. **Upload your XLSX file:** Click the "Upload your XLSX file" button or drag and drop your file onto the designated area.
2. **Sanitize:** Click the "Sanitize File" button.
3. **Download:** A sanitized version of your file will be generated and a download link will appear. Click the link to download the sanitized file.

<br>

> **Note:** This project is licensed under the **GNU Affero General Public License Version 3 (AGPLv3)**. See the LICENSE file for details.

<br>

> **Embedded Library:** This project utilizes the `xlsx.full.min.js` library located in the `/assets/js/` directory, which is licensed under the **Apache License 2.0**.
