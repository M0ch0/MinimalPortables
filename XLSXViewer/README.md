# XLSX Viewer

![Screenshot](https://github.com/user-attachments/assets/709c3c07-e8e5-43e9-a392-dc7451c93d91)

This is a simple XLSX viewer that allows you to view and interact with spreadsheet data directly in your web browser. It's built using JavaScript and the SheetJS library (xlsx.min.js).

## Features

* **The minimum functionality required for a portable device:** Multiple sheets, range selection, copy with header, font size adjustment, dark mode, and reset view.
* **Light load and lightweight:** The entire project is less than 500KB.

## How to Use

1. **Installation:**
    * **Option 1: Open index.html**
        Open index.html in your default browser, even if it's on your phone, or without internet connection, this works.
    * **Option 2: Place index.html on your server**
        Also works.
        

2. **View and Interact:**
    * Once the file is loaded, use the sheet selector to choose the sheet you want to view.
    * You can enable range selection to select and copy data from specific cells.
    * You can adjust the font size using the slider.
    * You can toggle dark mode using the "Toggle Dark Mode" button.
    * You can use the "Reset View" button to clear the current sheet.

<br>

> **Note:** This project is licensed under the **GNU Affero General Public License Version 3 (AGPLv3)**. See the LICENSE file for details.

<br>

> **Embedded Library:** This project utilizes the `xlsx.min.js` library located in the `/assets/js/` directory, which is licensed under the **Apache License 2.0**.
