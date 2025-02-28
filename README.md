# Web Application Mimicking Google Sheets
This project is a web application that mimics the user interface and core functionalities of Google Sheets. The application is built using HTML, CSS, and JavaScript and supports key features such as mathematical functions,  dynamic cell interactions, and more.

---
Live Demo : https://github.com/RamuMiryala/Google-sheet-Clone

## Features

### 1. Spreadsheet Interface
- **Dynamic Grid:** A dynamic spreadsheet grid with editable cells.
- **Headers:** Column and row headers to navigate cells easily.
- **Cell Interaction:** Supports adding, deleting, and resizing rows and columns dynamically.
- **Formatting Options:** Apply bold, italic, underline,standard and text color formatting to cells.
- **Drag-and-Drop:** Copy formulas and content by dragging across cells.

### 2. Mathematical Functions
The following mathematical functions are supported:
- **SUM:** Calculate the sum of a range of cells (e.g., `=SUM(A1:A5)`).
- **AVERAGE:** Calculate the average of a range of cells (e.g., `=AVERAGE(A1:A5)`).


### 3. Data Quality Functions

- **UPPER:** Convert text in a cell to uppercase.
- **LOWER:** Convert text in a cell to lowercase.
- **Captailze:** Fisrt letter capital.

---

## Technologies Used

### Frontend
- **HTML:** Structure of the application.
- **CSS:** Styling and layout.
- **JavaScript:** Functionality, formula parsing, and dynamic updates.

### Libraries and Tools
- **[Chart.js](https://www.chartjs.org/):** For data visualization.
- **[SheetJS (xlsx)](https://sheetjs.com/):** For saving and loading Excel files.

---

## Installation and Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/RamuMiryala/Google-sheet-Clone
   ```

2. Navigate to the project directory:
   ```bash
   cd Google-sheet clone
   ```

3. Open `index.html` in a browser to run the application.

---

## Usage

1. **Editing Cells:** Click on any cell to edit its content or enter formulas.
2. **Applying Formulas:** Use the formula bar to enter mathematical or data quality functions (e.g., `=SUM(A1:A5)`).
3. **Formatting Cells:** Use the toolbar to apply text formatting and colors.
4. **Saving Data:** Click the `Save Spreadsheet` button to download the spreadsheet as an Excel file.
5. **Charts:** Select data and click `Create Chart` to generate a visual chart.
6. **Testing:** Click `Run Tests` to execute predefined test cases.

---

## File Structure

```
/
├── index.html        # Main HTML file
├── styles.css        # CSS file for styling
├── script.js         # JavaScript file for functionality
└── README.md         # Project documentation (this file)
```

---

## Future Enhancements

1. **Complex Formulas:** Add support for relative and absolute cell references (e.g., `$A$1`).
2. **Auto-Update Dependencies:** Ensure formulas update automatically when dependent cells change.
3. **Advanced Charting:** Allow users to customize chart types and ranges.
4. **Improved Data Validation:** Add stricter validation rules for specific data types.

