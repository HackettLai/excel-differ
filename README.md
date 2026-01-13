# Excel Differ 📊

A web-based Excel file comparison tool that helps you identify differences between two Excel files quickly and easily.

![Excel Differ](https://img.shields.io/badge/version-1.1.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Static Badge](https://img.shields.io/badge/AI%20Assist-Claude%20Sonnet%204.5-orange)

| <img width="600"  alt="screenshort-edited-1" src="https://github.com/user-attachments/assets/2b3b1061-5a6e-4c22-bf84-3aa64f44ed77" />  | <img width="600"  alt="screenshort-edited-2" src="https://github.com/user-attachments/assets/6fc5efb0-dc1a-44f2-95d2-e71ba522be6b" /> | <img width="600"  alt="screenshort-edited-3" src="https://github.com/user-attachments/assets/08de3c2a-baa9-426d-990e-35984c9c65dd" /> |
|:---:|:---:|:---:|
| <img width="600"  alt="screenshort-edited-4" src="https://github.com/user-attachments/assets/87b570f6-2067-44d7-bbe4-0c30a56e1c80" />  | <img width="600"  alt="screenshort-edited-5" src="https://github.com/user-attachments/assets/abffa303-e828-403f-9d5b-823d4cdd81bb" /> |  |


## Features ✨

- **📁 Drag & Drop Support** - Simply drag and drop Excel files to compare
- **🔍 Sheet-Level Comparison** - Identifies added, removed, renamed, and modified sheets
- **📊 Cell-Level Diff** - Highlights individual cell changes with detailed tooltips
- **🔄 Smart Rename Detection** - Automatically detects renamed sheets based on content similarity
- **↔️ Synchronized Scrolling** - Side-by-side view with synchronized horizontal and vertical scrolling
- **🎯 Change Navigation** - Jump between changes with Previous/Next buttons or keyboard shortcuts (P/N)
- **💡 Visual Indicators** - Color-coded cells for easy identification of changes
- **📱 Responsive Design** - Works on desktop and mobile devices
- **🚀 No Server Required** - Runs entirely in your browser

## Demo 🎬

[Live Demo](https://excel-differ.hackettlai.com)

## Supported File Formats 📋

- `.xlsx` - Excel 2007+ files
- `.xls` - Excel 97-2003 files
- `.csv` - Comma-separated values files

**File Size Limit:** 50MB per file

## Getting Started 🚀

### Prerequisites

No installation required! Just a modern web browser:

- Chrome (recommended)
- Firefox
- Safari
- Edge

### Installation

1. Clone the repository:

```bash
git clone https://github.com/HackettLai/excel-differ.git
cd excel-differ
```

2. Open `index.html` in your web browser:

```bash
# On macOS
open index.html

# On Linux
xdg-open index.html

# On Windows
start index.html
```

That's it! No build process or dependencies to install.

## Usage 📖

### Basic Workflow

1. **Upload Files**

   - Click "Choose File" or drag & drop your Excel files into the upload areas
   - File A: Original/older version
   - File B: New/updated version

2. **Compare**

   - Click the "Compare Files" button
   - Wait for the comparison to complete

3. **Review Summary**

   - View the summary of all sheet changes
   - See statistics for added, removed, modified, and renamed sheets
   - Click on any sheet to view detailed cell-level differences

4. **Inspect Details**
   - Navigate between sheets using the tab bar
   - Hover over highlighted cells to see old vs new values
   - Use synchronized scrolling to compare side-by-side
   - **Jump to changes:**
     - Click "Previous" or "Next" buttons to navigate between changes
     - Use keyboard shortcuts: `P` for previous, `N` for next
     - Track your position with the change counter (e.g., "5 / 23")

### Understanding the Results

#### Sheet Status Indicators

- ✅ **Unchanged** - Sheet exists in both files with no changes
- ✏️ **Modified** - Sheet has cell-level changes
- ➕ **Added** - New sheet only in File B
- ❌ **Removed** - Sheet only exists in File A
- 🔄 **Renamed** - Sheet was renamed (detected by content similarity)

#### Cell Highlighting

- 🟢 **Green** - Cells added in File B
- 🔴 **Red** - Cells removed from File A
- 🟡 **Yellow** - Cells with modified values
- ⚪ **White** - Unchanged cells

## Project Structure 📁

```
excel-differ/
├── index.html              # Main HTML file
├── styles.css              # Styling and layout
├── js/
│   ├── main.js            # Application entry point
│   ├── fileHandler.js     # File upload and reading
│   ├── excelParser.js     # Excel parsing using SheetJS
│   ├── diffEngine.js      # Diff algorithm core
│   ├── summaryView.js     # Summary view renderer
│   ├── diffViewer.js      # Detailed diff viewer
│   ├── navBar.js          # Sheet navigation bar
│   └── copyright.js       # Keep copyright year up to date
└── README.md              # This file
```

## Technologies Used 🛠️

- **Pure JavaScript** - No frameworks, vanilla JS only
- **[SheetJS (xlsx)](https://sheetjs.com/)** - Excel file parsing
- **HTML5** - Modern web standards
- **CSS3** - Styling and animations
- **File API** - Browser-native file handling

## Browser Compatibility 🌐

| Browser | Version | Status             |
| ------- | ------- | ------------------ |
| Chrome  | 90+     | ✅ Fully Supported |
| Firefox | 88+     | ✅ Fully Supported |
| Safari  | 14+     | ✅ Fully Supported |
| Edge    | 90+     | ✅ Fully Supported |

## Features in Detail 🔬

### Smart Rename Detection

The tool uses content-based similarity comparison to detect renamed sheets:

- Compares sheet dimensions (rows/columns)
- Analyzes first 10 rows of data
- Calculates similarity score (85% threshold)
- Shows confidence percentage in results

### Change Navigation

Quickly navigate through all cell-level changes:

- **Previous/Next Buttons** - Navigate sequentially through changes
- **Keyboard Shortcuts** - Press `P` for previous, `N` for next
- **Change Counter** - Shows current position and total changes (e.g., "5 / 23")
- **Auto-Scroll** - Automatically scrolls changed cells into view
- **Visual Highlight** - Briefly highlights the target cell for easy identification
- **Auto-Enable Sync** - Automatically enables synchronized scrolling after navigation

### Synchronized Scrolling

- **Vertical Sync** - Scrolls to matching row numbers
- **Horizontal Sync** - Maintains column alignment
- **Toggle Control** - Enable/disable sync as needed
- **Safari Compatible** - Special handling for Safari rendering

### Performance Optimizations

- Efficient diff algorithm with O(n×m) complexity
- Lazy rendering for large spreadsheets
- Debounced scroll events
- Virtual scrolling for better performance

## Limitations ⚠️

- **File Size:** Maximum 50MB per file
- **Cell Count:** Very large sheets (>100,000 cells) may be slow
- **Formulas:** Only computed values are compared, not formula logic
- **Formatting:** Does not compare cell formatting/styles
- **Charts/Images:** Visual elements are not compared

## Contributing 🤝

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License 📄

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments 🙏

- [SheetJS](https://sheetjs.com/) for the excellent Excel parsing library
- Inspired by diff tools like Beyond Compare and WinMerge

---

⭐ If you find this tool useful, please consider giving it a star on GitHub!
