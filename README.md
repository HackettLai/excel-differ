# Excel Differ 📊

A web-based Excel file comparison tool that helps you identify differences between two Excel files quickly and easily. Runs completely in your browser with no server uploads required.

![Excel Differ](https://img.shields.io/badge/version-2.0.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Static Badge](https://img.shields.io/badge/AI%20Assist-Claude%20Sonnet%204.5-orange)

| Upload Interface | Sheet Selection | Unified Diff View |
| :--------------: | :-------------: | :---------------: |
| <img width="600" alt="Drag and drop Excel files to upload" src="https://upload.hackettlai.com/default/2026/s1-1768564290238.jpg" /> | <img width="600" alt="Select sheets from dropdown menus" src="https://upload.hackettlai.com/default/2026/s2-1768564290242.jpg" /> | <img width="600" alt="View differences in unified table with color-coded changes" src="https://upload.hackettlai.com/default/2026/s3-1768564290245.jpg" /> |

## Features ✨

- **📁 Drag & Drop Support** - Simply drag and drop Excel files to compare
- **🔍 Unified Table View** - View differences in a single unified table with old/new row indices
- **🎯 Smart Column Matching** - Intelligently matches columns by header content, not position
- **📊 Cell-Level Diff** - Highlights individual cell changes with old → new value display
- **🔄 Column Reordering Handling** - Correctly matches columns even when reordered (won't falsely report as added/deleted)
- **🎯 Change Navigation** - Jump between changes with Previous/Next buttons or keyboard shortcuts (P/N)
- **🖱️ Click-to-Navigate** - Click any changed cell to jump to that change
- **📍 Visual Change Counter** - Track your position through changes (e.g., "5 / 23")
- **💡 Visual Indicators** - Color-coded cells for easy identification of changes
- **🔒 100% Local** - All processing happens in your browser, no data leaves your device
- **🚀 No Server Required** - Runs entirely client-side

## Demo 🎬

[Live Demo](https://excel-differ.hackettlai.com)

## What's New in Version 2.0.0 🎉

### Major Changes

- **Unified Table View**: Replaced side-by-side comparison with a single unified table
- **Old/New Row Indices**: Each row shows both its old (File A) and new (File B) row numbers
- **Smart Column Matching**: Columns are matched by header content, not position
  - Handles column reordering correctly
  - Detects truly added/deleted columns
  - Preserves column relationships even when columns are moved
- **Enhanced Cell Display**: Modified cells show "old value → new value" inline
- **Click-to-Navigate**: Click any changed cell to jump to that change in the navigation sequence
- **Improved Change Detection**: More accurate detection of column additions/deletions

### Breaking Changes from v1.x

- Removed synchronized scrolling (replaced with unified table)
- Removed side-by-side sheet view (replaced with single table view)
- Removed sheet rename detection (simplified to focus on content comparison)
- Changed sheet selection workflow (now manual selection from dropdowns)

## Supported File Formats 📋

- `.xlsx` - Excel 2007+ files
- `.xls` - Excel 97-2003 files

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

   - Click "Select File" or drag & drop your Excel files into the upload areas
   - File A: Original/older version
   - File B: New/updated version

2. **Start Comparison**

   - Click the "Start Comparing" button
   - Wait for files to be parsed and compared

3. **Select Sheets to Compare**

   - If sheets with matching names are found, they'll be auto-selected
   - Otherwise, manually select sheets from the dropdowns
   - Click "Compare" to view differences

4. **Review Differences**

   - View all changes in a unified table
   - See old/new row indices for each row
   - Modified cells show "old value → new value"
   - Added columns marked with green header (+B)
   - Deleted columns marked with red header (−D)

5. **Navigate Changes**
   - Click "Previous" or "Next" buttons to jump between changes
   - Use keyboard shortcuts: `P` for previous, `N` for next
   - Click directly on any changed cell to navigate to it
   - Track your position with the change counter (e.g., "5 / 23")

### Understanding the Results

#### Column Headers (Two Rows)

- **Row 1**: Column letters with indicators
  - Normal columns: `A`, `B`, `C`
  - Added columns: `+B` (green background)
  - Deleted columns: `−D` (red background)
- **Row 2**: Header content
  - Shows actual header text from Excel
  - `(Blank Column)` for columns without headers

#### Row Indices

- **Old Column**: Row number in File A (or `-` if row was added)
- **New Column**: Row number in File B (or `-` if row was deleted)

#### Cell Highlighting

- 🟢 **Green Background** - Cell in added column or added row
- 🔴 **Red Background** - Cell in deleted column or deleted row
- 🟡 **Yellow Background** - Cell value was modified (shows old → new)
- ⚪ **White Background** - Unchanged cell

#### Cell Content Display

- **Modified cells**: `old value → new value`
- **Empty cells**: Shown as `Blank` in italic gray text
- **Normal cells**: Display current value

## Project Structure 📁

```
excel-differ/
├── index.html              # Main HTML file
├── styles.css              # Styling and layout
├── js/
│   ├── main.js            # Application entry point and controller
│   ├── fileHandler.js     # File upload, drag-and-drop, validation
│   ├── excelParser.js     # Excel parsing using SheetJS
│   ├── diffEngine.js      # Core comparison algorithm
│   ├── diffViewer.js      # Unified table renderer and navigation
│   └── copyright.js       # Copyright year management
└── README.md              # This file
```

## Technologies Used 🛠️

- **Pure JavaScript (ES6+)** - No frameworks, vanilla JS with modules
- **[SheetJS (xlsx)](https://sheetjs.com/)** - Excel file parsing
- **HTML5** - Modern web standards
- **CSS3** - Styling, animations, and responsive design
- **File API** - Browser-native file handling
- **Keyboard API** - Keyboard shortcut support

## Browser Compatibility 🌐

| Browser | Version | Status             |
| ------- | ------- | ------------------ |
| Chrome  | 90+     | ✅ Fully Supported |
| Firefox | 88+     | ✅ Fully Supported |
| Safari  | 14+     | ✅ Fully Supported |
| Edge    | 90+     | ✅ Fully Supported |

## Features in Detail 🔬

### Smart Column Matching

The tool matches columns by header content, not position:

- **Header-Based Matching**: Columns with identical header text are matched
  - Example: If "Email Address" moves from column G to H, it's still matched correctly
- **Reordering Tolerance**: Column position changes don't trigger false positives
  - Cells are compared based on what column they logically belong to
  - Reordered columns are NOT marked as added/deleted
- **True Add/Delete Detection**: Only reports genuine column additions/deletions
  - Added column: Header exists in File B but not in File A
  - Deleted column: Header exists in File A but not in File B

### Change Navigation

Quickly navigate through all cell-level changes:

- **Previous/Next Buttons** - Navigate sequentially through changes
- **Keyboard Shortcuts** - Press `P` for previous, `N` for next
  - Only active when not typing in input fields
  - Works globally across the page
- **Click Navigation** - Click any changed cell to jump to that change
  - Updates current position in navigation sequence
  - Automatically scrolls cell into view
- **Change Counter** - Shows current position and total changes (e.g., "5 / 23")
  - Updates dynamically as you navigate
  - Displays `0 / 0` when no changes exist
- **Visual Highlight** - Briefly highlights the target cell when navigating
  - 2-second highlight with CSS animation
  - Smooth scroll with centering

### Row Matching

Rows are matched using column A as the primary key:

- **Key-Based Matching**: Uses column A value to identify matching rows
  - Example: Row with "John Doe" in column A matches across files
- **Fallback Keys**: For rows without column A values, uses position-based keys
  - `old-{index}` for File A rows
  - `new-{index}` for File B rows
- **Add/Delete Detection**:
  - Added row: Key exists in File B but not in File A
  - Deleted row: Key exists in File A but not in File B

### Performance Optimizations

- Efficient diff algorithm with minimal memory footprint
- Lazy DOM manipulation for large spreadsheets
- Optimized scroll event handling
- Smart cell collection for navigation

## How It Works 🔍

Excel Differ uses a multi-step pipeline to detect changes:

1. **File Parsing** - Reads Excel files using SheetJS into JavaScript objects
2. **Column Matching** - Matches columns by header content (not position)
3. **Row Matching** - Matches rows using Column A as unique identifier
4. **Cell Comparison** - Compares matched cells and detects modifications

**Key Features:**
- ✅ Handles column reordering (matches by header name)
- ✅ Handles row reordering (matches by Column A value)
- ✅ Detects added/deleted columns and rows
- ✅ Highlights modified cells with old → new display

**Important Design Decision:**
- **Row matching relies on Column A values** - This enables accurate detection of row reordering, but requires Column A to contain stable unique identifiers
- If Column A values change, rows cannot be matched correctly (see [Limitations](#limitations-️) for details)

📖 **[Read detailed technical documentation →](TECHNICAL.md)**

## Keyboard Shortcuts ⌨️

| Key | Action          | Description                   |
| --- | --------------- | ----------------------------- |
| `P` | Previous Change | Jump to previous changed cell |
| `N` | Next Change     | Jump to next changed cell     |

_Note: Shortcuts are disabled when typing in input/textarea/select elements_

## Limitations ⚠️

### Row Matching Dependency
**Issue:** Rows are matched using Column A values only

**Impact:**
- If Column A values change between files, rows will be incorrectly matched
- All cells in that row may show as modified even if unchanged

**Example:**
```
File A: | 1 | peter | 34 |
File B: | 2 | peter | 34 |
```
**Result:** Treated as different rows  
**Shows:** "1 → 2" and all cells modified ❌


**Solutions:**
1. ✅ **Use stable IDs in Column A** (employee ID, product code, order number)
2. ✅ **Add an ID column** if your data doesn't have one
3. ✅ **Pre-sort both files** by the same key before comparison
4. ⚠️ **Avoid using auto-incremented numbers** that change between versions

---

### File Size Limit
**Issue:** Maximum 50MB per file

**Reason:** Browser memory limitations

**Solutions:**
1. Split large files into smaller chunks
2. Filter to relevant date ranges before comparison
3. Remove unnecessary columns before uploading

---

### What's NOT Compared
- ❌ Cell formatting (colors, fonts, borders)
- ❌ Formulas (only computed values)
- ❌ Charts and images
- ❌ Merged cell information
- ❌ Data validation rules
- ❌ Conditional formatting rules

**Alternative:** Use Excel's built-in "Compare and Merge Workbooks" for format comparison

## Use Cases 💼

### Perfect For

- **Version Control**: Compare different versions of Excel reports
- **Data Auditing**: Verify data changes in financial spreadsheets
- **Quality Assurance**: Validate data migrations or transformations
- **Collaboration**: Review changes made by team members
- **Configuration Management**: Track changes in Excel-based config files

### Not Suitable For

- **Format Comparison**: Use Excel's built-in compare for formatting changes
- **Large Dataset Analysis**: Consider database tools for millions of rows
- **Real-Time Collaboration**: Use Google Sheets or Excel Online instead
- **Formula Debugging**: Use Excel's formula auditing tools

## Troubleshooting 🔧

### Common Issues

**Files won't upload**

- Check file size (must be under 50MB)
- Ensure file format is .xlsx or .xls
- Try re-saving the file in Excel

**Comparison is slow**

- Large files may take 10-30 seconds to process
- Close other browser tabs to free up memory
- Try comparing smaller sections of data

**Changes not detected**

- Ensure both files have matching sheet names
- Check that column headers match exactly
- Verify data types are consistent

**Navigation buttons disabled**

- No changes detected in selected sheets
- Try selecting different sheets to compare

## Contributing 🤝

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Guidelines

- Use ES6+ JavaScript with modules
- Follow existing code structure and naming conventions
- Add comprehensive comments for complex logic
- Test with various Excel file formats and sizes
- Ensure browser compatibility

## License 📄

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments 🙏

- [SheetJS](https://sheetjs.com/) for the excellent Excel parsing library
- Inspired by diff tools like Beyond Compare, WinMerge, and Git diff
- Built with assistance from Claude AI (Anthropic)

## Support 💬

If you encounter issues or have questions:

- 🐛 [Report a bug](https://github.com/HackettLai/excel-differ/issues)
- 💡 [Request a feature](https://github.com/HackettLai/excel-differ/issues)

---

⭐ If you find this tool useful, please consider giving it a star on GitHub!

**Made with ❤️ by Hackett.Lai**
