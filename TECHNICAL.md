# Technical Documentation

## Table of Contents

1. [Processing Pipeline Overview](#processing-pipeline-overview)
2. [File Reading & Parsing](#1-file-reading--parsing)
3. [Header Row Configuration](#2-header-row-configuration-introduced-in-v210)
4. [Position-Based Comparison](#3-position-based-comparison-new-in-v220)
5. [Column Matching (Header-Based)](#4-column-matching-header-based)
6. [Row Matching (Key Column-Based)](#5-row-matching-key-column-based)
7. [Cell Comparison](#6-cell-comparison)
8. [Complete Example](#complete-example)
9. [Implementation Details](#implementation-details)
10. [Edge Cases & Limitations](#edge-cases--limitations)
11. [Performance Considerations](#performance-considerations)

---

## Processing Pipeline Overview

```
┌─────────────┐ ┌─────────────┐
│    File A   │ │    File B   │
│   (Excel)   │ │   (Excel)   │
└──────┬──────┘ └──────┬──────┘
       │               │
       └──────┬────────┘
              ▼
       ┌───────────────┐
       │    SheetJS    │
       │     Parser    │
       └───────┬───────┘
               ▼
       ┌───────────────┐
       │   JavaScript  │
       │     Object    │
       └───────┬───────┘
               │
         ┌─────┴─────┐
         ▼           ▼
    ┌─────────┐ ┌─────────┐
    │ Header  │ │   Key   │
    │   Row   │ │ Column  │
    │ Select  │ │ Select  │
    └────┬────┘ └────┬────┘
         │           │
         └─────┬─────┘
               ▼
      ┌────────┴────────┐
      ▼                 ▼
┌─────────────┐ ┌─────────────┐
│   Column    │ │     Row     │
│   Matching  │ │   Matching  │
└──────┬──────┘ └──────┬──────┘
       │               │
       └───────┬───────┘
               ▼
       ┌───────────────┐
       │      Cell     │
       │   Comparison  │
       └───────┬───────┘
               ▼
       ┌───────────────┐
       │  Diff Result  │
       │    (Viewer)   │
       └───────────────┘
```

**Data Flow:**
1. Excel/CSV files → SheetJS → JS Objects
2. User selects Header Rows & Key Column
3. Extract headers & build column maps
4. Extract Key Column values & build row maps
5. Compare cells at intersections
6. Render unified diff table

---

## 1. File Reading & Parsing

```
Excel/CSV File → SheetJS (xlsx) → JavaScript Object
```

**Input Files:**

- File A (old version) and File B (new version)
- Supported formats: `.xlsx`, `.xls`, `.csv` ⭐ NEW

**CSV File Handling (NEW in v2.2.0):**
- ✅ Automatic UTF-8 encoding detection
- ✅ Fallback to GBK/Big5 for Chinese characters
- ⚠️ One-time session warning on first CSV upload
- 💡 Warning stored in `sessionStorage` (cleared on page refresh)

**Parsing Result:**

```javascript
{
  fileName: "data.xlsx",
  sheets: {
    "Sheet1": {
      name: "Sheet1",
      data: [
        { A: "Name", B: "Email", C: "Phone" },           // Row 1 (headers)
        { A: "John Doe", B: "john@email.com", C: "123" }, // Row 2
        { A: "Jane Doe", B: "jane@email.com", C: "456" }  // Row 3
      ],
      rowCount: 3,
      colCount: 3
    }
  },
  sheetNames: ["Sheet1"]
}
```

**Key Points:**
- Each row is an object with column letters as keys (A, B, C, ...)
- First row (index 0) is typically headers, but user can configure this
- All rows are preserved (no assumptions about which row is headers)

---

## 2. Header Row Configuration (introduced in v2.1.0)

**Before comparison, users configure:**

### Step 2.1: Select Header Rows

Users specify which row contains column headers for each file:

```javascript
{
  headerRowA: 3,  // File A headers are in Row 3
  headerRowB: 1   // File B headers are in Row 1 (default)
}
```

**Supported Range:** Rows 1-50

### Step 2.2: Adjust Data Based on Header Row

Tool extracts data starting from the selected header row:

```javascript
// Original data (10 rows total)
data = [
  { A: "Company Name", B: "ABC Corp" },      // Row 1
  { A: "Report Date", B: "2024-01-15" },     // Row 2
  { A: "Name", B: "Email", C: "Phone" },     // Row 3 ← Headers!
  { A: "John", B: "john@email.com", C: "123" }, // Row 4 (data)
  { A: "Jane", B: "jane@email.com", C: "456" }  // Row 5 (data)
]

// After selecting Header Row = 3
adjustedData = [
  { A: "Name", B: "Email", C: "Phone" },        // Row 0 (headers)
  { A: "John", B: "john@email.com", C: "123" }, // Row 1 (data)
  { A: "Jane", B: "jane@email.com", C: "456" }  // Row 2 (data)
]
```

### Step 2.3: Calculate Excel Row Numbers

Excel row numbers are calculated as:

```
Excel Row Number = Header Row + 1 + Data Row Index
```

**Example:**
- Header Row = 3
- First data row (index 0) → Excel Row 4
- Second data row (index 1) → Excel Row 5

**Why This Matters:**
- Different files may have different header row positions
- Tool needs to correctly identify where data starts
- Row numbers in diff view must match actual Excel row numbers

---

### Step 2.4: Auto-Detect Common Columns (NEW)

After headers are extracted, tool automatically detects columns that exist in both files:

```javascript
File A Headers (Row 3): { A: "Name", B: "Employee ID", C: "Email" }
File B Headers (Row 1): { A: "Employee ID", B: "Name", C: "Phone", D: "Email" }

// Auto-detected common columns:
commonColumns = [
  { name: "Name", colIndexA: "A", colIndexB: "B" },
  { name: "Employee ID", colIndexA: "B", colIndexB: "A" },
  { name: "Email", colIndexA: "C", colIndexB: "D" }
]
```

**Matching Rules:**
- ✅ Case-insensitive ("Email" matches "email")
- ✅ Trim whitespace ("Name " matches "Name")
- ✅ Exact match required after normalization

### Step 2.5: User Selects Key Column

User chooses which common column to use for row matching:

```javascript
selectedKeyColumn = "B"  // User selects "Employee ID"
```

**Dropdown Shows:**
```
[ ] (Use Row Position)     ← NEW in v2.2.0
[ ] Employee ID  ← Default (first common column)
[ ] Name
[ ] Email
```

**Error Handling:**
- ⚠️ If no common columns found → Show warning
- ⚠️ User must select a key column before comparison (or use Position mode)

---

## 3. Position-Based Comparison (NEW in v2.2.0)

Position-based mode allows rows to be matched purely by their order when a reliable key column is not available. Unlike key-column matching, this mode does not require the user to select a column; all rows are compared sequentially. This is useful for simple spreadsheets where row identity is implied by position or where all rows share identical structure.

### Behavior
- Rows are paired one-to-one by index after header adjustment.
- Added or deleted rows are detected when one file has more rows than the other.
- Cell comparisons proceed across matched indexes using the same column mapping logic.
- In the diff viewer, changes are rendered similarly but without row-key labels.

### Limitations

#### Issue: Rows are matched only by their order

- Works without choosing a key column, but any insertion or deletion in the middle of the dataset will shift all subsequent row pairings and result in a cascade of false positives.
- Added/deleted rows are detected, but the remaining rows remain aligned by index rather than by a stable identifier.
- Users should only enable position-based mode for simple lists where row order is guaranteed to be consistent across versions.

**Example:**

```
File A rows: [A, B, C]
File B rows: [A, X, B, C]

// Position-based result: Row 2 (B) compared with X → flagged as change
// all subsequent rows misaligned
```

**Best Practice:**
- Only use position mode when no reliable key column exists.
- If data may have insertions/deletions, prefer key‑column matching.

*The remaining sections describe key-column matching which remains the default mode.*

---

## 4. Column Matching (Header-Based)

### Step 4.1: Extract Headers from Selected Header Row

```javascript
File A (Header Row = 3):
  Headers: { A: "Name", B: "Employee ID", C: "Email", D: "Phone" }

File B (Header Row = 1):
  Headers: { A: "Employee ID", B: "Name", C: "Phone", D: "Email", E: "Department" }
```

### Step 4.2: Build Header-to-Column Maps

```javascript
// Map: header content → column letter
File A: {
  "Name" → "A",
  "Employee ID" → "B",
  "Email" → "C",
  "Phone" → "D"
}

File B: {
  "Employee ID" → "A",
  "Name" → "B",
  "Phone" → "C",
  "Email" → "D",
  "Department" → "E"
}
```

### Step 4.3: Detect Column Changes

```
For each header in File A:
  ├─ If header exists in File B → MATCHED ✓
  └─ If header NOT in File B → DELETED ❌

For each header in File B:
  └─ If header NOT in File A → ADDED ✅
```

**Example Result:**

```javascript
{
  matched: ["Name", "Employee ID", "Email", "Phone"],
  added: ["Department"],
  deleted: []
}
```

**Key Point:** Columns are matched by **header content**, not position.

- "Email" in column C (File A) matches "Email" in column D (File B) ✓
- Column reordering does NOT trigger false positives

---

## 5. Row Matching (Key Column-Based)

**NEW in v2.1.0:** Users select which column to use as the Key Column for row matching.

### Step 5.1: User Selects Key Column

```javascript
// User selects Key Column from dropdown (auto-detected common columns)
selectedKeyColumn = "B"  // "Employee ID" column
```

### Step 5.2: Extract Row Keys from Selected Key Column

```javascript
File A (Header Row = 1, Key Column = B):
  Row 2: key = "E12345"  (from B2)
  Row 3: key = "E67890"  (from B3)
  Row 4: key = "E11111"  (from B4)

File B (Header Row = 1, Key Column = A):  ← Note: Same header, different column!
  Row 2: key = "E67890"  (from A2) ← Moved up
  Row 3: key = "E12345"  (from A3) ← Moved down
  Row 4: key = "E99999"  (from A4) ← New employee
```

**Important:** Key Column is matched by **header content**, not column letter!
- File A: "Employee ID" is in column B
- File B: "Employee ID" is in column A (reordered)
- Tool correctly extracts keys from both locations

### Step 5.3: Build Row Maps

```javascript
File A: {
  "E12345" → { row: {...}, index: 2 },
  "E67890" → { row: {...}, index: 3 },
  "E11111" → { row: {...}, index: 4 }
}

File B: {
  "E67890" → { row: {...}, index: 2 },
  "E12345" → { row: {...}, index: 3 },
  "E99999" → { row: {...}, index: 4 }
}
```

### Step 5.4: Detect Row Changes

```
For each row key in File A:
  ├─ If key exists in File B → MATCHED ✓
  └─ If key NOT in File B → DELETED ❌

For each row key in File B:
  └─ If key NOT in File A → ADDED ✅
```

**Example Result:**

```javascript
{
  matched: [
    { key: "E12345", oldRow: 2, newRow: 3 },
    { key: "E67890", oldRow: 3, newRow: 2 }
  ],
  added: [
    { key: "E99999", newRow: 4 }
  ],
  deleted: [
    { key: "E11111", oldRow: 4 }
  ]
}
```

**Key Points:**
- ✅ Rows are matched by **selected Key Column value**, not position
- ✅ Key Column can be ANY column (A, B, C, etc.), not limited to Column A
- ✅ Tool auto-detects common columns between files
- ✅ Users choose which column best identifies their rows
- ⚠️ If Key Column values change, rows cannot be matched

**Example:**
- Row with "E12345" in B2 (File A) matches "E12345" in A3 (File B) ✓
- Even though:
  - Row positions differ (2 vs 3)
  - Column positions differ (B vs A)
- They're matched by Key Column **value** ("E12345")

### Step 5.5: Handle Empty Key Values (Fallback)

For rows with empty Key Column values:

```javascript
File A Row 5: { A: "", B: "Peter", C: "34" }  ← Empty key
File B Row 6: { A: "", B: "Peter", C: "34" }  ← Empty key
```

**Fallback Matching:**
1. Use position-based unique key: `__blank_old_5`, `__blank_new_6`
2. Compare **all columns** to check if rows are identical
3. If all columns match → Merge rows
4. If any column differs → Treat as separate rows

**Example:**
```javascript
// Scenario 1: All columns match
File A: { A: "", B: "Peter", C: "34" }
File B: { A: "", B: "Peter", C: "34" }
Result: MATCHED ✓

// Scenario 2: One column differs
File A: { A: "", B: "Peter", C: "34" }
File B: { A: "", B: "Peter", C: "35" }  ← Age changed
Result: Treated as DELETED (old) + ADDED (new) ❌
```

---

## 6. Cell Comparison

### Step 6.1: For Each MATCHED Row

```
For each MATCHED column header:
  1. Find the column letter in File A (via header map)
  2. Find the column letter in File B (via header map)
  3. Get cell value from File A at (matched row, File A column letter)
  4. Get cell value from File B at (matched row, File B column letter)
  5. Compare the two values
     ├─ If same → No change
     └─ If different → MODIFIED 🔄
```

### Step 6.2: Example Cell Comparison

**Scenario:**

```
Row Key: "E12345" (from Key Column "Employee ID")
Matched Row: File A Row 2 ↔ File B Row 3
Column: "Email"
```

**Step-by-Step:**

```javascript
Step 1: Find column for "Email" in File A
  → Header Map A: "Email" → Column C

Step 2: Find column for "Email" in File B
  → Header Map B: "Email" → Column D (moved!)

Step 3: Get File A value at (Row 2, Column C)
  → File A[2]["C"] = "john@old.com"

Step 4: Get File B value at (Row 3, Column D)
  → File B[3]["D"] = "john@new.com"

Step 5: Compare values
  → "john@old.com" ≠ "john@new.com"
  → Mark as MODIFIED 🔄
```

### Step 6.3: Cell Change Result

```javascript
{
  rowKey: "E12345",
  header: "Email",
  oldCol: "C",
  newCol: "D",
  oldValue: "john@old.com",
  newValue: "john@new.com",
  changeType: "modified",
  row: 2  // Excel row number in File A
}
```

**Key Points:**
- ✅ Rows are matched by Key Column (user-selected)
- ✅ Columns are matched by header content
- ✅ Cell comparison happens at the intersection of matched row + matched column
- ✅ Handles both row and column reordering correctly
- ✅ Shows exact Excel row numbers in diff view

---

## Complete Example

### User Configuration

```javascript
{
  headerRowA: 1,           // File A headers in Row 1
  headerRowB: 1,           // File B headers in Row 1
  keyColumn: "A"           // Use Column A ("Employee ID") for matching
}
```

### File A

```
| A: Employee ID | B: Name     | C: Email          | D: Phone    |
|----------------|-------------|-------------------|-------------|
| E001           | John Doe    | john@old.com      | 111-1111    |
| E002           | Jane Doe    | jane@email.com    | 222-2222    |
| E003           | Bob Smith   | bob@email.com     | 333-3333    |
```

**Internal Representation:**

```javascript
[
  { A: "Employee ID", B: "Name", C: "Email", D: "Phone" },        // Row 1 (headers)
  { A: "E001", B: "John Doe", C: "john@old.com", D: "111-1111" }, // Row 2
  { A: "E002", B: "Jane Doe", C: "jane@email.com", D: "222-2222" }, // Row 3
  { A: "E003", B: "Bob Smith", C: "bob@email.com", D: "333-3333" }  // Row 4
]
```

### File B

```
| A: Employee ID | B: Name     | C: Phone    | D: Email          | E: Department |
|----------------|-------------|-------------|-------------------|---------------|
| E002           | Jane Doe    | 222-2222    | jane@email.com    | Sales         |
| E001           | John Doe    | 111-1111    | john@new.com      | IT            |
| E004           | Alice Wong  | 444-4444    | alice@email.com   | HR            |
```

**Internal Representation:**

```javascript
[
  { A: "Employee ID", B: "Name", C: "Phone", D: "Email", E: "Department" }, // Row 1 (headers)
  { A: "E002", B: "Jane Doe", C: "222-2222", D: "jane@email.com", E: "Sales" }, // Row 2
  { A: "E001", B: "John Doe", C: "111-1111", D: "john@new.com", E: "IT" },      // Row 3
  { A: "E004", B: "Alice Wong", C: "444-4444", D: "alice@email.com", E: "HR" }  // Row 4
]
```

---

### Processing Steps

#### 1. Column Matching

**Extract Headers:**
```javascript
File A: { A: "Employee ID", B: "Name", C: "Email", D: "Phone" }
File B: { A: "Employee ID", B: "Name", C: "Phone", D: "Email", E: "Department" }
```

**Build Header Maps:**
```javascript
File A: { "Employee ID" → "A", "Name" → "B", "Email" → "C", "Phone" → "D" }
File B: { "Employee ID" → "A", "Name" → "B", "Phone" → "C", "Email" → "D", "Department" → "E" }
```

**Detect Changes:**
- ✅ Matched: "Employee ID", "Name", "Email", "Phone"
- ✅ Added: "Department"
- ❌ Deleted: (none)

**Note:** "Email" moved from column C → D, but NOT marked as added/deleted ✓

---

#### 2. Row Matching (using Key Column "Employee ID")

**Extract Keys:**
```javascript
File A Keys: { "E001" → Row 2, "E002" → Row 3, "E003" → Row 4 }
File B Keys: { "E002" → Row 2, "E001" → Row 3, "E004" → Row 4 }
```

**Detect Changes:**
- ✅ Matched:
  - "E001" (File A Row 2 ↔ File B Row 3)
  - "E002" (File A Row 3 ↔ File B Row 2)
- ✅ Added: "E004" (Alice Wong, Row 4)
- ❌ Deleted: "E003" (Bob Smith, Row 4)

**Note:** "E001" moved from Row 2 → 3, but still matched correctly ✓

---

#### 3. Cell Comparison for "E001" (John Doe)

**Matched Row:** File A Row 2 ↔ File B Row 3

**For each matched column:**

| Header      | File A Col | File B Col | File A Value   | File B Value   | Result       |
|-------------|------------|------------|----------------|----------------|--------------|
| Employee ID | A          | A          | E001           | E001           | ✓ No change  |
| Name        | B          | B          | John Doe       | John Doe       | ✓ No change  |
| Email       | C          | D          | john@old.com   | john@new.com   | 🔄 MODIFIED  |
| Phone       | D          | C          | 111-1111       | 111-1111       | ✓ No change  |

**Cell Differences Found:**
```javascript
[
  {
    rowKey: "E001",
    header: "Email",
    oldCol: "C",
    newCol: "D",
    oldValue: "john@old.com",
    newValue: "john@new.com",
    row: 2  // Excel row number
  }
]
```

---

#### 4. Cell Comparison for "E002" (Jane Doe)

**Matched Row:** File A Row 3 ↔ File B Row 2

**For each matched column:**

| Header      | File A Col | File B Col | File A Value     | File B Value     | Result      |
|-------------|------------|------------|------------------|------------------|-------------|
| Employee ID | A          | A          | E002             | E002             | ✓ No change |
| Name        | B          | B          | Jane Doe         | Jane Doe         | ✓ No change |
| Email       | C          | D          | jane@email.com   | jane@email.com   | ✓ No change |
| Phone       | D          | C          | 222-2222         | 222-2222         | ✓ No change |

**Cell Differences Found:** None

**Note:** Despite row moving from 3 → 2, all cells are unchanged ✓

---

### Final Diff Summary

```javascript
{
  columnChanges: {
    matched: ["Employee ID", "Name", "Email", "Phone"],
    added: ["Department"],
    deleted: []
  },
  
  rowChanges: {
    matched: [
      { key: "E001", oldRow: 2, newRow: 3 },
      { key: "E002", oldRow: 3, newRow: 2 }
    ],
    added: [
      { key: "E004", newRow: 4, row: { A: "E004", B: "Alice Wong", ... } }
    ],
    deleted: [
      { key: "E003", oldRow: 4, row: { A: "E003", B: "Bob Smith", ... } }
    ]
  },
  
  cellDifferences: [
    {
      rowKey: "E001",
      header: "Email",
      oldValue: "john@old.com",
      newValue: "john@new.com",
      row: 2
    }
  ],
  
  summary: {
    totalColumns: 5,
    columnsAdded: 1,
    columnsDeleted: 0,
    rowsMatched: 2,
    rowsAdded: 1,
    rowsDeleted: 1,
    cellsModified: 1
  }
}
```

**Human-Readable Summary:**
- ✅ 1 column added ("Department")
- ✅ 1 row added ("E004" - Alice Wong)
- ❌ 1 row deleted ("E003" - Bob Smith)
- 🔄 1 cell modified (E001's Email: john@old.com → john@new.com)

---

## Implementation Details

### Data Structures

#### 1. Parsed Sheet Data

```javascript
{
  name: "Sheet1",
  data: [
    { A: "Header1", B: "Header2", ... },  // Row 0 (after adjustment)
    { A: "Value1", B: "Value2", ... },    // Row 1
    ...
  ],
  rowCount: 10,
  colCount: 5
}
```

#### 2. Header Maps

```javascript
{
  headerToCol: Map {
    "Employee ID" → "A",
    "Name" → "B",
    "Email" → "C"
  },
  colToHeader: Map {
    "A" → "Employee ID",
    "B" → "Name",
    "C" → "Email"
  }
}
```

#### 3. Row Maps

```javascript
Map {
  "E001" → {
    row: { A: "E001", B: "John", C: "john@email.com" },
    index: 2  // Excel row number
  },
  "E002" → {
    row: { A: "E002", B: "Jane", C: "jane@email.com" },
    index: 3
  }
}
```

---

### Key Algorithms

#### 1. Find Common Columns

```javascript
function findCommonColumns(headersA, headersB) {
  const commonColumns = [];
  
  headersA.forEach((headerA, indexA) => {
    const normalizedA = String(headerA).trim().toLowerCase();
    
    headersB.forEach((headerB, indexB) => {
      const normalizedB = String(headerB).trim().toLowerCase();
      
      if (normalizedA === normalizedB && normalizedA !== '') {
        const alreadyExists = commonColumns.some(col => col.name === headerA);
        
        if (!alreadyExists) {
          commonColumns.push({
            name: headerA,
            colIndex: getColumnLetter(indexA),
            indexA: indexA,
            indexB: indexB
          });
        }
      }
    });
  });
  
  return commonColumns;
}
```

#### 2. Match Rows by Key Column

```javascript
function matchRowsByKey(rowsA, rowsB, keyColumn, headerRowA, headerRowB) {
  const mapA = new Map();
  const mapB = new Map();
  
  // Build map for File A
  rowsA.forEach((row, index) => {
    const keyValue = String(row[keyColumn] || '').trim();
    if (keyValue) {
      mapA.set(keyValue, {
        row: row,
        index: headerRowA + 1 + index  // Excel row number
      });
    }
  });
  
  // Build map for File B
  rowsB.forEach((row, index) => {
    const keyValue = String(row[keyColumn] || '').trim();
    if (keyValue) {
      mapB.set(keyValue, {
        row: row,
        index: headerRowB + 1 + index  // Excel row number
      });
    }
  });
  
  // Find matched, added, deleted rows
  const matched = [];
  const added = [];
  const deleted = [];
  
  mapA.forEach((dataA, key) => {
    if (mapB.has(key)) {
      matched.push({
        key: key,
        oldRow: dataA.index,
        newRow: mapB.get(key).index
      });
    } else {
      deleted.push({
        key: key,
        oldRow: dataA.index,
        row: dataA.row
      });
    }
  });
  
  mapB.forEach((dataB, key) => {
    if (!mapA.has(key)) {
      added.push({
        key: key,
        newRow: dataB.index,
        row: dataB.row
      });
    }
  });
  
  return { matched, added, deleted };
}
```

#### 3. Compare Cells

```javascript
function compareCells(oldRow, newRow, headerToOldCol, headerToNewCol) {
  const differences = [];
  
  headerToOldCol.forEach((oldCol, headerContent) => {
    const newCol = headerToNewCol.get(headerContent);
    
    if (newCol) {
      const oldVal = oldRow[oldCol];
      const newVal = newRow[newCol];
      
      if (oldVal !== newVal) {
        differences.push({
          header: headerContent,
          oldCol: oldCol,
          newCol: newCol,
          oldValue: oldVal,
          newValue: newVal
        });
      }
    }
  });
  
  return differences;
}
```

---

## Edge Cases & Limitations

### 1. Key Column Requirements (NEW in v2.1.0)

#### Issue: Row matching requires a stable Key Column

**Requirements:**
- Key Column must exist in both files
- Key Column should contain unique identifiers
- Key values should not change between versions

**What Happens If:**

| Scenario | Result | Example |
|----------|--------|---------|
| ❌ No common columns | Cannot compare rows (error shown) | File A has "ID", File B has "Code" |
| ❌ Key Column values change | Rows treated as deleted + added | ID changes from "101" → "102" |
| ❌ Key Column has duplicates | Only first occurrence is matched | Two rows with ID "101" |
| ⚠️ Key Column is empty | Row matched by position (fallback) | ID column has blank cells |

**Example:**

```javascript
// File A
{ A: "101", B: "Peter", C: "34" }

// File B (ID changed)
{ A: "102", B: "Peter", C: "34" }

// Result: Treated as TWO different rows
// ❌ Deleted: "101" (Peter, 34)
// ✅ Added: "102" (Peter, 34)
```

**Solutions:**
1. ✅ Use stable IDs (employee ID, product code, order number)
2. ✅ Add an ID column if your data doesn't have unique keys
3. ✅ Pre-process data to ensure consistent IDs before comparison
4. ✅ Select the correct Key Column that best identifies your rows
5. ⚠️ Avoid using auto-incremented numbers that change between versions

---

### 1.1 Position-Based Mode Limitations (NEW in v2.2.0)

#### Issue: Rows are matched only by their order

- Works without choosing a key column, but any insertion or deletion in the middle of the dataset will shift all subsequent row pairings and result in a cascade of false positives.
- Added/deleted rows are detected, but the remaining rows remain aligned by index rather than by a stable identifier.
- Users should only enable position-based mode for simple lists where row order is guaranteed to be consistent across versions.

**Example:**

```
File A rows: [A, B, C]
File B rows: [A, X, B, C]

// Position-based result: Row 2 (B) compared with X → flagged as change
// all subsequent rows misaligned
```

**Best Practice:**
- Only use position mode when no reliable key column exists.
- If data may have insertions/deletions, prefer key‑column matching.

---

### 2. Header Row Position

#### Issue: Different files may have headers in different rows

**Handling:**
- ✅ Users select header row for each file independently
- ✅ Tool adjusts data parsing based on selected header row
- ✅ Row numbers calculated relative to header row
- ⚠️ If wrong header row selected, comparison will be incorrect

**Example:**

```javascript
// File A structure
Row 1: Company Name: ABC Corp
Row 2: Report Date: 2024-01-15
Row 3: Name | Email | Phone  ← Headers!
Row 4: John | john@email.com | 123
Row 5: Jane | jane@email.com | 456

// Correct configuration:
headerRowA = 3

// Wrong configuration:
headerRowA = 1  ❌
// Result: Treats "Company Name" as header, comparison fails
```

**Best Practice:**
- Always verify header row selection in preview
- Look for row with column names (not data values)
- If file has multiple header rows, select the last one

---

### 3. Empty Key Values

#### Issue: Some rows may have empty Key Column values

**Handling:**
- Uses position-based fallback matching
- Compares all columns to ensure exact match
- Empty-key rows are matched only if all columns are identical

**Example 1: Exact Match**

```javascript
File A: | (empty) | Peter | 34 |
File B: | (empty) | Peter | 34 |

Result: MATCHED ✓
// All columns identical, treated as same row
```

**Example 2: Partial Match**

```javascript
File A: | (empty) | Peter | 34 |
File B: | (empty) | Peter | 35 |  ← Age different

Result: Treated as different rows ❌
// Deleted: Row with (empty, Peter, 34)
// Added: Row with (empty, Peter, 35)
```

**Example 3: Multiple Empty Keys**

```javascript
File A:
| (empty) | Peter | 34 |
| (empty) | Alice | 28 |

File B:
| (empty) | Alice | 28 |
| (empty) | Peter | 34 |

Result: Position-based matching
// Row 1 (Peter, 34) vs Row 1 (Alice, 28) → Different → Not matched
// Row 2 (Alice, 28) vs Row 2 (Peter, 34) → Different → Not matched
// Final: 2 deleted + 2 added (because position-based fallback failed)
```

**Best Practice:**
- Avoid empty Key Column values when possible
- If unavoidable, ensure rows are in same order in both files
- Consider adding unique IDs to your data

---

### 4. Duplicate Key Values

#### Issue: Key Column may contain duplicate values

**Current Behavior:**
- Only first occurrence of each key is matched
- Subsequent duplicates are treated as separate rows

**Example:**

```javascript
File A:
Row 2: | E001 | John | john@email.com |
Row 3: | E001 | Jane | jane@email.com |  ← Duplicate ID!

File B:
Row 2: | E001 | John | john@new.com |
Row 3: | E001 | Jane | jane@email.com |

Result:
// Only first "E001" (John) is matched
// Second "E001" (Jane) treated as deleted + added
```

**Solutions:**
1. ✅ Ensure Key Column has unique values
2. ✅ Use composite keys if needed (manually concatenate columns)
3. ⚠️ If duplicates are expected, select a more unique column

---

### 5. Column Reordering

#### Issue: Columns may be reordered between files

**Handling:**
- ✅ Columns matched by header content, not position
- ✅ Reordered columns are NOT marked as added/deleted
- ✅ Cell values correctly compared even after reordering

**Example:**

```javascript
File A: | A: Name | B: Email | C: Phone |
File B: | A: Name | B: Phone | C: Email |  ← Email and Phone swapped

Result:
// ✅ All columns matched (Name, Email, Phone)
// ✅ No columns marked as added/deleted
// ✅ Cells compared correctly using header mapping
```

**Key Point:** This is intentional behavior to handle column reordering gracefully!

---

### 6. CSV Encoding and Character Sets (NEW in v2.2.0)

#### Issue: CSV files may use non‑UTF‑8 encodings

- Many Chinese users save CSV in GBK/GB2312, which appears garbled when read as UTF‑8.
- The parser automatically detects UTF‑8; if decoding fails it attempts GBK/Big5 and falls back to a non‑fatal UTF‑8 decode.
- A one‑time session warning alerts users to save their CSV as UTF‑8 if garbling occurs.

**Behavior:**
- If UTF‑8 decode succeeds without replacement characters, data is used directly.
- If UTF‑8 fails, the parser logs a warning and retries with common alternatives.
- If all attempts fail, an error is thrown.

**Warning Display Logic (NEW in v2.2.0):**
- ✅ Warning shows **only once per browser session** using `sessionStorage`
- ✅ After first display, subsequent CSV uploads skip the warning
- ✅ Refreshing the page clears `sessionStorage`, warning will appear again
- ✅ Opening a new tab/window will show the warning again (independent sessions)

**Implementation:**
```javascript
// Check if warning already shown in this session
const hasSeenCSVWarning = sessionStorage.getItem('csvEncodingWarningShown');

if (!hasSeenCSVWarning) {
  // Show warning popup
  const userConfirmed = confirm('⚠️ CSV File Encoding Notice...');
  
  if (userConfirmed) {
    // Mark warning as shown
    sessionStorage.setItem('csvEncodingWarningShown', 'true');
  }
}
```

**Recommendation:**
- Always save CSV files in UTF-8 when possible.
- If garbled text appears, re-export the file with the correct encoding before reuploading.

**User Impact:**
- ✅ Reduces annoyance from repeated warnings
- ✅ Warning persists throughout comparison workflow
- ✅ Clears naturally on page refresh (appropriate reset point)

---

### 7. Data Type Inconsistencies

#### Issue: Same column may have different data types

**Current Behavior:**
- All values converted to strings for comparison
- Numeric 123 and string "123" are considered equal
- Dates compared as formatted strings

**Example:**

```javascript
File A: | Age | 25    |  ← Number
File B: | Age | "25"  |  ← String

Result: No change detected (25 == "25") ✓
```

```javascript
File A: | Date | 2024-01-15          |  ← Date object
File B: | Date | "2024-01-15"        |  ← String

Result: Depends on date formatting
// If formatted as same string → No change
// If formatted differently → Change detected
```

**Limitations:**
- Cannot detect type changes (number → string)
- Date comparison depends on formatting
- Boolean true vs string "true" may not match

---

### 8. Large Files

#### Issue: Browser memory limitations

**Current Limits:**
- Maximum file size: 50MB
- Recommended: < 10MB for best performance
- Large files (> 10MB) may cause slowness

**Performance Tips:**
1. ✅ Filter data before comparison (date ranges, specific rows)
2. ✅ Remove unnecessary columns
3. ✅ Split large files into smaller chunks
4. ✅ Use desktop Excel for very large datasets

**Warning Signs:**
- ⚠️ Comparison takes > 30 seconds
- ⚠️ Browser tab becomes unresponsive
- ⚠️ "Out of memory" errors

---

### 9. Special Characters in Headers

#### Issue: Headers may contain special characters

**Handling:**
- ✅ Whitespace is trimmed
- ✅ Case-insensitive matching
- ⚠️ Special characters must match exactly

**Example:**

```javascript
File A: "Email Address"
File B: "Email  Address"  ← Extra space

Result: MATCHED ✓ (whitespace trimmed)
```

```javascript
File A: "Email Address"
File B: "email address"  ← Lowercase

Result: MATCHED ✓ (case-insensitive)
```

```javascript
File A: "Email_Address"
File B: "Email-Address"  ← Different delimiter

Result: NOT matched ❌
// Treated as two different columns
```

---

## Performance Considerations

### Time Complexity

| Operation | Complexity | Notes |
|-----------|------------|-------|
| File Parsing | O(n × m) | n = rows, m = columns |
| Column Matching | O(c²) | c = column count (typically small) |
| Row Matching | O(r) | r = row count (using Map) |
| Cell Comparison | O(r × c) | For all matched rows and columns |
| Total | O(n × m) | Linear with data size |

### Space Complexity

| Structure | Space | Notes |
|-----------|-------|-------|
| Parsed Data | O(n × m) | Stores entire sheet in memory |
| Header Maps | O(c) | One entry per column |
| Row Maps | O(r) | One entry per row |
| Diff Results | O(k) | k = number of changes (typically < n × m) |
| Total | O(n × m) | Dominated by parsed data |

### Optimization Strategies

1. **Lazy Rendering**
   - Don't render all rows at once
   - Use virtual scrolling for large tables
   - Render only visible rows

2. **Efficient Data Structures**
   - Use Map for O(1) lookups (not arrays)
   - Store row maps instead of scanning arrays
   - Pre-compute header mappings

3. **Early Termination**
   - Skip unchanged rows when possible
   - Don't compare cells in added/deleted rows
   - Short-circuit comparison on first difference

4. **Memory Management**
   - Clear old diff results before new comparison
   - Release DOM elements when switching views
   - Use WeakMap for temporary references

---

### Typical Performance

| File Size | Rows | Columns | Parse Time | Compare Time | Total Time |
|-----------|------|---------|------------|--------------|------------|
| Small | 100 | 10 | < 1s | < 1s | < 2s |
| Medium | 1,000 | 20 | 1-2s | 2-3s | 3-5s |
| Large | 10,000 | 30 | 5-10s | 10-15s | 15-25s |
| Very Large | 50,000+ | 50 | 20-30s | 30-60s | 50-90s |

**Note:** Actual performance depends on:
- Browser (Chrome is fastest)
- Device CPU/RAM
- File complexity (formulas, formatting)
- Number of changes detected

---

## Conclusion

Excel Differ v2.2.0 provides a robust, header-based comparison system with the following key features:

✅ **Flexible Header Row Selection** - Users choose where headers are located  
✅ **Position-Based Comparison Mode** - NEW in v2.2.0: Match rows by position when key column unavailable  
✅ **Custom Key Column Selection** - Any common column can be used for row matching  
✅ **Smart Column Matching** - Handles column reordering gracefully  
✅ **Accurate Row Matching** - Uses Key Column values, not positions  
✅ **Cell-Level Precision** - Detects individual cell changes  
✅ **Unified Diff View** - All changes in one table with old/new indices  
✅ **CSV Support with Encoding Detection** - NEW in v2.2.0: UTF-8, GBK, Big5 with one-time session warning

**Best Practices:**
1. Ensure Key Column contains stable, unique identifiers (or use Position mode for simple lists)
2. Verify header row selection before comparison
3. Review auto-detected common columns
4. Pre-process data to ensure consistency
5. Use appropriate file sizes for browser performance
6. Save CSV files in UTF-8 encoding when possible

For more information, see [README.md](README.md).

---

**Version:** 2.2.0  
**Last Updated:** February 2026  
**Author:** Hackett.Lai
