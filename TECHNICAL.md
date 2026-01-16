# Technical Documentation

## Table of Contents

1. [Processing Pipeline Overview](#processing-pipeline)
2. [File Reading & Parsing](#1-file-reading--parsing)
3. [Column Matching (Header-Based)](#2-column-matching-header-based)
4. [Row Matching (Column A-Based)](#3-row-matching-column-a-based)
5. [Cell Comparison](#4-cell-comparison)
6. [Complete Example](#complete-example)
7. [Implementation Details](#implementation-details)
8. [Edge Cases & Limitations](#edge-cases--limitations)
9. [Performance Considerations](#performance-considerations)

---

## Architecture Overview
```
┌─────────────┐ ┌─────────────┐
│    File A   │ │    File B   │
│   (Excel)   │ │   (Excel)   │
└──────┬──────┘ └──────┬──────┘
       │               │
       └──────┬────────┘
              ▼
       ┌───────────────┐
       │    SheetJS.   │
       │     Parser    │
       └───────┬───────┘
               ▼
       ┌───────────────┐
       │   JavaScript  │
       │      Object   │
       └───────┬───────┘
               │
      ┌────────┴────────┐
      ▼                 ▼
┌─────────────┐ ┌─────────────┐
│   Column    │ │     Row.    │
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
1. Excel files → SheetJS → JS Objects
2. Extract headers & build column maps
3. Extract Column A values & build row maps
4. Compare cells at intersections
5. Render unified diff table

---

## Processing Pipeline

### 1. File Reading & Parsing

```
Excel/CSV File → SheetJS (xlsx) → JavaScript Object
```

**Input Files:**

- File A (old version) and File B (new version)
- Supported formats: `.xlsx`, `.xls`, `.csv`

**Parsing Result:**

```javascript
{
  "A1": "Name",
  "B1": "Email",
  "C1": "Phone",
  "A2": "John Doe",
  "B2": "john@email.com",
  "C2": "123-456",
  // ... more cells
}
```

---

### 2. Column Matching (Header-Based)

**Step 2.1: Extract Headers (Row 1)**

```javascript
File A Headers: { A: "Name", B: "Email", C: "Phone" }
File B Headers: { A: "Name", B: "Phone", C: "Email" }
```

**Step 2.2: Build Header-to-Column Maps**

```javascript
// Map header content → column letter
File A: { "Name" → "A", "Email" → "B", "Phone" → "C" }
File B: { "Name" → "A", "Phone" → "B", "Email" → "C" }
```

**Step 2.3: Detect Column Changes**

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
  matched: ["Name", "Email", "Phone"],
  added: [],
  deleted: []
}
```

**Key Point:** Columns are matched by **header content**, not position.

- "Email" in column B (File A) matches "Email" in column C (File B) ✓

---

### 3. Row Matching (Column A-Based)

**Step 3.1: Extract Row Keys from Column A**

```javascript
File A:
  Row 2: key = "John Doe"  (from A2)
  Row 3: key = "Jane Doe"  (from A3)

File B:
  Row 2: key = "Jane Doe"  (from A2) ← Moved up
  Row 3: key = "John Doe"  (from A3) ← Moved down
```

**Step 3.2: Build Row Maps**

```javascript
File A: { "John Doe" → Row 2, "Jane Doe" → Row 3 }
File B: { "Jane Doe" → Row 2, "John Doe" → Row 3 }
```

**Step 3.3: Detect Row Changes**

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
    { key: "John Doe", oldRow: 2, newRow: 3 },
    { key: "Jane Doe", oldRow: 3, newRow: 2 }
  ],
  added: [],
  deleted: []
}
```

**Key Point:** Rows are matched by **Column A value**, not position.

- Row with "John Doe" in A2 (File A) matches "John Doe" in A3 (File B) ✓

---

### 4. Cell Comparison

**For each MATCHED row:**

```
For each MATCHED column header:
  1. Find the column letter in File A (via header map)
  2. Find the column letter in File B (via header map)
  3. Get cell value from File A at (row, column A)
  4. Get cell value from File B at (row, column B)
  5. Compare the two values
     ├─ If same → No change
     └─ If different → MODIFIED 🔄
```

**Example:**

```
Row Key: "John Doe"
Header: "Email"

Step 1: Find column for "Email" in File A → Column B
Step 2: Find column for "Email" in File B → Column C (moved!)
Step 3: Get File A value at (Row 2, Column B) → "john@old.com"
Step 4: Get File B value at (Row 3, Column C) → "john@new.com"
Step 5: Compare → DIFFERENT → Mark as MODIFIED
```

**Cell Change Result:**

```javascript
{
  rowKey: "John Doe",
  header: "Email",
  oldValue: "john@old.com",
  newValue: "john@new.com",
  changeType: "modified"
}
```

---

### Complete Example

**File A:**

```
| A: Name     | B: Email          | C: Phone    |
|-------------|-------------------|-------------|
| John Doe    | john@old.com      | 111-1111    |
| Jane Doe    | jane@email.com    | 222-2222    |
| Bob Smith   | bob@email.com     | 333-3333    |
```

**File B:**

```
| A: Name     | B: Phone    | C: Email          | D: Department |
|-------------|-------------|-------------------|---------------|
| Jane Doe    | 222-2222    | jane@email.com    | Sales         |
| John Doe    | 111-1111    | john@new.com      | IT            |
| Alice Wong  | 444-4444    | alice@email.com   | HR            |
```

**Processing Steps:**

1. **Column Matching:**

   - ✓ Matched: "Name", "Email", "Phone"
   - ✅ Added: "Department"
   - ❌ Deleted: (none)

2. **Row Matching:**

   - ✓ Matched: "John Doe" (Row 2→3), "Jane Doe" (Row 3→2)
   - ✅ Added: "Alice Wong"
   - ❌ Deleted: "Bob Smith"

3. **Cell Comparison for "John Doe":**

   - Name: "John Doe" = "John Doe" → No change
   - Email: "john@old.com" ≠ "john@new.com" → 🔄 MODIFIED
   - Phone: "111-1111" = "111-1111" → No change

4. **Cell Comparison for "Jane Doe":**
   - All cells unchanged (despite row moving from 3→2)

**Final Diff Summary:**

- 1 column added
- 1 row added, 1 row deleted
- 1 cell modified
