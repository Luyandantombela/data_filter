# Data Filter — Excel Add-in

Filter spreadsheet rows by column values and export matching rows to a new sheet.

## Files to upload to this repo

Upload all files **keeping the same folder structure**:

```
data_filter/
├── manifest.xml          ← sideload this into Excel
├── taskpane.html
├── taskpane.js
├── commands.html
├── .nojekyll
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png
```

## Setup steps

### 1. Enable GitHub Pages

1. Go to your repo → **Settings** → **Pages**
2. Under **Source**, choose **Deploy from a branch**
3. Select branch: `main`, folder: `/ (root)`
4. Click **Save**
5. Wait ~1 minute. Your site will be live at:  
   `https://luyandantombela.github.io/data_filter/`

### 2. Sideload the add-in into Excel

**On Excel for Windows (desktop):**
1. Open Excel
2. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Click **Trusted Add-in Catalogs**
4. Add catalog URL: `https://luyandantombela.github.io/data_filter/`
5. Check **Show in Menu**, click **OK**, restart Excel
6. Go to **Insert** → **My Add-ins** → **Shared Folder** → select **Data Filter**

**Alternatively — quick sideload via manifest file:**
1. Go to **Insert** → **My Add-ins** → **Upload My Add-in**
2. Browse to and select `manifest.xml`
3. Click **Upload**

### 3. Use the add-in

1. Click **Open Data Filter** from the **Home** tab ribbon
2. **Step 1** — Select your data (including header row) and click **Capture Selection**
3. **Step 2** — Click **+ Add Filter**, choose a column, type the value to match
   - Add more filters for more columns (AND logic)
   - Add the same column again with a different value (OR logic)
4. Click **✦ Apply Filters & Create New Sheet**
5. A new sheet named `Filtered_HHMMSS` is created with all matching rows

## How filtering works

| Scenario | Logic |
|----------|-------|
| Filter A (Country = ZA) + Filter B (Status = Active) | AND — row must match both |
| Filter A (Country = ZA) + Filter B (Country = UK) | OR — row matches if country is ZA or UK |
| Mix of same + different columns | AND between columns, OR within each column |
