# 📊 CSV Filter Tool

A simple and powerful tool to **clean and filter your CSV files** — no coding skills needed!

Built with [Streamlit](https://streamlit.io/) and [Pandas](https://pandas.pydata.org/). Runs entirely on your computer — **your data never leaves your machine**.


## 🚀 Features

After uploading a CSV file, the app will automatically:

1. **Remove double quotes** from the data.
2. **Rename and clean up** columns.
3. **Filter rows** based on specific conditions:
   - Excludes rows where **`Förv/Bolag`** is `KulturN`.
   - Excludes rows where **`Semgrp`** is `1`, `8`, `9`, or `22`.
   - Keeps only rows where **`Arbhel`** starts with `35` or `40`.
   - Removes rows where **`Gfom`** is a **weekend** or **public holiday**.
4. Automatically handles decimal formatting and rounding.
5. Provides options to **download the filtered data** as CSV or Excel.



## 🧑‍💻 How to Use (No Coding Required!)

### ✅ Option 1: Just Run the App (Windows Users)

1. Download the project folder.
2. Double-click the **`launch.exe`** file.
3. Your browser will open with the app running locally.
4. Upload your CSV file and let the app do the rest.
5. Download the cleaned file when you're done!

> [!TIP]
> If it's your first time running it, the app will automatically install any missing Python packages.



### ⚙️ Option 2: Run from Source (For Developers / Advanced Users)

#### 1. Install Python

Download Python from: https://www.python.org/downloads/

> Make sure to check **"Add Python to PATH"** during installation.

#### 2. Download the code

Clone or download this repo to your computer.

#### 3. Install required packages

```bash
pip install -r requirements.txt
```

#### 4. Run the app

```bash
streamlit run main.py
```



## 📁 CSV File Format

Your input CSV file should:

- Use `;` (semicolon) as a separator.
- Contain the following columns:
  - `Förv/Bolag`
  - `Semgrp`
  - `Arbhel`
  - `Gfom` (must be in `YYYYMMDD` format)



## 🎄 Public Holidays

The app already includes common Swedish public holidays for 2025 and 2026.

To add more dates, open `main.py` and edit the `PUBLIC_HOLIDAYS` list:

```python
PUBLIC_HOLIDAYS = ["20250101", "20251225", ...]
```



## 💾 Output

After filtering, you can download the result as:

- ✅ CSV file
- ✅ Excel file (.xlsx)

The cleaned file will retain the original filename with `-result` added.



## 🙋 Need Help?

If something doesn’t work:

- Check your CSV file format and column names.
- Ensure `Gfom` dates are in the correct format (`YYYYMMDD`).
- Try running the app again — it will guide you with clear messages.
