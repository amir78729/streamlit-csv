# 📊 CSV Filter Tool

This is a simple web app that helps you filter and clean your CSV files — now with Excel output support! You can use it without writing any code.

Built with [Streamlit](https://streamlit.io/) and [Pandas](https://pandas.pydata.org/).


## 🚀 What does it do?

After uploading a CSV file, the app will:

1. **Remove double quotes** from the data.
2. **Drop** the `Orsak` column (if present).
3. **Rename** the `Orstxt` column to `Frånvaro`.
4. **Filter out rows** where:
   - `Förv/Bolag` is `KulturN`
   - `Semgrp` is `1`, `8`, `9`, or `22`
   - `Arbhel` does **not** start with `35` or `40`
   - `Gfom` falls on a **weekend** or **public holiday**
   - `Frånvaro` contains specific unwanted values
5. **Convert commas to dots** in numeric columns: `Arbhel`, `Omf`, `Specantal`, `Antal`
6. **Round** numeric values to two decimal places
7. Show both the **raw** and **filtered** results
8. Let you **download** the cleaned data as either **CSV** or **Excel**

> 💡 All processing happens locally on your device. Your data is never uploaded or shared.


## 🛠 How to run it

### Option 1: Manual setup

#### 1. Install Python

Download and install Python from: https://www.python.org/downloads/

Make sure to check the box **"Add Python to PATH"** during installation.

#### 2. Download the project

You can download or clone this project to your computer.

#### 3. Install the required libraries

Open your terminal or command prompt, go to the project folder, and run:

```bash
pip install -r requirements.txt
```

Alternatively, for Python 3 explicitly:

```bash
pip3 install -r requirements.txt
```

#### 4. Start the app

```bash
streamlit run csv-filter.py
```

The app will launch in your browser! 🎉


### Option 2: Use Windows `.bat` files

For Windows users, you can use these helper scripts:

- `run-app.bat` — Starts the app automatically
- `run-installation.bat` — Installs required libraries using `pip`
- `run-installation-v3.bat` — Installs required libraries using `pip3`


## 📁 Your CSV File

- The CSV must use `;` (semicolon) as a separator.
- It should include the following columns:
  - `Förv/Bolag`
  - `Semgrp`
  - `Arbhel`
  - `Gfom` (date in `YYYYMMDD` format)
  - `Orsak`, `Orstxt`, `Frånvaro`, `Omf`, `Specantal`, `Antal` (optional, but recommended)


## 🎄 Public Holidays

To recognize public holidays, you can edit the `PUBLIC_HOLIDAYS` list in the script:

```python
PUBLIC_HOLIDAYS = [
    "20250101", "20250418", "20251225", ...
]
```

Dates must be in `YYYYMMDD` format.


## 💾 Output

After filtering, you’ll be able to download the cleaned data as:

- A `CSV` file
- An `Excel` file (`.xlsx`) with a sheet named `FilteredData`


## 🙋 Need Help?

If something doesn’t work, double-check:

- Your file uses `;` as a separator
- The required columns exist and are spelled correctly
- The `Gfom` column uses the `YYYYMMDD` format
```
