# 📊 CSV Filter Tool

This is a simple web app that helps you filter and clean your CSV files. You can use it without knowing how to code!

Built with [Streamlit](https://streamlit.io/) and [Pandas](https://pandas.pydata.org/).


## 🚀 What does it do?

After uploading a CSV file, the app will:

1. **Remove double quotes** from the data.
2. **Filter out rows** based on values in specific columns:
   - Removes rows where **`Förv/Bolag`** is `KulturN`.
   - Removes rows where **`Semgrp`** is `1`, `8`, `9`, or `22`.
   - Removes rows where **`Arbhel`** is `35` or `40`.
   - Removes rows where **`Gfom`** is a **weekend** or **public holiday**.
3. Show the **filtered results**.
4. Let you **download the cleaned file**.


## 🛠 How to run it

### 1. Install Python (if you haven’t already)

Download and install Python from: https://www.python.org/downloads/

Make sure to check the box **"Add Python to PATH"** during installation.


### 2. Download the code

You can download or clone this project to your computer.


### 3. Install the required libraries

Open your terminal or command prompt, go to the project folder, and run:

```bash
pip install -r requirements.txt
```


### 4. Start the app

Run this command in your terminal inside the project folder:

```bash
streamlit run csv-filter.py
```

The app will open in your browser! 🎉


## 📁 Your CSV File

- Make sure your CSV uses `;` (semicolon) as a separator.
- It should have the following columns:
  - `Förv/Bolag`
  - `Semgrp`
  - `Arbhel`
  - `Gfom` (date column in `YYYYMMDD` format)


## 🎄 Public Holidays

To make it work with public holidays, add them to the `PUBLIC_HOLIDAYS` list in the script like this:

```python
PUBLIC_HOLIDAYS = ["20231225", "20240101"]
```


## 💾 Output

After filtering, you’ll be able to download a new, cleaned CSV file.


## 🙋 Need Help?

If something doesn’t work, double-check:
- Your file format
- Column names
- Date format in the `Gfom` column

