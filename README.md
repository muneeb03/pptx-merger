# 🚀 Bulk PPT to PDF Converter 🖥️➡️📄

## 📌 Overview
This project provides a **GUI and CLI tool** to convert **PowerPoint files** (`.ppt` and `.pptx`) into **PDFs**. 
- The **GUI version** is built using **PyQt5** 🎨.
- The **CLI version** processes multiple files in a directory 📂.

## ✨ Features
✅ Convert multiple PowerPoint files to PDFs.
✅ GUI-based conversion using PyQt5.
✅ Command-line interface (CLI) for bulk conversion.
✅ Progress tracking using multi-threading.

## ⚙️ Installation
Make sure you have Python installed, then install dependencies:
```sh
pip install -r requirements.txt
````

## 🚀 Usage

### 🖥️ GUI Mode

Run the following command to start the GUI version:

```sh
python frontend.py
```

### 💻 CLI Mode

To convert all PowerPoint files in a directory:

```sh
python pptx.py
```

You'll be prompted to enter the directory path.

## 📦 Dependencies

- 🖌️ `PyQt5` - GUI Interface.
- 🔄 `comtypes` & `pywin32` - PowerPoint Automation.
- 📊 `tqdm` - Progress Tracking in CLI Mode.

## ⚠️ Notes

- **Microsoft PowerPoint must be installed** on Windows. 🏢
- **Close PowerPoint** before starting the conversion. ❌📊
- Works **only on Windows** due to PowerPoint COM dependency. 🖥️

```
```
