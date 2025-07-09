# 🗡️ SWORD SCHOOL - Student Report Card Generator

A modern and classy GUI-based Python app to read student data from Excel and display individual **report cards** with sleek progress bars, navigation, and one-click **PDF export**.

<p align="center">
  <img src="logo.png" width="100" />
</p>

## 🔥 Features

- 🎓 School Title: `SWORD SCHOOL` with subtitle `SPSAC`
- 📖 Reads student data from `students.xlsx`
- 📊 Visual progress bars for subjects (English, Math, Science)
- ⏮️ ⏭️ Navigate students with **Previous** / **Next** buttons
- 📄 Export each report card as a clean PDF
- ✍️ Signed footer by: **add your name**
- 💅 Aesthetic and modern GUI built with Tkinter

---

## 📁 Files Structure

📦 SWORD-Report-Card
┣ 🖼️ logo.png # School logo
┣ 📊 students.xlsx # Student data in Excel format
┣ 🐍 report_card_gui.py # Main GUI application
┗ 📄 README.md # This file


---
![project](https://github.com/user-attachments/assets/be750dd1-162b-4506-93ff-f79cc4711b4a)


## 📥 Installation

Make sure you have Python 3 installed.

### 1. Install dependencies:

pip install pandas pillow fpdf openpyxl

2. Run the App:
python report_card_gui.py

📘 Sample Excel Format

| Name        | Roll No | Class | English | Math | Science | Total | Grade |
|-------------|---------|-------|---------|------|---------|-------|-------|
| Ahmed Ali   | 101     | 9th   | 85      | 90   | 78      | 253   | A     |

📤 Exported PDF Sample
The exported PDF includes:

Student Info

Subject Marks Table

Total & Grade

Footer: Signed by (your name) 

📃 License
MIT License — Use freely, modify, and share 🙌

Built with ❤️ by Sohaib for  SWORD SCHOOL - SPSAC (sample name) 



