import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import ImageTk, Image
import os
from fpdf import FPDF

df = pd.read_excel("students.xlsx")
students = df.to_dict(orient="records")
index = 0

root = tk.Tk()
root.title("SWORD SCHOOL - SPSAC")
root.geometry("560x750")
root.configure(bg="#fafafa")

# 🔥 SCHOOL HEADER
tk.Label(root, text="SWORD SCHOOL", font=("Georgia", 24, "bold"), bg="#fafafa", fg="#D7263D").pack(pady=(15, 0))
tk.Label(root, text="SPSAC", font=("Segoe UI", 14, "italic"), bg="#fafafa", fg="#555").pack(pady=(0, 5))

# 🔥 LOGO
try:
    logo_img = Image.open("logo.png").resize((65, 65))
    logo = ImageTk.PhotoImage(logo_img)
    tk.Label(root, image=logo, bg="#fafafa").pack(pady=5)
except:
    tk.Label(root, text="Logo Missing", fg="red", bg="#fafafa").pack()

ttk.Separator(root, orient='horizontal').pack(fill='x', pady=10)

# 🔐 Vars
name_var = tk.StringVar()
roll_var = tk.StringVar()
class_var = tk.StringVar()
english_var = tk.StringVar()
math_var = tk.StringVar()
science_var = tk.StringVar()
total_var = tk.StringVar()
grade_var = tk.StringVar()

# 🔍 Update GUI
def show_student(i):
    student = students[i]
    name_var.set(student['Name'])
    roll_var.set(student['Roll No'])
    class_var.set(student['Class'])
    english_var.set(student['English'])
    math_var.set(student['Math'])
    science_var.set(student['Science'])
    total_var.set(student['Total'])
    grade_var.set(student['Grade'])
    update_bar(eng_bar, student['English'])
    update_bar(math_bar, student['Math'])
    update_bar(sci_bar, student['Science'])

def next_student():
    global index
    if index < len(students) - 1:
        index += 1
        show_student(index)

def prev_student():
    global index
    if index > 0:
        index -= 1
        show_student(index)

# 🧾 Info section
card = tk.Frame(root, bg="white", bd=2, relief="groove")
card.pack(padx=30, pady=10, fill="x")

def info_row(label, var):
    frame = tk.Frame(card, bg="white")
    frame.pack(anchor="w", padx=20, pady=5)
    tk.Label(frame, text=label, font=("Helvetica", 12, "bold"), bg="white", fg="#3c3c3c").pack(side="left", padx=(0, 10))
    tk.Label(frame, textvariable=var, font=("Helvetica", 12), bg="white", fg="#1f6f8b").pack(side="left")

info_row("Name:", name_var)
info_row("Roll No:", roll_var)
info_row("Class:", class_var)

# 📊 Marks section
marks = tk.LabelFrame(root, text="Performance", font=("Arial", 12, "bold"), bg="white", fg="#D7263D", padx=20, pady=10)
marks.pack(padx=30, pady=15, fill="x")

def marks_row(label, var):
    tk.Label(marks, text=label, font=("Arial", 11), bg="white").pack(anchor="w", pady=2)
    bar = ttk.Progressbar(marks, orient="horizontal", length=400, mode="determinate")
    bar.pack(pady=(0, 5))
    return bar

tk.Label(marks, text="English", bg="white").pack(anchor="w")
eng_bar = ttk.Progressbar(marks, orient="horizontal", length=400, mode="determinate")
eng_bar.pack(pady=3)

tk.Label(marks, text="Math", bg="white").pack(anchor="w")
math_bar = ttk.Progressbar(marks, orient="horizontal", length=400, mode="determinate")
math_bar.pack(pady=3)

tk.Label(marks, text="Science", bg="white").pack(anchor="w")
sci_bar = ttk.Progressbar(marks, orient="horizontal", length=400, mode="determinate")
sci_bar.pack(pady=3)

info_row("Total:", total_var)
info_row("Grade:", grade_var)

# 🔄 Bar Updater
def update_bar(bar, value):
    try:
        val = min(100, int((value / 100) * 100))
        bar['value'] = val
    except:
        bar['value'] = 0

# 🖨️ Export to PDF
def export_pdf():
    student = students[index]
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 10, "SWORD SCHOOL - SPSAC", ln=True, align="C")
    pdf.ln(10)
    pdf.set_font("Arial", "", 12)
    pdf.cell(50, 10, f"Name: {student['Name']}", ln=True)
    pdf.cell(50, 10, f"Roll No: {student['Roll No']}", ln=True)
    pdf.cell(50, 10, f"Class: {student['Class']}", ln=True)
    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(50, 10, "Subject", 1)
    pdf.cell(40, 10, "Marks", 1, ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.cell(50, 10, "English", 1)
    pdf.cell(40, 10, str(student['English']), 1, ln=True)
    pdf.cell(50, 10, "Math", 1)
    pdf.cell(40, 10, str(student['Math']), 1, ln=True)
    pdf.cell(50, 10, "Science", 1)
    pdf.cell(40, 10, str(student['Science']), 1, ln=True)
    pdf.ln(10)
    pdf.cell(50, 10, f"Total: {student['Total']}")
    pdf.ln(5)
    pdf.cell(50, 10, f"Grade: {student['Grade']}")
    pdf.ln(20)
    pdf.cell(50, 10, "Signed by: Saad Ahmed")
    filename = f"{student['Name'].replace(' ', '_')}_report.pdf"
    pdf.output(filename)
    messagebox.showinfo("Saved", f"{filename} exported!")

# 🔘 Buttons
btn_frame = tk.Frame(root, bg="#fafafa")
btn_frame.pack(pady=10)

ttk.Button(btn_frame, text="⏪ Previous", command=prev_student).grid(row=0, column=0, padx=10)
ttk.Button(btn_frame, text="Next ⏩", command=next_student).grid(row=0, column=1, padx=10)
ttk.Button(root, text="📄 Export as PDF", command=export_pdf).pack(pady=10)

# 🖊️ Signature
tk.Label(root, text="Signed by: Saad Ahmed", font=("Segoe UI", 10, "italic"), bg="#fafafa", fg="#555").pack(pady=10)

# 👀 Start
show_student(index)
root.mainloop()
