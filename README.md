# 🐼 Panda Employee Attendance System

An executable employee attendance system powered by **face recognition**, designed with a simple GUI for local use. This application allows employers to manage employee attendance records, apply direct deductions, and export monthly PDF reports — all while using a webcam for real-time recognition.

---

## ✨ Features

- 🧑‍💼 Add, update, and delete employees with face encoding
- 🎥 Real-time face recognition-based attendance logging
- 📝 Manual attendance input & direct deductions
- 📟 Monthly PDF report generation in **Arabic** with employee records and salary placeholders
- 💾 Data stored locally in CSV files (`Employees.csv`, `Monthly_Data.csv`)
- 🔒 Secured admin actions (password: `2255`)
- 📷 Saves employee face snapshots and uses them for recognition
- 🍍 Arabic language support in GUI and reports (right-to-left PDF rendering)

---

## 💥 Tech Stack

- Python 3.10+
- Tkinter (GUI)
- OpenCV & Dlib (Face recognition)
- Pandas & CSV (Data storage)
- ReportLab (PDF generation)
- Arabic reshaper + Bidi (for Arabic support in PDFs)

---

## 📁 Project Structure

```
📆 Panda_Employee_System/
├\2500 Panda_Employee_System.py         # Main application script
├\2500 data/
│   ├\2500 Employees.csv                # Employee database
│   └\2500 Monthly_Data.csv            # Daily attendance records
├\2500 images/
│   └\2500 [national_id].png           # Employee face snapshots
├\2500 reports/
│   └\2500 [national_id].pdf           # Monthly exported reports
├\2500 face_recognition_models/
│   └\2500 models/                     # Dlib model files
└\2500 fonts/
    └\2500 Amiri-Regular.ttf           # Arabic font for PDF
```

---

## ⚙️ Requirements

Install dependencies using:
```bash
pip install -r requirements.txt
```

Example dependencies:
```txt
opencv-python
face_recognition
numpy
pandas
reportlab
Pillow
arabic_reshaper
python-bidi
```

---

## 🚀 How to Run

Run the main application:
```bash
python Panda_Employee_System.py
```

To compile it into a standalone `.exe` (for Windows deployment):
```bash
pyinstaller --onefile --windowed Panda_Employee_System.py
```

> Make sure to include required model files and fonts in your PyInstaller bundle!

---

## 🔒 License

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International License**.  
See the [LICENSE](./LICENSE) file for full terms.

Commercial use is **strictly prohibited** without written permission.

---

## 👤 Author

Developed by **Mahmoud Elgendy**  
For a custom client project under specific constraints.

