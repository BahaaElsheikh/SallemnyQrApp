# SallemnyQrApp
📦 Sallemny
Sallemny is a lightweight Python application designed to streamline the manual process of collecting and recording assignment submissions in universities using QR codes.
🚀 Features
Student data registration with unique QR code generation.
Simple mobile-like interface using Kivy.
Built-in camera scanner using OpenCV for QR code detection.
Auto-save scanned data (Name – Student ID – Submission Time) into a local database.
Export submissions as Excel or Word reports.
📱 App Structure
🧑‍🎓 User App
Enter student details (Name, Faculty, ID).
Generate a QR code.
Save QR for scanning.
🧑‍💼 Admin App
Open camera and scan student QR codes.
Automatically log submission info in SQLite database.
View submission history.
Export reports (.xlsx / .docx).
⚙️ Tech Stack
Python
Kivy
OpenCV
Pyzbar
SQLite3
xlsxwriter / python-docx / pandas
📌 Future Improvements
Google Sheets or Firebase DB integration.
Student login using university email.
Real-time validation & duplicate prevention.
Full web or native mobile version.
