# 📊 Excel Pivot Table Generator & Allocation Reporting Tool

A powerful **desktop-based workforce analytics application** built using **Python, PyQt5, and Pandas** to automate allocation tracking, generate insightful reports, and streamline reporting workflows with Excel and Outlook integration.

---

## 🚀 Features

### 📌 Data Processing & Automation

* Reads and processes Excel-based allocation data
* Intelligent column mapping and normalization
* Handles inconsistent formats and missing data gracefully

### 📊 Report Generation

Generates **7 comprehensive reports**:

1. **Client Allocation Summary**
2. **Monthly Headcount Summary**
3. **Total Headcount Summary**
4. **Allocation Comparison (Day & Week trends)**
5. **Billed Employee Details**
6. **Unbilled Employee Details**
7. **Allocations Ending This Month**

---

### 📈 Advanced Analytics

* Tracks historical allocation changes
* Compares:

  * Today vs Previous Working Day
  * Today vs Previous Week
* Calculates:

  * Utilization %
  * Billed %
  * Allocation trends

---

### 🧠 Smart Data Handling

* Automatic SOW categorization (Billed, Bench, PreSow, etc.)
* Skill data integration via `skills.csv`
* Deduplication and validation of historical records

---

### 🖥️ User Interface (PyQt5)

* Interactive desktop GUI
* File browser for Excel upload
* Date selector for custom reporting
* Real-time table rendering (HTML view)
* Auto-refresh on file changes

---

### 📤 Export & Sharing

* Export reports to formatted Excel files
* Auto-adjusted column widths and styling
* Color-coded client rows
* Send reports via **Microsoft Outlook integration**

---

## 🛠️ Tech Stack

* **Language:** Python 3.x
* **Libraries:**

  * Pandas
  * PyQt5
  * OpenPyXL / XlsxWriter
  * Win32com (Outlook Automation)
* **UI Framework:** PyQt5

---

## 📂 Project Structure

```
.
├── main.py                     # Main application
├── skills.csv                 # Optional skills mapping file
├── allocation_history.xlsx    # Auto-generated history tracking
├── icon.png                   # App icon
└── README.md
```

---

## ⚙️ Setup & Installation

### 1️⃣ Clone the Repository

```
git clone https://github.com/your-username/excel-pivot-app.git
cd excel-pivot-app
```

### 2️⃣ Install Dependencies

```
pip install pandas pyqt5 openpyxl xlsxwriter pywin32
```

### 3️⃣ Run the Application

```
python main.py
```

---

## 📥 Input Requirements

Your Excel file should contain (minimum required columns):

* `Date`
* `Client`
* `Sow`
* `No. of Employees`
* `Start Date`
* `End Date`
* `Employee Code`

Optional but recommended:

* `% Allocation`
* Project & employee details
* Skills data (via `skills.csv`)

---

## 📊 Skills Integration (Optional)

Place a `skills.csv` file in the same directory with columns:

* Employee Code
* Skill Item
* Skill Item Category
* Skill Item Category Group

The app will automatically merge and display skills in reports.

---

## 📧 Email Automation

* Sends reports via **Microsoft Outlook**
* Configurable recipients:

  * Company-specific recipients
  * Universal recipients
* Attaches generated Excel report

---

## ⚠️ Known Limitations

* Outlook must be installed for email functionality
* Designed for Windows (due to `win32com`)
* Large files may impact performance
* Requires consistent Excel structure for best results

---

## 🔮 Future Enhancements

* 📉 Data visualization (charts & graphs)
* 🌐 Web-based version (Flask/React)
* 📊 Dashboard with filters & drill-downs
* 🌙 Dark mode UI
* 📄 PDF export support

---

## 💡 Use Case

Ideal for:

* Workforce allocation tracking
* Resource management teams
* Delivery & operations reporting
* Business analytics automation

---

## 👩‍💻 Author

**Samridhi Aggarwaal**
BTech CSE | Data & Software Enthusiast

---

## ⭐ If you like this project

Give it a ⭐ on GitHub and feel free to contribute!

---
