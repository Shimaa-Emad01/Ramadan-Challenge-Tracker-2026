# 🌙 Ramadan Challenge Tracker 2025 ✨

A creative and spiritual way to make Ramadan more productive and rewarding.

---

## 🕌 Project Overview

**Ramadan Challenge Tracker 2025** is an interactive Excel workbook designed to help you maintain **consistency and motivation** in your worship and good deeds throughout the blessed month of Ramadan.  
Instead of traditional manual tracking, this workbook provides a **clean and modern Dark Mode interface** powered by **advanced Excel formulas and VBA macros**, turning daily tracking into a **fun, colorful, and meaningful experience**.

---

## ✨ Core Features

This project combines the power of advanced Excel functions with VBA to deliver an engaging experience:

- ✅ **Interactive Checkboxes:** Simply click to mark completion; each checkbox automatically links to the backend sheet (`True_False`).
- 🔢 **Automatic Achievement Calculation:** Uses advanced formulas like **`SUMPRODUCT`** to calculate progress per day and per category.
- 🎨 **Dynamic Conditional Formatting:** Cells and charts change color dynamically with your progress using **gradient color scales**.
- 💬 **Random Motivational Messages:** Weekly goal section shows random encouraging messages using **`INDEX + RANDBETWEEN`** for motivation.
- ⚙️ **Advanced VBA Modules:** Two modules manage:
  - Auto creation and linking of checkboxes.
  - Dynamic visual glow effects on progress cells.
- 🌈 **Aesthetic “Dark Mode” Design:** Visually calming design highlighting early, middle, and late Ramadan progress.

---

## ⚙️ How It Works

The structure is simple and intuitive:

1. **Daily Tracking:** Each column represents a good deed, and each row represents a Ramadan day (1–30).  
2. **Automatic Update:** Clicking a checkbox (✅) records `TRUE` in the `True_False` sheet.  
3. **Weekly Goals:** Formulas automatically calculate your weekly progress and display motivational feedback like:
   - “🩷 Keep going, you’re doing amazing!”
   - “💫 Don’t worry, next week is a new chance!”

---

## 🧠 VBA & Formulas Used

| Type | Formula / Code | Function |
| :--- | :--- | :--- |
| **Total Calculation** | `=SUMPRODUCT(--(True_False!B16:AE200=TRUE))` | Counts all completed tasks during Ramadan. |
| **Weekly Evaluation** | `=IF(SUMPRODUCT(--(True_False!B16:H20=TRUE))=35, "Motivational message", "Try again next week!")` | Checks if weekly goal (e.g., 35 actions) is achieved. |
| **Random Quotes** | `=INDEX({...}, RANDBETWEEN(1,3))` | Displays random motivational messages. |
| **VBA Modules** | `Module1` & `Module2` | Handle checkbox creation, linking, and dynamic effects. |

---

## 📂 File Structure

- **Sheet1 (Dashboard):** Main interface showing the tracking table, total, and motivation.  
- **True_False (Data):** Hidden sheet storing TRUE/FALSE values for checkboxes.  
- **Module1 & Module2 (VBA):** Backend automation for interactivity and visuals.

---

## 🚀 How to Use

1. **Download the file:** `finalProject_تحدى_شهر_رمضان.xlsm`
2. **Enable Macros:** When opening in Excel, click **“Enable Content/Macros”** to allow interactivity.  
3. **Start Tracking:** Mark each daily checkbox to log your progress.  
4. **Explore the Code (optional):** Press `Alt + F11` to open the VBA editor.

---

## 💡 Future Improvements

- 📊 **Additional Dashboard:** Add more charts to visualize progress trends.  
- 🧮 **Custom Points System:** Assign custom points per deed (e.g., Fajr = 2 points, Sunnah = 1).  
- ☁️ **Cloud Integration:** Connect to Google Sheets via Apps Script for online sharing.

---

## 🩵 Credits

**Designed by Shimaa Emad — Ramadan Challenge Tracker 2025 🌙✨**
