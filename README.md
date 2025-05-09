# spreadsheet-showcase
A modern Vue 3 + Handsontable spreadsheet showcase featuring formatting, persistence, custom toolbars, and enhanced UX patterns — designed to demonstrate an interactive data table experience in a Vue app.

# 🧮 Spreadsheet Showcase App

This project is a **Vue 3 + Handsontable** demonstration of building a powerful, user-friendly spreadsheet experience for the web. It showcases advanced UI techniques, live formatting, persistent settings, and reactive integration with modern frontend practices.

---

## ✨ Features

- ✅ **Live Spreadsheet Editing** powered by Handsontable
- 🎨 **Custom Formatting Tools** (bold, italic, underline, strikethrough, cell/background color, font size)
- 📌 **Freeze Top Row / First Column** toggle
- 🔍 **Search and Highlight** keywords and matches
- ⚠️ **Error highlighting** for specific columns
- 📊 **Column Type Analysis**
- 🦓 **Zebra striping** for alternating row backgrounds
- 💾 **Auto-Save + Persistent State** for:
  - Data values
  - Formatting styles
  - Row and column order
- 🔁 **Undo-friendly Paste** with custom row parsing
- 🧪 **Mock Data Generator** (30k rows for stress testing)
- 🧱 **Border Selection UI** with visual controls
- ⬇️ **CSV Export** for output spreadsheet
- Supports building as .exe via electron

---

## 🔧 Tech Stack

- [Vue 3](https://vuejs.org/) – Frontend framework
- [Handsontable](https://handsontable.com/) – Spreadsheet engine
- [Vite](https://vitejs.dev/) – Lightning-fast dev tooling

---

## 🗂️ Project Structure

```bash
src/
├── assets/
│   ├── css/     
│   │   ├── global-handsontable-styles.css      # global css for sheet formatting
├── components/
│   ├── SpreadsheetComponent.vue      # Core spreadsheet wrapper using Handsontable
│   ├── BorderModal.vue               # Visual modal for applying border styles
├── router/
│   ├── index.js                      # vue router to support navbar and future tools
├── pages/
│   ├── HomePage.vue                  # Landing page with links
│   ├── AboutPage.vue                 # About this app
│   ├── ContactPage.vue               # Contact info
│   ├── SpreadsheetShowcase.vue       # Main showcase tool page
├── utils/
│   └── SpreadsheetUtils.js           # All formatting and logic functions
