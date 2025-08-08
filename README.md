# 📊 Split Excel Data into Separate Workbooks Using Power Automate Desktop

## 🔍 Overview

This Power Automate Desktop flow automates the process of splitting a multi-sheet Excel workbook into multiple separate Excel files. Each new file corresponds to a different worksheet from the source workbook.

Ideal for reporting, distribution, or archiving, this solution removes the need for manual copying or complex VBA scripting.

---

## ⚙️ Features

✅ Automatically opens a source Excel workbook  
✅ Loops through each worksheet  
✅ Copies worksheet content into a new workbook  
✅ Saves each worksheet as its own Excel file (named after the sheet)  
✅ Optionally deletes processed sheets from the source file  
✅ 100% no-code — built entirely with Power Automate Desktop actions

---

## 🔁 How It Works

1. **Launch Excel:** Open the main workbook  
2. **Loop Worksheets:** For each sheet:
   - Get worksheet name
   - Launch a new Excel instance
   - Copy all cells
   - Paste into a blank workbook
   - Save the file using the sheet name
   - Optionally delete the original sheet
3. **Cleanup:** Close all Excel files

---

## 🧠 Visual Workflow

![Power Automate Desktop Flow Screenshot](<img width="1600" height="900" alt="Screenshot (167)" src="https://github.com/user-attachments/assets/fa04e0e8-c8ec-4705-a0ed-9a4f6670ffa2" />
)

---

## 🚀 How to Use

1. Update the source file path in the **Launch Excel** action  
2. Set your desired **output folder** and **file naming pattern**  
3. (Optional) Modify loop range or add filters for sheet names  
4. Run the flow — each worksheet will be saved as an individual Excel file

---

## ✅ Benefits

- ❌ No VBA or coding needed  
- ✅ Reduces manual effort  
- ✅ Prevents copy-paste errors  
- ✅ Reusable for multiple reporting tasks

---

## 💡 Customization Ideas

- Export only specific worksheets (e.g., based on prefix or name)
- Add email/Teams notifications after completion
- Extend to SharePoint or OneDrive cloud flows

---

## 🏁 Status

✅ Tested and working with sample data  
📂 Compatible with Excel Desktop via Power Automate Desktop

---

