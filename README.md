# 📝 Markdown to PDF Converter
**Effortlessly convert Markdown files into high-quality PDFs using Pandoc & Microsoft Word Automation.**

## 🚀 Features
- ✅ **Batch conversion** of `.md` to `.pdf`
- 📝 Converts **Markdown → DOCX → PDF**
- 🖥 **Automatic dependency handling** (Python, `pywin32`)
- 📁 **Organized output** in `DOCX_Files` & `PDF_Files`
- 🔧 **Runs via a simple command**

---

## 📦 Requirements
- **Windows OS**
- **[Pandoc](https://pandoc.org/installing.html)**
- **Microsoft Word (Installed)**
- **Python 3.x** (with `pywin32`)

---

## ⏳ Installation & Usage

### **1⃣ Run via PowerShell (One-Line Command)**
To **download & run** directly (In the dir where .md files are present):
```powershell
irm "https://raw.githubusercontent.com/MuhammadSohaibHassan/mdToPdf/refs/heads/main/mdToPdf.ps1" | iex
```
### **2⃣ Manual Usage**
1. **Download** `mdToPdf.ps1`
2. Open PowerShell in the directory where .md files are present.
3. Run:
   ```powershell
   .\mdToPdf.ps1
   ```

---

## 📚 How It Works
1. **Finds all `.md` files** in the directory.
2. Uses **Pandoc** to convert each file to `.docx`.
3. Uses **Microsoft Word** to convert `.docx` to `.pdf`.
4. Saves files in:
   - 📂 `DOCX_Files/`
   - 📂 `PDF_Files/`

---

## 🛠 Troubleshooting

### 🔹 **Pandoc Not Found?**
Ensure **Pandoc is installed & added to the system `PATH`.**
Test by running:
```powershell
pandoc --version
```

### 🔹 **Python Module `pywin32` Missing?**
The script **automatically installs** missing dependencies.

If manual installation is needed:
```powershell
python -m pip install pywin32
```

### 🔹 **Microsoft Word Issues?**
- Ensure **Word is installed** and can open `.docx` files.
- Check that **no conflicting Word processes** are running.

---

## 💚 Contributions
Pull requests & suggestions are welcome! 🎉🚀

---

## 🐟 License
MIT License - Free for use and modification.

---

## 💡 Author
Developed by **MuhammadSohaibHassan**.

