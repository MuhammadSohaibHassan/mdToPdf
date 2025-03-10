# PowerShell Script: Markdown to PDF Converter
# Author: Your Name (or GitHub Username)
# Usage: Run via `irm "RAW_SCRIPT_URL" | iex`
# Dependencies: Pandoc, Python (with pywin32), Microsoft Word

# -------------------------------
# 📂 Define Output Directories
# -------------------------------
$docxDir = "$PWD\DOCX_Files"
$pdfDir = "$PWD\PDF_Files"

# Ensure output directories exist
New-Item -ItemType Directory -Force -Path $docxDir | Out-Null
New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null

# -------------------------------
# 🐍 Define Python Script for DOCX → PDF
# -------------------------------
$pythonScript = "$PWD\convert_docx_to_pdf.py"

# Check & Create Python Script if Missing
if (!(Test-Path $pythonScript)) {
    @'
import sys
import os
import subprocess

# Install missing dependencies if needed
try:
    import win32com.client
except ImportError:
    print("📦 Installing missing dependency: pywin32...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pywin32"], check=True)
    import win32com.client

def docx_to_pdf(docx_path, pdf_path):
    """ Converts a DOCX file to PDF using Microsoft Word Automation """
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Keep Word hidden
        docx_path = os.path.abspath(docx_path)
        pdf_path = os.path.abspath(pdf_path)

        if not os.path.exists(docx_path):
            print(f"❌ [ERROR] DOCX file not found: {docx_path}")
            return

        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
        print(f"✅ [SUCCESS] {os.path.basename(docx_path)} → {os.path.basename(pdf_path)}")

    except Exception as e:
        print(f"❌ [ERROR] {e}")

if __name__ == "__main__":
    docx_to_pdf(sys.argv[1], sys.argv[2])
'@ | Out-File -Encoding utf8 $pythonScript
}

# -------------------------------
# 🔎 Find Markdown Files
# -------------------------------
$mdFiles = Get-ChildItem -Path . -Filter "*.md"

# Track Progress
$totalFiles = $mdFiles.Count
$count = 0

Write-Host "`n🔄 Starting conversion of $totalFiles Markdown files..." -ForegroundColor Cyan

# -------------------------------
# 🔄 Convert Each Markdown File
# -------------------------------
foreach ($mdFile in $mdFiles) {
    $count++
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($mdFile.Name)
    $docxFile = "$docxDir\$baseName.docx"
    $pdfFile = "$pdfDir\$baseName.pdf"

    Write-Host "📖 Processing [$count/$totalFiles]: $($mdFile.Name)" -ForegroundColor Yellow

    # Convert Markdown to DOCX using Pandoc
    Start-Process -NoNewWindow -Wait -FilePath "pandoc" -ArgumentList "`"$($mdFile.FullName)`" -o `"$docxFile`" -V fontsize=12pt -V margin=1in"

    # Ensure DOCX conversion succeeded
    if (!(Test-Path $docxFile)) {
        Write-Host "❌ [ERROR] Failed to convert $mdFile.Name to DOCX" -ForegroundColor Red
        continue
    }

    # Convert DOCX to PDF using Python
    $pythonOutput = python "$pythonScript" "$docxFile" "$pdfFile" 2>&1

    # Ensure PDF file exists before moving
    if (Test-Path $pdfFile) {
        Move-Item -Path $pdfFile -Destination $pdfDir -Force
        Write-Host "📄 $baseName.md → $baseName.docx → $baseName.pdf ✅" -ForegroundColor Green
    } else {
        Write-Host "❌ [ERROR] Conversion failed for $baseName.docx" -ForegroundColor Red
    }
}

Write-Host "✅ All conversions completed! 🎉`n" -ForegroundColor Cyan
