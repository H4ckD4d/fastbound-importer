# 🧾 FastBound Importer – ATF A&D Integration Tool

**Created for Knights Arms USA by Chris Cruz – h4ckd4d**

A professional Python utility built to simplify and automate **ATF A&D record management** for **FFLs, gunsmiths, and compliance officers**.

This tool reads your **ATF A&D Record spreadsheet** and automatically converts it into the **FastBound import format**, preserving all required compliance fields and generating detailed mapping and audit reports.

---

## ⚙️ Main Features

- 🔍 **Intelligent Field Mapping**  
  Automatically matches column names from ATF → FastBound (using direct, alias, and fuzzy matching).

- 🧩 **Manual Override Support**  
  Accepts custom mapping via CSV / JSON / YAML.

- 📊 **Professional Excel Output**  
  Includes three sheets:  
  - **FastBoundImport:** Ready for FastBound upload  
  - **Mapping Report:** Source tracking for every field  
  - **Missing & Guidance:** Lists blank fields + how to locate that info (4473, FFL, invoice, etc.)

- ⚙️ **Cross-Platform CLI** – works on **macOS, Linux, and Windows**  
- 🔐 Designed for **ATF record compliance and internal audit traceability**

---

## 🧠 Example Usage

```bash
python fastbound_importer.py \
  --atf "ATF-Firearms-AD-Record.xlsx" --atf-sheet "ATF A&D Record" \
  --fastbound "FastBoundImport Live - By Chris.xlsx" \
  --fastbound-sheet "FastBoundImport Live - By Chris" \
  --out "FastBoundImport_Populado.xlsx" \
  --verbose
