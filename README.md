📐 AutoCAD Nesting Automation

Automates nesting drawings from Excel BOM + DXF files.

🚀 How to Use

1. Install requirements

pip install -r requirements.txt

2. Run the program

python nesting.py

Or double-click:

▶️ run_nesting.bat

3. In the application

📊 Select the Excel BOM file

📑 Select the required sheet

📁 Select the folder containing DXF files

📤 Select the output folder

💾 Select DXF / DWG / Both

▶️ Click Generate Nesting

📋 Excel Requirements

The Excel file should contain:

PartNo

CutlistName

DXF_File

Material

Thickness_mm

Qty

The DXF_File name must match the actual DXF filename.

🏗️ DWG Output

DWG generation requires AutoCAD with accoreconsole.exe available.

BOM → DXF Parts → Nesting → DXF/DWG 🏭