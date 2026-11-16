# Warehouse Cycle Count Excel Auto-Combiner

A simple Excel VBA tool that combines multiple warehouse cycle count Excel files into one workbook with a single click.

# Noted
- Please edit the **Folder path** in VBA code
- Edit the **Create the 3 target tabs**
  
# What It Does
- Reads all .xlsx files inside the source path folder.
- Combines files into **3 separate tabs** in one new Excel file:
  1. Range_of_Locations_Detail
  2. Pallet_Detail
  3. Scanned_Detail
- Keeps the original headers and appends extra files automatically without duplicates.

# How to Use
1. Put all **.xlsx files** into source path.
2. Create **xlsm file** and put the shape act as UI
3. Right click and 'Assign Macro' and import the **ExampleModule1.bas**
4. It will pop up the new **xlsx Files**.
5. Save the new combined Excel file.

# Files Included
- **Example.xlsm** = Main Excel tool with the one-click button.
- **ExampleModule1.bas** = VBA macro source code.

# Tools Used
- Microsoft Excel
- Excel VBA (Visual Basic for Applications)
