🔍 MDB Reader & Combiner

A Python utility for reading Microsoft Access database files (.mdb/.accdb) and combining TankRecords data from multiple files into a single CSV.
🚀 Quick Start
1. Install Dependencies
bash

pip install pyodbc pandas

2. Install Microsoft Access Database Engine

Download from: Microsoft Official Download

    Choose the version (32-bit or 64-bit) that matches your Python installation

    Run the installer as Administrator

3. Clone and Use python

# Save the code as mdb_reader.py and use:

from mdb_reader import combine_mdb_files_to_single_csv

# Combine all TankRecords from .mdb files in a folder
combine_mdb_files_to_single_csv(
    root_folder="C:/Your/Folder/With/MDB/Files",
    output_file="combined_tank_data.csv"
)

📁 Project Structure
text

mdb_reader.py          # Main utility file
├── MDBReader class    # Database connection handler
└── combine_mdb_files_to_single_csv()  # Batch processor

🛠️ Usage Examples
Example 1: Single File Processing
python

from mdb_reader import MDBReader

# Process a single .mdb file
with MDBReader("database.mdb") as reader:
    # See available tables
    tables = reader.get_tables()
    print(f"Tables: {tables}")
    
    # Read TankRecords table
    data = reader.read_table_data("TankRecords")
    
    # Save to CSV
    reader.save_to_csv("TankRecords", "single_file_output.csv")

Example 2: Batch Process Folder
python

from mdb_reader import combine_mdb_files_to_single_csv

# Process all .mdb files in folder and subfolders
combine_mdb_files_to_single_csv(
    root_folder="C:/TankData/2024",
    output_file="all_tank_records_2024.csv"
)

Example 3: Custom Columns
python

# Read specific columns only
custom_columns = ["TANK_NAME", "PRODUCT_NAME", "PRODUCT_TEMP", "GSV"]

with MDBReader("data.mdb") as reader:
    df = reader.read_table_data("TankRecords", columns=custom_columns)

⚙️ Configuration

Edit the main section at the bottom of the script:
python

if __name__ == "__main__":
    search_folder = r"C:\Your\Data\Folder"    # ← Change this path
    output_csv = "Combined_Tank_Records.csv"  # ← Output filename
    
    combine_mdb_files_to_single_csv(search_folder, output_csv)

📊 Default Data Columns

The tool extracts these columns from TankRecords table:

    BACKGROUND_TIME_STAMP

    TANK_NAME

    PRODUCT_NAME

    PRODUCT_TEMP (rounded to 2 decimal places)

    CORRECTION_FACTOR

    GSV (converted to integers)

    PRODUCT_LEVEL

🐛 Troubleshooting
Common Issues & Solutions

❌ Error: "Microsoft Access Driver not found"
python

# Check available drivers
import pyodbc
print(pyodbc.drivers())

✅ Solution: Install correct Microsoft Access Database Engine

❌ Error: "File not found"
✅ Solution: Ensure the path is correct and file isn't open in Access

❌ Error: Architecture mismatch
✅ Solution: Use 32-bit Python with 32-bit Driver, or 64-bit with 64-bit

❌ Error: Permission denied
✅ Solution: Run as Administrator or close the .mdb file in other programs
📋 Prerequisites Checklist

    Python 3.6+ installed

    pip install pyodbc pandas

    Microsoft Access Database Engine installed

    .mdb files accessible

    Sufficient disk space for output CSV

🔧 Advanced Features
Context Manager

The MDBReader class supports context management for automatic connection handling:
python

with MDBReader("file.mdb") as reader:
    # Connection automatically opened
    data = reader.read_table_data("TankRecords")
    # Connection automatically closed

Custom Output

Modify the combine_mdb_files_to_single_csv function to:

    Change extracted columns

    Add custom data filtering

    Modify output format

📝 License

This project is provided as-is for working with Microsoft Access databases.
🤝 Contributing

Feel free to submit issues and enhancement requests!
