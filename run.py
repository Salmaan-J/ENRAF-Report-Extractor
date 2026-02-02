import pyodbc
from pathlib import Path
import pandas as pd
import os
from warnings import filterwarnings

class MDBReader:
    def __init__(self, mdb_path):
        """
        Initialize the MDB reader with database path
        param mdb_path: Path to .mdb/.accdb file
        Noted that this is like the normal process of initialising the class from C++ and other where you need to referrence itself.
        """
        self.mdb_path = mdb_path
        self.conn = None
        self.cursor = None
        filterwarnings('ignore', category=UserWarning)  # Suppress pandas warning

    def __enter__(self):
        """
        Looks like the connect to DB
        
        """
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """DB  manager exit"""
        self.close_connection()

    def connect(self):
        """Establish database connection to DB from path."""
        if not os.path.exists(self.mdb_path):
            raise FileNotFoundError(f"MDB file not found at: {self.mdb_path}")

        try:
            print("Available ODBC Drivers:", pyodbc.drivers())
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                r'DBQ=' + self.mdb_path + ';'
            )
            self.conn = pyodbc.connect(conn_str, timeout=30)
            self.cursor = self.conn.cursor()
        except pyodbc.Error as e:
            raise ConnectionError(f"Failed to connect to database: {str(e)}")

    def close_connection(self):
        """Close database connection if active"""
        if self.conn:
            self.conn.close()
            self.conn = None
            self.cursor = None

    def get_tables(self):
        """Return list of tables in database"""
        if not self.conn:
            raise ConnectionError("Database not connected")
        return [t.table_name for t in self.cursor.tables(tableType='TABLE')]

    def read_table_data(self, table_name, columns=None):
        """
        Read data from specified table with formatting
        :param table_name: Name of table to read
        :param columns: List of columns to select (None for all)
        :return: Formatted DataFrame
        """
        if not self.conn:
            raise ConnectionError("Database not connected")

        # Validate table exists
        tables = self.get_tables()
        if table_name not in tables:
            raise ValueError(f"Table '{table_name}' not found. Available tables: {tables}")

        # Build query with MS Access compatible formatting
        column_list =  ", ".join(col for col in columns)
       # print(column_list)
        query = f"SELECT {column_list} FROM [{table_name}]"
        
        try:
            # Read into DataFrame with proper typing
            df = pd.read_sql(query, self.conn)
            
         
            if "PRODUCT_TEMP" in df.columns:
                df["PRODUCT_TEMP"] = df["PRODUCT_TEMP"].round(2)
            if "GSV" in df.columns:
                df["GSV"] = df["GSV"].astype(int)
            if "BACKGROUND_TIME_STAMP" in df.columns:
                df["BACKGROUND_TIME_STAMP"] = pd.to_datetime(df["BACKGROUND_TIME_STAMP"])
                # Round down to minute precision
                df["BACKGROUND_TIME_STAMP"] = df["BACKGROUND_TIME_STAMP"].dt.round('2min')
            return df
            
        except pyodbc.Error as e:
            raise RuntimeError(f"Query failed: {str(e)}")

    def save_to_csv(self, columns):
        """
        Prompt user to select fuel grade(s) and generate reports
        """

        print("\nSelect grade(s) to extract:")
        print("1 - D50")
        print("2 - ULP")
        print("3 - KERO")
        print("4 - JET A1")
        print("5 - All Grades")
        print("You can select multiple grades (comma separated), e.g. 1,2")

        choice = input("Enter your choice: ").strip()

        grade_map = {
            "1": "DIESEL",
            "2": "ULP",
            "3": "KERO",
            "4": "JET A1",
            "5": "All"
        }

        selected_grades = []
        if grade_map.get(choice) == "All":
            selected_grades = ["DIESEL", "ULP", "KERO", "JET A1"]
        else:
            for c in choice.split(","):
                c = c.strip()
                print(c)
                if c in grade_map:
                    selected_grades.append(grade_map[c])

        if not selected_grades:
            print("No valid grade selected. Exiting.")
            return

        for grade in selected_grades:
            print(f"Extracting {grade} report...")
            self.Grade_Extract(columns, grade)


    def Grade_Extract(self, combined_df, grade_name):
        """
        Extract ULP data from combined DataFrame
        
    """

        # Ensure correct ordering
        
        ulp_df = combined_df[combined_df['PRODUCT_NAME'].str.contains(grade_name, case=False, na=False)].copy()
        ulp_df.sort_values(by=["BACKGROUND_TIME_STAMP", "TANK_NAME"], inplace=True)
        tank_cols = [
            "PRODUCT_NAME",
            "PRODUCT_TEMP",
            "CORRECTION_FACTOR",
            "GSV",
            "PRODUCT_LEVEL",
        ]

        # Fixed tank order (critical)
        tank_order = sorted(ulp_df["TANK_NAME"].unique())

        rows = []

        for ts in sorted(ulp_df["BACKGROUND_TIME_STAMP"].unique()):
            row = [ts]

            for tank in tank_order:
                rec = ulp_df[
                    (ulp_df["BACKGROUND_TIME_STAMP"] == ts) &
                    (ulp_df["TANK_NAME"] == tank)
                ]

                if not rec.empty:
                    row.append(tank)  # keep tank name aligned
                    row.extend(rec.iloc[0][tank_cols].tolist())
                else:
                    # missing tank at this timestamp → placeholders
                    row.append(tank)
                    row.extend([None] * len(tank_cols))

            rows.append(row)

        # Build columns
        columns = ["BACKGROUND_TIME_STAMP"]
        for _ in tank_order:
            columns.extend(
                ["TANK_NAME"] + tank_cols
            )

        wide_df = pd.DataFrame(rows, columns=columns)
        wide_df.to_csv(f"{grade_name} Grade_report.csv", index=False)
        print("Data saved to Grade_report.csv")
            
        
def combine_mdb_files_to_single_csv(root_folder,output_file):
    """
    Combines TankRecords from all .mdb files into one CSV file
    
    Args:
        root_folder (str): Folder to search for .mdb files
        output_file (str): Path for combined output CSV file
    """
    columns = [
        "BACKGROUND_TIME_STAMP",
        "TANK_NAME",
        "PRODUCT_NAME",
        "PRODUCT_TEMP",
        "CORRECTION_FACTOR",
        "GSV",
        "PRODUCT_LEVEL"
    ]
    
    all_data = []
    processed_files = 0
    
    # Find all .mdb files recursively
    for root, _, files in os.walk(root_folder):
        for file in files:
            if file.lower().endswith('.mdb'):
                mdb_path = os.path.join(root, file)
                try:
                    with MDBReader(mdb_path) as reader:
                        print(f"Processing: {file}")
                        df = reader.read_table_data("TankRecords", columns)
                        all_data.append(df)
                        processed_files += 1
                except Exception as e:
                    print(f"Error processing {file}: {str(e)}")
                    continue
    
    if not all_data:
        print("No valid .mdb files found with TankRecords table")
        return
    
    # Combine all DataFrames
    combined_df = pd.concat(all_data, ignore_index=True)
    print(combined_df)
    print(f"Total records: {len(combined_df)}")
    reader.save_to_csv(combined_df)
  
    


if __name__ == "__main__":
    # Configure these paths as needed:
    search_folder = r"C:\Users\Salmaan\Documents\ENRAF Report Extractor\ENRAF REPORTS"
    
    output_csv = "Combined Tank records.csv"
    
    # Process all files and combine into one CSV
    combine_mdb_files_to_single_csv(search_folder, output_csv)

