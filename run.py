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
            
            # Ensure formatting (redundant safety)
            if "PRODUCT_TEMP" in df.columns:
                df["PRODUCT_TEMP"] = df["PRODUCT_TEMP"].round(2)
            if "GSV" in df.columns:
                df["GSV"] = df["GSV"].astype(int)
                
            return df
            
        except pyodbc.Error as e:
            raise RuntimeError(f"Query failed: {str(e)}")

    def save_to_csv(self, table_name, output_path, columns=None):
        """
        Save table data to CSV with formatting
        :param table_name: Table to export
        :param output_path: Output CSV path
        :param columns: Optional list of columns to export
        """
        df = self.read_table_data(table_name, columns)
        df.to_csv(output_path, index=False)
        print(f"Data saved to {output_path}")




def combine_mdb_files_to_single_csv(root_folder, output_file):
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
    
    # Save to single CSV
    combined_df.to_csv(output_file, index=False)
    print(f"\nSuccess! Combined data from {processed_files} files into {output_file}")
    print(f"Total records: {len(combined_df)}")


if __name__ == "__main__":
    # Configure these paths as needed:
    search_folder = r"\"
    output_csv = "Combined Tank records.csv"
    
    # Process all files and combine into one CSV
    combine_mdb_files_to_single_csv(search_folder, output_csv)



