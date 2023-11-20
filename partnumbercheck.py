import pandas as pd
import os
import pyodbc

class DatabaseConnection:
    def __init__(self):
        self.server = 'EV1VSQL08.US.Bosch.com,56482'
        self.database = 'DB_PART_RELEASE_SQL'
        self.conn = self.connect()

    def connect(self):
        return pyodbc.connect(f'DRIVER={{ODBC Driver 13 for SQL Server}};SERVER={self.server};DATABASE={self.database};Trusted_Connection=yes')

    def close(self):
        if self.conn:
            self.conn.close()

    def execute_query(self, query):
        cursor = self.conn.cursor()
        cursor.execute(query)
        return cursor.fetchall()


def is_new_fg_packaged_in_atl(part_number_b):
    connection = DatabaseConnection()
    result = None

    try:
        query = f"SELECT [BOM Item Text] FROM [DB_Part_Release_SQL].[dbo].[Download_BOM] WHERE [MATERIAL] = '{part_number_b}'"

        rows = connection.execute_query(query)  # Fetch all rows at once

        # Debug: Print the count of rows fetched
        print(f"Retrieved {len(rows)} rows from the database")

        relabel_needed = False
        found_part = False

        for index, row in enumerate(rows, start=1):
            print(f"Row {index}: {row}")  # Debug: Print the entire row
            found_part = True
            bom_item_text = row[0]

            # Check if the value is None or blank, and if so, skip to the next iteration
            if not bom_item_text:
                print("Encountered empty or null BOM Item Text. Skipping this row.")
                continue

            print(f"Checking BOM Item Text: {bom_item_text}")  # Print every BOM Item Text for checking

            if bom_item_text in ["BIRE LABEL", "BIRELABEL"]:
                relabel_needed = True
                break

        # Decision based on the BOM items found and relabeling requirements
        if found_part:
            if relabel_needed:
                result = "Relabel only"
            else:
                result = "Yes"

    except pyodbc.Error as e:
        print(f"Error: {str(e)}")
    finally:
        connection.close()

    return result if result is not None else "No data"



def fetch_umrez(part_number_a, part_number_b):
    connection = DatabaseConnection()

    try:
        # Initialize variables
        umrez_a = umrez_b = None
        gro_found_a = gro_found_b = False

        # Fetch all UMREZ and MEINH for part_number_a
        query_a = f"SELECT UMREZ, MEINH FROM [DB_Part_Release_SQL].[dbo].[Download_MARM] WHERE MATNR = '{part_number_a}'"
        results_a = connection.execute_query(query_a)

        # Loop through all results for part_number_a to find 'GRO'
        for row in results_a:
            if row[1].strip() == 'GRO':
                umrez_a = row[0]
                gro_found_a = True
                break

        # Fetch all UMREZ and MEINH for part_number_b
        query_b = f"SELECT UMREZ, MEINH FROM [DB_Part_Release_SQL].[dbo].[Download_MARM] WHERE MATNR = '{part_number_b}'"
        results_b = connection.execute_query(query_b)

        # Loop through all results for part_number_b to find 'GRO'
        for row in results_b:
            if row[1].strip() == 'GRO':
                umrez_b = row[0]
                gro_found_b = True
                break

        # Handle cases where 'GRO' is not found
        if not gro_found_a:
            return "No value found (A)"
        if not gro_found_b:
            return "No value found (B)"
        umrez_a = int(umrez_a) if umrez_a is not None else None
        umrez_b = int(umrez_b) if umrez_b is not None else None

        # Continue with UMREZ logic using the 'GRO' values
        if umrez_a == 1:
            print(umrez_a, umrez_b)
            return "No, Check with PM" if umrez_b != 1 else "Yes"
        elif umrez_a > umrez_b:
            print(umrez_a, umrez_b)
            return "Yes" if umrez_b == 1 else "No, Check with PM"

        else:
            print(umrez_a, umrez_b)
            return "Yes" if umrez_a == umrez_b else "No, Check with PM"


    finally:
        connection.close()


'''
def is_supersession_between_old_and_new(part_number_a, part_number_b):
    connection = DatabaseConnection()

    try:
        print("Establishing database connection...")
        connection.connect()  # Connect to the database
        print("Connection established successfully.")

        query = f"""
        SELECT [Reason Code]
        FROM [DB_PART_RELEASE_SQL].[dbo].[Download_YMTK326_Ext]
        WHERE LEFT([13 Digit Predecessor], 10) = '{part_number_a[:10]}'
        AND LEFT([13 Digit Successor], 10) = '{part_number_b[:10]}'
        """

        print(f"Executing query: {query}")
        rows = connection.execute_query(query)  # Fetch all rows at once

        for row in rows:
            print("Checking Row:", row)
            if row[0] == '03':
                return "Yes"

        # If none of the rows had a 'Reason Code' of 1
        return "No, check with PM"

    except pyodbc.Error as e:
        print(f"Error: {str(e)}")
        return f"Error: {str(e)}"
    finally:
        print("Closing database connection...")
        connection.close()
        print("Connection closed successfully.")
'''

def is_coo_same(pkg_idx_a, pkg_idx_b):
    excel_file_path = os.path.join(os.getcwd(), 'CoO.xlsx')

    if os.path.exists(excel_file_path):
        try:
            df = pd.read_excel(excel_file_path, engine='openpyxl')
            country_a = df[df['Packaging Index'] == pkg_idx_a]['Country'].values[0]
            country_b = df[df['Packaging Index'] == pkg_idx_b]['Country'].values[0]
            print(country_a)
            print(country_b)
            return country_a == country_b
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            return False
    else:
        print("Excel file 'CoO.xlsx' not found in the local directory")
        return False
    
def get_part_number_results(part_number_a, part_number_b):
    results = {}

    # Check if the 10-digit part number is the same.
    results['ten_digit_same'] = "Yes" if part_number_a[:10] == part_number_b[:10] else "No"

    # Check if the new F/G is packaged in AtL.
    results['atl_result'] = is_new_fg_packaged_in_atl(part_number_b)

    # Check if there is a supersession between the old and new part numbers.
    #results['supersession_result'] = is_supersession_between_old_and_new(part_number_a, part_number_b)

    # Check if the PAK is the same.
    results['pak_same'] = fetch_umrez(part_number_a, part_number_b)

    # Check if the CoO is the same.
    pkg_idx_a = part_number_a[10:]
    pkg_idx_b = part_number_b[10:]
    results['coo_same'] = "Yes" if is_coo_same(pkg_idx_a, pkg_idx_b) else "No, Check with PM"

    # Determine the overall kit-to-kit potential.
    if results['atl_result'] == "Relabel only" and results['ten_digit_same'] == "Yes" and results['coo_same'] == "Yes" and results['pak_same'] == "Yes":
        results['kit_to_kit_potential'] = "YES"
    elif results['atl_result'] == "Yes":
        results['kit_to_kit_potential'] = "YES" if all(criterion == "Yes" for criterion in [results['ten_digit_same'], results['coo_same'], results['pak_same']]) else "NO"
    else:
        results['kit_to_kit_potential'] = "NO"

    return results