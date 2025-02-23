import pandas as pd
import tkinter as tk
from tkinter import simpledialog, filedialog
from fuzzywuzzy import process
import re
from datetime import datetime

# Open file selection window
root = tk.Tk()
root.withdraw()  # Hide the root window
file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])

if not file_path:
    print("No file selected. Exiting.")
    exit()

# Load Excel sheet
sheet_name = 'Sheet1'  # Change if needed
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Standardize column names (strip spaces and fix case)
df.columns = df.columns.str.strip()

# Extract column names for fuzzy matching
df_columns = df.columns.tolist()

# Fuzzy matching helper function
def fuzzy_match(query, choices, threshold=80):  # Lower threshold for better matching
    match, score = process.extractOne(query, choices)
    return match if score >= threshold else None

# Function to convert natural language queries into Pandas queries
def translate_to_query(nl_query, df):
    nl_query = nl_query.lower()
    nl_query = re.sub(r'[^\w\s+-]', '', nl_query)  # Remove punctuation
    df_columns = df.columns.tolist()  # Refresh column names
    conditions = []

    # Handling sorting queries (e.g., "sorted by age in ascending order")
    match = re.search(r"sorted by (\w+) in (ascending|descending) order", nl_query)
    if match:
        column = fuzzy_match(match.group(1).strip().title(), df_columns)
        order = match.group(2).strip().lower()
        if column:
            return f"df.sort_values(by='{column}', ascending={order == 'ascending'})"

    # Handling max queries (e.g., "highest total marks")
    match = re.search(r"highest (.+)", nl_query)
    if match:
        column = fuzzy_match(match.group(1).strip().title(), df_columns)
        if column:
            return f"df[df['{column}'] == df['{column}'].max()]"

    # Handling min queries (e.g., "lowest IA1 marks")
    match = re.search(r"lowest (.+)", nl_query)
    if match:
        column = fuzzy_match(match.group(1).strip().title(), df_columns)
        if column:
            return f"df[df['{column}'] == df['{column}'].min()]"

    # Handling conditions like "older than 19"
    match = re.search(r"(?:older|greater) than (\d+)", nl_query)
    if match:
        conditions.append(f"(df['Age'] > {match.group(1).strip()})")

    # Handling class queries (e.g., "students from class B")
    match = re.search(r"class (\w+)", nl_query)
    if match:
        class_name = match.group(1).strip().upper()
        conditions.append(f"(df['Class'] == '{class_name}')")

    # Handling admission year queries (e.g., "admission year 2023")
    match = re.search(r"admission year (\d{4})", nl_query)
    if match:
        conditions.append(f"(df['Admission Year'] == {match.group(1).strip()})")

    # Handling marks greater than (e.g., "marks more than 40")
    match = re.search(r"(?:total marks|marks) more than (\d+)", nl_query)
    if match:
        conditions.append(f"(df['Total Marks'] > {match.group(1).strip()})")

    # Handling sports marks condition (e.g., "sports marks more than 5")
    match = re.search(r"sports marks more than (\d+)", nl_query)
    if match:
        conditions.append(f"(df['Sports Marks'] > {match.group(1).strip()})")

    # Handling blood group queries (case-insensitive, e.g., "blood group A+")
    match = re.search(r"blood group ([A|B|AB|O][+-])", nl_query, re.IGNORECASE)
    if match:
        blood_group = match.group(1).strip().upper()
        conditions.append(f"(df['Blood Group'].str.upper() == '{blood_group}')")

    # Handling activity grade queries (e.g., "got an A+ in Activities Grade")
    match = re.search(r"got an ['\"]?(A\+?)['\"]? in activities grade", nl_query)
    if match:
        conditions.append(f"(df['Activities Grade'] == '{match.group(1).strip()}')")

    # Handling queries like "List all students along with phone numbers"
    match = re.search(r"list all students along with (.+)", nl_query)
    if match:
        columns = match.group(1).strip().split(" and ")  # Extract requested columns
        column_names = [fuzzy_match(col.title(), df_columns) for col in columns]  # Match column names
        column_names = [col for col in column_names if col]  # Remove None matches

        # Always include "Name" if listing phone numbers
        if "Phone Number" in column_names:
            name_col = fuzzy_match("Name", df_columns)
            if name_col and name_col not in column_names:
                column_names.insert(0, name_col)  # Ensure "Name" is first

        if column_names:
            return f"df[['{column_names[0]}', {', '.join(map(repr, column_names[1:]))}]]"
        else:
            return "df"

    # Combine conditions using '&' (AND operation in pandas)
    if conditions:
        query_string = " & ".join(conditions)
        return f"df[{query_string}]"

    return "print('Query not understood')"

# Function to execute Pandas query safely
def execute_query(query, df):
    try:
        result = eval(query)

        if isinstance(result, pd.DataFrame):
            missing_columns = [col for col in result.columns if col not in df.columns]
            if missing_columns:
                print(f"Error: Columns not found: {missing_columns}")
                return

        if not result.empty:
            print(result)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f'query_result_{timestamp}.xlsx'
            result.to_excel(filename, index=False)
            print(f"Results saved to '{filename}'")
        else:
            print("No valid results found. Nothing saved.")

    except Exception as e:
        print(f"Error executing query: {e}")

# Main loop for user queries
if __name__ == "__main__":
    while True:
        nl_query = simpledialog.askstring("Input", "Enter your query (or 'exit' to quit):")
        if nl_query is None or nl_query.lower() == 'exit':
            break

        pandas_query = translate_to_query(nl_query, df)
        print(f"Generated Pandas query: {pandas_query}")
        execute_query(pandas_query, df)
