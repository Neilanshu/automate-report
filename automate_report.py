import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Load and Process the Excel File
file_path = '/Users/nilanshubhandare/Documents/kpi_dashboard/Sales Data.xls'

# Load the Excel file
df = pd.read_excel(file_path)

# Display the first few rows to understand its structure
print("Data Preview:")
print(df.head())

# Example extraction: Calculate total sales and find top products
try:
    total_sales = df['TotalSales'].sum()
except KeyError:
    print("Error: Column 'TotalSales' not found in the data.")
    total_sales = None

try:
    top_products = df.groupby('ProductID')['TotalSales'].sum().sort_values(ascending=False)
except KeyError:
    print("Error: Column 'ProductID' or 'TotalSales' not found in the data.")
    top_products = pd.Series()

print(f"Total Sales: {total_sales}")
print("Top Products:")
print(top_products)

# Upload Data to Google Sheets
def update_google_sheet(sheet_name, data):
    # Define the scope and authorize
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/nilanshubhandare/Downloads/phonic-agility-434200-h8-978b0740112e.json', scope)
    client = gspread.authorize(creds)

    # Open the sheet
    sheet = client.open(sheet_name).sheet1

    # Clear existing data
    sheet.clear()

    # Prepare data for batch update
    rows = [['Total Sales'], [data['Total Sales']]]
    rows.append(['Top Products'])
    rows.append(['ProductID', 'TotalSales'])  # Header for Top Products

    for product, sales in data['Top Products'].items():
        rows.append([product, sales])

    # Use batch_update to reduce the number of write requests
    sheet.batch_update([{
        'range': 'A1',  # Starting cell to write data
        'values': rows
    }])

    print("Google Sheet updated")

# Process and update Google Sheet
processed_data = {
    'Total Sales': total_sales,
    'Top Products': top_products.to_dict()
}

update_google_sheet('Automated Report Data', processed_data)