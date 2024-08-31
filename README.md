# Automate-Report 

## Overview

The Automate Report Tool is a Python script that automates the process of downloading a report, extracting and processing data, and uploading the processed data to a Google Sheet. This tool helps in maintaining updated reports and data analysis without manual intervention.

## Features

1. **Report Download:** Automatically downloads a report from a specified URL.
2. **Data Extraction:** Extracts specific data points from the downloaded report.
3. **Data Processing:** Processes the extracted data, such as calculating totals and identifying top products.
4. **Google Sheets Integration:** Uploads the processed data to a Google Sheet.
5. **Automation:** The script can be scheduled to run at regular intervals.

## Prerequisites

- Python 3.12 or later
- Required Python libraries:
  - `pandas`
  - `gspread`
  - `oauth2client`
  - `xlrd`

You can install the necessary libraries using pip:

```bash
pip install pandas gspread oauth2client xlrd
```

## Setup

### 1. Google Sheets API Access

1. **Create a Google Cloud Project:**
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a new project or select an existing one.

2. **Enable Google Sheets and Drive APIs:**
   - Go to the [Google Sheets API page](https://console.developers.google.com/apis/library/sheets.googleapis.com).
   - Click "Enable".
   - Go to the [Google Drive API page](https://console.developers.google.com/apis/library/drive.googleapis.com).
   - Click "Enable".

3. **Create Credentials:**
   - Go to the [Credentials page](https://console.developers.google.com/apis/credentials).
   - Click "Create Credentials" and choose "Service Account".
   - Follow the prompts and create a service account key in JSON format.
   - Download the JSON key file.

4. **Share Google Sheet:**
   - Open the Google Sheet you want to update.
   - Share it with the email address from your service account (found in the JSON key file).

### 2. Configure the Script

- Place the downloaded JSON key file in the project directory.
- Update the `file_path` variable in `automate_report.py` with the path to your report file.
- Update the `sheet_name` variable in `update_google_sheet` function with your Google Sheet name.

### 3. Run the Script

To run the script, use:

```bash
python automate_report.py
```

### 4. Schedule Automation

You can schedule the script to run at regular intervals using a task scheduler:

- **On Windows:**
  - Use Task Scheduler to create a new task and set the script to run at your preferred intervals.

- **On macOS/Linux:**
  - Use `cron` to schedule the script. For example, to run the script daily at 2 AM, add the following line to your crontab (`crontab -e`):

    ```bash
    0 2 * * * /path/to/your/venv/bin/python /path/to/automate_report.py
    ```

## Example Output

- **Console Output:**
  - Displays a preview of the data, total sales, and top products.

- **Google Sheet:**
  - The sheet will be updated with total sales and top products.

## Troubleshooting

- If you encounter errors related to missing dependencies, ensure all required libraries are installed.
- For quota errors with Google Sheets API, check your usage limits in the Google Cloud Console.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

If you have suggestions or improvements, please fork the repository and create a pull request.
