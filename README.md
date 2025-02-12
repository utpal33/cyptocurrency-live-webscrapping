<h1>Web Scraping with Live Excel Updates</h1>

<h3>Project Overview</h3>

This project fetches real-time cryptocurrency market data using the CoinGecko API and updates an Excel file dynamically using Python. The data includes cryptocurrency names, symbols, prices, market caps, volumes, and 24-hour price changes.

Technologies Used

Python

Requests (for API calls)

Pandas (for data processing)

Xlwings (for Excel automation)

Time (for periodic updates)

OS (for system operations)

Features

Fetches live cryptocurrency data from the CoinGecko API.

Stores the data in a structured Pandas DataFrame.

Updates an active Excel file in real-time.

Runs continuously with periodic updates.

Installation & Setup

Prerequisites

Ensure you have the following installed:

Python 3.x

Required Python packages (install using pip):

pip install requests pandas xlwings

Running the Script

Open an Excel application and ensure an active workbook is available.

Run the Python script:

python webScrappoingWithLiveExcelUpdates.py

The script will fetch and update cryptocurrency data in the Excel file at regular intervals.

Notes

Ensure Excel is open before running the script.

The script uses a while loop to update data continuously.

Modify the update frequency in the script if needed.


