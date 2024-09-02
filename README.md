# IntroPythonForDataAnalytics
This repository hosts the materials for an introductory course to Python. The course covers essential topics in Data Analytics, including Data Types, Data Structures, Iterations, and working with OpenPyXL. 

Unemployment Rate Data Analysis

Overview

This project fetches and analyzes unemployment rate data from the Bureau of Labor Statistics (BLS) and the U.S. Census Bureau, calculates various unemployment rates, performs correlation analysis, and generates an Excel report with charts for data visualization.

Prerequisites

Ensure you have the following libraries installed:
- requests
- openpyxl
- dotenv
- scipy

You can install the required packages using pip:

bash
pip install requests openpyxl python-dotenv scipy

Environment Variables

Create a `.env` file in the project directory with the following content:

plaintext
BLS_API_KEY=your_bls_api_key_here
CENSUS_API_KEY=your_census_api_key_here

Replace `your_bls_api_key_here` and `your_census_api_key_here` with your actual API keys.

Code Description

Functions

- fetch_bls_data(): Retrieves data from the BLS API for national and veteran unemployment rates.
- fetch_census_data(): Retrieves data from the Census API.
- calculate_veteran_unemployment_rate(census_data): Calculates the veteran unemployment rate.
- calculate_civilian_unemployment_rate(census_data): Calculates the civilian unemployment rate.
- calculate_correlation_analysis(national_rates, veteran_rates): Computes the Pearson correlation between national and veteran unemployment rates.
- add_charts(workbook): Adds various charts to the Excel workbook for data visualization.
- write_to_excel(bls_data, census_data, veteran_unemployment_rate, civilian_unemployment_rate, correlation, filename): Writes the fetched and calculated data to an Excel file and adds charts.
- main(): The main function that orchestrates data fetching, calculation, and writing to Excel.

3. The script will fetch data, perform calculations, and save the results to an Excel file. 

Excel Workbook

The generated Excel workbook will contain the following sheets:
- National Unemployment Rate: Contains BLS national unemployment rate data.
- Veteran Unemployment Rate: Contains BLS veteran unemployment rate data.
- Census Veteran Data (2021): Displays Census data related to veteran demographics.
- Trend Analysis: Compares national and veteran unemployment rates with line and scatter charts.
- Unemployment Rate Summary: Summarizes civilian and veteran unemployment rates.



