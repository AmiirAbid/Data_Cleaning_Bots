# Real Estate Data Cleaner

## Description
Real Estate Data Cleaner is a collection of Python scripts designed to simplify and automate the process of cleaning and transforming raw real estate listings data stored in Excel files. The project aims to provide a quick and efficient way to prepare data for further analysis or visualization, making it a valuable tool for real estate professionals, data analysts, and enthusiasts. 

With user-friendly features and customizable options, these scripts meticulously scan through each listing, identifying and correcting common inconsistencies, errors, and formatting issues in the data.

Whether you are dealing with missing values, or irregular text entries, Real Estate Data Cleaner offers a set of robust and reliable solutions to cleanse and enhance your data, saving time and effort while improving the accuracy and reliability of your real estate datasets.

## Key Features
- Combining data from multiple files into one file
- Automated cleaning of missing or inconsistent data entries.
- Transformation and normalization of text data to ensure uniformity and consistency.
- Customizable cleaning rules and guidelines to suit different data requirements and specifications.
- Calculations of average prices and average prices per m² in each city and state
- Calculations of each type of listings in each city and state

## Usage
Designed with simplicity and efficiency in mind, the Real Estate Data Cleaner requires minimal setup and configuration. Users can quickly get started by running the provided Python scripts on their target Excel files, and letting the tool handle the rest. The scripts are related so they have to be used in their order. Scripts with the same order are not related, each one does a seperate job.

With a set of configurable parameters, users can identify their input and output file paths, specify the structure of their data and define their own cleaning criteria, addressing the specific keywords of the data they want to extract from their real estate listings data. This level of customization ensures that the tool is flexible and adaptable, providing valuable assistance regardless of the structure of your data.

## Installation
To utilize the Real Estate Data Cleaner, clone this repository, navigate to the project directory, and install the necessary Python dependencies:

- **openpyxl library:** reads and writes Excel 2010 xlsx/xlsm/xltx/xltm files
`pip install openpyxl`