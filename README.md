# Py Business Search

Py Business Search is a Python application designed to streamline the process of finding and cataloging business information based on specific keywords and locations. Using Selenium for web scraping and pandas for data handling, this tool is an invaluable asset for businesses and researchers looking to compile comprehensive lists of companies in particular niches or areas.

## Features

- **Keyword and Location Based Search:** Input keywords (e.g., "Stone Countertops") and a city or ZIP code to search for relevant businesses.
- **Data Compilation into Excel:** Automatically adds business names and addresses to an Excel file.
- **Duplicate Management:** Cross-references new names and addresses with existing entries to avoid duplicates.
- **Address Validation:** Ensures collected addresses are in a valid format.
- **Data Cleansing:** Removes entries with empty addresses from the list.
- **User Interface:** Simple UI built with Tkinter for easy interaction.

## Requirements

To run Py Business Search, you will need the following:

- Python 3.x
- Selenium (`pip install selenium`)
- Pandas (`pip install pandas`)
- Geopy (`pip install geopy`)
- Tkinter (usually comes pre-installed with Python)
- ttkthemes (`pip install ttkthemes`)

You will also need to have Google Chrome This is necessary for Selenium to interact with Google Chrome.

## Installation

1. Clone the repository or download the files.
2. Ensure all the required modules listed above are installed.
3. Run the script with Python.

## Usage

1. Open the application.
2. Enter the search keyword and location in the provided fields.
3. Click the "Set Excel File Path" button to set the path to your Excel file where the data will be stored.
4. Click the "Search" button to start the search process.
5. The results will be automatically saved in the specified Excel file, updating it with new entries and avoiding duplicates.

## Disclaimer

- Py Business Search is a web scraping tool. Web scraping may be subject to legal and ethical considerations depending on the data being scraped and how it is used. Users of Py Business Search should ensure that they are in compliance with any relevant laws and website terms of service.
- **Legal Responsibility:** The creators of Py Business Search are not responsible for any legal consequences that arise from the use of this software. Users assume full responsibility for the use of this tool and any legal implications their actions may incur.

## Contributions

Contributions to Py Business Search are welcome! If you have suggestions for improvements or encounter any issues, please feel free to open an issue or submit a pull request on GitHub.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contact

For any inquiries or further assistance, please contact the repository owner.

Happy searching! 🚀
