# Py Business Search

Py Business Search is a Python application designed to streamline the process of finding and cataloging business information based on specific keywords and locations. Using Selenium for web automation and pandas for data handling, this tool is an invaluable asset for businesses and researchers looking to compile comprehensive lists of companies in particular niches or areas.

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

You will also need to have ChromeDriver compatible with your version of Chrome. This is necessary for Selenium to interact with the browser.

## Installation

1. Download the newest release.
2. Extract the installation folder.
3. Run the `Setup.EXE` installer.

## Usage

1. Open the application.
2. Enter the search keyword and location in the provided fields.
3. Click the "Set Excel File Path" button to set the path to your Excel file where the data will be stored.
4. Click the "Search" button to start the search process.
5. The results will be automatically saved in the specified Excel file, updating it with new entries and avoiding duplicates.

## Disclaimer

- **Ethical and Legal Considerations:** Py Business Search is a tool that automates the process of collecting publicly available business information. Users should ensure they comply with relevant laws and website terms of service. This software is provided for educational and research purposes only. It is the user's responsibility to ensure that their use of the tool adheres to all applicable laws and ethical guidelines.
- **Legal Responsibility:** The creators of Py Business Search are not responsible for any legal consequences that arise from the use of this software. Users assume full responsibility for their use of this tool and any legal implications that may result.

## Contributions

Contributions to Py Business Search are welcome! If you have suggestions for improvements or encounter any issues, please feel free to open an issue or submit a pull request on GitHub.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contact

For any inquiries or further assistance, please contact the repository owner.

Happy searching! 🚀
