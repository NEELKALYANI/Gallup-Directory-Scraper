# Gallup-Directory-Scraper

A powerful and efficient Python scraper for extracting consultant profile details from Gallup profiles. This tool automates data collection, processes multiple URLs, and exports results into a structured Excel file. Ideal for researchers, recruiters, or data analysts looking to gather consultant information efficiently.
```
🚀 Features:

✅ Flexible Input: Reads URLs directly from an Excel file (links.xlsx).
✅ Comprehensive Data Extraction: Captures multiple key details including:
Name
Email
Country
Expertise
Availability
Method
Language
About Me

✅ Robust Error Handling: Handles broken links, missing data, and unexpected page structures without breaking the workflow.
✅ Smart Delays: Introduces randomized delays between requests to prevent server blocking.
✅ Debug Mode: Captures HTML structures for challenging profiles to improve scraping accuracy.
✅ Detailed Reporting: Generates an organized Excel file with completion rates for each field to ensure transparency.
```
📋 Prerequisites:

Ensure you have the following dependencies installed:

```
pip install requests beautifulsoup4 pandas lxml openpyxl
```

📄 Usage:

1. Place your list of URLs in an Excel file (e.g., links.xlsx) with a column named "URL".

2. Run the script using:

```
python Main_Gallup_scraper.py
```

3. The extracted data will be saved as Scraped2.xlsx.

🧩 How It Works:

The script follows these steps:

1. Read URLs from Excel: URLs are validated to ensure they start with http or https.
2. Data Extraction: Multiple strategies are applied to improve data collection:

    BeautifulSoup for HTML parsing.

    lxml with XPath for structured content retrieval.

    Keyword-based pattern searches for complex fields like expertise and language.

3. Error Handling: If a URL fails to load or data is incomplete, the script logs warnings and continues processing.
4. Debug Mode (Optional): Saves challenging HTML pages for later analysis if data is missing.
5. Excel Output: Results are compiled into a structured Excel file with completion rates for each field.


💡 Future Improvements:

1. Implement parallel processing for faster scraping.
2. Enhance email detection with improved regex patterns.
3. Add more flexible output formats (CSV, JSON).

🏆 Contributing:

Contributions are welcome! Feel free to submit issues, suggestions, or pull requests to improve the scraper.


