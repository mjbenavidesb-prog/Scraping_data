# Scraping_data
The goal of this task is to automatically extract the admission exam results from Universidad Nacional Mayor de San Marcos using Python and Selenium, career by career, and consolidate all the information into an Excel file.

# What does the project do?
The purpose of this project is to automate the extraction of admission exam results from the UNMSM website. Using Python, the script identifies all available career links, enters each one, retrieves the applicants' information, and saves everything in one consolidated Excel document.
# How to install the dependencies?
To execute the project correctly, install these libraries first: pip install pandas selenium webdriver-manager
# How to run the script?
Open the terminal in the repository folder and execute: python scraper.py
# What does the output contain? 
The final result is stored in: output/resultados_sanmarcos.xlsx
The Excel file includes the complete consolidated dataset obtained from the admission results pages of the UNMSM website.

