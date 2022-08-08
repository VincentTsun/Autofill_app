The program is split into 3 functions: data extraction from Excel files (that use the specified template), adding all purchase orders to the existing booking, and filling extracted data onto the Maersk system.

First, the program will ask the user to input all the purchase order IDs, their relative port code, and their descriptions. It then extracts all Excel sheets that match the purchase order IDs from the specified folder, and imports them as Pandas data frames. After, it extracts all key data by locating keywords, and outputs the compiled data into an Excel file for inspection.

After the inspection, the software is now ready for the next step: adding purchase orders to the booking. Using Selenium, the software opens a Chrome browser, takes the purchase order IDs from user input and add all of them to the booking on Maersk's website.

Finally, the software fills the booking with the extracted data.
