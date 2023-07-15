---------------------------------------------
|                                           |
|                  README                   |
|                                           |
---------------------------------------------

Product List Generator
======================

This script generates a product list in a Microsoft Word document based on data from an Excel or CSV file. It uses the Tkinter library for the user interface, the pandas library for reading data from files, and the docx library for creating and manipulating Word documents.

Installation:
--------------

1. Clone or download the repository to your local machine.
2. Make sure you have Python 3 installed.
3. Install the required dependencies by running the following command:

pip install tkinter pandas docx

Usage:
-------

1. Run the script by executing the following command:

python product_list_generator.py

2. The script will open a GUI window.
3. Click on the "Upload File" button to select an Excel or CSV file containing the product data.
4. Once a file is selected, the script will read the data and generate a product list document.
5. The output Word document will be saved in the same directory as the script with the name "output.docx".

Customization:
--------------

You can customize the template used for the product list by modifying the `TEMPLATE_FILE` constant in the script. By default, it is set to `'template_productlist.docx'`. You can replace it with your own template file, but make sure it has the necessary placeholders for the product data.

Dependencies:
--------------

- Python 3
- tkinter
- pandas
- docx

You can install the required dependencies using the command mentioned in the Installation section.

Contributing:
-------------

Contributions are welcome! If you have any suggestions, bug reports, or feature requests, please open an issue or submit a pull request.

License:
--------

This project is licensed under the MIT License. Feel free to use and modify the code according to your needs.
