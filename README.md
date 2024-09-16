# excel_to_json_converter
Data converter with .xlsx files to .json

Here is a list of the libraries we have used in the Excel to JSON conversion script:

1. json
tkinter
openpyxl

Explanation of each:
1. json: This is a standard Python library. It is used to work with data in JSON format, allowing encoding and decoding of JSON.
2. tkinter: This is also a standard Python library. It is used to create graphical user interfaces (GUI). In our case, we specifically used it to open a file selection dialog.
3. openpyxl: This is an external library used to read and write Excel 2010 xlsx/xlsm/xltx/xltm files. It is the one we need to install additionally with pip.
   
To make sure that all these libraries are installed and working properly, you can run the following commands:

python -c "import json; print('json OK')"
python -c "import tkinter; print('tkinter OK')"
python -c "import openpyxl; print('openpyxl OK')"

If any of these commands produce an error, you will know which library needs attention. Remember that only openpyxl needs to be installed manually using pip, as the other two are part of the Python standard library.
