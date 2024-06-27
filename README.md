# STARMANS Pressure Data to Excel Processor

This script reads pressure data from a microcontroller and populates it into specific cells of an Excel file. Additionally, it creates a folder named with the current date and saves the Excel file in this folder. The Excel file must be empty and present in the same directory as the script each time the script is run.

## Requirements

Ensure you have Python installed on your system. Additionally, install the required libraries:

- `pyserial`
- `openpyxl`

You can install these libraries using pip:

```sh
pip install pyserial openpyxl
```

## Setup Instructions
Download the script_name.py script and save it in a directory of your choice.

Ensure your microcontroller is connected to the computer via USB.
Modify USB Port (if needed): The USB port may vary from computer to computer. If necessary, modify the serial port path in line 59 of script_name.py to match your system.

Run the Script:
Open a terminal or command prompt.
Navigate to the directory where script_name.py and data_template.xlsx are located.
Execute the script using Python with the following command:

```sh
python script_name.py
```
## Excel Output

Once data is successfully saved, the script will create a folder named with the current date and save the Excel file in this folder.
The script will print a message confirming the Excel file has been updated and saved in the date-named folder.

## Notes
The script assumes a specific layout for data placement in the Excel file. Modify the script's cell positions if your Excel layout differs.
