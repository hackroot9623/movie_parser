# Folder List Script

This script generates an XLS file containing a list of folders in a specified directory. The folder names are parsed to extract the year, movie name, and director information. The extracted data is then saved in an XLS file with three columns: "Year," "Name of the Movie," and "Director."

## Requirements

- Python 3.x
- xlwt library

## Usage

1. Clone the repository or download the script directly.

2. Install the required library using the following command:

   ```shell
   pip install xlwt
   
3. Run the script using the following command:
   ```shell
   python folder_list.py
4. Enter the directory path when prompted.
5. The script will recursively process the folders in the specified directory and generate an XLS file named "folder_list.xls" containing the extracted data.

## Folder Naming Convention
   ```scss
   (year)name of the movie[Director and lead actor]

- The year should be enclosed in parentheses.
