import os
import xlwt

def process_folder(directory):
    folder_name = os.path.basename(directory)
    year = ""
    movie_name = ""
    director = ""

    # Extracting year, movie name, and director information from folder name
    if "(" in folder_name and ")" in folder_name:
        year_start = folder_name.index("(") + 1
        year_end = folder_name.index(")")
        year = folder_name[year_start:year_end]
        movie_info = folder_name[year_end + 1:]
    else:
        movie_info = folder_name
    
    # Extracting movie name and director information from folder name
    if "[" in movie_info and "]" in movie_info:
        movie_start = movie_info.index("[") + 1
        movie_end = movie_info.index("]")
        movie_name = movie_info[movie_start:movie_end]
        director = movie_info[:movie_start - 1] + movie_info[movie_end + 1:]
    else:
        movie_name = movie_info
    
    return year, movie_name, director

def list_folders_recursively(directory):
    folders_data = []

    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        if os.path.isdir(item_path):
            folder_data = process_folder(item_path)
            folders_data.append(folder_data)
            folders_data.extend(list_folders_recursively(item_path))

    return folders_data

# Main script
target_directory = input("Enter the directory path: ")

if not os.path.isdir(target_directory):
    print("Error: Directory '{}' does not exist.".format(target_directory))
    exit(1)

# Collecting folder data recursively
folders_data = list_folders_recursively(target_directory)

# Creating an XLS workbook and adding a worksheet
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet("Folder List")

# Writing column headers
worksheet.write(0, 0, "Year")
worksheet.write(0, 1, "Name of the Movie")
worksheet.write(0, 2, "Director")

# Writing folder data to rows
for i, data in enumerate(folders_data, start=1):
    worksheet.write(i, 0, data[0])
    worksheet.write(i, 1, data[1])
    worksheet.write(i, 2, data[2])

# Saving the XLS file
xls_file_path = os.path.join(target_directory, "folder_list.xls")
workbook.save(xls_file_path)

print("XLS file 'folder_list.xls' has been generated successfully.")
