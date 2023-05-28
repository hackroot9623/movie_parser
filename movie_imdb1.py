import os
import re
import xlwt
from imdb import IMDb

def process_folder(directory, ia, movie_cache):
    folder_name = os.path.basename(directory)
    year = ""
    movie_name = ""
    director = ""
    rating = ""
    cast = ""
    storyline = ""

    # Extracting year, movie name, and director information from folder name
    pattern = r"\((\d+)\)\s*(.*?)\s*\[Dir\.\s*(.*?)\]$"
    match = re.match(pattern, folder_name)
    if match:
        year = match.group(1)
        movie_name = match.group(2)
        director = match.group(3)
    else:
        # Rename the folder with the correct formatting
        new_folder_name = f"({year}) {movie_name} [Dir. {director}]"
        new_directory = os.path.join(os.path.dirname(directory), new_folder_name)
        os.rename(directory, new_directory)
        folder_name = new_folder_name
        year = folder_name[1:5]
        movie_name = folder_name[6:-1]
        director = folder_name.split("[Dir. ")[1][:-1]

    # Search for movie details on IMDb
    if movie_name:
        if movie_name in movie_cache:
            # Retrieve movie details from cache
            movie_details = movie_cache[movie_name]
        else:
            # Perform IMDb search
            search_results = ia.search_movie(movie_name)
            if search_results:
                movie_id = search_results[0].movieID
                movie = ia.get_movie(movie_id)
                movie_details = {}

                if 'rating' in movie:
                    movie_details['rating'] = movie['rating']

                if 'cast' in movie:
                    cast_list = movie['cast']
                    movie_details['cast'] = ', '.join([actor['name'] for actor in cast_list])

                if 'plot outline' in movie:
                    movie_details['storyline'] = movie['plot outline']

                # Cache movie details for future use
                movie_cache[movie_name] = movie_details
        # Retrieve movie details from cache
        rating = movie_details.get('rating', '')
        cast = movie_details.get('cast', '')
        storyline = movie_details.get('storyline', '')

    return year, movie_name, director, rating, cast, storyline, folder_name

def list_folders_recursively(directory, ia):
    folders_data = []
    movie_cache = {}

    try:
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            if os.path.isdir(item_path):
                folder_data = process_folder(item_path, ia, movie_cache)
                folders_data.append(folder_data)
                folders_data.extend(list_folders_recursively(item_path, ia))
    except FileNotFoundError:
        print(f"Warning: Directory '{directory}' does not exist.")

    return folders_data

# Main script
target_directory = input("Enter the directory path: ")

if not os.path.isdir(target_directory):
    print(f"Error: Directory '{target_directory}' does not exist.")
    exit(1)

# Create an instance of the IMDb class
ia = IMDb()

# Collect folder data recursively
folders_data = list_folders_recursively(target_directory, ia)

# Create an XLS workbook and add a worksheet
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet("Folder List")

# Write column headers
worksheet.write(0, 0, "Year")
worksheet.write(0, 1, "Name of the Movie")
worksheet.write(0, 2, "Director")
worksheet.write(0, 3, "Rating")
worksheet.write(0, 4, "Cast")
worksheet.write(0, 5, "Storyline")
worksheet.write(0, 6, "Folder Name")

# Write folder data to rows
for i, data in enumerate(folders_data, start=1):
    worksheet.write(i, 0, data[0])
    worksheet.write(i, 1, data[1])
    worksheet.write(i, 2, data[2])
    worksheet.write(i, 3, data[3])
    worksheet.write(i, 4, data[4])
    worksheet.write(i, 5, data[5])
    worksheet.write(i, 6, data[6])

# Save the XLS file
xls_file_path = os.path.join(target_directory, "folder_list.xls")
workbook.save(xls_file_path)

print(f"XLS file 'folder_list.xls' has been generated successfully.")
