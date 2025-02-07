import os
import shutil
import pandas as pd
from fuzzywuzzy import fuzz
import string
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import re

# Set the default source folder for Source 1
SOURCE_FOLDER_1 = "C:/Users/hfreeman/Desktop/masterfootnotes"

# Maximum allowed filename length
MAX_FILENAME_LENGTH = 250

# Create a function to open a folder dialog and set the selected path to an entry field
def browse_directory(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

# Create a function to open a file dialog for Excel files
def browse_excel_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

# Create a function to run the file matching process
def run_file_matching():
    source_folder_1 = source_folder_entry.get()
    source_folder_2 = source_folder_entry_2.get()
    target_folder = target_folder_entry.get()
    excel_file_path = excel_file_path_entry.get()

    if not source_folder_1 or not source_folder_2 or not target_folder or not excel_file_path:
        status_label.config(text="Please fill in all fields.")
        return

    try:
        df = pd.read_excel(excel_file_path, header=None)

        # Remove the first four characters from each value in the first column
        df[0] = df[0].str[4:]

    except Exception as e:
        status_label.config(text=f"Error reading Excel file: {str(e)}")
        return

    # Create a dictionary of synonyms
    synonym_dictionary = {
        "TMB": "Texas Medical Board",
        "ABPN": "American Board of Psychiatry and Neurology",
        "AB": "American Board",
        "ABOS": "American Board of Orthopaedic Surgery",
        "LCP": "Life Care Plan",
        "PVA": "Present Value Assessment",
        "IME": "Independent Medical Examination",
        "VOC": "Vocational Evaluation Report",
        "VOC": "Vocational Rehabilitation Assessment",
        "LEC": "Loss of Earning Capacity",
        "ABNS": "American Board of Neurological Surgery",
        "Healthy Weight, Nutrition, and Physical Activity": "Centers for Disease Control and Prevention",
        "ABPP": "American Board of Professional Psychology",
        "ABPMR": "American Board of Physical Medicine and Rehabilitation",
        "ABIM": "American Board of Internal Medicine",
        "TBHEC": "Texas Behavioral Health Executive Council"
    }

    def sanitize_filename(filename):
        valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)

        # Remove "http" and everything following it
        sanitized = re.sub(r'http.*', '', filename)

        # Remove other invalid characters
        sanitized = ''.join(c for c in sanitized if c in valid_chars)

        # Truncate filename if it's too long
        if len(sanitized) > MAX_FILENAME_LENGTH:
            sanitized = sanitized[:MAX_FILENAME_LENGTH]

        return sanitized

    def find_most_similar_filename(filename, source_folders):
        best_match = None
        best_similarity = -1

        # Sanitize the filename before matching
        sanitized_filename = sanitize_filename(filename)

        for source_folder in source_folders:
            for root, _, files in os.walk(source_folder):
                for item in files:
                    full_path = os.path.join(root, item)
                    similarity = fuzz.token_sort_ratio(sanitized_filename, item)
                    if similarity > best_similarity:
                        best_similarity = similarity
                        best_match = full_path

        return best_match, best_similarity

    min_similarity_threshold = 60

    for index, row in df.iterrows():
        number = str(index + 1)  # Start numbering with 1
        filename = str(row[0])  # Use the modified value from the first column

        # Check if the filename contains acronyms and replace them with synonyms
        for acronym, synonym in synonym_dictionary.items():
            if synonym in filename:
                filename = sanitize_filename(filename.replace(synonym, acronym))

        similar_filename, similarity = find_most_similar_filename(filename, [source_folder_1, source_folder_2])

        _, source_extension = os.path.splitext(similar_filename)
        source_extension = source_extension.lower()

        # If the file is a .mp4 or .mov, create an empty text file and continue
        if source_extension in ['.mp4', '.mov']:
            empty_text_file_path = os.path.join(target_folder, f"{number}.txt")
            with open(empty_text_file_path, 'w') as blank_file:
                pass  # The text file is left empty
            status_label.config(text=f"Found video file ('{similar_filename}'). Created placeholder text file: '{number}.txt'")
            continue

        if similarity < min_similarity_threshold:
            blank_text_file_path = os.path.join(target_folder, f"{number}.txt")
            with open(blank_text_file_path, 'w') as blank_file:
                pass  # The text file is left empty but you can optionally write something here

            status_label.config(text=f"No similar filename found in source folders for '{filename}' with sufficient similarity. Created a blank text file with the number.")
        else:
            # Remove all numbers and periods before the first letter of the matched filename
            matched_filename = re.sub(r'^[0-9.\s-]+', '', os.path.basename(similar_filename))

            # Add the prefix to the matched filename
            target_filename = f"{number}. {matched_filename}"

            # Ensure the filename is exactly 120 characters by replacing characters from the middle
            while len(target_filename) > 120:
                middle_index = len(target_filename) // 2
                target_filename = f"{target_filename[:middle_index - 1]}{target_filename[middle_index + 1:]}"

            target_file_path = os.path.join(target_folder, target_filename)
            source_file_path = similar_filename

            try:
                shutil.copy(source_file_path, target_file_path)
                status_label.config(text=f"File copied: '{target_filename}'")
            except FileNotFoundError:
                status_label.config(text=f"File not found in source folders for '{filename}': {source_file_path}")

    status_label.config(text="File copying completed.")


# Create the main GUI window
root = tk.Tk()
root.title("Fuzzy Footnotes")
root.geometry("600x450")  # Set the initial window size

# Set a standard window icon (use one of the built-in icons)
# This icon may vary slightly depending on the system's OS
# Common icon names: "info", "warning", "error", "question", "hourglass", "new", "star", "grey25"
root.iconbitmap(default='info')

# Create and style GUI elements using ttk
style = ttk.Style()
style.configure("TButton", padding=(10, 5))

# Add padding to the outer frame
main_frame = ttk.Frame(root)
main_frame.pack(pady=20)

# Source Folder 1 input fields
source_folder_label = ttk.Label(main_frame, text="Source Folder 1:")
source_folder_label.grid(row=0, column=0, padx=10, pady=5)
source_folder_entry = ttk.Entry(main_frame)
source_folder_entry.grid(row=0, column=1, padx=10, pady=5)
source_folder_entry.insert(0, SOURCE_FOLDER_1)  # Preload the path for Source 1
source_folder_button = ttk.Button(main_frame, text="Browse", command=lambda: browse_directory(source_folder_entry))
source_folder_button.grid(row=0, column=2, padx=10, pady=5)

# Source Folder 2 input fields
source_folder_label_2 = ttk.Label(main_frame, text="Source Folder 2:")
source_folder_label_2.grid(row=1, column=0, padx=10, pady=5)
source_folder_entry_2 = ttk.Entry(main_frame)
source_folder_entry_2.grid(row=1, column=1, padx=10, pady=5)
source_folder_button_2 = ttk.Button(main_frame, text="Browse", command=lambda: browse_directory(source_folder_entry_2))
source_folder_button_2.grid(row=1, column=2, padx=10, pady=5)

# Target Folder input fields
target_folder_label = ttk.Label(main_frame, text="Target Folder:")
target_folder_label.grid(row=2, column=0, padx=10, pady=5)
target_folder_entry = ttk.Entry(main_frame)
target_folder_entry.grid(row=2, column=1, padx=10, pady=5)
target_folder_button = ttk.Button(main_frame, text="Browse", command=lambda: browse_directory(target_folder_entry))
target_folder_button.grid(row=2, column=2, padx=10, pady=5)

# Excel file input fields
excel_file_label = ttk.Label(main_frame, text="Footnotes in Excel:")
excel_file_label.grid(row=3, column=0, padx=10, pady=5)
excel_file_path_entry = ttk.Entry(main_frame)
excel_file_path_entry.grid(row=3, column=1, padx=10, pady=5)
excel_file_button = ttk.Button(main_frame, text="Browse", command=lambda: browse_excel_file(excel_file_path_entry))
excel_file_button.grid(row=3, column=2, padx=10, pady=5)

# Run button
run_button = ttk.Button(main_frame, text="Run", command=run_file_matching)
run_button.grid(row=4, column=1, padx=10, pady=20)

# Status label
status_label = ttk.Label(main_frame, text="", wraplength=500)
status_label.grid(row=5, column=0, columnspan=3, padx=10, pady=5)

root.mainloop()
