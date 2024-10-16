import os 
import tkinter as tk 
import threading 
import time 
from tkinter import ttk, messagebox 
from tkinter import filedialog 
from PIL import Image, ImageTk 
from openpyxl import load_workbook


# Global variables
current_row = None
current_image = None
retry_count = 0
FOLDER_PATH = ""
EXCEL_PATH = ""
preloaded_image = None
image_size = (1536, 864)  # Default image size


# Function to prompt the user to select window size
def choose_window_size():
    size_prompt = tk.Tk()
    size_prompt.title("Choose Window Size")
    size_prompt.geometry("300x150")
    
    def set_normal_size():
        global image_size, window_size
        window_size = "1550x1200"
        image_size = (1536, 864)  # Set image size for normal window
        size_prompt.destroy()

    def set_small_size():
        global image_size, window_size
        window_size = "1163x900"
        image_size = (1152, 648)  # Set image size for small window
        size_prompt.destroy()

    tk.Label(size_prompt, text="Select Window Size").pack(pady=10)
    
    normal_button = tk.Button(size_prompt, text="Normal (1550x1200)", command=set_normal_size)
    normal_button.pack(pady=5)
    
    small_button = tk.Button(size_prompt, text="Small (1163x900)", command=set_small_size)
    small_button.pack(pady=5)
    
    size_prompt.grab_set()  # Prevent interacting with other windows until this is closed
    size_prompt.mainloop()  # Wait until the size is selected

# Call the window size prompt when the program starts
choose_window_size()

# Initialize the main window after choosing the size
window = tk.Tk()
window.title("Zooniverse Validator v.1.08")
window.geometry(window_size)  # Set the window size based on user selection

# Function to choose folder and Excel file
def choose_folder_and_file():
    global FOLDER_PATH, EXCEL_PATH, wb, ws

    # Ask the user to choose the image folder
    FOLDER_PATH = filedialog.askdirectory(title="Select Image Folder")
    if not FOLDER_PATH:
        messagebox.showerror("Error", "You must select an image folder.")
        return False

    # Ask the user to select the Excel file
    EXCEL_PATH = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not EXCEL_PATH:
        messagebox.showerror("Error", "You must select an Excel file.")
        return False

    # Load the workbook and worksheet
    try:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active  # Assuming the filenames are in the first sheet
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open Excel file: {e}")
        return False

    return True  # Successfully chosen folder and file



# Find the next unvalidated row with 'FALSE' in column A
def find_next_unvalidated():
    for row in range(2, ws.max_row + 1):  # Start from row 2 (assuming row 1 is a header)
        status = ws[f'A{row}'].value  # Check the value in column A
        if status is None:
            status = "FALSE"  # Treat missing values as unvalidated (set as 'FALSE')
        status = str(status).strip().upper()  # Normalize the value for comparison
        print(f"Row {row} status: {status}")  # Debug print statement to track status
        
        if status == "FALSE":  # If unvalidated (FALSE), return the row number
            return row

    return None  # Return None if no unvalidated rows remain



# Add a label to display the original classifications below the image
original_info_label = tk.Label(window, text="", justify=tk.LEFT, wraplength=1000)
original_info_label.grid(row=5, column=0, columnspan=4)


# Function to search for the original species information in column L for the current image filename
def find_original_classification(filename):
    original_classifications = []
    for row in range(2, ws.max_row + 1):
        original_filename = ws[f'L{row}'].value  # Assuming original filenames are in column L
        if original_filename == filename:
            species1_orig = ws[f'M{row}'].value or "NONE"
            species1_count_orig = ws[f'N{row}'].value or "NONE"
            species2_orig = ws[f'O{row}'].value or "NONE"
            species2_count_orig = ws[f'P{row}'].value or "NONE"
            original_classifications.append(
                (species1_orig, species1_count_orig, species2_orig, species2_count_orig)
            )
    return original_classifications


# Function to display the original classifications under the image
def display_original_classifications(classifications):
    # Clear previous text
    original_info_label.config(text="")
    
    # Calculate the width of each column
    column_widths = [30, 20, 30, 20]
    total_width = sum(column_widths) + len(column_widths) - 1
    
    # Header for clarity, use fixed-width alignment
    header = f"{'Species 1':<{column_widths[0]}} {'Count 1':<{column_widths[1]}} {'Species 2':<{column_widths[2]}} {'Count 2':<{column_widths[3]}}\n"
    info_text = header
    
    # Loop through classifications and format them into columns
    for (species1, count1, species2, count2) in classifications:
        line = f"{species1:<{column_widths[0]}} {str(count1):<{column_widths[1]}} {species2:<{column_widths[2]}} {str(count2):<{column_widths[3]}}\n"
        info_text += line
    
    # Display the formatted text in the label using a fixed-width font
    original_info_label.config(text=info_text, font=("Courier", 11))



# Function to show the image
def show_image(image_path, preloaded=False):
    global current_image, preloaded_image, image_size

    if preloaded and preloaded_image:
        current_image = preloaded_image
        preloaded_image = None
    else:
        try:
            image = Image.open(image_path)
            image = image.resize(image_size, Image.Resampling.LANCZOS)  # Resize image based on window size
            current_image = ImageTk.PhotoImage(image)
            print(f"Displaying image: {image_path}")  # Debug message
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open image: {e}")
            print(f"Error loading image: {e}")  # Debug print for error
    image_label.config(image=current_image)


# Preloading Function
def preload_next_image(next_row):
    print(f"Preloading image for row {next_row}...")  # Debug message
    global preloaded_image, image_size
    filename = ws[f'C{next_row}'].value
    if filename:
        image_path = find_image_in_subfolders(filename)
        if image_path:
            try:
                image = Image.open(image_path)
                image = image.resize(image_size, Image.Resampling.LANCZOS)  # Resize based on window size
                preloaded_image = ImageTk.PhotoImage(image)
                print(f"Finished preloading image for row {next_row}.")  # Debug message
            except Exception as e:
                print(f"Failed to preload image: {e}")  # Error handling and debug


# Function to search for an image in the folder and subfolders, case-insensitive
def find_image_in_subfolders(filename):
    start_time = time.time()
    print(f"Searching for {filename} in subfolders...")
    
    # Convert the target filename to lowercase for comparison
    target_filename = filename.lower()
    
    for root, dirs, files in os.walk(FOLDER_PATH):
        # Loop through each file in the current directory
        for file in files:
            # Compare the filenames case-insensitively
            if file.lower() == target_filename:
                elapsed_time = time.time() - start_time
                print(f"Found {file} in {elapsed_time:.2f} seconds.")  # Debug message
                return os.path.join(root, file)
    
    elapsed_time = time.time() - start_time
    print(f"Could not find {filename}. Search took {elapsed_time:.2f} seconds.")  # Debug message
    return None


# Count number of rows in column A
# Yes I know what it's called.
def get_total_rows_in_column_c():
    total_rows = 0
    for row in range(2, ws.max_row + 1):  # Starting from row 2 (assuming row 1 is header)
        if ws[f'A{row}'].value is not None:  # If Column A has any value (string, number, boolean, etc.)
            total_rows += 1
    return total_rows


# Function to count rows with the exact string "TRUE" in Column A
def get_true_rows_in_column_A():
    true_rows = 0
    for row in range(2, ws.max_row + 1):  # Starting from row 1 (including header)
        if ws[f'A{row}'].value == "TRUE":  # Check if the value in column A is exactly "TRUE"
            true_rows += 1
    return true_rows



# Add a label for displaying the filename and image number
image_info_label = tk.Label(window, text="", font=("Courier", 10))
image_info_label.grid(row=2, column=0, columnspan=4)  # Positioned under "Next" button

# Function to update the filename and image count display
def update_image_info(filename, row_number):
    total_rows = get_total_rows_in_column_c()  # Get total rows in Column A
    true_rows = get_true_rows_in_column_A()    # Get number of TRUE rows in Column A
    # Update the label to show the current filename and the true rows/total rows
    image_info_label.config(text=f"{filename} ({true_rows}/{total_rows})")
    print(f"Updated image info: {filename}, TRUE Rows: {true_rows}/{total_rows}")  # Debugging print statement




#Load next image function with updating file data functions
def load_next_image():
    global current_row, retry_count, preloaded_image
    
    if current_row is None:
        messagebox.showinfo("Done", "All images validated!")
        return
    
    # Initialize retry_count if it's the first image in the batch
    if retry_count is None:
        retry_count = 0

    # Show a loading message
    original_info_label.config(text="Loading image...")
    window.update_idletasks()

    def load_image_thread():
        global retry_count  # Declare retry_count as global
        
        print(f"Starting to load image for row {current_row}...")
        if preloaded_image:
            window.after(1, lambda: show_image(None, preloaded=True))  # Display the preloaded image directly
            retry_count = 0  # Reset retry count
            
            filename = ws[f'C{current_row}'].value
            # Update the filename and image count display
            window.after(1, lambda: update_image_info(filename, current_row - 1))

            # Find and display the original classification for the current image
            original_classifications = find_original_classification(filename)
            if original_classifications:
                window.after(1, lambda: display_original_classifications(original_classifications))
                
            else:
                window.after(1, lambda: original_info_label.config(text="No original classifications found for this image."))
            
            # Preload the next image
            next_row = current_row + 1
            if next_row <= ws.max_row:
                preload_next_image(next_row)
        
        else:
            filename = ws[f'C{current_row}'].value
            if filename:
                image_path = find_image_in_subfolders(filename)
                if image_path:
                    window.after(1, lambda: show_image(image_path))  # Display current image
                    retry_count = 0  # Reset retry count
                    
                    # Update the filename and image count display
                    window.after(1, lambda: update_image_info(filename, current_row - 1))

                    
                    # Find and display the original classification for the current image
                    original_classifications = find_original_classification(filename)
                    if original_classifications:
                        window.after(1, lambda: display_original_classifications(original_classifications))
                    else:
                        window.after(1, lambda: original_info_label.config(text="No original classifications found for this image."))
                    
                    # Preload the next image
                    next_row = current_row + 1
                    if next_row <= ws.max_row:
                        preload_next_image(next_row)
                    
                else:
                    retry_count += 1
                    if retry_count >= 3:
                        window.after(1, lambda: messagebox.showinfo("Stopping", "Image could not be found - Please verify filename in excel sheet and image folder."))
                        window.after(1, window.quit)
                        return
                    window.after(1, load_next_image)
            else:
                retry_count += 1
                if retry_count >= 3:
                    window.after(1, lambda: messagebox.showinfo("Stopping", "Image could not be found - Please verify filename in excel sheet and image folder."))
                    window.after(1, window.quit)
                    return
                window.after(1, load_next_image)
    
    # Run image loading in a background thread
    threading.Thread(target=load_image_thread).start()
    print(f"Finished loading image for row {current_row}.")


# Add this function to set the current row and the previous row to FALSE and load the previous row
def go_back():
    global current_row, preloaded_image

    print(f"Requesting to return to previous image.")

    # Clear the preloaded image to prevent skipping issues
    preloaded_image = None
    print(f"Clearing preloaded images to prevent skipping.")

    # If there's no row selected or we're at the first image, show a message
    if current_row is None or current_row <= 2:  # Assuming row 2 is the first data row
        messagebox.showinfo("Info", "Cannot go back; this is the first image or no row selected.")
        return

    # Move to the previous row
    previous_row = current_row - 1

    #Set previous row to false
    ws[f'A{previous_row}'].value = "FALSE"
    print(f"Set previous row to FALSE")

    # Search for the previous row (status == "FALSE")
    while previous_row > 1:  # Stay above the header row (assumed to be row 1)
        status = ws[f'A{previous_row}'].value

        if status is None or status.strip().upper() == "FALSE":  # Unvalidated row found
    
            # Set the previous row status to "FALSE" and save the workbook
            wb.save(EXCEL_PATH)
            print(f"Workbook saved after setting previous row {previous_row} to FALSE.")

            # Update the current row to the previous row
            current_row = previous_row
            print(f"Setting current row to {current_row}.")

            # Load the image for the previous row
            load_image(current_row)
            print(f"Loading image for row {current_row}")
            return  # Successfully loaded previous row

        # If the row is validated, keep moving backward
        previous_row -= 1

    messagebox.showinfo("Info", "Error in handling back command (no previous row)")


def load_image(row):
    """Function to load the image based on the row."""
    if row is None:
        messagebox.showinfo("Done", "All images validated!")
        return

    # Show a loading message
    original_info_label.config(text="Loading image...")
    window.update_idletasks()

    # Get the filename from column C in the current row
    filename = ws[f'C{row}'].value  # Assuming column C holds the image filenames

    if filename:
        image_path = find_image_in_subfolders(filename)
        if image_path:
            show_image(image_path)  # Display the image
            update_image_info(filename, row - 1)  # Update the filename and image count display

            # Find and display the original classification for the current image
            original_classifications = find_original_classification(filename)
            if original_classifications:
                display_original_classifications(original_classifications)
            else:
                original_info_label.config(text="No original classifications found for this image.")
        else:
            messagebox.showerror("Error", f"File {filename} not found in the selected folder or its subfolders!")
    else:
        messagebox.showerror("Error", "No filename found in the current row.")



# Create Back button
back_button = tk.Button(window, text="Back", command=go_back)
back_button.grid(row=4, column=1, padx=0, sticky=tk.N)

# Function to save data and move to next row
def save_and_next():
    global current_row
    if current_row is None:
        return
    
    # Get the species and count data
    ws[f'D{current_row}'].value = species1_var.get()
    ws[f'F{current_row}'].value = species2_var.get()
    ws[f'E{current_row}'].value = species1_count_var.get()  
    ws[f'G{current_row}'].value = species2_count_var.get()  

    print(f"Saving data for row {current_row}...")
    
    # Mark row as validated
    ws[f'A{current_row}'].value = "TRUE"
    
    # Save the workbook
    wb.save(EXCEL_PATH)
    print(f"Data saved for row {current_row}.")

    # Move to next unvalidated row
    current_row = find_next_unvalidated()
    if current_row is None:
        messagebox.showinfo("Done", "All images validated!")
    else:
        load_next_image()  # Load the next image (preloaded)


# Custom autocomplete Combobox class
class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list)  # Sort the list
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = self._completion_list  # Set the values

    def autocomplete(self, delta=0):
        if delta:  # Move the index in the list based on the delta
            self._hit_index += delta
            if self._hit_index == len(self._hits):
                self._hit_index = 0
            elif self._hit_index == -1:
                self._hit_index = len(self._hits) - 1
        if self._hits:
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down'):
            return

        self.position = self.index(tk.END)
        _input = self.get()

        if _input == '':
            self._hits = self._completion_list
        else:
            self._hits = [item for item in self._completion_list if item.lower().startswith(_input.lower())]

        if self._hits:
            self._hit_index = 0
            self.autocomplete()

# Adjust grid column weights for better alignment
window.grid_columnconfigure(0, weight=0)  # Column for labels (left)
window.grid_columnconfigure(1, weight=0)  # Column for dropdowns (middle left)
window.grid_columnconfigure(2, weight=0)  # Column for count labels (middle right)
window.grid_columnconfigure(3, weight=0)  # Column for entry boxes (right)

# Labels for species and count fields
species1_label = tk.Label(window, text="Species 1: ")
species1_label.grid(row=1, column=0, sticky="e", padx=0, pady=5)  # Aligned right, next to species1 dropdown

species1_count_label = tk.Label(window, text="Species 1 Count: ")
species1_count_label.grid(row=1, column=2, sticky="e", padx=0, pady=5)  # Aligned right, next to species1_count entry

species2_label = tk.Label(window, text="Species 2: ")
species2_label.grid(row=2, column=0, sticky="e", padx=0, pady=5)  # Aligned right, next to species2 dropdown

species2_count_label = tk.Label(window, text="Species 2 Count: ")
species2_count_label.grid(row=2, column=2, sticky="e", padx=0, pady=5)  

# Create species dropdowns with the Autocomplete feature
species_options = ["NONE", "HUMAN", "MULEDEER", "DESERTCOTTONTAIL", "COYOTE", "RODENT", "VIRGINIAOPOSSUM", "EASTERNFOXSQUIRREL", "CALIFORNIAGROUNDSQUIRREL", "STRIPEDSKUNK", "BIRD", "BLACKBEAR", "RACCOON", "BOBCAT", "SNAKE", "DOMESTICDOG", "DOMESTICCAT"]  # Add more species here

# Create species count options
count_options = ["NONE", "1", "2", "3", "4ORMORE"]


species1_var = tk.StringVar(value="NONE")
species2_var = tk.StringVar(value="NONE")

# Set the width parameter to accommodate long species names
species1_dropdown = AutocompleteCombobox(window, textvariable=species1_var, width=30)
species1_dropdown.set_completion_list(species_options)
species1_dropdown.grid(row=1, column=1, padx=0, pady=5, sticky="w")  # Positioned next to species1_label

species2_dropdown = AutocompleteCombobox(window, textvariable=species2_var, width=30)
species2_dropdown.set_completion_list(species_options)
species2_dropdown.grid(row=2, column=1, padx=0, pady=5, sticky="w")  # Positioned next to species2_label

# Create dropdowns for species counts with predefined options
species1_count_var = tk.StringVar(value="NONE")
species1_count_dropdown = ttk.Combobox(window, textvariable=species1_count_var, values=count_options, width=20, state="normal")
species1_count_dropdown.grid(row=1, column=3, padx=0, pady=5, sticky="w")

species2_count_var = tk.StringVar(value="NONE")
species2_count_dropdown = ttk.Combobox(window, textvariable=species2_count_var, values=count_options, width=20, state="normal")
species2_count_dropdown.grid(row=2, column=3, padx=0, pady=5, sticky="w")


# Image display area
image_label = tk.Label(window)
image_label.grid(row=0, column=0, columnspan=4)

# Buttons for navigation
next_button = tk.Button(window, text="Next", command=save_and_next)
next_button.grid(row=4, column=1, padx=0, columnspan=2)

# Button to set all species data and counts to "NONE"
set_all_none_button = tk.Button(window, text="Set All to None", command=lambda: (
    species1_var.set("NONE"),  # Set species1 to NONE
    species2_var.set("NONE"),  # Set species2 to NONE
    species1_count_var.set("NONE"),  # Set first species count to NONE
    species2_count_var.set("NONE")   # Set second species count to NONE
))
set_all_none_button.grid(row=3, column=1, columnspan=2)


# This is now called immediately when the program starts to ask the user for the folder and file
if choose_folder_and_file():
    current_row = find_next_unvalidated()  # Find the first unvalidated row after loading the file
    if current_row:
        load_next_image()  # Load the first image if unvalidated rows exist
    else:
        messagebox.showinfo("Done", "All images are already validated.")
else:
    window.destroy()  # Close the window if the user cancels or an error occurs

# Start the Tkinter event loop
window.mainloop()
