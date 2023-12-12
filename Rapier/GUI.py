import customtkinter
from customtkinter import *
from tkinter import *
import tkinter.messagebox
from Rapier_Functions import *

output = ["Excel File", "New Window"]

# Define variables to store file paths
file1 = None
file2 = None
file3 = None


# Function to initiate the file comparison process
def initiate_comparison():

    global file1, file2, file3
    
    # Check if file1_placed has a file path
    if file1 is None:
        tkinter.messagebox.showerror(title="No 'Building Materials List' file Selected", message="Please ensure you have selected a 'Building Materials List' file before continuing.")
        file1_placed.configure(text="\t\t\t\t\t")
        my_canvas.create_text(345,338, text="Select BOM File!", font=("Helvetica", 9), fill="red")

    # Check if file2_placed has a file path
    if file2 is None:
        tkinter.messagebox.showerror(title="No 'Pick and Place' file Selected", message="Please ensure you have selected a 'Pick and Place' file before continuing.")
        file2_placed.configure(text="\t\t\t\t\t")
        my_canvas.create_text(345,438, text="Select PnP File!", font=("Helvetica", 9), fill="red")
        
    # Check if file2 is equal to file3
    if file2 == file3 and file2 != None:
        tkinter.messagebox.showerror(title="Duplicate Files Selected", message="Please ensure that you have selected distinct files for comparison. To compare against a single file, select it once.")
        file3_placed.configure(text="\t\t\t\t\t")

    # Get the selected output option
    selected_option = selected_output.get()

    if selected_option == "Export in...":
       tkinter.messagebox.showerror(title="No Export Option Selected", message="Please choose an export option for your file comparison.")
       my_canvas.create_text(190, 580, text="Select Export!", font=("Helvetica", 9), fill="red")

    # Check if all files are selected before initiating the comparison
    if file1 and file2:
        compareFiles(file1, file2, file3, selected_option)


# Function to open a file dialog and return the selected file path
def find_File():
    file = root.filename = filedialog.askopenfilename(initialdir="/Rapier", title="Select a file",
                                                      filetypes=(("Text files", "*.txt"), ("Excel files", "*.xlsx"), ("All files", "*.*")))
    if file:
        file = file.strip()  # Remove extra spaces and whitespace from the file path
        return file

# Function to update the displayed file path with a shortened version
def update_file_label(label, file_path):
    max_length = 46  # Adjust the maximum length as needed
    if len(file_path) > max_length:
        shortened_path = "..." + file_path[-(max_length - 3):] + " "
    else:
        shortened_path = file_path

    # Update the label's associated file path
    label.file_path = file_path

    label.configure(text=shortened_path)

# Function to select the first file
def select_file1():
    global file1 
    file1 = find_File()
    update_file_label(file1_placed, file1)

# Function to select the second file
def select_file2():
    global file2
    file2 = find_File()
    update_file_label(file2_placed, file2)

# Function to select the third file
def select_file3():
    global file3
    file3 = find_File()
    update_file_label(file3_placed, file3)

# Set the appearance mode and default color theme for customtkinter
customtkinter.set_appearance_mode("white")
customtkinter.set_default_color_theme("green")

# Create the main window
root = customtkinter.CTk()
root.title("Rapier Electronics Manufacturing 2023")
root.iconbitmap("Images/Rapier-Icon.ico")
root.geometry("700x700")

# Define images
backG = PhotoImage(file="Images/background.png")
Title_icon = PhotoImage(file="Images/Title_image_rapier.png")

# Create a canvas
my_canvas = CTkCanvas(root, width=600, height=600)
my_canvas.pack(fill="both", expand=True)

# Set a background image on the canvas
my_canvas.create_image(0, 0, image=backG, anchor="nw")
my_canvas.create_image(195, 60, image=Title_icon, anchor="nw")

# Add labels and text to the canvas
my_canvas.create_text(360, 125, text="Rapier Electronics", font=("Helvetica", 25), fill="white")
my_canvas.create_text(400, 155, text="BOM/Pick And Place", font=("Helvetica", 14), fill="white")

# Create a text label for the BOM File section
BOM_window = my_canvas.create_text(370, 275, text="Building Materials", font=("Helvetica", 12), fill="white")

# Create a button to search for the BOM file
entry1 = customtkinter.CTkButton(master=my_canvas, text="Search BOM", text_color="white",
                                 background_corner_colors=("#041c4a", "#041c4a", "#041c4a", "#041c4a"), command=select_file1)

entry1_window = my_canvas.create_window(150, 300, anchor="nw", window=entry1)

# Create a label to display the selected file name for File 1 (initially empty)
file1_placed = customtkinter.CTkLabel(master=my_canvas, text="\t\t\t\t\t", text_color="black")
file1_window = my_canvas.create_window(300, 300, anchor="nw", window=file1_placed)

# Create a text label for the Pick and Place Files section
BOM_window = my_canvas.create_text(370, 375, text="Pick and Place - Top and Bottom", font=("Helvetica", 12), fill="white")

# Create a button to search for the PnP Top file
entry2 = customtkinter.CTkButton(master=my_canvas, text="Search PnP Top", text_color="white",
                                 background_corner_colors=("#041c4a", "#041c4a", "#041c4a", "#041c4a"), command=select_file2)

entry2_window = my_canvas.create_window(150, 400, anchor="nw", window=entry2)

# Create a label to display the selected file name for PnP Top
file2_placed = customtkinter.CTkLabel(master=my_canvas, text="\t\t\t\t\t", text_color="black")
file2_window = my_canvas.create_window(300, 400, anchor="nw", window=file2_placed)

# Create a button to search for the PnP Bottom file
entry3 = customtkinter.CTkButton(master=my_canvas, 
                                 text="Search PnP Bottom", 
                                 text_color="white",
                                 background_corner_colors=("#041c4a", "#041c4a", "#041c4a", "#041c4a"), 
                                 command=select_file3)

entry3_window = my_canvas.create_window(150, 450, anchor="nw", 
                                        window=entry3)

# Create a label to display the selected file name for PnP Bottom
file3_placed = customtkinter.CTkLabel(master=my_canvas, 
                                      text="\t\t\t\t\t", 
                                      text_color="black")

file3_window = my_canvas.create_window(300, 450, 
                                       anchor="nw", 
                                       window=file3_placed)


# Create an option menu (dropdown) for selecting the output option

selected_output = StringVar()
selected_output.set("Export in...")  # Set the default value

option_menu = customtkinter.CTkOptionMenu(
    master=my_canvas,
    values=output,
    variable=selected_output)
    
option_menu_window = my_canvas.create_window(150, 540, anchor="nw", window=option_menu)

# Initialize the comparing process button
button = customtkinter.CTkButton(
    master=my_canvas,
    text="Compare Files",
    command=initiate_comparison,
    background_corner_colors=("#041c4a", "#041c4a", "#041c4a", "#041c4a")
)

button3_window = my_canvas.create_window(440, 540, anchor="nw", window=button)

# Start the main GUI loop
root.mainloop()