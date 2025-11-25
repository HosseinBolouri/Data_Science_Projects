# Inserting relative libraries
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import pandas as pd

# Suppress chained assignment warnings
pd.set_option('future.no_silent_downcasting', True)
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image, ImageTk

# The base path for locating resources, compatible with both .py and .exe files
import os
import sys
base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

# DPI handling to manage scalling
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 1 = SYSTEM_DPI_AWARE, 2 = PER_MONITOR_DPI_AWARE (for multiple_monitors)
except Exception:
    pass

# Encapsulate global variables
class AppData:
    def __init__(self):

        # Variables to track file upload for AMAZON
        self.file_uploaded = False
        self.df_v1_read_point_24h = None
        self.df_testing_24h = None
        self.df_general_24h = None
        self.df_analysis_24h = None
        self.df_general_24h_help_1 = None
        self.df_general_24h_help_2 = None
        self.df_general_24h_help_3 = None
        self.graph_24h = None
        self.failure_24h = None
        self.df_limit_24h = None
        self.read_point_graphs= []
        self.final_table = None
        self.overview_generated = False
        self.graph_24h_generated = False
        self.excessive_units = None
        self.df_limit_24h_generated = False
        self.analysis_point_graphs = []
        self.full_limit_amazon_merged = None
        self.full_limit_amazon_generated = False
        self.excessive_units_generated = False
        self.plot_df_amazon = None
        self.plot_df_amazon_generated = False
        self.full_distribution_amazon_df = None
        self.full_distribution_amazon_generated = False
        
        # Global Variables for IRIS
        self.file_uploaded_iris = False
        self.df_v1_read_point_24h_iris = None
        self.df_testing_24h_iris = None
        self.df_general_24h_iris = None
        self.df_analysis_24h_iris = None
        self.df_general_24h_help_1_iris = None
        self.df_general_24h_help_2_iris = None
        self.df_general_24h_help_3_iris = None
        self.graph_24h_iris = None
        self.failure_24h_iris = None
        self.df_limit_24h_iris = None
        self.read_point_graphs_iris = []
        self.final_table_iris = None
        self.overview_generated_iris = False
        self.graph_24h_generated_iris = False
        self.excessive_units_iris = None
        self.df_limit_24h_iris_generated= False
        self.analysis_point_graphs_iris = []
        self.excessive_units_iris_generated = False
        self.analysis_point_graphs_iris_merged = None
        self.full_graph_iris_generated = False
        self.full_graph_iris_df = None
        self.plot_df_generated = False
        self.plot_df = None
        self.full_limit_generated = False


# Function to show the tooltip dynamically
def show_dynamic_tooltip(event, tooltip_text, parent_window):
    global dynamic_tooltip
    dynamic_tooltip = tk.Toplevel(parent_window)
    dynamic_tooltip.wm_overrideredirect(True)
    dynamic_tooltip.wm_attributes("-topmost", True)

    # Create the label for the tooltip
    label = tk.Label(dynamic_tooltip, text=tooltip_text, fg="black", bg="white", wraplength=320)
    label.pack()

    # Wait for the label to be drawn so we can get its size
    dynamic_tooltip.update_idletasks()
    tooltip_width = label.winfo_reqwidth()
    tooltip_height = label.winfo_reqheight()

    # Get screen dimensions
    screen_width = parent_window.winfo_screenwidth()
    screen_height = parent_window.winfo_screenheight()

    # Calculate initial position
    x = event.x_root + 10
    y = event.y_root + 10

    # Adjust position if the tooltip exceeds screen boundaries
    if x + tooltip_width > screen_width:
        x = screen_width - tooltip_width - 10
    if y + tooltip_height > screen_height:
        y = screen_height - tooltip_height - 10

    # Set the adjusted position
    dynamic_tooltip.geometry(f"{tooltip_width}x{tooltip_height}+{x}+{y}")

# Function to hide the tooltip
def hide_dynamic_tooltip(event):
    if dynamic_tooltip:
        dynamic_tooltip.destroy()

# Create the main window
root = tk.Tk()
root.title("MDRF Q&R Data Analysis Control Tower")

# Set the dimensions of the main window
window_width = 700
window_height = 500
root.geometry(f"{window_width}x{window_height}")

# Set the application icon
root.iconbitmap(os.path.join(base_path, "resources", "Q&R.ico"))

# Initialize shared data
app_data = AppData()

# Create a style object for ttk
style = ttk.Style()
style.configure("TFrame", background="white")

# Calculate the width of each frame based on the given ratios
frame1_width = window_width * 15 / (35 + 12)
frame2_width = window_width * 12 / (35 + 12)

# Create the first frame
frame1 = ttk.Frame(root, width=frame1_width, height=window_height, relief="flat")
frame1.grid(row=0, column=0, sticky="nsew")

# Create the second frame with a dark blue theme
frame2 = ttk.Frame(root, width=frame2_width, height=window_height, relief="flat", style="DarkBlue.TFrame")
frame2.grid(row=0, column=1, sticky="nsew")

# Configure grid weights to ensure frames expand properly
root.grid_columnconfigure(0, weight=35)
root.grid_columnconfigure(1, weight=12)
root.grid_rowconfigure(0, weight=1)

# Configure the style for the dark blue theme
style.configure("DarkBlue.TFrame", background="#1C2541")  # Dark blue color

# Construct the path to the GIF image for Frame 1 using the base_path for .exe file
image_path_frame1 = os.path.join(base_path, "resources", "Data_6.jpg")
photo_frame1 = Image.open(image_path_frame1)
photo_frame1_tk = ImageTk.PhotoImage(photo_frame1)
label_frame1 = tk.Label(frame1, image=photo_frame1_tk)
label_frame1.image = photo_frame1_tk  # Keep reference to avoid garbage collection
label_frame1.place(x=0, y=0, relwidth=1, relheight=1)

def resize_image(event):
    new_width = event.width
    new_height = event.height
    resized_image = photo_frame1.resize((new_width, new_height), Image.Resampling.LANCZOS)
    photo = ImageTk.PhotoImage(resized_image)
    label_frame1.config(image=photo)
    label_frame1.image = photo  # Keep reference

frame1.bind("<Configure>", resize_image)

# Add a text box at the top of Frame 2
header_font = ("Helvetica", 22, "bold")
text_label = tk.Label(frame2, text="Q&R Data Analysis Control Tower", font=header_font, fg="white", bg="#1C2541")
text_label.pack(pady=10, anchor='n')

# Add the "Developed by Eng. Hossein Bolouri" text under the header
developer_font = ("Helvetica", 15, "italic")
developer_label = tk.Label(frame2, text="Designed and Developed by Eng. Hossein Bolouri", font=developer_font, fg="white", bg="#1C2541")
developer_label.pack(pady=5, anchor='n')

# Add the software version text under the developer label (under Hossein Bolouri section)
version_font = ("Helvetica", 12, "italic")
version_label = tk.Label(frame2, text="Version: V1.1", font=version_font, fg="white", bg="#1C2541")
version_label.pack(pady=2, anchor='n')

# Create a frame (button frame) for the buttons to center them vertically in Frame 2
button_frame = tk.Frame(frame2, bg="#1C2541")
button_frame.pack(expand=True)

# Create sub-option frame for AMAZON MAKALU. This is the child of button_frame in Frame 2 of the main window
makalu_sub_option_frame = tk.Frame(button_frame, bg="#1C2541")

# Flag to check if sub-options for AMAZON MAKALU are visible
makalu_sub_options_visible = False

# Function to toggle sub-options for AMAZON MAKALU
def toggle_makalu_options():
    global makalu_sub_options_visible
    if makalu_sub_options_visible:
        makalu_sub_option_frame.pack_forget()
    else:
        makalu_sub_option_frame.pack(pady=5, after=button1)
    makalu_sub_options_visible = not makalu_sub_options_visible

# Create sub-options for AMAZON Reliability Analysis. This is the child of makalu_sub_option_frame since it is a sub-category of AMAZON MAKALU
reliability_sub_option_frame = tk.Frame(makalu_sub_option_frame, bg="#1C2541")

# Flag to check if AMAZON Reliability sub-options are visible
reliability_sub_options_visible = False

# Create sub-option frame for CIENA IRIS. This is the child of button_frame in Frame 2 of the main window
ciena_sub_option_frame = tk.Frame(button_frame, bg="#1C2541")
 
# Flag to check if sub-options for CIENA IRIS is visible
ciena_sub_options_visible = False

# Function to toggle sub-options for CIENA IRIS
def toggle_ciena_options():
    global ciena_sub_options_visible
    if ciena_sub_options_visible:
        ciena_sub_option_frame.pack_forget()
    else:
        ciena_sub_option_frame.pack(pady=5, after=button2)
    ciena_sub_options_visible = not ciena_sub_options_visible

# Function to toggle IRIS sub-options (Garph and IRIS Excel Macro) for IRIS Reliability Analysis of CIENA
def toggle_ciena_reliability_options():
    global ciena_reliability_sub_options_visible
    if ciena_reliability_sub_options_visible:
        ciena_reliability_sub_option_frame.pack_forget()
    else:
        ciena_reliability_sub_option_frame.pack(pady=5, after=sub_button2_iris)
    ciena_reliability_sub_options_visible = not ciena_reliability_sub_options_visible

# Create sub-options for CIENA Reliability Analysis. This is the child of ciena_sub_option frame
ciena_reliability_sub_option_frame = tk.Frame(ciena_sub_option_frame, bg="#1C2541")

# Flag to check if CIENA IRIS Reliability sub-options Analysis is visible
ciena_reliability_sub_options_visible = False

# Define button style with 3D effect, white font color, and bold text
button_style = {
    'font': ('Verdana', 12, 'bold'),  # Bold font
    'fg': 'white',  # White font color
    'bg': "#1C2541",
    'relief': 'raised',  # 3D effect
    'bd': 3,  # Border width for 3D effect
    'width': 20  # Fixed width for uniform button size
}

# Define italic button style for sub-options (without width)
italic_button_style = {
    'font': ('Verdana', 12, 'italic'),  # Italic font
    'fg': 'white',  # White font color
    'bg': "#1C2541",
    'relief': 'raised',  # 3D effect
    'bd': 3  # Border width for 3D effect
}

# Create 'AMAZON MAKALU' button and put it in the middle of Frame 2 using pack geometry manager
button1 = tk.Button(button_frame, text="AMAZON",command=toggle_makalu_options, **button_style)
button1.pack(side='top', pady=5)

# Create 'CIENA IRIS' button in Frame 2
button2 = tk.Button(button_frame, text="CIENA", command=toggle_ciena_options, **button_style)
button2.pack(side='top', pady=5)

# Create 'OTHER PRODUCTS' button with specified styles (button_style)
button3 = tk.Button(button_frame, text="OTHER PRODUCTS", **button_style)

# Locate 'CIENA IRIS' and 'OTHER PRODUCTS' vertically with equal spacing using pack geometry manager
button3.pack(side='top', pady=5)

# Sub-options for AMAZON MAKALU (Screening and Reliability features)
sub_button1 = tk.Button(makalu_sub_option_frame, text="Screening Analysis", **italic_button_style)
sub_button2 = tk.Button(makalu_sub_option_frame, text="Reliability Analysis", **italic_button_style)

# Using pack geometry manager to locate the buttons vertically in the middle of Frame 2
sub_button1.pack(side='top', pady=5)
sub_button2.pack(side='top', pady=5)

# Create sub-options for CIENA IRIS button (Screening and Reliability features)
sub_button1_iris = tk.Button(ciena_sub_option_frame, text="Screening Analysis", **italic_button_style)
sub_button2_iris = tk.Button(ciena_sub_option_frame, text="Reliability Analysis", command=toggle_ciena_reliability_options,**italic_button_style)

# Using pack geometry manager to locate the buttons vertically in the middle of Frame 2
sub_button1_iris.pack(side='top', pady=5)
sub_button2_iris.pack(side='top', pady=5)

# Calculate half the width of the Reliability Analysis button
half_button_width = button_style['width'] // 2

# Add a frame at the bottom of Frame 2 for the footer (footer design)
footer_frame = tk.Frame(frame2, bg="#1C2541")
footer_frame.pack(pady=10, anchor='s', side='bottom')

# Construct the path to the OUTLOOK logo image using the base_path for .exe file
logo_path = os.path.join(base_path, "resources", "outlooklogo.png")

# Load and resize the OUTLOOK logo image using Pillow library
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((25, 25), Image.Resampling.LANCZOS)  # Resize the image to fit nicely in the footer of the main window
logo_image_tk = ImageTk.PhotoImage(logo_image)  # Convert the image to a format Tkinter can use (compatibility)

# Add the logo to the footer (without Pillow library)
logo_label = tk.Label(footer_frame, image=logo_image_tk, bg="#1C2541")  # Use logo_image_tk
logo_label.image = logo_image_tk  # Keep a reference to avoid garbage collection
logo_label.pack(side='left', padx=3)

# Add the contact information text next to the logo
footer_font = ("Helvetica", 10, 'bold')
footer_label = tk.Label(footer_frame, text="For support, contact: hossein.bolouri@st.com", font=footer_font, fg="white", bg="#1C2541")
footer_label.pack(side='left')

# Load the resized ST logo image
logo_path = os.path.join(base_path, "resources", "STlogo.gif_min.gif")
logo_image = tk.PhotoImage(file=logo_path)

# Create a label to display the logo in the bottom right corner
logo_label = tk.Label(frame2, image=logo_image, bg="#1C2541")
logo_label.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)  # Adjust x and y for padding and image adjustment

# Function to show AMAZON Excel Macro window and message
def show_excel_macro_message():
    user_response = messagebox.askokcancel("Tip", "Upload your original STDF file in (.xlsx) format and then run the Macro by pressing ctrl+m.")
    if user_response:  # Check if the user clicked "OK"
        macro_path = os.path.join(base_path, "resources", "Reliability_Makalu_Macro.xlsm")
        os.startfile(macro_path)  # Open the macro file

# Sub-options for Reliability Analysis (Excel Macro)
reliability_sub_button1 = tk.Button(reliability_sub_option_frame, text="Excel Macro", command=show_excel_macro_message, **italic_button_style, width=half_button_width)
reliability_sub_button1.pack(side='top', pady=5)

# Function to show IRIS Excel Macro window and message
def iris_show_excel_macro_message():
    user_response = messagebox.askokcancel("Tip", "Upload your original STDF file in (.xlsx) format and then run the Macro by pressing ctrl+m.")
    if user_response:  # Check if the user clicked "OK"
        macro_path = os.path.join(base_path, "resources", "CIENA_IRIS_HOSSEIN_TOOL.xlsm")
        os.startfile(macro_path)  # Open the macro file

# Sub-option button for IRIS Reliability Analysis (Excel Macro)
iris_reliability_sub_button1 = tk.Button(ciena_reliability_sub_option_frame, text="Excel Macro", command=iris_show_excel_macro_message, **italic_button_style, width=half_button_width)
iris_reliability_sub_button1.pack(side='top', pady=5)

# Function to open the IRIS Failure Detection window for CIENA
def iris_open_failure_detection_window():
    iris_failure_window = tk.Toplevel(root)
    iris_failure_window.title("CIENA")
    iris_failure_window.geometry("900x800")
    iris_failure_window.configure(bg="#1C2541")

    # Set the icon for the CIENA window
    ciena_icon_path = os.path.join(base_path, "resources", "CIENA.ico")
    iris_failure_window.iconbitmap(ciena_icon_path)
    
# Create a canvas and both scrollbars
    canvas = tk.Canvas(iris_failure_window, bg="#1C2541", height=280, width=380)
    v_scrollbar = tk.Scrollbar(iris_failure_window, orient="vertical", command=canvas.yview, width=14)
    h_scrollbar = tk.Scrollbar(iris_failure_window, orient="horizontal", command=canvas.xview, width=14)

    # Create a frame inside the canvas
    scrollable_frame = tk.Frame(canvas, bg="#1C2541")

    # Bind the scrollbars to the canvas
    canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

    # Add the scrollable frame to the canvas
    scrollable_frame_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # After creating scrollable_frame and scrollable_frame_id
    def resize_scrollable_frame(event):
        # Set the width of the scrollable_frame to match the canvas
        canvas.itemconfig(scrollable_frame_id)

    canvas.bind("<Configure>", resize_scrollable_frame)
    
    # Configure the scrollable frame to update the canvas scroll region
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    # Grid the canvas and scrollbars
    canvas.grid(row=0, column=0, sticky="nsew")
    v_scrollbar.grid(row=0, column=1, sticky="ns", padx=1)
    h_scrollbar.grid(row=1, column=0, sticky="ew", padx=1)

    # Configure grid weights for resizing
    iris_failure_window.grid_rowconfigure(0, weight=1)
    iris_failure_window.grid_columnconfigure(0, weight=1)

    # Bind the mouse wheel event to the canvas for vertical scrolling
    def on_mouse_wheel(event):
        direction = -1 if event.delta > 0 else 1
        canvas.yview_scroll(direction, "units")

    # Bind the mouse wheel event to the canvas for horizontal scrolling (Shift+Wheel)
    def on_shift_mouse_wheel(event):
        direction = -1 if event.delta > 0 else 1
        canvas.xview_scroll(direction, "units")

    # Bind the mouse wheel events to the IRIS window and its components
    iris_failure_window.bind("<MouseWheel>", on_mouse_wheel)
    canvas.bind("<MouseWheel>", on_mouse_wheel)
    scrollable_frame.bind("<MouseWheel>", on_mouse_wheel)
    v_scrollbar.bind("<MouseWheel>", on_mouse_wheel)

    iris_failure_window.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)
    canvas.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)
    scrollable_frame.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)
    h_scrollbar.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)

    # Configure the IRIS window grid
    iris_failure_window.grid_rowconfigure(0, weight=1)
    iris_failure_window.grid_columnconfigure(0, weight=1)

    
    ##### Functions for data analysis of IRIS
    
    # Function to upload the file for IRIS product
    def upload_file_iris(app_data):
        try:
            # Open file dialog to select the file
            file_path = filedialog.askopenfilename(parent=scrollable_frame, filetypes=[("Excel files", "*.xlsx")])

            if file_path:
                # Update the entry field directly
                entry_field.delete(0, tk.END)
                entry_field.insert(0, file_path)

                # Initialize variables before processing the file
                app_data.df_testing_24h_iris = None
                app_data.df_v1_read_point_24h_iris = None
                app_data.df_general_24h_iris = None
                app_data.df_analysis_24h_iris = None

                # Reset flags to ensure proper workflow
                app_data.overview_generated_iris = False
                app_data.graph_24h_generated_iris = False

                # Create a progress bar window
                progress_window = tk.Toplevel(scrollable_frame)
                progress_window.title("Uploading File")
                progress_window.iconbitmap(os.path.join(base_path, "resources", "CIENA.ico"))
                progress_window.geometry("400x200")
                progress_window.configure(bg="#1C2541")

                # Create a frame for the logo and title at the top
                top_frame = tk.Frame(progress_window, bg="#1C2541")
                top_frame.pack(pady=(10, 0))

                # Add CIENA logo at the top of the progress window
                ciena_logo_path = os.path.join(base_path, "resources", "CIENA.ico")
                ciena_logo_img = Image.open(ciena_logo_path)
                ciena_logo_img = ciena_logo_img.resize((48, 48), Image.Resampling.LANCZOS)
                ciena_logo_tk = ImageTk.PhotoImage(ciena_logo_img)
                logo_label = tk.Label(progress_window, image=ciena_logo_tk, bg="#1C2541")
                logo_label.image = ciena_logo_tk  # Keep reference to avoid garbage collection
                logo_label.pack(pady=(10, 0))

                progress_label = tk.Label(progress_window, text="Uploading file, please wait...", bg="#1C2541", fg="white", font=("Helvetica", 12))
                progress_label.pack(pady=10)

                progress_bar = ttk.Progressbar(progress_window, orient="horizontal", mode="determinate", length=250)
                progress_bar.pack(pady=10)

                def simulate_progress(step=1):
                    if step <= 100:
                        progress_bar["value"] = step
                        progress_window.after(20, lambda: simulate_progress(step + 1))
                    else:
                        process_file()

                def process_file():
                    try:
                        # Perform file processing
                        df_read_point_24h_iris = pd.read_excel(file_path)
                        app_data.df_v1_read_point_24h_iris = df_read_point_24h_iris.iloc[:, [0, 1, 2, 9, 10, 11]]
                        app_data.df_testing_24h_iris = df_read_point_24h_iris.iloc[:, 27:]
                        app_data.df_general_24h_iris = pd.concat([app_data.df_v1_read_point_24h_iris, app_data.df_testing_24h_iris], axis=1)
                        app_data.df_analysis_24h_iris = app_data.df_general_24h_iris.copy()
                        app_data.file_uploaded_iris = True

                        # Close the progress bar window
                        progress_window.destroy()
                        messagebox.showinfo("Success", "File uploaded successfully!", parent=scrollable_frame)
                    except Exception as e:
                        progress_window.destroy()
                        messagebox.showerror("Error", f"File upload failed. Please ensure you upload the correct output file, Data Analysis sheet, generated after running the Macro. The file must be in '.xlsx' format and contain the required data structure.", parent=scrollable_frame)

                simulate_progress()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function to retrieve the test name by inserting the relative test number for IRIS
    def get_test_names_iris(app_data):
        
        if not app_data.file_uploaded_iris:
            messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
            return

        test_numbers_input = test_number_field.get().strip()
        if not test_numbers_input:  # Check if the field is empty
            messagebox.showerror("Error", "Please enter test number(s) separated by commas.", parent=scrollable_frame)
            return

        try:
            # Convert test numbers to integers
            test_numbers = [int(num.strip()) for num in test_numbers_input.split(',')]

            # Extract the test names and include the test numbers
            test_names = app_data.df_testing_24h_iris.loc[0, test_numbers]

            # Combine test numbers and test names into a formatted string
            test_name_output = "Test Number, Test Name:\n"
            for test_number, test_name in zip(test_numbers, test_names):
                test_name_output += f"{test_number}, {test_name}\n"

            # Prompt the user to select a location to save the .txt file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")],
                title="Save Test Names As",
                parent=scrollable_frame
            )

            if save_path:
                # Save the test numbers and names to the .txt file
                with open(save_path, 'w') as file:
                    file.write(test_name_output)

                # Notify the user that the file has been saved
                messagebox.showinfo("Success", f"Test name text file (.txt) saved successfully at this path: {save_path}", parent=scrollable_frame)

                # Automatically open the .txt file
                os.startfile(save_path)

        except KeyError:
            messagebox.showerror("Error", "One or more test number(s) do not exist in the STDF file. Ensure you insert correct test number(s).", parent=scrollable_frame)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid test numbers separated by commas.", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function to create general overview DataFrame and Excel file for IRIS
    def generate_overview_iris(app_data):
        if not app_data.file_uploaded_iris:
            messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
            return

        try:
            # Validate that the dataframes are not empty
            if app_data.df_v1_read_point_24h_iris is None or app_data.df_testing_24h_iris is None:
                raise ValueError("The uploaded file does not contain valid data.")

            # Check if both fields are empty
            read_point = read_point_field.get().strip()
            sheet_name = sheet_name_field.get().strip()

            # Handle the three scenarios
            if not read_point and not sheet_name:  # If both fields are empty
                messagebox.showerror("Error", "Please fill in the related blank fields to generate the Overview Excel file.", parent=scrollable_frame)
                return
            
            if not read_point:
                messagebox.showerror("Error", "Please enter a number (integer), indicating read point number of the analysis.", parent=scrollable_frame)    
                return
            
            # Ensure the sheet name is provided
            if not sheet_name:  # If the sheet name is empty
                messagebox.showerror("Error", "Please provide a sheet name to generate the Overview Excel file.", parent=scrollable_frame)
                return

            # Combine dataframes and generate overview
            app_data.df_analysis_24h_iris = pd.concat([app_data.df_v1_read_point_24h_iris, app_data.df_testing_24h_iris], axis=1)
            app_data.df_general_24h_iris = app_data.df_analysis_24h_iris.copy().iloc[3:, :]  # Update df_general_24h_iris globally

            # Set the overview_generated_iris flag to True
            app_data.overview_generated_iris = True

            # Prompt user to select save location
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], parent=scrollable_frame)
            if not save_path:
                return

            # Create a progress bar window
            progress_window = tk.Toplevel(scrollable_frame)
            progress_window.title("Saving Overview")
            progress_window.iconbitmap(os.path.join(base_path, "resources", "CIENA.ico"))
            progress_window.geometry("400x200")
            progress_window.configure(bg="#1C2541")

            # Create a frame for the logo and title at the top
            top_frame = tk.Frame(progress_window, bg="#1C2541")
            top_frame.pack(pady=(10, 0))
            
            # Add CIENA logo at the top of the progress window
            ciena_logo_path = os.path.join(base_path, "resources", "CIENA.ico")
            ciena_logo_img = Image.open(ciena_logo_path)
            ciena_logo_img = ciena_logo_img.resize((48, 48), Image.Resampling.LANCZOS)
            ciena_logo_tk = ImageTk.PhotoImage(ciena_logo_img)
            logo_label = tk.Label(progress_window, image=ciena_logo_tk, bg="#1C2541")
            logo_label.image = ciena_logo_tk  # Keep reference to avoid garbage collection
            logo_label.pack(pady=(10, 0))
            
            progress_label = tk.Label(progress_window, text="Saving the file, please wait...", bg="#1C2541", fg="white", font=("Helvetica", 12))
            progress_label.pack(pady=10)

            progress_bar = ttk.Progressbar(progress_window, orient="horizontal", mode="determinate", length=250)
            progress_bar.pack(pady=10)

            def simulate_progress(step=1):
                if step <= 100:
                    progress_bar["value"] = step
                    progress_window.after(20, lambda: simulate_progress(step + 1))
                else:
                    save_overview()

            def save_overview():
                try:
                    # Save the Excel file
                    app_data.df_general_24h_iris.to_excel(save_path, sheet_name=sheet_name, index=False)
                    messagebox.showinfo("Success", f"Overview Excel file saved successfully at this path: {save_path}", parent=scrollable_frame)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save the Overview Excel file: {str(e)}", parent=scrollable_frame)
                finally:
                    # Destroy the progress bar window
                    progress_window.destroy()

            # Start the progress simulation
            simulate_progress()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function definition for generating graph table dataset (preview) for IRIS
    def generate_graph_table_preview_iris(app_data):
        try:
            # Ensure if the Excel file uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return
            
            # Ensure if the user has already created the Overview dataset of the STDF file
            if app_data.df_general_24h_iris is None or not hasattr(app_data, "overview_generated_iris") or not app_data.overview_generated_iris:
                messagebox.showerror("Error", "Please complete the Overview Generation step first.", parent=scrollable_frame)
                return

            # Validate test number input
            test_number = graph_table_test_field.get()
            if not test_number.strip():  # Check if the field is empty or contains only whitespace
                messagebox.showerror("Error", "Please fill in the test number field with valid test number(s), separated by commas.", parent=scrollable_frame)
                return

            # Parse test numbers into a list of integers
            test_numbers = [int(num.strip()) for num in test_number.split(',')]

            # Validate sheet name input and its related error handling part
            sheet_name = graph_table_sheet_field.get()
            if not sheet_name.strip():  # Check if the sheet name field is empty or contains only whitespace
                messagebox.showerror("Error", "Please specify the sheet name for the Graph Table Excel file to be generated.", parent=scrollable_frame)
                return

            # Generate the required dataframes
            app_data.df_general_24h_help_1_iris = app_data.df_general_24h_iris.iloc[2:, :]  # Slice the dataframe
            app_data.df_general_24h_help_2_iris = app_data.df_general_24h_help_1_iris.iloc[:, :6]

            # Validate that all test numbers exist in the dataframe's columns
            invalid_test_numbers = [num for num in test_numbers if num not in app_data.df_general_24h_help_1_iris.columns]
            if invalid_test_numbers:
                raise KeyError(f"Invalid test numbers: {', '.join(map(str, invalid_test_numbers))}")

            # Extract the corresponding columns for all valid test numbers
            app_data.df_general_24h_help_3_iris = app_data.df_general_24h_help_1_iris.loc[:, test_numbers]

            # Combine dataframes and add the "Hatrick" column
            app_data.graph_24h_iris = pd.concat([app_data.df_general_24h_help_2_iris, app_data.df_general_24h_help_3_iris], axis=1)
            app_data.graph_24h_iris['Hatrick'] = app_data.graph_24h_iris[15303006].astype(str) + '_' + app_data.graph_24h_iris[15303007].astype(str) + '_' + app_data.graph_24h_iris[15303008].astype(str)

            # Set the flag to indicate that the graph_24h dataset has been generated
            app_data.graph_24h_generated_iris = True

            # Prompt user to select save location
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], parent=scrollable_frame)
            if save_path:
                app_data.graph_24h_iris.to_excel(save_path, sheet_name=sheet_name, index=False)
                messagebox.showinfo("Success", f"Graph Table Excel file saved successfully at this path: {save_path}", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", "You inserted invalid test number(s).", parent=scrollable_frame)
        except ValueError as e:
            messagebox.showerror("Error", f"{str(e)}", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate Graph Table Excel File.", parent=scrollable_frame)
    
    # Define the function to generate a graph with limits
    def generate_complete_table_iris(app_data):
        try:
            # Check if the file has been uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Get the test numbers from the first entry field of the "Graph Table" label
            test_numbers_input = graph_table_test_field.get().strip()
            if not test_numbers_input:  # Check if the test_number entry field is empty
                messagebox.showerror("Error", "Please fill in the test number entry field first.", parent=scrollable_frame)
                return

            # Check if test number and sheet name entry fields are empty or not
            if not test_numbers_input and not sheet_name:
                messagebox.showerror("Error", "Please fill in all required entry fields first.", parent=scrollable_frame)
                return

            # Validate sheet name input
            sheet_name = graph_table_sheet_field.get().strip()
            if not sheet_name:  # Check if the sheet name field is empty to show an error message
                messagebox.showerror("Error", "Please specify the sheet name for the Complete Table Excel file to be generated.", parent=scrollable_frame)
                return
            
            # Parse the test numbers into a list of integers
            test_numbers = [int(num.strip()) for num in test_numbers_input.split(',')]
            
            # Generate the required dataframes
            df_limit_24h_help_1_iris = pd.concat([app_data.df_analysis_24h_iris.iloc[[0]], app_data.df_analysis_24h_iris.iloc[3:]])
            df_limit_24h_help_1_iris.at[0, 'PID_Number'] = 'Test_Name'
            df_limit_24h_help_1_iris.at[0, 15303006] = ''
            df_limit_24h_help_1_iris.at[0, 15303007] = ''
            df_limit_24h_help_1_iris.at[0, 15303008] = ''
            df_limit_24h_help_2_iris = df_limit_24h_help_1_iris.reset_index(drop=True)
            df_limit_24h_help_3_iris = df_limit_24h_help_2_iris.iloc[:, :6]
            df_limit_24h_help_3_iris['Hatrick'] = (
                df_limit_24h_help_3_iris[15303006].astype(str) + '_' +
                df_limit_24h_help_3_iris[15303007].astype(str) + '_' +
                df_limit_24h_help_3_iris[15303008].astype(str)
            )
            df_limit_24h_help_3_iris.at[0, 'Hatrick'] = ''
            df_limit_24h_help_3_iris.at[1, 'Hatrick'] = ''
            df_limit_24h_help_3_iris.at[2, 'Hatrick'] = ''

            # Select the test numbers dynamically based on user input
            df_limit_24h_help_4_iris = df_limit_24h_help_2_iris.loc[:, test_numbers]

            # Combine the dataframes
            app_data.df_limit_24h_iris = pd.concat([df_limit_24h_help_3_iris, df_limit_24h_help_4_iris], axis=1)

            # Set the flag for df_limit_24h_iris_generated to True
            app_data.df_limit_24h_iris_generated = True
                        
            # Open a "Save As" dialog to let the user specify the file name and location
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Complete Table As",
                parent=scrollable_frame
            )
            if not save_path:  # If the user cancels the save dialog
                return

            # Save the file to the specified path
            app_data.df_limit_24h_iris.to_excel(save_path, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", f"Complete Table Excel file (with limitation) saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", "You inserted invalid test number(s).", parent=scrollable_frame)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid test numbers separated by commas.", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function for 'Limit' button, generting a dataset with test names and limits for iris CIENA (HighL and LowL)
    def limit_with_test_name_iris(app_data):
        try:
            # Ensure the Excel file is uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return

            # Ensure the Complete Table is generated
            if app_data.df_limit_24h_iris is None or not getattr(app_data, "df_limit_24h_iris_generated", False):
                messagebox.showerror("Error", "Please generate the Complete Table step first to be able to create 'Analysis' dataset.", parent=scrollable_frame)
                return

            # Validate test number input
            test_number = graph_table_test_field.get()
            if not test_number.strip():
                messagebox.showerror("Error", "Please fill in the test number entry field with valid test number(s), separated by commas.", parent=scrollable_frame)
                return

            # Parse test numbers into a list of integers
            test_numbers = [int(num.strip()) for num in test_number.split(',')]

            # Validate sheet name input
            sheet_name = graph_table_sheet_field.get()
            if not sheet_name.strip():
                messagebox.showerror("Error", "Please specify the sheet name for the Limit Excel file with only limits and test names to be generated.", parent=scrollable_frame)
                return

            # Get the read point number (number_h) from the General View entry field
            read_point = read_point_field.get().strip()
            if not read_point:
                messagebox.showerror("Error", "Please enter the read point number (only an integer) in the 'General View' section.", parent=scrollable_frame)
                return

            # Data manipulation using Pandas
            complete_table = app_data.df_limit_24h_iris.copy() # To avoid change in this dataset (df_limit_24h_iris)
            complete_table.at[0, 'Hatrick'] = 'Test_Name'
            complete_table.at[1, 'Hatrick'] = 'HighL'
            complete_table.at[2, 'Hatrick'] = 'LowL'
            start_idx = complete_table.columns.get_loc('Hatrick')
            filtered = complete_table.loc[:, complete_table.columns[start_idx:]]
            filtered_limit = filtered[filtered['Hatrick'].isin(['Test_Name', 'HighL', 'LowL'])]

            # Rename columns after 'Hatrick' column
            new_columns = list(filtered_limit.columns)
            for i, col in enumerate(new_columns):
                if i == 0:
                    continue  # 'Hatrick' column
                try:
                    col_int = int(col)
                    new_columns[i] = f"{col_int}_{read_point}h"
                except Exception:
                    pass  # leave as is if not integer
            filtered_limit.columns = new_columns

            # Cache the analysis table
            if not hasattr(app_data, "analysis_point_graphs_iris"):
                app_data.analysis_point_graphs_iris = []
            app_data.analysis_point_graphs_iris.append(filtered_limit)

            # Inform the user about the creation of the analysis table
            analysis_number = len(app_data.analysis_point_graphs_iris)
            messagebox.showinfo(
                "Success",
                f"{analysis_number} Limit Table(s) with only test names and limits either HighL or LowL are created.",
                parent=scrollable_frame
            )

            # Ask the user if they want to save the Excel file
            user_response = messagebox.askyesno(
                "Save Excel File",
                "Do you want to generate the Excel file including only the test names and limits?",
                parent=scrollable_frame
            )
            if user_response:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Analysis Table As",
                    parent=scrollable_frame
                )
                if save_path:
                    filtered_limit.to_excel(save_path, sheet_name=sheet_name, index=False)
                    messagebox.showinfo("Success", f"Excel file (with only limitation and test names) saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", "You inserted invalid test number(s). Make sure you select the same test number(s) you have already used to generate 'Complete Table' and 'Graph Table'.", parent=scrollable_frame)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid test numbers separated by commas.", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    
    # Function for concatenation of the tables with LowL, HighL and Test_Names for all read points
    def full_analysis(app_data):
        try:
            # Ensure we have enough analysis_point_graphs_iris datasets to merge
            if not app_data.analysis_point_graphs_iris or len(app_data.analysis_point_graphs_iris) < 2:
                raise ValueError("Please generate at least two Limit Tables or datasets before creating the Full Limit Excel file.") 

            # Merging DataFrames
            merged = app_data.analysis_point_graphs_iris[0]
            for df in app_data.analysis_point_graphs_iris[1:]:
                merged = merged.merge(df, on="Hatrick", how="inner")

            # Cache (store) the generated DataFrame in a new attribute
            app_data.analysis_point_graphs_iris_merged = merged
            
            # Check to see if the merged dataset has been created 
            app_data.full_limit_generated = True
            
            # Ask the user to save the merged file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Full Limit Table As",
                parent=scrollable_frame
            )
            if save_path:
                merged.to_excel(save_path, index=False, sheet_name='full_limit')
                messagebox.showinfo("Success", f"Full Limit Table Excel file saved successfully at this path: {save_path}.", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"{str(e)}", parent=scrollable_frame)
    
    # Function for graph with HighL and LowL
    def plot_full_graph_iris(app_data):
        try:
            # 1. Check if excessive units dataset is generated
            if not getattr(app_data, "excessive_units_iris_generated", False):
                messagebox.showerror(
                    "Error",
                    "You need to first generate the excessive units dataset to be able to create the graph with LowL and HighL.",
                    parent=scrollable_frame
                )
                return

            # 2. Check if full analysis dataset that includes HighL, LowL and Test_Names is generated that include all read points with HighL and LowL
            merged = getattr(app_data, "analysis_point_graphs_iris_merged", None)
            if merged is None or isinstance(merged, bool):
                messagebox.showerror(
                    "Error",
                    "Please generate the 'Full Limit' dataset including the LowL and HighL and test name sections.",
                    parent=scrollable_frame
                )
                return

            # 3. Get user inputs from the entry fields
            test_number = final_chart_test_field.get().strip()
            colors = final_chart_color_field.get().strip()
            labels = final_chart_read_points_field.get().strip()
            title = final_chart_title_field.get().strip()
            ylabel = final_chart_ylabel_field.get().strip()
            width = final_chart_width_field.get().strip()
            height = final_chart_height_field.get().strip()

            # Check if any field is empty
            if not all([test_number, colors, labels, title, ylabel, width, height]):
                messagebox.showerror("Error", "Please ensure all required entry fields are filled before proceeding.", parent=scrollable_frame)
                return

            # Validate width and height as positive numbers
            try:
                width = int(width)
                height = int(height)
                if width <= 0 or height <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid positive number to specify the width and height of your line chart.", parent=scrollable_frame)
                return

            # 4. Subset merged DataFrame for the test number columns (format: '2010_24h', etc.)
            test_number_str = str(test_number)
            matching_cols = [col for col in merged.columns if col.startswith(test_number_str + '_')]
            if not matching_cols:
                messagebox.showerror("Input Error", f"No columns found for test number {test_number}.", parent=scrollable_frame)
                return

            # Validate if the number of colors matches the number of matching columns
            color_list = [c.strip() for c in colors.split(',')]
            if len(color_list) != len(matching_cols):
                messagebox.showerror("Input Error", "The number of colors must match the number of read points.", parent=scrollable_frame)
                return

            # Validate if the number of labels matches the number of matching columns
            label_list = [l.strip() for l in labels.split(',')]
            if len(label_list) != len(matching_cols):
                messagebox.showerror("Error", "The number of labels must match the number of read points.", parent=scrollable_frame)
                return

            # Always include 'Hatrick' and any other needed columns
            subset_limit = merged[['Hatrick'] + matching_cols].copy()

            # 5. Merge vertically with excessive units DataFrame
            excessive_df = getattr(app_data, "excessive_units_iris", None)
            if excessive_df is None:
                messagebox.showerror("Error", "Excessive units dataset not found.", parent=scrollable_frame)
                return

            # Align columns for concat (should have same columns)
            subset_limit = subset_limit.reset_index(drop=True)
            excessive_df = excessive_df.reset_index(drop=True)
            subset_limit = subset_limit.loc[:, excessive_df.columns]

            # Concatenate vertically
            charting_limit = pd.concat([subset_limit, excessive_df], axis=0, ignore_index=True)

            # 6. Find the column with the highest hour value (e.g., '2010_1000h')
            def extract_hour(col):
                try:
                    return int(col.split('_')[1].replace('h', ''))
                except Exception:
                    return -1

            highest_column = max(matching_cols, key=extract_hour)

            # 7. Filter for plotting: exclude 'Test_Name' and rows with NaN in the highest hour column
            plot_df = charting_limit[
                (charting_limit['Hatrick'] != 'Test_Name') &
                (charting_limit[highest_column].notna())
            ].copy()

            # Falg to check for the plot_df
            app_data.plot_df = plot_df
            app_data.plot_df_generated = True
            
            # 8. Prepare for plotting
            plt.figure(figsize=(width, height))
            for col, color, label in zip(matching_cols, color_list, label_list):
                # Exclude HighL and LowL from the x-axis
                mask = ~plot_df['Hatrick'].isin(['HighL', 'LowL'])
                x = plot_df.loc[mask, 'Hatrick']
                y = pd.to_numeric(plot_df.loc[mask, col].replace(' ', np.nan), errors='coerce')
                plt.plot(x, y, marker='o', linestyle='-', color=color, label=label)

            # Plot HighL and LowL as horizontal lines (once per column, across all Hatrick values)
            for col in matching_cols:
                highl = charting_limit.loc[charting_limit['Hatrick'] == 'HighL', col].replace(' ', np.nan)
                lowl = charting_limit.loc[charting_limit['Hatrick'] == 'LowL', col].replace(' ', np.nan)
                if not highl.empty and pd.notna(highl.values[0]):
                    plt.axhline(y=float(highl.values[0]), color='red', linestyle='--', label=f'HighL {col}')
                if not lowl.empty and pd.notna(lowl.values[0]):
                    plt.axhline(y=float(lowl.values[0]), color='darkred', linestyle='--', label=f'LowL {col}')

            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Save as PDF
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Full Graph As PDF",
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"The line graph of all read points, including HighL and LowL saved successfully at this path: {save_path}", parent=scrollable_frame)
            plt.close()

            # Ask user to save the full dataset (with Test_Name, HighL, LowL)
            user_response = messagebox.askyesno(
                "Save Full Dataset",
                "Do you want to save the full dataset (with Test_Name, HighL, LowL, and all Hatricks) as an Excel file?",
                parent=scrollable_frame
            )
            if user_response:
                excel_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Full Dataset As Excel",
                    parent=scrollable_frame
                )
                if excel_path:
                    charting_limit.to_excel(excel_path, index=False, sheet_name='final')
                    messagebox.showinfo("Success", f"Excel file (full dataset of all read points) saved successfully at this path: {excel_path}", parent=scrollable_frame)

            # Ask user to save the dataset without the Test_Name row
            user_response_plot = messagebox.askyesno(
                "Save Dataset Without Test_Name",
                "Do you want to save the dataset without the Test_Name row (includes HighL, LowL, and all Hatrick data rows) as an Excel file?",
                parent=scrollable_frame
            )
            if user_response_plot:
                excel_path_plot = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Dataset Without Test_Name As Excel",
                    parent=scrollable_frame
                )
                if excel_path_plot:
                    plot_df.to_excel(excel_path_plot, index=False, sheet_name='final')
                    messagebox.showinfo("Success", f"Excel file (without Test_Name row) saved successfully at this path: {excel_path_plot}", parent=scrollable_frame)

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Function for Read Point Table for IRIS
    def generate_read_point_graph_iris(app_data):
        try:
            # Ensure the General View and Graph Table sections are completed
            if not app_data.file_uploaded_iris or app_data.df_general_24h_iris is None or not app_data.overview_generated_iris or not app_data.graph_24h_generated_iris:
                raise ValueError("Please ensure the Excel file is uploaded and both the 'Overview Generation' and 'Graph Table' steps are completed before proceeding.")

            # Get the test numbers from the Graph Table entry field
            test_numbers_input = graph_table_test_field.get().strip()
            if not test_numbers_input:
                raise ValueError("Please enter test numbers in the Graph Table section.")

            # Convert test numbers to strings
            test_numbers = [int(num.strip()) for num in test_numbers_input.split(',')]
            test_numbers_str = [str(num) for num in test_numbers]

            # Get the read point number from the General View entry field
            read_point = read_point_field.get().strip()
            if not read_point:
                raise ValueError("Please enter the read point number in the General View section.")

            # Create a copy of graph_24h_iris and rename its columns
            graph_read_point_help_iris = app_data.graph_24h_iris.copy()
            graph_read_point_help_iris.columns = graph_read_point_help_iris.columns.astype(str)  # Convert column names to strings
            graph_read_point_iris = graph_read_point_help_iris.rename(columns={num: f"{num}_{read_point}h" for num in test_numbers_str})

            # Cache the graph_read_point with a unique name
            if not hasattr(app_data, "read_point_graphs_iris"):
                app_data.read_point_graphs_iris = []
            app_data.read_point_graphs_iris.append(graph_read_point_iris)

            # Inform the user about the creation of the graph
            graph_number = len(app_data.read_point_graphs_iris)
            messagebox.showinfo(
                "Success",
                f"{graph_number} Read Point Graph Table(s) is (are) created.",
                parent=scrollable_frame
            )

            # Ask the user if they want to save the Excel file
            user_response = messagebox.askyesno(
                "Save Excel File",
                "Do you want to generate the Excel file for the single Read Point?",
                parent=scrollable_frame
            )
            if user_response:
                # Prompt the user to save the Excel file
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Read Point Graph As",
                    parent=scrollable_frame
                )
                if save_path:
                    # Save the DataFrame to Excel with the sheet name "Read_Point"
                    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                        graph_read_point_iris.to_excel(writer, index=False, sheet_name="Read_Point")
                    messagebox.showinfo("Success", f"Read Point Excel file saved successfully at this path: {save_path}", parent=scrollable_frame)

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", "Please ensure both the 'General View' and 'Graph Table' Excel files are generated before proceeding.", parent=scrollable_frame)
    
    # Function to plot the line graph with limits of a single read point for IRIS
    def plot_line_limit_chart_iris(app_data):
        try:
            # Check if the Excel file has been uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Check if the complete table with limits is generated
            if not getattr(app_data, "df_limit_24h_iris_generated", False):
                messagebox.showerror(
                    "Error",
                    "Please first generate the complete table that includes the limits to be able to have a line graph with HighL and LowL for a single read point.",
                    parent=scrollable_frame
                )
                return

            # Retrieve all field values
            test_number_input = line_chart_test_field.get().strip()
            read_point = read_point_field.get().strip()
            color = line_chart_color_field.get().strip()
            legend = line_chart_legend_field.get().strip()
            title = line_chart_title_field.get().strip()
            ylabel = line_chart_ylabel_field.get().strip()
            width_input = line_chart_width_field.get().strip()
            height_input = line_chart_height_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, read_point, color, legend, title, ylabel, width_input, height_input]):
                messagebox.showerror("Error", "Please fill all required entry fields before proceeding.", parent=scrollable_frame)
                return

            # Validate integer fields
            try:
                test_number = int(test_number_input)
                width = int(width_input)
                height = int(height_input)
            except ValueError:
                messagebox.showerror("Error", "Test number, width, and height must be valid integers.", parent=scrollable_frame)
                return

            # Validate width and height as positive numbers
            if width <= 0 or height <= 0:
                messagebox.showerror("Error", "Please enter a valid positive number to specify the width and height of your chart.", parent=scrollable_frame)
                return

            # Additional check: Ensure the test number exists in the dataset for plotting
            df_limit_24h_iris = app_data.df_limit_24h_iris
            if test_number not in df_limit_24h_iris.columns:
                messagebox.showerror(
                    "Error",
                    f"Please use a test number that you have already used to generate the table with limits in step 4 by pressing 'Complete Table' button.",
                    parent=scrollable_frame
                )
                return

            # Data Manipulation and copy creation to avoid the change in the original dataset
            df_single_limit_plot_1 = df_limit_24h_iris.copy()
            df_single_limit_plot_1.at[1, 'Hatrick'] = 'HighL'
            df_single_limit_plot_1.at[2, 'Hatrick'] = 'LowL'

            # Get all columns from 'Hatrick' onwards
            start_col = df_single_limit_plot_1.columns.get_loc('Hatrick')
            df_single_limit_plot_2 = df_single_limit_plot_1.iloc[1:, start_col:].reset_index(drop=True)

            # Replace ' ' with np.nan in the first and second rows (index 0 and 1)
            df_single_limit_plot_2.iloc[0] = df_single_limit_plot_2.iloc[0].replace(' ', np.nan)
            df_single_limit_plot_2.iloc[1] = df_single_limit_plot_2.iloc[1].replace(' ', np.nan)

            # Prepare for plotting
            plt.figure(figsize=(width, height))
            # Exclude HighL and LowL from the x-axis
            mask = ~df_single_limit_plot_2['Hatrick'].isin(['HighL', 'LowL'])
            x = df_single_limit_plot_2.loc[mask, 'Hatrick']
            y = pd.to_numeric(df_single_limit_plot_2.loc[mask, test_number].replace(' ', np.nan), errors='coerce')
            plt.plot(x, y, marker='o', linestyle='-', color=color, label=legend)

            # Plot HighL and LowL as horizontal lines (from the manipulated DataFrame)
            highl = df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'HighL', test_number].replace(' ', np.nan)
            lowl = df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'LowL', test_number].replace(' ', np.nan)
            if not highl.empty and pd.notna(highl.values[0]):
                plt.axhline(y=float(highl.values[0]), color='red', linestyle='--', label=f'HighL {test_number}')
            if not lowl.empty and pd.notna(lowl.values[0]):
                plt.axhline(y=float(lowl.values[0]), color='darkred', linestyle='--', label=f'LowL {test_number}')

            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Save the chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Line chart with HighL and LowL for a single read point saved successfully at this path: {save_path}.", parent=scrollable_frame)
            plt.close()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function for Distribution Plot with Limits and Average
    def dist_iris_single_read_point(app_data):
        try:
            # Check if the Excel file has been uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Check if the complete table with limits is generated
            if not getattr(app_data, "df_limit_24h_iris_generated", False):
                messagebox.showerror(
                    "Error",
                    "Please first generate the complete table that includes the limits to be able to have a distribution chart with HighL and LowL.",
                    parent=scrollable_frame
                )
                return

            # Retrieve all field values
            test_number_input = histogram_test_field.get().strip()
            bins_input = histogram_bins_field.get().strip()
            histogram_title = histogram_title_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, bins_input, histogram_title]):
                messagebox.showerror("Error", "Please fill all required entry fields before proceeding.", parent=scrollable_frame)
                return

            # Validate integer fields
            try:
                test_number = int(test_number_input)
                bins = int(bins_input)
                if bins <= 0 or test_number <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Test number and number of bins must be positive integers.", parent=scrollable_frame)
                return

            # Additional check: Ensure the test number exists in the dataset for plotting
            df_limit_24h_iris = app_data.df_limit_24h_iris
            if test_number not in df_limit_24h_iris.columns:
                messagebox.showerror(
                    "Error",
                    f"Please use a test number that you have already used to generate the table with limits in step 4 by pressing 'Complete Table' button.",
                    parent=scrollable_frame
                )
                return

            # Data Manipulation and copy creation to avoid the change in the original data
            df_single_limit_plot_1 = df_limit_24h_iris.copy()
            df_single_limit_plot_1.at[1, 'Hatrick'] = 'HighL'
            df_single_limit_plot_1.at[2, 'Hatrick'] = 'LowL'

            # Get all columns from 'Hatrick' onwards
            start_col = df_single_limit_plot_1.columns.get_loc('Hatrick')
            df_single_limit_plot_2 = df_single_limit_plot_1.iloc[1:, start_col:].reset_index(drop=True)

            # Replace ' ' with np.nan in the first and second rows (index 0 and 1)
            df_single_limit_plot_2.iloc[0] = df_single_limit_plot_2.iloc[0].replace(' ', np.nan)
            df_single_limit_plot_2.iloc[1] = df_single_limit_plot_2.iloc[1].replace(' ', np.nan)

            # Prepare values for histogram
            mask = ~df_single_limit_plot_2['Hatrick'].isin(['HighL', 'LowL'])
            all_values = pd.to_numeric(df_single_limit_plot_2.loc[mask, test_number], errors='coerce').dropna().tolist()
            highl_values = pd.to_numeric(df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'HighL', test_number], errors='coerce').dropna().tolist()
            lowl_values = pd.to_numeric(df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'LowL', test_number], errors='coerce').dropna().tolist()

            # Plot histogram for all values
            plt.figure(figsize=(10, 6))
            plt.hist(all_values, bins=bins, alpha=0.7, label=histogram_title)

            # Calculate mean (average) excluding HighL and LowL rows
            mean_value = np.mean(all_values)

            # Plot vertical lines for HighL, LowL, and mean
            for hl in highl_values:
                plt.axvline(x=hl, color='red', linestyle='--', label=f'HighL: {hl}')
            for ll in lowl_values:
                plt.axvline(x=ll, color='darkred', linestyle='--', label=f'LowL: {ll}')
            plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')

            plt.title(histogram_title)
            plt.xlabel('Test Value')
            plt.ylabel('Frequency')
            plt.legend()
            plt.tight_layout()

            # Ensure all lines are visible
            all_x = all_values + highl_values + lowl_values + [mean_value]
            plt.xlim(min(all_x) - 1, max(all_x) + 1)

            # Save the chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Distribution chart with HighL and LowL for a single read point saved successfully at this path: {save_path}", parent=scrollable_frame)
            plt.close()

            # Ask the user if they want to generate a chart excluding HighL and LowL
            user_response = messagebox.askyesno(
                "Generate Focused Distribution",
                "HighL and LowL may distort the chart. Would you like to generate a distribution chart excluding HighL and LowL (showing only the average and actual data)?",
                parent=scrollable_frame
            )
            if user_response:
                plt.figure(figsize=(10, 6))
                plt.hist(all_values, bins=bins, alpha=0.7, label=histogram_title + " (No HighL/LowL)")
                plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')
                plt.title(histogram_title + " (No HighL/LowL)")
                plt.xlabel('Test Value')
                plt.ylabel('Frequency')
                plt.legend()
                plt.tight_layout()
                plt.xlim(min(all_values) - 1, max(all_values) + 1)
                save_path_focused = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")],
                    title="Save Focused Distribution Chart As PDF",
                    parent=scrollable_frame
                )
                if save_path_focused:
                    plt.savefig(save_path_focused)
                    messagebox.showinfo("Success", f"Focused distribution chart saved successfully at this path: {save_path_focused}", parent=scrollable_frame)
                plt.close()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot the distribution chart: {str(e)}", parent=scrollable_frame)
    
    # Function to plot the line graph of a single read point
    def plot_line_chart_iris(app_data):
        try:
            # Check if the Excel file has been uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Ensure the graph_24h DataFrame is available
            if app_data.graph_24h_iris is None or app_data.graph_24h_iris.empty or not hasattr(app_data, "graph_24h_generated_iris") or not app_data.graph_24h_generated_iris:
                raise ValueError(
                    "To plot the line chart for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )

            # Retrieve all field values
            test_number_input = line_chart_test_field.get().strip()
            color = line_chart_color_field.get().strip()
            legend = line_chart_legend_field.get().strip()
            title = line_chart_title_field.get().strip()
            ylabel = line_chart_ylabel_field.get().strip()
            width_input = line_chart_width_field.get().strip()
            height_input = line_chart_height_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, color, legend, title, ylabel, width_input, height_input]):
                raise ValueError("Please fill all required entry fields before proceeding.")

            # Validate integer fields
            try:
                test_number = int(test_number_input)
                width = int(width_input)
                height = int(height_input)
            except ValueError:
                raise ValueError("Test number, width, and height must be valid integers.")

            # Validate width and height as positive numbers
            if width <= 0 or height <= 0:
                raise ValueError("Please enter a valid positive number to specify the width and height of your chart.")

            # Ensure the test number exists in the dataset
            if test_number not in app_data.graph_24h_iris.columns:
                raise KeyError("Please enter the test number you selected earlier to generate the 'Graph Table' in Step 4.")

            # Replace spaces (' ') with numpy.nan in the selected column
            app_data.graph_24h_iris[test_number] = app_data.graph_24h_iris[test_number].replace(' ', np.nan).astype(float)

            # Plot the line chart
            plt.figure(figsize=(width, height))
            plt.plot(
                app_data.graph_24h_iris['Hatrick'],
                app_data.graph_24h_iris[test_number],
                marker='o',
                linestyle='-',
                color=color,
                label=legend
            )
            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Save the chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Line chart graph without any limits saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except KeyError as ke:
            messagebox.showerror("Error", str(ke), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function for plotting the histogram for IRIS
    def plot_distribution_chart_iris(app_data):
        try:
            # Ensure if the Excel file is uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return

            # Ensure the graph_24h DataFrame is available and generated in step 4
            if app_data.graph_24h_iris is None or app_data.graph_24h_iris.empty or not hasattr(app_data, "graph_24h_generated_iris") or not app_data.graph_24h_generated_iris:
                raise ValueError(
                    "To plot the distribution chart (histogram) for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )

            # Retrieve all field values
            test_number_input = histogram_test_field.get().strip()
            bins_input = histogram_bins_field.get().strip()
            title = histogram_title_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, bins_input, title]):
                raise ValueError("Please fill all required entry fields before proceeding.")

            # Validate integer fields
            try:
                test_number = int(test_number_input)
                bins = int(bins_input)
                if bins <= 0:
                    raise ValueError("The value entered for the number of bins and test numbers in your distribution chart must be a positive integer. Please correct your input and try again.")
            except ValueError as ve:
                raise ValueError(str(ve))

            # Ensure the test number exists in the dataset
            if test_number not in app_data.graph_24h_iris.columns:
                raise KeyError("Please enter the test number you selected earlier to generate the 'Graph Table' in Step 4.")

            # Replace spaces (' ') with numpy.nan in the selected column
            app_data.graph_24h_iris[test_number] = app_data.graph_24h_iris[test_number].replace(' ', np.nan).astype(float)

            # Plot the distribution chart, skipping NaN values automatically
            plt.figure()
            plt.hist(app_data.graph_24h_iris[test_number].dropna().values, bins=bins, label=f"Test {test_number}", alpha=0.7)
            plt.title(title)
            plt.xlabel("Test Value")
            plt.ylabel("Frequency")
            plt.tight_layout()

            # Save the chart
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], parent=scrollable_frame)
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Distribution chart saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except KeyError as ke:
            messagebox.showerror("Error", str(ke), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot the distribution chart. Please enter the same test numbers used in step 4.", parent=scrollable_frame)
    
    # Function to get test statistics for IRIS
    def generate_test_statistics_iris(app_data):
        try:
            # Ensure if the Excel file is uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return

            # Ensure the graph_24h DataFrame is available and generated in step 4
            if app_data.graph_24h_iris is None or app_data.graph_24h_iris.empty or not hasattr(app_data, "graph_24h_generated_iris") or not app_data.graph_24h_generated_iris:
                raise ValueError(
                    "To have test statistics for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )

            # Retrieve the test number from the entry field
            test_number_input = test_statistics_field.get().strip()

            # Ensure the field is filled
            if not test_number_input:
                raise ValueError("Please fill the required entry field before proceeding.")

            # Validate the test number as an integer
            try:
                test_number = int(test_number_input)
            except ValueError:
                raise ValueError("Test number must be a valid number (integer).")

            # Validate that the test number exists in the DataFrame
            if test_number not in app_data.graph_24h_iris.columns:
                raise KeyError("Please enter the test number you selected earlier to generate the 'Graph Table' in Step 4.")

            # Extract the relevant column and calculate statistics
            real_stat_24h = pd.to_numeric(app_data.graph_24h_iris[test_number].copy(), errors='coerce')
            stats = real_stat_24h.describe()

            # Prompt the user to save the statistics to a .txt file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")],
                title="Save Test Statistics As",
                parent=scrollable_frame
            )

            if save_path:
                # Save the statistics to the .txt file
                with open(save_path, 'w') as file:
                    file.write(f"Statistics for Test Number {test_number}:\n")
                    file.write(stats.to_string())

                # Notify the user that the file has been saved
                messagebox.showinfo("Success", f"Test statistics text file (.txt) saved successfully at this path: {save_path}", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", f"{str(e)}", parent=scrollable_frame)
        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function for failure finding for IRIS product
    def generate_failure_table_iris(app_data):
        try:
            # Ensure if the Excel file is uploaded
            if not app_data.file_uploaded_iris:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return 
            
            # Ensure the graph_24h DataFrame is available and generated in step 4
            if app_data.graph_24h_iris is None or app_data.graph_24h_iris.empty or not hasattr(app_data, "graph_24h_generated_iris") or not app_data.graph_24h_generated_iris:
                raise ValueError(
                "To have units with particular HBIN(s) for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )     
            
            # Ensure the 'HBIN' column exists in the DataFrame
            if 'HBIN' not in app_data.graph_24h_iris.columns:
                raise KeyError("The 'HBIN' column is missing in the STDF dataset.")
            
            # Get the operand and HBIN number from the entry fields
            operand = failure_operand_field.get().strip()
            hbin_number = failure_hbin_field.get().strip()
            
            # Check if both entry fields are empty
            if not operand or not hbin_number:
                raise ValueError("Please ensure both the operand and HBIN number fields are filled before proceeding.")

            # Validate the HBIN number
            if not hbin_number.isdigit():
                raise ValueError("HBIN must be a positive number (integer).")
            hbin_number = int(hbin_number)

            # Validate the operand
            if operand not in ['<', '<=', '>', '>=', '=']:
                raise ValueError("Invalid operand. Please enter one of these operands: <, <=, >, >=, or =.")

            # Convert '=' to '==' for the query
            if operand == '=':
                operand = '=='

            # Dynamically construct the condition based on the operand and HBIN number
            condition = f"HBIN {operand} {hbin_number}"

            # Query the DataFrame
            app_data.failure_24h_iris = app_data.graph_24h_iris.query(condition)

            # Check if the query returned any results
            if app_data.failure_24h_iris.empty:
                raise ValueError(f"No data matches the specified condition: {condition}.")

            # Prompt the user to select a save location
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Failure Table As",
                parent=scrollable_frame
            )

            if save_path:
                # Save the failure table to the specified location
                app_data.failure_24h_iris.to_excel(save_path, sheet_name='Detection', index=False)
                messagebox.showinfo("Success", f"Excel file with specific HBINs saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except ValueError as ve:
            messagebox.showerror("Error", f"{str(ve)}", parent=scrollable_frame)
        except Exception as ve:
            messagebox.showerror("Error", f"{str(ve)}", parent=scrollable_frame)
    
    # Function to generate Final Table Dataset for IRIS
    def generate_final_table_iris(app_data):
        try:
            # Ensure read_point_graphs_iris has enough data
            if not app_data.read_point_graphs_iris or len(app_data.read_point_graphs_iris) < 2:
                raise ValueError("Please generate at least two Read Point Tables or datasets before creating the Final Table Excel file or dataset.")

            # Dynamically select columns for merging based on 'Hatrick' and columns between '15303008' and 'Hatrick'
            merged_table = app_data.read_point_graphs_iris[0][['Hatrick']]
            for graph in app_data.read_point_graphs_iris:
                # Find the positions of "15303008" and "Hatrick" in the columns
                columns = graph.columns.tolist()
                if "15303008" in columns and "Hatrick" in columns:
                    start_index = columns.index("15303008") + 1  # Start after "15303008"
                    end_index = columns.index("Hatrick")  # End before "Hatrick"
                    columns_to_merge = ["Hatrick"] + columns[start_index:end_index]
                else:
                    raise ValueError("Columns '15303008' and 'Hatrick' are not found in the 'read_point_graphs' dataset.")

                # Perform the merge
                merged_table = merged_table.merge(graph[columns_to_merge], on='Hatrick', how='outer')

            # Reset the index for the final table
            merged_table.reset_index(drop=True, inplace=True)

            # Ask the user if they want to save the Excel file
            user_response = messagebox.askyesno(
                "Save Final Table",
                "Do you want to generate the Excel file for the Final Table?",
                parent=scrollable_frame
            )

            if user_response:
                # Prompt the user to save the Excel file
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Final Table As",
                    parent=scrollable_frame
                )
                if not save_path:
                    messagebox.showinfo("Canceled", "The operation was canceled. No Excel file was saved.", parent=scrollable_frame)
                    return
                
                # Save the DataFrame to Excel with the sheet name "Final_Table"
                merged_table.to_excel(save_path, sheet_name="Final_Table", index=False)
                messagebox.showinfo("Success", f"Final Table Excel file with all test numbers and read points saved successfully at this path: {save_path}", parent=scrollable_frame)

            # Cache the final table for later use
            app_data.final_table_iris = merged_table

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function to generate the final graph for IRIS
    def generate_final_graph_iris(app_data):
        try:
            # Ensure the final_table_iris DataFrame is available
            if not hasattr(app_data,'final_table_iris') or app_data.final_table_iris is None:
                raise ValueError("Unable to create the Final Graph with all read points. Please ensure the Final Table is generated first.")

            # Get user inputs from the entry fields
            test_number = final_chart_test_field.get().strip()
            colors = final_chart_color_field.get().strip()
            labels = final_chart_read_points_field.get().strip()
            title = final_chart_title_field.get().strip()
            ylabel = final_chart_ylabel_field.get().strip()
            width = final_chart_width_field.get().strip()
            height = final_chart_height_field.get().strip()

            # Check if any field is empty
            if not all([test_number, colors, labels, title, ylabel, width, height]):
                raise ValueError("Please ensure all required entry fields are filled before proceeding.")

            # Validate width and height as positive numbers
            try:
                width = int(width)
                height = int(height)
                if width <= 0 or height <= 0:
                    raise ValueError
            except ValueError:
                raise ValueError("Please enter a valid positive number to specify the width and height of your chart.")

            # Convert test_number to string for matching column names
            test_number_str = str(test_number)

            # Filter columns that match the test number
            matching_columns = [col for col in app_data.final_table_iris.columns if col.startswith(test_number_str + '_')]

            # Check if the test number exists in the final_table
            if not matching_columns:
                raise ValueError("Ensure that the test number you entered matches one of the test numbers imported during step 4 of the analysis.")

            # Validate the number of colors matches the number of matching columns
            color_list = colors.split(',')
            if len(color_list) != len(matching_columns):
                raise ValueError("The number of colors you specify in the color entry field must match the number of read points you wish to analyze.")

            # Validate the number of labels matches the number of matching columns
            label_list = labels.split(',')
            if len(label_list) != len(matching_columns):
                raise ValueError(f"Number of labels ({len(label_list)}) does not match the number of read points.")

            # Handle empty cells in the matching columns
            for col in matching_columns:
                app_data.final_table_iris[col] = app_data.final_table_iris[col].replace(' ', np.nan).astype(float)

            # Plot the graph
            plt.figure(figsize=(width, height))
            for col, color, label in zip(matching_columns, color_list, label_list):
                plt.plot(app_data.final_table_iris['Hatrick'], app_data.final_table_iris[col], marker='o', linestyle='-', color=color.strip(), label=label.strip())
            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Prompt user to select save location for the PDF file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Final Graph As PDF",
                parent=scrollable_frame
            )

            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Final line graph with all read points saved successfully at this path: {save_path}", parent=scrollable_frame)

            # Close the plot to avoid displaying it
            plt.close()

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function to plot excessive units for IRIS
    def plot_excessive_units_iris(app_data):
        try:
            # Ensure the final_table_iris DataFrame is available
            if not hasattr(app_data, 'final_table_iris') or app_data.final_table_iris is None:
                raise ValueError("Unable to create the Excessive Units Graph. Please ensure the Final Table is generated first.")

            # Get user inputs from the entry fields
            test_number = final_chart_test_field.get().strip()
            colors = final_chart_color_field.get().strip()
            labels = final_chart_read_points_field.get().strip()
            title = final_chart_title_field.get().strip()
            ylabel = final_chart_ylabel_field.get().strip()
            width = final_chart_width_field.get().strip()
            height = final_chart_height_field.get().strip()

            # Check if any field is empty
            if not all([test_number, colors, labels, title, ylabel, width, height]):
                raise ValueError("Please ensure all required entry fields are filled before proceeding.")

            # Validate width and height as positive numbers
            try:
                width = int(width)
                height = int(height)
                if width <= 0 or height <= 0:
                    raise ValueError
            except ValueError:
                raise ValueError("Please enter a valid positive number to specify the width and height of your chart.")

            # Convert test_number to string for matching column names
            test_number_str = str(test_number)

            # Filter columns that match the test number
            matching_columns = [col for col in app_data.final_table_iris.columns if col.startswith(test_number_str + '_')]

            # Check if the test number exists in the final_table_iris
            if not matching_columns:
                raise ValueError("Ensure that the test number you entered matches one of the test numbers imported during step 4 of the analysis.")

            # Validate the number of colors matches the number of matching columns
            color_list = colors.split(',')
            if len(color_list) != len(matching_columns):
                raise ValueError("The number of colors you specify in the color entry field must match the number of read points you wish to analyze.")

            # Validate the number of labels matches the number of matching columns
            label_list = labels.split(',')
            if len(label_list) != len(matching_columns):
                raise ValueError(f"Number of labels ({len(label_list)}) does not match the number of read points.")

            # Handle empty cells in the matching columns
            for col in matching_columns:
                app_data.final_table_iris[col] = app_data.final_table_iris[col].replace(' ', np.nan).astype(float)

            # Identify the column with the highest numeric value before "h"
            highest_column = max(matching_columns, key=lambda col: int(col.split('_')[1].replace('h', '')))

            # Create a copy of the final_table_iris to avoid modifying the original DataFrame
            excessive_units_iris = app_data.final_table_iris.copy()

            # Remove rows with NaN values in the highest_column
            excessive_units_iris = excessive_units_iris.dropna(subset=[highest_column])

            # Flag to check for excessive units is being generated or not
            app_data.excessive_units_iris_generated = True
            app_data.excessive_units_iris = excessive_units_iris[['Hatrick'] + matching_columns]
            
            # Plot the line graph for all matching columns
            plt.figure(figsize=(width, height))
            for col, color, label in zip(matching_columns, color_list, label_list):
                plt.plot(
                    excessive_units_iris['Hatrick'],  # X-axis
                    excessive_units_iris[col],  # Y-axis
                    marker='o',
                    linestyle='-',
                    color=color.strip(),
                    label=label.strip()
                )
            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Ask the user if they want to save the dataset as an Excel file
            user_response = messagebox.askyesno(
                "Save Dataset",
                "Do you want to save the dataset used for this graph as an Excel file?",
                parent=scrollable_frame
            )

            if user_response:
                # Prompt the user to select a save location for the Excel file
                save_excel_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Dataset As",
                    parent=scrollable_frame
                )
                if save_excel_path:
                    # Save the dataset to the specified location
                    excessive_units_iris[['Hatrick'] + matching_columns].to_excel(save_excel_path, index=False, sheet_name='Excessive_Units')
                    messagebox.showinfo("Success", f"Dataset saved successfully at this path: {save_excel_path}", parent=scrollable_frame)
                else:
                    messagebox.showinfo("Canceled", "The operation of creating the Excel file was canceled.", parent=scrollable_frame)

            # Prompt user to select save location for the PDF file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Excessive Units Graph As PDF",
                parent=scrollable_frame
            )

            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Excessive Units graph saved successfully at this path: {save_path}", parent=scrollable_frame)

            # Close the plot to avoid displaying it
            plt.close()

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Function for 'Full Distribution' button for CIENA that includes all read points
    def full_distribution_iris(app_data):
        try:
            # 1. Check if plot_df dataset has been created and is valid
            plot_df = getattr(app_data, "plot_df", None)
            if plot_df is None or plot_df.empty:
                messagebox.showerror(
                    "Error",
                    "Please first complete the 'Full Graph' section by clicking on the corresponding button to be able to have the full distribution chart of all read points.",
                    parent=scrollable_frame
                )
                return

            # 2. Check if all entry fields are filled
            test_number_input = histogram_test_field.get().strip()
            bin_number_input = histogram_bins_field.get().strip()
            histogram_title = histogram_title_field.get().strip()
            if not all([test_number_input, bin_number_input, histogram_title]):
                messagebox.showerror(
                    "Error",
                    "Please fill in all entry fields at step '6' to be able to generate the full distribution chart of all read points.",
                    parent=scrollable_frame
                )
                return

            # 3. Validate integer fields
            try:
                test_number = int(test_number_input)
                bin_number = int(bin_number_input)
                if bin_number <= 0 or test_number <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "Error",
                    "Test number and number of bins must be positive integers.",
                    parent=scrollable_frame
                )
                return

            # 4. Find all columns in plot_df matching the test number (format: "testnumber_#h")
            matching_cols = [col for col in plot_df.columns if col.startswith(f"{test_number}_")]
            if not matching_cols:
                messagebox.showerror(
                    "Error",
                    "Please enter a valid test number you have already used in step '9' to generate the distribution graph of all read points with Test_Name, HighL and LowL (Full Graph With All Data).",
                    parent=scrollable_frame
                )
                return

            # 5. For each matching column, plot its histogram and collect HighL/LowL values
            plt.figure(figsize=(10, 6))
            all_values = []
            highl_values = []
            lowl_values = []
            for col in matching_cols:
                # Exclude rows with 'HighL' and 'LowL' in 'Hatrick'
                mask = ~plot_df['Hatrick'].isin(['HighL', 'LowL'])
                values = pd.to_numeric(plot_df.loc[mask, col], errors='coerce').dropna()
                all_values.extend(values)

                # Get HighL and LowL for this column
                highl = plot_df.loc[plot_df['Hatrick'] == 'HighL', col]
                lowl = plot_df.loc[plot_df['Hatrick'] == 'LowL', col]
                highl_values.extend(pd.to_numeric(highl, errors='coerce').dropna())
                lowl_values.extend(pd.to_numeric(lowl, errors='coerce').dropna())

            # Plot histogram for all values
            plt.hist(all_values, bins=bin_number, alpha=0.7, label=histogram_title)

            # 6. Calculate mean (average) excluding HighL and LowL rows
            mean_value = np.mean(all_values)

            # 7. Plot vertical lines for HighL, LowL, and mean
            for hl in highl_values:
                plt.axvline(x=hl, color='red', linestyle='--', label=f'HighL: {hl}')
            for ll in lowl_values:
                plt.axvline(x=ll, color='darkred', linestyle='--', label=f'LowL: {ll}')
            plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')

            plt.title(histogram_title)
            plt.xlabel('Test Value')
            plt.ylabel('Frequency')
            plt.legend()
            plt.tight_layout()

            # Ensure all lines are visible
            all_x = all_values + highl_values + lowl_values + [mean_value]
            plt.xlim(min(all_x) - 1, max(all_x) + 1)

            # Save the chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Distribution chart with HighL and LowL of all read points saved successfully at this path: {save_path}", parent=scrollable_frame)
            plt.close()

            # Ask the user if they want to generate a chart excluding HighL and LowL
            user_response = messagebox.askyesno(
                "Generate Focused Distribution",
                "HighL and LowL may distort the chart. Would you like to generate a distribution chart of all read points excluding HighL and LowL (showing only the average and actual data)?",
                parent=scrollable_frame
            )
            if user_response:
                plt.figure(figsize=(10, 6))
                plt.hist(all_values, bins=bin_number, alpha=0.7, label=histogram_title + " (No HighL/LowL)")
                plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')
                plt.title(histogram_title + " (No HighL/LowL)")
                plt.xlabel('Test Value')
                plt.ylabel('Frequency')
                plt.legend()
                plt.tight_layout()
                plt.xlim(min(all_values) - 1, max(all_values) + 1)
                save_path_focused = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")],
                    title="Save Focused Distribution Chart As PDF",
                    parent=scrollable_frame
                )
                if save_path_focused:
                    plt.savefig(save_path_focused)
                    messagebox.showinfo("Success", f"Focused distribution chart saved successfully at this path: {save_path_focused}", parent=scrollable_frame)
                plt.close()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot the full distribution chart: {str(e)}", parent=scrollable_frame)


    # Function for 'Clear All' button for IRIS CIENA
    def clear_all_fields_and_cache_iris(app_data):
        # Show warning prompt before clearing
        user_response = messagebox.askyesno(
            "Warning",
            "Are you sure you want to clear all entries? If you click 'YES', all data will be reset and you will need to restart the analysis from the beginning.",
            parent=scrollable_frame,
            icon= "warning"
        )
        if not user_response:
            return  # Do nothing if user selects NO

        try:
            # Clear all entry fields in the IRIS window
            entry_field.delete(0, tk.END)
            test_number_field.delete(0, tk.END)
            read_point_field.delete(0, tk.END)
            sheet_name_field.delete(0, tk.END)
            graph_table_test_field.delete(0, tk.END)
            graph_table_sheet_field.delete(0, tk.END)
            line_chart_test_field.delete(0, tk.END)
            line_chart_color_field.delete(0, tk.END)
            line_chart_legend_field.delete(0, tk.END)
            line_chart_title_field.delete(0, tk.END)
            line_chart_ylabel_field.delete(0, tk.END)
            line_chart_width_field.delete(0, tk.END)
            line_chart_height_field.delete(0, tk.END)
            histogram_test_field.delete(0, tk.END)
            histogram_bins_field.delete(0, tk.END)
            histogram_title_field.delete(0, tk.END)
            test_statistics_field.delete(0, tk.END)
            failure_operand_field.delete(0, tk.END)
            failure_hbin_field.delete(0, tk.END)
            final_chart_test_field.delete(0, tk.END)
            final_chart_color_field.delete(0, tk.END)
            final_chart_read_points_field.delete(0, tk.END)
            final_chart_title_field.delete(0, tk.END)
            final_chart_ylabel_field.delete(0, tk.END)
            final_chart_width_field.delete(0, tk.END)
            final_chart_height_field.delete(0, tk.END)

            # Reset all variables and cache specific to IRIS
            app_data.file_uploaded_iris = False
            app_data.df_v1_read_point_24h_iris = None
            app_data.df_testing_24h_iris = None
            app_data.df_general_24h_iris = None
            app_data.df_analysis_24h_iris = None
            app_data.df_general_24h_help_1_iris = None
            app_data.df_general_24h_help_2_iris = None
            app_data.df_general_24h_help_3_iris = None
            app_data.graph_24h_iris = None
            app_data.failure_24h_iris = None
            app_data.df_limit_24h_iris = None
            app_data.read_point_graphs_iris = []
            app_data.final_table_iris = None
            app_data.overview_generated_iris = False
            app_data.graph_24h_generated_iris = False
            app_data.excessive_units_iris = None
            app_data.df_limit_24h_iris_generated= False
            app_data.analysis_point_graphs_iris = []
            app_data.excessive_units_iris_generated = False
            app_data.analysis_point_graphs_iris_merged = None
            app_data.full_graph_iris_generated = False
            app_data.full_graph_iris_df = None
            app_data.plot_df_generated = False
            app_data.plot_df = None
            app_data.full_limit_generated = False
            
            # Close all matplotlib plots in the main thread
            plt.close('all')

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while clearing: {str(e)}", parent=scrollable_frame)
    
    # Function to clear all fields in the Final Chart section of the IRIS window
    def clear_final_chart_fields_iris():
        final_chart_test_field.delete(0, tk.END)
        final_chart_color_field.delete(0, tk.END)
        final_chart_read_points_field.delete(0, tk.END)
        final_chart_title_field.delete(0, tk.END)
        final_chart_ylabel_field.delete(0, tk.END)
        final_chart_width_field.delete(0, tk.END)
        final_chart_height_field.delete(0, tk.END)

    # Clear function for the Failure Detection part of the IRIS window
    def clear_failure_detection_fields_iris():
        failure_operand_field.delete(0, tk.END)
        failure_hbin_field.delete(0, tk.END)    

    # Function for Clear button for Test Statistics of IRIS
    def clear_test_statistics_field_iris():
        test_statistics_field.delete(0, tk.END)

    # Function for Clear button for the Distribution chart of IRIS
    def clear_distribution_fields_iris():
        histogram_test_field.delete(0, tk.END)
        histogram_bins_field.delete(0, tk.END)
        histogram_title_field.delete(0, tk.END)
    
    # Function to clear entries for the line chart (section 5) for IRIS
    def clear_line_chart_fields_iris():
        line_chart_test_field.delete(0, tk.END)
        line_chart_color_field.delete(0, tk.END)
        line_chart_legend_field.delete(0, tk.END)
        line_chart_title_field.delete(0, tk.END)
        line_chart_ylabel_field.delete(0, tk.END)
        line_chart_width_field.delete(0, tk.END)
        line_chart_height_field.delete(0, tk.END)
    
    # Function to clear graph Table section (4)
    def clear_graph_table_fields_iris():
        graph_table_test_field.delete(0, tk.END)
        graph_table_sheet_field.delete(0, tk.END)
    
    # Function to Clear General View Section (3)
    def clear_general_view():
        read_point_field.delete(0, tk.END)
        sheet_name_field.delete(0, tk.END)

    # Function to clear "Test Name" entry field of section (2)
    def clear_test_name_fields_iris():
        test_number_field.delete(0, tk.END)

    # Function for the 'Clear' button next to the upload button at the beginning
    def clear_upload_field_iris(app_data):
        try:
            entry_field.delete(0, tk.END)

            # Reset all variables and cache related to the uploaded file
            app_data.file_uploaded_iris = False
            app_data.df_v1_read_point_24h_iris = None
            app_data.df_testing_24h_iris = None
            app_data.df_general_24h_iris = None
            app_data.df_analysis_24h_iris = None
            app_data.df_general_24h_help_1_iris = None
            app_data.df_general_24h_help_2_iris = None
            app_data.df_general_24h_help_3_iris = None
            app_data.df_limit_24h_iris_generated = False
            app_data.df_limit_24h_iris = None
            app_data.graph_24h_iris = None
            app_data.failure_24h_iris = None
            app_data.overview_generated_iris = False
            app_data.graph_24h_generated_iris = False

            # Close all matplotlib plots
            plt.close('all')

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while clearing: {str(e)}", parent=scrollable_frame)

    ### Widgets for CIENA window and tooltip management
    
    # Entry field for Excel file
    entry_label = tk.Label(scrollable_frame, text="1. Insert Excel File:", bg="#1C2541", fg="white")
    entry_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    entry_field = tk.Entry(scrollable_frame, width=30)
    entry_field.grid(row=0, column=1, padx=10, pady=10)

    # Tooltip for the entry field
    tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    tooltip_label.grid(row=0, column=2, padx=10, pady=10)

    # Add the "Upload" button next to the question mark
    upload_button_iris = tk.Button(scrollable_frame, text="Upload", command=lambda: upload_file_iris(app_data), bg="#1C2541", fg="white", width=10)
    upload_button_iris.grid(row=0, column=3, padx=10, pady=10)

    # Add the "Clear" button next to the "Upload" button
    clear_button_iris = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_upload_field_iris(app_data),bg="#1C2541", fg="white", width=10)
    clear_button_iris.grid(row=0, column=4, padx=10, pady=10)
     
    # Function to show the tooltip
    tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(
            event,
            "Insert the Excel file (Data Analysis sheet), obtained after running the Macro.",
            scrollable_frame
        )
    )
    tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Test Number" label and entry field
    test_number_label = tk.Label(scrollable_frame, text="2. Test Number:", bg="#1C2541", fg="white")
    test_number_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
    test_number_field = tk.Entry(scrollable_frame, width=30)
    test_number_field.grid(row=1, column=1, padx=10, pady=10)

    # Add the tooltip for the "Test Number" entry field
    test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    test_tooltip_label.grid(row=1, column=2, padx=10, pady=10)

    # Function to show the tooltip for "Test Number"
    test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(
            event,
            "Insert test number(s), separated by commas, to obtain their relevant test name(s).",
            scrollable_frame
        )
    )
    test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Test Name" button
    test_name_button = tk.Button(scrollable_frame, text="Test Name", command= lambda: get_test_names_iris(app_data), bg="#1C2541", fg="white", width=10)
    test_name_button.grid(row=1, column=3, padx=10, pady=10)

    # Add the "Clear" button for the "Test Number" entry field
    clear_test_button = tk.Button(scrollable_frame, text="Clear", command= clear_test_name_fields_iris, bg="#1C2541", fg="white", width=10)
    clear_test_button.grid(row=1, column=4, padx=10, pady=10)

    # Add the "General View" label
    general_view_label = tk.Label(scrollable_frame, text="3. General View:", bg="#1C2541", fg="white")
    general_view_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

    # Add the "Read Point" entry field
    read_point_field = tk.Entry(scrollable_frame, width=30)
    read_point_field.grid(row=2, column=1, padx=10, pady=10)

    # Add the tooltip for the "Read Point" entry field
    read_point_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    read_point_tooltip_label.grid(row=2, column=2, padx=10, pady=10)

    # Function to show the tooltip for "Read Point"
    read_point_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(
            event,
            "Enter the read point number (only an integer) you wish to analyze (e.g., 24, 168, 1000, etc.).",
            scrollable_frame
        )
    )
    read_point_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Sheet Name" entry field
    sheet_name_field = tk.Entry(scrollable_frame, width=30)
    sheet_name_field.grid(row=2, column=3, padx=10, pady=10)

    # Add the tooltip for the "Sheet Name" entry field
    sheet_name_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    sheet_name_tooltip_label.grid(row=2, column=4, padx=10, pady=10)

    # Function to show the tooltip for "Sheet Name"
    sheet_name_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(
            event,
            "Specify the sheet name for the Excel file you are creating (e.g., general_view).",
            scrollable_frame
        )
    )
    sheet_name_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Clear" button for the General View section
    clear_general_button = tk.Button(scrollable_frame, text="Clear", command=clear_general_view, bg="#1C2541", fg="white", width=10)
    clear_general_button.grid(row=2, column=5, padx=10, pady=10)

    # Function to clear the General View fields
    def clear_general_view():
        read_point_field.delete(0, tk.END)
        sheet_name_field.delete(0, tk.END)

    # Bind the clear function to the button
    clear_general_button.config(command=clear_general_view)

    # Add the "Overview Generation" button
    overview_button = tk.Button(scrollable_frame, text="Overview Generation", command=lambda: generate_overview_iris(app_data), bg="#1C2541", fg="white", width=20)
    overview_button.grid(row=2, column=6, padx=10, pady=10)

    # Add the "Graph Table" label
    graph_table_label = tk.Label(scrollable_frame, text="4. Graph Table:", bg="#1C2541", fg="white")
    graph_table_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")

    # Add the "Test Number" entry field
    graph_table_test_field = tk.Entry(scrollable_frame, width=30)
    graph_table_test_field.grid(row=3, column=1, padx=10, pady=10)

    # Add the tooltip for the "Test Number" entry field
    graph_table_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    graph_table_test_tooltip_label.grid(row=3, column=2, padx=10, pady=10)

    # Function to show the tooltip for "Test Number" for the Graph Table section
    graph_table_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Insert the test number(s) (integers) you wish to analyze.", scrollable_frame)
    )
    graph_table_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Sheet Name" entry field
    graph_table_sheet_field = tk.Entry(scrollable_frame, width=30)
    graph_table_sheet_field.grid(row=3, column=3, padx=10, pady=10)

    # Add the tooltip for the "Sheet Name" entry field
    graph_table_sheet_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    graph_table_sheet_tooltip_label.grid(row=3, column=4, padx=10, pady=10)

    # Function to show the tooltip for "Sheet Name"
    graph_table_sheet_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the sheet name for the Excel file you are creating (e.g., graph_view).", scrollable_frame)
    )
    graph_table_sheet_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Clear" button for the Graph Table section
    clear_graph_table_button = tk.Button(scrollable_frame, text="Clear", command=clear_graph_table_fields_iris, bg="#1C2541", fg="white", width=10)
    clear_graph_table_button.grid(row=3, column=5, padx=10, pady=10)

    # Add the "Graph Table" button
    graph_table_button = tk.Button(scrollable_frame, text="Graph Table", command=lambda: generate_graph_table_preview_iris(app_data),bg="#1C2541", fg="white", width=20)
    graph_table_button.grid(row=3, column=6, padx=10, pady=10)

    # Add the "Complete Table" button
    complete_table_button = tk.Button(scrollable_frame, text="Complete Table", command=lambda: generate_complete_table_iris(app_data),bg="#1C2541", fg="white", width=13)
    complete_table_button.grid(row=3, column=7, padx=10, pady=10)

    # Add the "Read Point" button
    read_point_button = tk.Button(scrollable_frame, text="Read Point", command= lambda: generate_read_point_graph_iris(app_data) ,bg="#1C2541", fg="white", width=13)
    read_point_button.grid(row=3, column=8, padx=10, pady=10)

    # Add the "Line Chart" label
    line_chart_label = tk.Label(scrollable_frame, text="5. Line Chart:", bg="#1C2541", fg="white")
    line_chart_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")

    # Entry field for test number (1st entry field)
    line_chart_test_field = tk.Entry(scrollable_frame, width=30)
    line_chart_test_field.grid(row=4, column=1, padx=10, pady=10)

    # Tooltip for the test number field
    line_chart_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_test_tooltip_label.grid(row=4, column=2, padx=10, pady=10)

    # Entry field for line color (2nd entry field)
    line_chart_color_field = tk.Entry(scrollable_frame, width=30)
    line_chart_color_field.grid(row=4, column=3, padx=10, pady=10)

    # Tooltip for the line color field
    line_chart_color_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_color_tooltip_label.grid(row=4, column=4, padx=10, pady=10)

    # Entry field for legend label (3rd entry field)
    line_chart_legend_field = tk.Entry(scrollable_frame, width=30)
    line_chart_legend_field.grid(row=4, column=5, padx=10, pady=10)

    # Tooltip for the legend label field
    line_chart_legend_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_legend_tooltip_label.grid(row=4, column=6, padx=10, pady=10)

    # Entry field for chart title (4th entry field)
    line_chart_title_field = tk.Entry(scrollable_frame, width=30)
    line_chart_title_field.grid(row=5, column=1, padx=10, pady=10)

    # Tooltip for the chart title field
    line_chart_title_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_title_tooltip_label.grid(row=5, column=2, padx=10, pady=10)

    # Entry field for y-axis label (5th entry field)
    line_chart_ylabel_field = tk.Entry(scrollable_frame, width=30)
    line_chart_ylabel_field.grid(row=5, column=3, padx=10, pady=10)

    # Tooltip for the y-axis label field
    line_chart_ylabel_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_ylabel_tooltip_label.grid(row=5, column=4, padx=10, pady=10)

    # Entry field for figure width (6th entry field)
    line_chart_width_field = tk.Entry(scrollable_frame, width=30)
    line_chart_width_field.grid(row=5, column=5, padx=10, pady=10)

    # Tooltip for the figure width field
    line_chart_width_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_width_tooltip_label.grid(row=5, column=6, padx=10, pady=10)

    # Entry field for figure height (7th entry field)
    line_chart_height_field = tk.Entry(scrollable_frame, width=30)
    line_chart_height_field.grid(row=6, column=1, padx=10, pady=10)

    # Tooltip for the figure height field
    line_chart_height_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_height_tooltip_label.grid(row=6, column=2, padx=10, pady=10)

    # Add the "Plot" button
    plot_button = tk.Button(scrollable_frame, text="Plot", command=lambda: plot_line_chart_iris(app_data), bg="#1C2541", fg="white", width=10)
    plot_button.grid(row=6, column=3, padx=10, pady=10)

    # Add the "Plot & Limit" button
    plot_button = tk.Button(scrollable_frame, text="Plot & Limit", command= lambda: plot_line_limit_chart_iris(app_data) ,bg="#1C2541", fg="white", width=10)
    plot_button.grid(row=6, column=4, padx=10, pady=10)
    
    
    # Add the "Clear" button for the Line Chart section
    clear_line_chart_button = tk.Button(scrollable_frame, text="Clear", command=clear_line_chart_fields_iris ,bg="#1C2541", fg="white", width=10)
    clear_line_chart_button.grid(row=6, column=5, padx=10, pady=10)

    # Tooltip for the "Test Number" entry field for the Line Chart in CIENA window
    line_chart_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number you wish to plot. Ensure you select the same test number(s) as those chosen in step 4.", scrollable_frame)
    )
    line_chart_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Tooltip for the "Line Color" entry field
    line_chart_color_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Insert one of these colors: 'darkred', 'olive', 'blue', 'darkorange', 'purple', 'brown', 'black', 'goldenrod', 'teal', 'darkmagenta', 'peru'.", scrollable_frame)
    )
    line_chart_color_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Tooltip for the "Legend Label" entry field
    line_chart_legend_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the read point number (e.g., 24H, 48H, 168H, etc.).", scrollable_frame)
    )
    line_chart_legend_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Tooltip for the "Chart Title" entry field
    line_chart_title_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Select the title for your chart with this naming convention: 'Test Number, Test Name'.", scrollable_frame)
    )
    line_chart_title_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Tooltip for the "Y-axis Label" entry field
    line_chart_ylabel_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number to label the Y-axis (e.g., Test Value #2006).", scrollable_frame)
    )
    line_chart_ylabel_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Tooltip for the "Figure Width" entry field
    line_chart_width_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the width of your line chart (default: 25).", scrollable_frame)
    )
    line_chart_width_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Tooltip for the "Figure Height" entry field
    line_chart_height_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the height of your line chart (default: 10).", scrollable_frame)
    )
    line_chart_height_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "6. Histogram" label
    histogram_label = tk.Label(scrollable_frame, text="6. Histogram:",bg="#1C2541", fg="white")
    histogram_label.grid(row=7, column=0, padx=10, pady=10, sticky="w")

    # Add the "Test Number" entry field (1st entry field)
    histogram_test_field = tk.Entry(scrollable_frame, width=30)
    histogram_test_field.grid(row=7, column=1, padx=10, pady=10)

    # Add the tooltip for the "Test Number" entry field
    histogram_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    histogram_test_tooltip_label.grid(row=7, column=2, padx=10, pady=10)

    # Tooltip for the "Test Number" entry field
    histogram_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number to analyze its distribution. Use the same test numbers as in step 4. If you wish to plot the distribution for all read points, you need to select a test number as of step 9.", scrollable_frame)
    )
    histogram_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Bins" entry field (2nd entry field)
    histogram_bins_field = tk.Entry(scrollable_frame, width=30)
    histogram_bins_field.grid(row=7, column=3, padx=10, pady=10)

    # Add the tooltip for the "Bins" entry field
    histogram_bins_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    histogram_bins_tooltip_label.grid(row=7, column=4, padx=10, pady=10)

    # Tooltip for the "Bins" entry field
    histogram_bins_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the number of bins for the histogram. Higher values provide more detail.", scrollable_frame)
    )
    histogram_bins_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Chart Title" entry field (3rd entry field)
    histogram_title_field = tk.Entry(scrollable_frame, width=30)
    histogram_title_field.grid(row=7, column=5, padx=10, pady=10)

    # Add the tooltip for the "Chart Title" entry field
    histogram_title_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    histogram_title_tooltip_label.grid(row=7, column=6, padx=10, pady=10)

    # Tooltip for the "Chart Title" entry field
    histogram_title_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the title for the histogram. Use the naming convention: 'Test Number, Test Name'.", scrollable_frame)
    )
    histogram_title_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Plot Histogram" button
    plot_histogram_button = tk.Button(scrollable_frame, text="Distribution Plot", command=lambda: plot_distribution_chart_iris(app_data), bg="#1C2541", fg="white", width=15)
    plot_histogram_button.grid(row=8, column=1, padx=10, pady=10)

    # Add the "Plot Histogram with limits" button
    plot_histogram_button = tk.Button(scrollable_frame, text="Distribution Plot with Limits", command= lambda: dist_iris_single_read_point(app_data) ,bg="#1C2541", fg="white", width=25)
    plot_histogram_button.grid(row=8, column=2, padx=10, pady=10)
    
    # Add the "Clear" button for the Histogram section
    clear_histogram_button = tk.Button(scrollable_frame, text="Clear", command=clear_distribution_fields_iris, bg="#1C2541", fg="white", width=10)
    clear_histogram_button.grid(row=8, column=4, padx=10, pady=10)

    # Add the "7. Test Statistics" label
    test_statistics_label = tk.Label(scrollable_frame, text="7. Test Statistics:", bg="#1C2541", fg="white")
    test_statistics_label.grid(row=9, column=0, padx=10, pady=10, sticky="w")

    # Add the entry field for test number
    test_statistics_field = tk.Entry(scrollable_frame, width=30)
    test_statistics_field.grid(row=9, column=1, padx=10, pady=10)

    # Add the tooltip for the test number field
    test_statistics_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    test_statistics_tooltip_label.grid(row=9, column=2, padx=10, pady=10)

    # Tooltip for the "Test Number" entry field
    test_statistics_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number you previously selected in step 4 to retrieve statistics.", scrollable_frame)
    )
    test_statistics_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Test Stats" button
    test_statistics_button = tk.Button(scrollable_frame, text="Test Stats", command=lambda: generate_test_statistics_iris(app_data),bg="#1C2541", fg="white", width=10)
    test_statistics_button.grid(row=9, column=3, padx=10, pady=10)

    # Add the "Clear" button next to the "Test Stats" button
    clear_test_statistics_button = tk.Button(scrollable_frame, text="Clear", command= clear_test_statistics_field_iris ,bg="#1C2541", fg="white", width=10)
    clear_test_statistics_button.grid(row=9, column=4, padx=10, pady=10)

    # Add the "8. Failure Detection" label
    failure_detection_label = tk.Label(scrollable_frame, text="8. Failure Detection:", bg="#1C2541", fg="white")
    failure_detection_label.grid(row=10, column=0, padx=10, pady=10, sticky="w")

    # Add the entry field for the operand
    failure_operand_field = tk.Entry(scrollable_frame, width=10)
    failure_operand_field.grid(row=10, column=1, padx=10, pady=10)

    # Add the tooltip for the operand field
    failure_operand_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    failure_operand_tooltip_label.grid(row=10, column=2, padx=10, pady=10)

    # Tooltip for the operand field
    failure_operand_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter one of these operands: <, <=, >, >=, or =.", scrollable_frame)
    )
    failure_operand_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the entry field for the HBIN number
    failure_hbin_field = tk.Entry(scrollable_frame, width=10)
    failure_hbin_field.grid(row=10, column=3, padx=10, pady=10)

    # Add the tooltip for the HBIN number field
    failure_hbin_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    failure_hbin_tooltip_label.grid(row=10, column=4, padx=10, pady=10)

    # Tooltip for the HBIN number field
    failure_hbin_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the HBIN number you wish to analyze.", scrollable_frame)
    )
    failure_hbin_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Failure Table Generation" button
    failure_table_button = tk.Button(scrollable_frame, text="Failure Table Generation", command=lambda:generate_failure_table_iris(app_data), bg="#1C2541", fg="white", width=20)
    failure_table_button.grid(row=10, column=5, padx=10, pady=10)

    # Add the "Clear" button for the Failure Detection section
    clear_failure_button = tk.Button(scrollable_frame, text="Clear", command= clear_failure_detection_fields_iris, bg="#1C2541", fg="white", width=10)
    clear_failure_button.grid(row=10, column=6, padx=10, pady=10)

    # Function to clear all fields in the Failure Detection section
    def clear_failure_detection_fields():
        failure_operand_field.delete(0, tk.END)
        failure_hbin_field.delete(0, tk.END)

    # Bind the clear function to the "Clear" button
    clear_failure_button.config(command=clear_failure_detection_fields)

    # Add the "9. Final Chart" label
    final_chart_label = tk.Label(scrollable_frame, text="9. Final Chart:", bg="#1C2541", fg="white")
    final_chart_label.grid(row=11, column=0, padx=10, pady=10, sticky="w")

    # Add the "Test Number" entry field (1st entry field)
    final_chart_test_field = tk.Entry(scrollable_frame, width=30)
    final_chart_test_field.grid(row=11, column=1, padx=10, pady=10)

    # Add the tooltip for the "Test Number" entry field
    final_chart_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_test_tooltip_label.grid(row=11, column=2, padx=10, pady=10)

    # Tooltip for the "Test Number" entry field
    final_chart_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number you wish to plot.", scrollable_frame)
    )
    final_chart_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Chart Color" entry field (2nd entry field)
    final_chart_color_field = tk.Entry(scrollable_frame, width=30)
    final_chart_color_field.grid(row=11, column=3, padx=10, pady=10)

    # Add the tooltip for the "Chart Color" entry field
    final_chart_color_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_color_tooltip_label.grid(row=11, column=4, padx=10, pady=10)

    # Tooltip for the "Chart Color" entry field
    final_chart_color_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the color of your chart. Ensure the number of color names matches the number of read points.", scrollable_frame)
    )
    final_chart_color_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Read Points" entry field (3rd entry field)
    final_chart_read_points_field = tk.Entry(scrollable_frame, width=30)
    final_chart_read_points_field.grid(row=11, column=5, padx=10, pady=10)

    # Add the tooltip for the "Read Points" entry field
    final_chart_read_points_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_read_points_tooltip_label.grid(row=11, column=6, padx=10, pady=10)

    # Tooltip for the "Read Points" entry field
    final_chart_read_points_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter all read points you wish to analyze (e.g., 24H, 68H, etc.).", scrollable_frame)
    )
    final_chart_read_points_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Chart Title" entry field (4th entry field)
    final_chart_title_field = tk.Entry(scrollable_frame, width=30)
    final_chart_title_field.grid(row=12, column=1, padx=10, pady=10)

    # Add the tooltip for the "Chart Title" entry field
    final_chart_title_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_title_tooltip_label.grid(row=12, column=2, padx=10, pady=10)

    # Tooltip for the "Chart Title" entry field
    final_chart_title_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the title of the chart using the naming convention 'Test Number: Test Name'.", scrollable_frame)
    )
    final_chart_title_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Y-axis Label" entry field (5th entry field)
    final_chart_ylabel_field = tk.Entry(scrollable_frame, width=30)
    final_chart_ylabel_field.grid(row=12, column=3, padx=10, pady=10)

    # Add the tooltip for the "Y-axis Label" entry field
    final_chart_ylabel_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_ylabel_tooltip_label.grid(row=12, column=4, padx=10, pady=10)

    # Tooltip for the "Y-axis Label" entry field
    final_chart_ylabel_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Label the Y-axis using this naming convention: Test Value #2006.", scrollable_frame)
    )
    final_chart_ylabel_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Chart Width" entry field (6th entry field)
    final_chart_width_field = tk.Entry(scrollable_frame, width=30)
    final_chart_width_field.grid(row=12, column=5, padx=10, pady=10)

    # Add the tooltip for the "Chart Width" entry field
    final_chart_width_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_width_tooltip_label.grid(row=12, column=6, padx=10, pady=10)

    # Tooltip for the "Chart Width" entry field
    final_chart_width_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the width of your line chart (default: 25).", scrollable_frame)
    )
    final_chart_width_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Chart Height" entry field (7th entry field)
    final_chart_height_field = tk.Entry(scrollable_frame, width=30)
    final_chart_height_field.grid(row=13, column=1, padx=10, pady=10)

    # Add the tooltip for the "Chart Height" entry field
    final_chart_height_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_height_tooltip_label.grid(row=13, column=2, padx=10, pady=10)

    # Tooltip for the "Chart Height" entry field
    final_chart_height_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the height of your line chart (default: 10).", scrollable_frame)
    )
    final_chart_height_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Final Graph" button
    final_graph_button = tk.Button(scrollable_frame, text="Final Graph", command=lambda: generate_final_graph_iris(app_data) ,bg="#1C2541", fg="white", width=15)
    final_graph_button.grid(row=13, column=4, padx=10, pady=10)

    # Add the "Final Table" button
    final_table_button = tk.Button(scrollable_frame, text="Final Table", command=lambda: generate_final_table_iris(app_data),bg="#1C2541", fg="white", width=15)
    final_table_button.grid(row=13, column=3, padx=10, pady=10)

    # Add the "Excessive Units" button
    excessive_units_button = tk.Button(scrollable_frame, text="Excessive Units", command=lambda: plot_excessive_units_iris(app_data),bg="#1C2541", fg="white", width=15)
    excessive_units_button.grid(row=14, column=1, padx=10, pady=10)

    # Add the "Full Distribution" button
    excessive_units_button = tk.Button(scrollable_frame, text="Full Distribution", command= lambda: full_distribution_iris(app_data), bg="#1C2541", fg="white", width=15)
    excessive_units_button.grid(row=8, column=3, padx=10, pady=10)
    
    # Add the "Clear All" button
    clear_all_button = tk.Button(scrollable_frame, text="Clear All", command= lambda:clear_all_fields_and_cache_iris(app_data), bg="#1C2541", fg="white", width=10)
    clear_all_button.grid(row=13, column=6, padx=10, pady=10)

    # Add the "Full Limit" button
    full_analysis_button = tk.Button(scrollable_frame, text="Full Limit", command= lambda: full_analysis(app_data), bg="#1C2541", fg="white", width=15)
    full_analysis_button.grid(row=14, column=2, padx=10, pady=10)
    
    # Add the "Full Graph" button
    full_graph_button = tk.Button(scrollable_frame, text="Full Graph", command= lambda: plot_full_graph_iris(app_data) ,bg="#1C2541", fg="white", width=15)
    full_graph_button.grid(row=14, column=3, padx=10, pady=10)
    
    # Add 'Limit' button
    analysis_button = tk.Button(scrollable_frame, text="Limit", command= lambda: limit_with_test_name_iris(app_data),bg="#1C2541", fg="white", width=10)
    analysis_button.grid(row=3, column=9, padx=10, pady=10 )

    # Construct the path to the Warning Sign Image using the base_path for .exe file application
    warning_logo_path = os.path.join(base_path, "resources", "warning.png")

    # Create a dummy image and label first (will be resized and updated later)
    warning_logo_image = Image.open(warning_logo_path)
    warning_logo_image = warning_logo_image.resize((24, 24), Image.Resampling.LANCZOS)  # Initial size, will be updated
    warning_logo_image_tk = ImageTk.PhotoImage(warning_logo_image)
    warning_logo_label = tk.Label(
        scrollable_frame,
        image=warning_logo_image_tk,
        bg="#1C2541"
    )
    warning_logo_label.image = warning_logo_image_tk  # Keep a reference to avoid garbage collection

    # Place the label initially (will be repositioned and resized by place_warning_logo)
    warning_logo_label.place(x=0, y=0)

    def place_warning_logo():
        button_height = clear_all_button.winfo_height()
        button_width = clear_all_button.winfo_width()
        ratio = 0.9
        img_height = int(button_height * ratio)
        img_width = int(button_height * ratio)
        warning_logo_image = Image.open(warning_logo_path)
        warning_logo_image = warning_logo_image.resize((img_width, img_height), Image.Resampling.LANCZOS)
        warning_logo_image_tk = ImageTk.PhotoImage(warning_logo_image)
        warning_logo_label.config(image=warning_logo_image_tk)
        warning_logo_label.image = warning_logo_image_tk  # Keep a reference
        warning_logo_label.place(
            x=clear_all_button.winfo_x() + button_width + 8,
            y=clear_all_button.winfo_y() + (button_height - img_height) // 2
        )

    scrollable_frame.after(200, place_warning_logo)
    
    # Add the "Clear" button
    clear_button_label_9 = tk.Button(scrollable_frame, text="Clear", command=clear_final_chart_fields_iris,bg="#1C2541", fg="white", width=10)
    clear_button_label_9.grid(row=13, column=5, padx=10, pady=10)

    # Construct the path to the ST logo image using the base_path for .exe file
    logo_path = os.path.join(base_path, "resources", "STlogo.gif_min.gif")

    # Load the ST logo image
    logo_image = tk.PhotoImage(file=logo_path)

    # Add the ST logo to the bottom-right corner of the IRIS window
    iris_logo_label = tk.Label(iris_failure_window, image=logo_image, bg="#1C2541")
    iris_logo_label.image = logo_image  # Keep a reference to avoid garbage collection
    iris_logo_label.place(relx=1.0, rely=1.0, anchor='se', x=-20, y= -20)  # Adjust x and y for padding


############ Sub-options for CIENA IRIS Reliability Analysis (Graph)
iris_reliability_sub_button2 = tk.Button(ciena_reliability_sub_option_frame, text="Graph", command= iris_open_failure_detection_window,**italic_button_style, width=half_button_width)
iris_reliability_sub_button2.pack(side='top', pady=5)



# Function to open the Failure Detection window for AMAZON PRODUCTS
def open_failure_detection_window():
    failure_window = tk.Toplevel(root)
    failure_window.title("AMAZON")
    failure_window.geometry("900x800")
    failure_window.configure(bg="#1C2541")

    # Set the icon for the AMAZON window
    amazon_icon_path = os.path.join(base_path, "resources", "amazon.ico")
    failure_window.iconbitmap(amazon_icon_path)
    
# Create a canvas and both scrollbars
    canvas = tk.Canvas(failure_window, bg="#1C2541", height=280, width=380)
    v_scrollbar = tk.Scrollbar(failure_window, orient="vertical", command=canvas.yview, width=14)
    h_scrollbar = tk.Scrollbar(failure_window, orient="horizontal", command=canvas.xview, width=14)

    # Create a frame inside the canvas
    scrollable_frame = tk.Frame(canvas, bg="#1C2541")

    # Bind the scrollbars to the canvas
    canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

    # Add the scrollable frame to the canvas
    scrollable_frame_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # After creating scrollable_frame and scrollable_frame_id
    def resize_scrollable_frame(event):
        # Set the width of the scrollable_frame to match the canvas
        canvas.itemconfig(scrollable_frame_id)

    canvas.bind("<Configure>", resize_scrollable_frame)
    
    # Configure the scrollable frame to update the canvas scroll region
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    # Grid the canvas and scrollbars
    canvas.grid(row=0, column=0, sticky="nsew")
    v_scrollbar.grid(row=0, column=1, sticky="ns", padx=1)
    h_scrollbar.grid(row=1, column=0, sticky="ew", padx=1)

    # Configure grid weights for resizing
    failure_window.grid_rowconfigure(0, weight=1)
    failure_window.grid_columnconfigure(0, weight=1)

    # Bind the mouse wheel event to the canvas for vertical scrolling
    def on_mouse_wheel(event):
        direction = -1 if event.delta > 0 else 1
        canvas.yview_scroll(direction, "units")

    # Bind the mouse wheel event to the canvas for horizontal scrolling (Shift+Wheel)
    def on_shift_mouse_wheel(event):
        direction = -1 if event.delta > 0 else 1
        canvas.xview_scroll(direction, "units")

    # Bind the mouse wheel events to the failure window and its components
    failure_window.bind("<MouseWheel>", on_mouse_wheel)
    canvas.bind("<MouseWheel>", on_mouse_wheel)
    scrollable_frame.bind("<MouseWheel>", on_mouse_wheel)
    v_scrollbar.bind("<MouseWheel>", on_mouse_wheel)

    failure_window.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)
    canvas.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)
    scrollable_frame.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)
    h_scrollbar.bind("<Shift-MouseWheel>", on_shift_mouse_wheel)
    
    # Configure the failure window grid
    failure_window.grid_rowconfigure(0, weight=1)
    failure_window.grid_columnconfigure(0, weight=1)
    
    # Entry field for Excel file name
    entry_label = tk.Label(scrollable_frame, text="1. Insert Excel File:", bg="#1C2541", fg="white")
    entry_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    entry_field = tk.Entry(scrollable_frame, width=30)
    entry_field.grid(row=0, column=1, padx=10, pady=10)

    # Tooltip for the entry field
    tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    tooltip_label.grid(row=0, column=2, padx=10, pady=10)

    tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Insert the Excel file (Data Analysis sheet), obtained after running the Macro.", scrollable_frame)
    )
    tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Entry field for Test Number
    test_number_label = tk.Label(scrollable_frame, text="2. Test Number:", bg="#1C2541", fg="white")
    test_number_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
    test_number_field = tk.Entry(scrollable_frame, width=30)
    test_number_field.grid(row=1, column=1, padx=10, pady=10)

    # Tooltip for the test number field
    test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    test_tooltip_label.grid(row=1, column=2, padx=10, pady=10)

    # Function to show the test number tooltip for AMAZON
    test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Insert test number(s), separated by commas, to obtain their relevant test name(s).", scrollable_frame)
    )
    test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # General View widget creation
    general_view_label = tk.Label(scrollable_frame, text="3. General View:", bg="#1C2541", fg="white")
    general_view_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
    read_point_field = tk.Entry(scrollable_frame, width=30)
    read_point_field.grid(row=2, column=1, padx=10, pady=10)
    sheet_name_field = tk.Entry(scrollable_frame, width=30)
    sheet_name_field.grid(row=2, column=3, padx=10, pady=10)

    # Tooltip for the read point of general view section
    read_point_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    read_point_tooltip_label.grid(row=2, column=2, padx=10, pady=10)

    # Function to show the read point tooltip for the general view
    read_point_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the read point number (only an integer) you wish to analyze (e.g., 24, 168, 1000, etc.).", scrollable_frame)
    )
    read_point_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Tooltip for the sheet name field
    sheet_name_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    sheet_name_tooltip_label.grid(row=2, column=4, padx=10, pady=10)

    # Function to show the sheet name tooltip for the general view
    sheet_name_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the sheet name for the Excel file you are creating (e.g., general_view).", scrollable_frame)
    )
    sheet_name_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # File upload function definition
    def upload_file(app_data):
        try:
            # Open file dialog to select the file
            file_path = filedialog.askopenfilename(parent=scrollable_frame, filetypes=[("Excel files", "*.xlsx")])

            if file_path:
                # Update the entry field directly
                entry_field.delete(0, tk.END)
                entry_field.insert(0, file_path)

                # Initialize variables before processing the file
                app_data.df_testing_24h = None
                app_data.df_v1_read_point_24h = None
                app_data.df_general_24h = None
                app_data.df_analysis_24h = None

                # Reset flags to ensure proper workflow
                app_data.overview_generated = False
                app_data.graph_24h_generated = False
                
                # Create a progress bar window
                progress_window = tk.Toplevel(scrollable_frame)
                progress_window.title("Uploading File")
                progress_window.iconbitmap(os.path.join(base_path, "resources", "amazon.ico"))
                progress_window.geometry("400x200")
                progress_window.configure(bg="#1C2541")

                # Top frame (optional placeholder if you want to add title text later)
                top_frame = tk.Frame(progress_window, bg="#1C2541")
                top_frame.pack(pady=(10, 0))

                # Add AMAZON logo at the top of the progress window
                amazon_logo_path = os.path.join(base_path, "resources", "amazon.ico")
                try:
                    amazon_logo_img = Image.open(amazon_logo_path)
                    amazon_logo_img = amazon_logo_img.resize((48, 48), Image.Resampling.LANCZOS)
                    amazon_logo_tk = ImageTk.PhotoImage(amazon_logo_img)
                    logo_label = tk.Label(progress_window, image=amazon_logo_tk, bg="#1C2541")
                    logo_label.image = amazon_logo_tk  # Keep reference
                    logo_label.pack(pady=(10, 0))
                except Exception:
                    # Fallback: show text if logo load fails
                    logo_label = tk.Label(progress_window, text="AMAZON", fg="white", bg="#1C2541", font=("Helvetica", 16, "bold"))
                    logo_label.pack(pady=(10, 0))

                progress_label = tk.Label(
                    progress_window,
                    text="Uploading file, please wait...",
                    bg="#1C2541",
                    fg="white",
                    font=("Helvetica", 12)
                )
                progress_label.pack(pady=10)

                progress_bar = ttk.Progressbar(
                    progress_window,
                    orient="horizontal",
                    mode="determinate",
                    length=250
                )
                progress_bar.pack(pady=10)

                def simulate_progress(step=1):
                    if step <= 100:
                        progress_bar["value"] = step
                        progress_window.after(20, lambda: simulate_progress(step + 1))
                    else:
                        process_file()

                def process_file():
                    try:
                        # Perform file processing
                        df_read_point_24h = pd.read_excel(file_path)
                        app_data.df_v1_read_point_24h = df_read_point_24h.iloc[:, [0, 1, 2, 8, 9, 10]]
                        app_data.df_testing_24h = df_read_point_24h.iloc[:, 25:]
                        app_data.df_general_24h = pd.concat([app_data.df_v1_read_point_24h, app_data.df_testing_24h], axis=1)
                        app_data.df_analysis_24h = app_data.df_general_24h.copy()
                        app_data.file_uploaded = True
                        
                        # Close the progress bar window
                        progress_window.destroy()
                        messagebox.showinfo("Success", "File uploaded successfully!", parent=scrollable_frame)
                    except Exception as e:
                        progress_window.destroy()
                        messagebox.showerror("Error", f"File upload failed. Please ensure you upload the correct output file, Data Analysis sheet, generated after running the Macro. The file must be in '.xlsx' format and contain the required data structure.", parent=scrollable_frame)

                simulate_progress()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Retreive the test name by inserting the relative test number 
    def get_test_names(app_data):
        if not app_data.file_uploaded:
            messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
            return

        test_numbers_input = test_number_field.get().strip()
        if not test_numbers_input:  # Check if the field is empty
            messagebox.showerror("Error", "Please enter test number(s) separated by commas.", parent=scrollable_frame)
            return

        try:
            # Convert test numbers to integers
            test_numbers = [int(num.strip()) for num in test_numbers_input.split(',')]

            # Extract the test names and include the test numbers
            test_names = app_data.df_testing_24h.loc[0, test_numbers]

            # Combine test numbers and test names into a formatted string
            test_name_output = "Test Number, Test Name:\n"
            for test_number, test_name in zip(test_numbers, test_names):
                test_name_output += f"{test_number}, {test_name}\n"

            # Prompt the user to select a location to save the .txt file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")],
                title="Save Test Names As",
                parent=scrollable_frame
            )

            if save_path:
                # Save the test numbers and names to the .txt file
                with open(save_path, 'w') as file:
                    file.write(test_name_output)

                # Notify the user that the file has been saved
                messagebox.showinfo("Success", f"Test name text file (.txt) saved successfully at this path: {save_path}", parent=scrollable_frame)

                # Automatically open the .txt file
                os.startfile(save_path)

        except KeyError:
            messagebox.showerror("Error", "One or more test number(s) do not exist in the STDF file. Ensure you insert correct test number(s).", parent=scrollable_frame)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid test numbers separated by commas.", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Function to clear the Excel file entry and reset related variables
    def clear_excel_entry(app_data):
        try:
            # Clear all entry fields in the failure detection window
            entry_field.delete(0, tk.END)

            # Reset all variables and cache
            app_data.file_uploaded = False
            app_data.df_v1_read_point_24h = None
            app_data.df_testing_24h = None
            app_data.df_general_24h = None
            app_data.df_analysis_24h = None
            app_data.df_general_24h_help_1 = None
            app_data.df_general_24h_help_2 = None
            app_data.df_general_24h_help_3 = None
            app_data.df_limit_24h_generated = False
            app_data.df_limit_24h = None
            app_data.graph_24h = None
            app_data.failure_24h = None
            app_data.overview_generated = False
            app_data.graph_24h_generated = False
            
            # Close all matplotlib plots in the main thread
            plt.close('all')
        
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while clearing: {str(e)}", parent=scrollable_frame)
    
    # Function to clear the General View entries
    def clear_general_view():
        read_point_field.delete(0, tk.END)
        sheet_name_field.delete(0, tk.END)
    
    # Function to create general overview DataFrame and Excel file
    def generate_overview(app_data):
        if not app_data.file_uploaded:
            messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
            return

        try:
            # Validate that the dataframes are not empty
            if app_data.df_v1_read_point_24h is None or app_data.df_testing_24h is None:
                raise ValueError("The uploaded file does not contain valid data.")

            # Check if both fields are empty
            read_point = read_point_field.get().strip()
            sheet_name = sheet_name_field.get().strip()

            # Handle the three scenarios
            if not read_point and not sheet_name:  # If both fields are empty
                messagebox.showerror("Error", "Please fill in the related blank fields to generate the Overview Excel file.", parent=scrollable_frame)
                return
            
            if not read_point:
                messagebox.showerror("Error", "Please enter a number (integer), indicating read point number of the analysis.", parent=scrollable_frame)    
                return
            
            # Ensure the sheet name is provided
            if not sheet_name:  # If the sheet name is empty
                messagebox.showerror("Error", "Please provide a sheet name to generate the Overview Excel file.", parent=scrollable_frame)
                return

            # Combine dataframes and generate overview
            app_data.df_analysis_24h = pd.concat([app_data.df_v1_read_point_24h, app_data.df_testing_24h], axis=1)
            app_data.df_general_24h = app_data.df_analysis_24h.copy().iloc[3:, :]  # Update df_general_24h globally

            # Set the overview_generated flag to True
            app_data.overview_generated = True

            # Prompt user to select save location
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], parent=scrollable_frame)
            if not save_path:
                return

            # Create a progress bar window for AMAZON
            progress_window = tk.Toplevel(scrollable_frame)
            progress_window.title("Saving Overview")
            progress_window.iconbitmap(os.path.join(base_path, "resources", "amazon.ico"))
            progress_window.geometry("400x200")
            progress_window.configure(bg="#1C2541")

            top_frame = tk.Frame(progress_window, bg="#1C2541")
            top_frame.pack(pady=(10, 0))

            amazon_logo_path = os.path.join(base_path, "resources", "amazon.ico")
            try:
                amazon_logo_img = Image.open(amazon_logo_path)
                amazon_logo_img = amazon_logo_img.resize((48, 48), Image.Resampling.LANCZOS)
                amazon_logo_tk = ImageTk.PhotoImage(amazon_logo_img)
                logo_label = tk.Label(progress_window, image=amazon_logo_tk, bg="#1C2541")
                logo_label.image = amazon_logo_tk
                logo_label.pack(pady=(10, 0))
            except Exception:
                logo_label = tk.Label(progress_window, text="AMAZON", fg="white", bg="#1C2541", font=("Helvetica", 16, "bold"))
                logo_label.pack(pady=(10, 0))

            progress_label = tk.Label(
                progress_window,
                text="Saving the file, please wait...",
                bg="#1C2541",
                fg="white",
                font=("Helvetica", 12)
            )
            progress_label.pack(pady=10)

            progress_bar = ttk.Progressbar(
                progress_window,
                orient="horizontal",
                mode="determinate",
                length=250
            )
            progress_bar.pack(pady=10)

            def simulate_progress(step=1):
                if step <= 100:
                    progress_bar["value"] = step
                    progress_window.after(20, lambda: simulate_progress(step + 1))
                else:
                    save_overview()

            def save_overview():
                try:
                    app_data.df_general_24h.to_excel(save_path, sheet_name=sheet_name, index=False)
                    messagebox.showinfo("Success", f"Overview Excel file saved successfully at this path: {save_path}", parent=scrollable_frame)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save the Overview Excel file: {str(e)}", parent=scrollable_frame)
                finally:
                    progress_window.destroy()

            simulate_progress()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Function to generate the Read Point Graph Table and its Excel file
    def generate_read_point_graph(app_data):
        try:
            # Ensure the General View and Graph Table sections are completed
            if not app_data.file_uploaded or app_data.df_general_24h is None or not app_data.overview_generated or not app_data.graph_24h_generated:
                raise ValueError("Please ensure the Excel file is uploaded and both the 'Overview Generation' and 'Graph Table' steps are completed before proceeding.")

            # Get the test numbers from the Graph Table entry field
            test_numbers_input = graph_table_test_field.get().strip()
            if not test_numbers_input:
                raise ValueError("Please enter test numbers in the Graph Table section.")

            # Convert test numbers to strings
            test_numbers = [int(num.strip()) for num in test_numbers_input.split(',')]
            test_numbers_str = [str(num) for num in test_numbers]

            # Get the read point number from the General View entry field
            read_point = read_point_field.get().strip()
            if not read_point:
                raise ValueError("Please enter the read point number in the General View section.")

            # Create a copy of graph_24h and rename its columns
            graph_read_point_help = app_data.graph_24h.copy()
            graph_read_point_help.columns = graph_read_point_help.columns.astype(str)  # Convert column names to strings
            graph_read_point = graph_read_point_help.rename(columns={num: f"{num}_{read_point}h" for num in test_numbers_str})

            # Cache the graph_read_point with a unique name
            if not hasattr(app_data, "read_point_graphs"):
                app_data.read_point_graphs = []
            app_data.read_point_graphs.append(graph_read_point)

            # Inform the user about the creation of the graph
            graph_number = len(app_data.read_point_graphs)
            messagebox.showinfo(
                "Success",
                f"{graph_number} Read Point Graph Table(s) is created.",
                parent=scrollable_frame
            )

            # Ask the user if they want to save the Excel file
            user_response = messagebox.askyesno(
                "Save Excel File",
                "Do you want to generate the Excel file for the Read Point?",
                parent=scrollable_frame
            )
            if user_response:
                # Prompt the user to save the Excel file
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Read Point Graph As",
                    parent=scrollable_frame
                )
                if save_path:
                    # Save the DataFrame to Excel with the sheet name "Read_Point"
                    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                        graph_read_point.to_excel(writer, index=False, sheet_name="Read_Point")
                    messagebox.showinfo("Success", f"Read Point Excel file saved successfully at this path: {save_path}", parent=scrollable_frame)

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", "Please ensure both the 'General View' and 'Graph Table' Excel files are generated before proceeding.", parent=scrollable_frame)
    
    # Create the "Read Point" button
    read_point_graph_button = tk.Button(
        scrollable_frame,
        text="Read Point",
        command=lambda: generate_read_point_graph(app_data),
        bg="#1C2541",
        fg="white",
        width=9,
        height= 1
    )
    read_point_graph_button.grid(row=3, column=8, padx=4, pady=10)

    # Button for Test Name
    get_test_name_button = tk.Button(scrollable_frame, text="Test Name", command=lambda: get_test_names(app_data), bg="#1C2541", fg="white", width=10)
    get_test_name_button.grid(row=1, column=3, padx=10, pady=10)

    # Button for Overview Generation
    overview_button = tk.Button(scrollable_frame, text="Overview Generation", command=lambda: generate_overview(app_data), bg="#1C2541", fg="white", width=20)
    overview_button.grid(row=2, column=6, padx=10, pady=10)

    # Button for Uploading the Excel file
    upload_button = tk.Button(scrollable_frame, text="Upload", command=lambda: upload_file(app_data), bg="#1C2541", fg="white", width=10)
    upload_button.grid(row=0, column=3, padx=10, pady=10)

    # Clear button for Test Number entry field
    clear_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_result(app_data), bg="#1C2541", fg="white", width=10)
    clear_button.grid(row=1, column=4, padx=10, pady=10)

    def clear_result(app_data):
        result_label.config(text="")
        test_number_field.delete(0, tk.END)

    # Clear button for Excel file entry field
    clear_excel_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_excel_entry(app_data), bg="#1C2541", fg="white", width=10)
    clear_excel_button.grid(row=0, column=4, padx=10, pady=10)

    # Clear button for General View field
    clear_general_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_general_view(app_data), bg="#1C2541", fg="white", width=10)
    clear_general_button.grid(row=2, column=5, padx=10, pady=10)

    def clear_general_view(app_data):
        read_point_field.delete(0, tk.END)
        sheet_name_field.delete(0, tk.END)

    # Label to display the graph table label
    result_label = tk.Label(scrollable_frame, text="", bg="#1C2541", fg="white")
    result_label.grid(row=3, column=0, columnspan=7, padx=10, pady=10)
    
    # Add the "4. Graph Table" label
    graph_table_label = tk.Label(scrollable_frame, text="4. Graph Table:", bg="#1C2541", fg="white")
    graph_table_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")

    # Add the first entry field for test number
    graph_table_test_field = tk.Entry(scrollable_frame, width=30)
    graph_table_test_field.grid(row=3, column=1, padx=10, pady=10)

    # Add the first tooltip for test number
    graph_table_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    graph_table_test_tooltip_label.grid(row=3, column=2, padx=10, pady=10)

    # Function to show the tooltip for test number for graph table section
    graph_table_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Insert the test number(s) (integers) you wish to analyze.", scrollable_frame)
    )
    graph_table_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the second entry field for sheet name
    graph_table_sheet_field = tk.Entry(scrollable_frame, width=30)
    graph_table_sheet_field.grid(row=3, column=3, padx=10, pady=10)

    # Add the second tooltip for sheet name
    graph_table_sheet_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    graph_table_sheet_tooltip_label.grid(row=3, column=4, padx=10, pady=10)

    # Function to show the tooltip for sheet name
    graph_table_sheet_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the sheet name for the Excel file you are creating (e.g., graph_view).", scrollable_frame)
    )
    graph_table_sheet_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Clear button creation for the Graph Table
    clear_graph_table_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_graph_table_fields(app_data), bg="#1C2541", fg="white", width=10)
    clear_graph_table_button.grid(row=3, column=5, padx=0, pady=0)

    # Function definition for the Clear button belong to Graph Table section (step 4)
    def clear_graph_table_fields(app_data):
        graph_table_test_field.delete(0, tk.END)
        graph_table_sheet_field.delete(0, tk.END)

    # Reset the relevant variables in app_data (class definition)
    app_data.df_general_24h_help_1 = None
    app_data.df_general_24h_help_2 = None
    app_data.df_general_24h_help_3 = None
    app_data.graph_24h = None

    # Function definition for generating graph table dataset (preview)
    def generate_graph_table_preview(app_data):
        try:
            # Ensure if the Excel file uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return
            
            # Ensure if the user has already created the Overview dataset of the STDF file
            if app_data.df_general_24h is None or not hasattr(app_data, "overview_generated") or not app_data.overview_generated:
                messagebox.showerror("Error", "Please complete the Overview Generation step first.", parent=scrollable_frame)
                return

            # Validate test number input
            test_number = graph_table_test_field.get()
            if not test_number.strip():  # Check if the field is empty or contains only whitespace
                messagebox.showerror("Error", "Please fill in the test number field with valid test number(s), separated by commas.", parent=scrollable_frame)
                return

            # Parse test numbers into a list of integers
            test_numbers = [int(num.strip()) for num in test_number.split(',')]

            # Validate sheet name input and its related error handling part
            sheet_name = graph_table_sheet_field.get()
            if not sheet_name.strip():  # Check if the sheet name field is empty or contains only whitespace
                messagebox.showerror("Error", "Please specify the sheet name for the Graph Table Excel file to be generated.", parent=scrollable_frame)
                return

            # Generate the required dataframes
            app_data.df_general_24h_help_1 = app_data.df_general_24h.iloc[2:, :]  # Slice the dataframe
            app_data.df_general_24h_help_2 = app_data.df_general_24h_help_1.iloc[:, :6]

            # Validate that all test numbers exist in the dataframe's columns
            invalid_test_numbers = [num for num in test_numbers if num not in app_data.df_general_24h_help_1.columns]
            if invalid_test_numbers:
                raise KeyError(f"Invalid test numbers: {', '.join(map(str, invalid_test_numbers))}")

            # Extract the corresponding columns for all valid test numbers
            app_data.df_general_24h_help_3 = app_data.df_general_24h_help_1.loc[:, test_numbers]

            # Combine dataframes and add the "Hatrick" column
            app_data.graph_24h = pd.concat([app_data.df_general_24h_help_2, app_data.df_general_24h_help_3], axis=1)
            app_data.graph_24h['Hatrick'] = app_data.graph_24h[150000112].astype(str) + '_' + app_data.graph_24h[150000113].astype(str) + '_' + app_data.graph_24h[150000114].astype(str)

            # Set the flag to indicate that the graph_24h dataset has been generated
            app_data.graph_24h_generated = True

            # Prompt user to select save location
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], parent=scrollable_frame)
            if save_path:
                app_data.graph_24h.to_excel(save_path, sheet_name=sheet_name, index=False)
                messagebox.showinfo("Success", f"Graph Table Excel file saved successfully at this path: {save_path}", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", "You inserted invalid test number(s).", parent=scrollable_frame)
        except ValueError as e:
            messagebox.showerror("Error", f"{str(e)}", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate Graph Table Excel File.", parent=scrollable_frame)

    # Graph_table button creation
    graph_table_preview_button = tk.Button(scrollable_frame, text="Graph Table", command=lambda: generate_graph_table_preview(app_data), bg="#1C2541", fg="white", width=20, height= 1)
    graph_table_preview_button.grid(row=3, column=6, padx=10, pady=10)

    # Function to generate the complete table
    def generate_complete_table(app_data):
        try:
            # Check if the file has been uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Get the test numbers from the first entry field of the "Graph Table" label
            test_numbers_input = graph_table_test_field.get().strip()
            if not test_numbers_input:  # Check if the test_number entry field is empty
                messagebox.showerror("Error", "Please fill in the test number entry field first.", parent=scrollable_frame)
                return

            # Check if test number and sheet name entry fields are empty or not
            if not test_numbers_input and not sheet_name:
                messagebox.showerror("Error", "Please fill in all required entry fields first.", parent=scrollable_frame)
                return

            # Validate sheet name input
            sheet_name= graph_table_sheet_field.get().strip()
            if not sheet_name: # Check if the sheet name field is empty to show an error massage
                messagebox.showerror("Error", "Please specify the sheet name for the Complete Table Excel file to be generated.", parent=scrollable_frame)
                return
            
            # Parse the test numbers into a list of integers
            test_numbers = [int(num.strip()) for num in test_numbers_input.split(',')]
            
            # Generate the required dataframes
            df_limit_24h_help_1 = pd.concat([app_data.df_analysis_24h.iloc[[0]], app_data.df_analysis_24h.iloc[3:]])
            df_limit_24h_help_1.at[0, 'PID_Number'] = 'Test_Name'
            df_limit_24h_help_1.at[0, 150000113] = ''
            df_limit_24h_help_1.at[0, 150000112] = ''
            df_limit_24h_help_1.at[0, 150000114] = ''
            df_limit_24h_help_2 = df_limit_24h_help_1.reset_index(drop=True)
            df_limit_24h_help_3 = df_limit_24h_help_2.iloc[:, :6]
            df_limit_24h_help_3['Hatrick'] = (
                df_limit_24h_help_3[150000112].astype(str) + '_' +
                df_limit_24h_help_3[150000113].astype(str) + '_' +
                df_limit_24h_help_3[150000114].astype(str)
            )
            df_limit_24h_help_3.at[0, 'Hatrick'] = ''
            df_limit_24h_help_3.at[1, 'Hatrick'] = ''
            df_limit_24h_help_3.at[2, 'Hatrick'] = ''

            # Select the test numbers dynamically based on user input
            df_limit_24h_help_4 = df_limit_24h_help_2.loc[:, test_numbers]

            # Combine the dataframes
            app_data.df_limit_24h = pd.concat([df_limit_24h_help_3, df_limit_24h_help_4], axis=1)
            app_data.df_limit_24h_generated = True
            
            # Open a "Save As" dialog to let the user specify the file name and location
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Complete Table As",
                parent=scrollable_frame
            )
            if not save_path:  # If the user cancels the save dialog
                return

            # Save the file to the specified path
            app_data.df_limit_24h.to_excel(save_path, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", f"Complete Table Excel file (with limitation) saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", "You inserted invalid test number(s).", parent=scrollable_frame)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid test numbers separated by commas.", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Function for generating the limit and test_name table of each read point for AMAZON
    def limit_with_test_name(app_data):
        try:
            # Ensure the Excel file is uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return

            # Ensure the Complete Table is generated
            if app_data.df_limit_24h is None or not getattr(app_data, "df_limit_24h_generated", False):
                messagebox.showerror("Error", "Please generate the Complete Table step first to be able to create 'Analysis' dataset.", parent=scrollable_frame)
                return

            # Validate test number input
            test_number = graph_table_test_field.get()
            if not test_number.strip():
                messagebox.showerror("Error", "Please fill in the test number entry field with valid test number(s), separated by commas.", parent=scrollable_frame)
                return

            # Parse test numbers into a list of integers
            test_numbers = [int(num.strip()) for num in test_number.split(',')]

            # Validate sheet name input
            sheet_name = graph_table_sheet_field.get()
            if not sheet_name.strip():
                messagebox.showerror("Error", "Please specify the sheet name for the Limit Excel file with only limits and test names to be generated.", parent=scrollable_frame)
                return

            # Get the read point number (number_h) from the General View entry field
            read_point = read_point_field.get().strip()
            if not read_point:
                messagebox.showerror("Error", "Please enter the read point number (only an integer) in the 'General View' section.", parent=scrollable_frame)
                return

            # Data manipulation using Pandas
            complete_table = app_data.df_limit_24h.copy()  # To avoid change in this dataset (df_limit_24h)
            complete_table.at[0, 'Hatrick'] = 'Test_Name'
            complete_table.at[1, 'Hatrick'] = 'HighL'
            complete_table.at[2, 'Hatrick'] = 'LowL'
            start_idx = complete_table.columns.get_loc('Hatrick')
            filtered = complete_table.loc[:, complete_table.columns[start_idx:]]
            filtered_limit = filtered[filtered['Hatrick'].isin(['Test_Name', 'HighL', 'LowL'])]

            # Rename columns after 'Hatrick' column
            new_columns = list(filtered_limit.columns)
            for i, col in enumerate(new_columns):
                if i == 0:
                    continue  # 'Hatrick' column
                try:
                    col_int = int(col)
                    new_columns[i] = f"{col_int}_{read_point}h"
                except Exception:
                    pass  # leave as is if not integer
            filtered_limit.columns = new_columns

            # Cache the analysis table
            if not hasattr(app_data, "analysis_point_graphs"):
                app_data.analysis_point_graphs = []
            app_data.analysis_point_graphs.append(filtered_limit)

            # Inform the user about the creation of the analysis table
            analysis_number = len(app_data.analysis_point_graphs)
            messagebox.showinfo(
                "Success",
                f"{analysis_number} Limit Table(s) with only test names and limits either HighL or LowL are created.",
                parent=scrollable_frame
            )

            # Ask the user if they want to save the Excel file
            user_response = messagebox.askyesno(
                "Save Excel File",
                "Do you want to generate the Excel file including only the test names and limits?",
                parent=scrollable_frame
            )
            if user_response:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Analysis Table As",
                    parent=scrollable_frame
                )
                if save_path:
                    filtered_limit.to_excel(save_path, sheet_name=sheet_name, index=False)
                    messagebox.showinfo("Success", f"Excel file (with only limitation and test names) saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", "You inserted invalid test number(s). Make sure you select the same test number(s) you have already used to generate 'Complete Table' and 'Graph Table'.", parent=scrollable_frame)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid test numbers separated by commas.", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Function for Full Limit of AMAZON Products
    def full_limit_amazon(app_data):
        try:
            # Ensure we have enough analysis_point_graphs datasets to merge
            if not app_data.analysis_point_graphs or len(app_data.analysis_point_graphs) < 2:
                raise ValueError("Please generate at least two Limit Tables or datasets before creating the Full Limit Excel file.")

            # Merging DataFrames
            merged = app_data.analysis_point_graphs[0]
            for df in app_data.analysis_point_graphs[1:]:
                merged = merged.merge(df, on="Hatrick", how="inner")

            # Cache (store) the generated DataFrame
            app_data.full_limit_amazon_merged = merged
            app_data.full_limit_amazon_generated = True

            # Ask the user to save the merged file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Full Limit Table As",
                parent=scrollable_frame
            )
            if save_path:
                merged.to_excel(save_path, index=False, sheet_name='full_limit')
                messagebox.showinfo("Success", f"Full Limit Table Excel file saved successfully at this path: {save_path}.", parent=scrollable_frame)
        except ValueError as ve:
            messagebox.showerror("Error", f"{str(ve)}", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"{str(e)}", parent=scrollable_frame)   
    
    # Function for graph with HighL and LowL for AMAZON
    def plot_full_graph_amazon(app_data):
        try:
            # 1. Check if excessive units dataset is generated
            if not getattr(app_data, "excessive_units_generated", False):
                messagebox.showerror(
                    "Error",
                    "You need to first generate the 'Excessive Units' dataset before creating the Full Graph with HighL and LowL.",
                    parent=scrollable_frame
                )
                return

            # 2. Check if Full Limit dataset is generated (merged limit tables)
            merged = getattr(app_data, "full_limit_amazon_merged", None)
            if merged is None or isinstance(merged, bool):
                messagebox.showerror(
                    "Error",
                    "Please generate the 'Full Limit' dataset (all read points with HighL, LowL, Test_Name) before proceeding.",
                    parent=scrollable_frame
                )
                return

            # 3. Get user inputs (AMAZON fields already created in section 9)
            test_number = final_chart_test_field.get().strip()
            colors = final_chart_color_field.get().strip()
            labels = final_chart_read_points_field.get().strip()
            title = final_chart_title_field.get().strip()
            ylabel = final_chart_ylabel_field.get().strip()
            width = final_chart_width_field.get().strip()
            height = final_chart_height_field.get().strip()

            if not all([test_number, colors, labels, title, ylabel, width, height]):
                messagebox.showerror("Error", "Please ensure all required entry fields are filled before proceeding.", parent=scrollable_frame)
                return

            try:
                width = int(width); height = int(height)
                if width <= 0 or height <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid positive number to specify the width and height of your line chart.", parent=scrollable_frame)
                return

            # 4. Subset merged for test number columns (e.g., 2010_24h)
            test_number_str = str(test_number)
            matching_cols = [col for col in merged.columns if col.startswith(test_number_str + '_')]
            if not matching_cols:
                messagebox.showerror("Input Error", f"No columns found for test number {test_number}.", parent=scrollable_frame)
                return

            color_list = [c.strip() for c in colors.split(',')]
            if len(color_list) != len(matching_cols):
                messagebox.showerror("Input Error", "The number of colors must match the number of read points.", parent=scrollable_frame)
                return

            label_list = [l.strip() for l in labels.split(',')]
            if len(label_list) != len(matching_cols):
                messagebox.showerror("Error", "The number of labels must match the number of read points.", parent=scrollable_frame)
                return

            subset_limit = merged[['Hatrick'] + matching_cols].copy()

            # 5. Merge vertically with excessive units DataFrame
            excessive_df = getattr(app_data, "excessive_units", None)
            if excessive_df is None:
                messagebox.showerror("Error", "Excessive units dataset not found.", parent=scrollable_frame)
                return

            subset_limit = subset_limit.reset_index(drop=True)
            excessive_df = excessive_df.reset_index(drop=True)

            # Ensure same columns ordering as excessive_df
            subset_limit = subset_limit.loc[:, excessive_df.columns]

            charting_limit = pd.concat([subset_limit, excessive_df], axis=0, ignore_index=True)

            # 6. Find highest hour column
            def extract_hour(col):
                try:
                    return int(col.split('_')[1].replace('h', ''))
                except Exception:
                    return -1
            highest_column = max(matching_cols, key=extract_hour)

            # 7. Filter for plotting
            plot_df_amazon = charting_limit[
                (charting_limit['Hatrick'] != 'Test_Name') &
                (charting_limit[highest_column].notna())
            ].copy()

            app_data.plot_df_amazon = plot_df_amazon
            app_data.plot_df_amazon_generated = True

            # 8. Plot
            plt.figure(figsize=(width, height))
            for col, color, label in zip(matching_cols, color_list, label_list):
                mask = ~plot_df_amazon['Hatrick'].isin(['HighL', 'LowL'])
                x = plot_df_amazon.loc[mask, 'Hatrick']
                y = pd.to_numeric(plot_df_amazon.loc[mask, col].replace(' ', np.nan), errors='coerce')
                plt.plot(x, y, marker='o', linestyle='-', color=color, label=label)

            # HighL / LowL horizontal lines
            for col in matching_cols:
                highl = charting_limit.loc[charting_limit['Hatrick'] == 'HighL', col].replace(' ', np.nan)
                lowl = charting_limit.loc[charting_limit['Hatrick'] == 'LowL', col].replace(' ', np.nan)
                if not highl.empty and pd.notna(highl.values[0]):
                    plt.axhline(y=float(highl.values[0]), color='red', linestyle='--', label=f'HighL {col}')
                if not lowl.empty and pd.notna(lowl.values[0]):
                    plt.axhline(y=float(lowl.values[0]), color='darkred', linestyle='--', label=f'LowL {col}')

            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Full Graph As PDF",
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success",
                                    f"The line graph of all read points, including HighL and LowL saved successfully at this path: {save_path}",
                                    parent=scrollable_frame)
            plt.close()

            # Save full dataset (with Test_Name, HighL, LowL)
            user_response = messagebox.askyesno(
                "Save Full Dataset",
                "Do you want to save the full dataset (with Test_Name, HighL, LowL, and all Hatricks) as an Excel file?",
                parent=scrollable_frame
            )
            if user_response:
                excel_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Full Dataset As Excel",
                    parent=scrollable_frame
                )
                if excel_path:
                    charting_limit.to_excel(excel_path, index=False, sheet_name='final')
                    messagebox.showinfo("Success",
                                        f"Excel file (full dataset of all read points) saved successfully at this path: {excel_path}",
                                        parent=scrollable_frame)

            # Save dataset without Test_Name row
            user_response_plot = messagebox.askyesno(
                "Save Dataset Without Test_Name",
                "Do you want to save the dataset without the Test_Name row (includes HighL, LowL, and all data rows) as an Excel file?",
                parent=scrollable_frame
            )
            if user_response_plot:
                excel_path_plot = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Dataset Without Test_Name As Excel",
                    parent=scrollable_frame
                )
                if excel_path_plot:
                    plot_df_amazon.to_excel(excel_path_plot, index=False, sheet_name='final')
                    messagebox.showinfo("Success",
                                        f"Excel file (without Test_Name row) saved successfully at this path: {excel_path_plot}",
                                        parent=scrollable_frame)

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Full distribution of all read points for AMAZON
    def full_distribution_amazon(app_data):
        try:
            # 1. Check if plot_df_amazon dataset has been created and is valid
            plot_df_amazon = getattr(app_data, "plot_df_amazon", None)
            if plot_df_amazon is None or plot_df_amazon.empty:
                messagebox.showerror(
                    "Error",
                    "Please first complete the 'Full Graph' section by clicking the corresponding button to generate the line graph with HighL and LowL for all read points.",
                    parent=scrollable_frame
                )
                return

            # 2. Check if all histogram entry fields are filled (step 6: Histogram)
            test_number_input = distribution_test_field.get().strip()
            bin_number_input = distribution_bins_field.get().strip()
            histogram_title = distribution_title_field.get().strip()
            if not all([test_number_input, bin_number_input, histogram_title]):
                messagebox.showerror(
                    "Error",
                    "Please fill in all required entry fields in step '6' to generate the full distribution chart of all read points.",
                    parent=scrollable_frame
                )
                return

            # 3. Validate integer fields
            try:
                test_number = int(test_number_input)
                bin_number = int(bin_number_input)
                if bin_number <= 0 or test_number <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "Error",
                    "Test number and number of bins must be positive integers.",
                    parent=scrollable_frame
                )
                return

            # 4. Find all columns in plot_df_amazon matching the test number (format: '2010_24h', etc.)
            matching_cols = [col for col in plot_df_amazon.columns if col.startswith(f"{test_number}_")]
            if not matching_cols:
                messagebox.showerror(
                    "Error",
                    "Please enter a valid test number you have already used in step '9' (Final Chart) to generate the Full Distribution chart of all read points.",
                    parent=scrollable_frame
                )
                return

            # 5. Aggregate values across all matching read points (excluding HighL/LowL rows)
            plt.figure(figsize=(10, 6))
            all_values = []
            highl_values = []
            lowl_values = []
            for col in matching_cols:
                mask = ~plot_df_amazon['Hatrick'].isin(['HighL', 'LowL'])
                values = pd.to_numeric(plot_df_amazon.loc[mask, col], errors='coerce').dropna()
                all_values.extend(values)

                # HighL / LowL for the specific column
                highl = plot_df_amazon.loc[plot_df_amazon['Hatrick'] == 'HighL', col]
                lowl = plot_df_amazon.loc[plot_df_amazon['Hatrick'] == 'LowL', col]
                highl_values.extend(pd.to_numeric(highl, errors='coerce').dropna())
                lowl_values.extend(pd.to_numeric(lowl, errors='coerce').dropna())

            # Plot histogram for all values
            plt.hist(all_values, bins=bin_number, alpha=0.7, label=histogram_title)

            # 6. Mean excluding HighL/LowL rows
            mean_value = np.mean(all_values) if all_values else np.nan

            # 7. Plot vertical lines (HighL, LowL, Average)
            for hl in highl_values:
                plt.axvline(x=hl, color='red', linestyle='--', label=f'HighL: {hl}')
            for ll in lowl_values:
                plt.axvline(x=ll, color='darkred', linestyle='--', label=f'LowL: {ll}')
            if not np.isnan(mean_value):
                plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')

            plt.title(histogram_title)
            plt.xlabel('Test Value')
            plt.ylabel('Frequency')
            plt.legend()
            plt.tight_layout()

            # Ensure all lines are visible
            all_x = all_values + highl_values + lowl_values + ([mean_value] if not np.isnan(mean_value) else [])
            if all_x:
                plt.xlim(min(all_x) - 1, max(all_x) + 1)

            # Save the chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo(
                    "Success",
                    f"Distribution chart with HighL and LowL of all read points saved successfully at this path: {save_path}",
                    parent=scrollable_frame
                )
            plt.close()

            # 8. Ask the user if they want a focused distribution (excluding HighL / LowL markers)
            user_response = messagebox.askyesno(
                "Generate Focused Distribution",
                "HighL and LowL may distort the chart. Would you like to generate a distribution chart of all read points excluding HighL and LowL (showing only the average and actual data)?",
                parent=scrollable_frame
            )
            if user_response:
                plt.figure(figsize=(10, 6))
                plt.hist(all_values, bins=bin_number, alpha=0.7, label=histogram_title + " (No HighL/LowL)")
                if not np.isnan(mean_value):
                    plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')
                plt.title(histogram_title + " (No HighL/LowL)")
                plt.xlabel('Test Value')
                plt.ylabel('Frequency')
                plt.legend()
                plt.tight_layout()
                if all_x:
                    plt.xlim(min(all_x) - 1, max(all_x) + 1)
                save_path_focused = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")],
                    title="Save Focused Distribution Chart As PDF",
                    parent=scrollable_frame
                )
                if save_path_focused:
                    plt.savefig(save_path_focused)
                    messagebox.showinfo(
                        "Success",
                        f"Focused distribution chart saved successfully at this path: {save_path_focused}",
                        parent=scrollable_frame
                    )
                plt.close()

            # 9. Cache dataset & flag (optional)
            app_data.full_distribution_amazon_df = plot_df_amazon.copy()
            app_data.full_distribution_amazon_generated = True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot the full distribution chart: {str(e)}", parent=scrollable_frame)
    
    # Function for Distribution Plot with Limits for single read points of AMAZON
    def dist_amazon_single_read_point(app_data):
        try:
            # Check if the Excel file has been uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Check if the complete table with limits is generated
            if not getattr(app_data, "df_limit_24h_generated", False):
                messagebox.showerror(
                    "Error",
                    "Please first generate the complete table that includes the limits to be able to have a distribution chart with HighL and LowL.",
                    parent=scrollable_frame
                )
                return

            # Retrieve all Histogram (Step 6) entry field values
            test_number_input = distribution_test_field.get().strip()
            bins_input = distribution_bins_field.get().strip()
            histogram_title = distribution_title_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, bins_input, histogram_title]):
                messagebox.showerror("Error", "Please fill all required entry fields before proceeding.", parent=scrollable_frame)
                return

            # Validate integer fields
            try:
                test_number = int(test_number_input)
                bins = int(bins_input)
                if bins <= 0 or test_number <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Test number and number of bins must be positive integers.", parent=scrollable_frame)
                return

            # Additional check: Ensure the test number exists in the dataset for plotting
            df_limit_24h = app_data.df_limit_24h
            if test_number not in df_limit_24h.columns:
                messagebox.showerror(
                    "Error",
                    "Please use a test number that you have already used to generate the table with limits in step 4 by pressing 'Complete Table' button.",
                    parent=scrollable_frame
                )
                return

            # Data Manipulation and copy creation to avoid modifying the original dataset
            df_single_limit_plot_1 = df_limit_24h.copy()
            df_single_limit_plot_1.at[1, 'Hatrick'] = 'HighL'
            df_single_limit_plot_1.at[2, 'Hatrick'] = 'LowL'

            # Get all columns from 'Hatrick' onwards
            start_col = df_single_limit_plot_1.columns.get_loc('Hatrick')
            df_single_limit_plot_2 = df_single_limit_plot_1.iloc[1:, start_col:].reset_index(drop=True)

            # Replace ' ' with np.nan in the first and second rows (index 0 and 1 after slicing)
            df_single_limit_plot_2.iloc[0] = df_single_limit_plot_2.iloc[0].replace(' ', np.nan)
            df_single_limit_plot_2.iloc[1] = df_single_limit_plot_2.iloc[1].replace(' ', np.nan)

            # Prepare values for histogram (exclude HighL / LowL rows)
            mask = ~df_single_limit_plot_2['Hatrick'].isin(['HighL', 'LowL'])
            all_values = pd.to_numeric(df_single_limit_plot_2.loc[mask, test_number], errors='coerce').dropna().tolist()
            highl_values = pd.to_numeric(df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'HighL', test_number], errors='coerce').dropna().tolist()
            lowl_values = pd.to_numeric(df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'LowL', test_number], errors='coerce').dropna().tolist()

            # Plot histogram for all values
            plt.figure(figsize=(10, 6))
            plt.hist(all_values, bins=bins, alpha=0.7, label=histogram_title)

            # Calculate mean (average) excluding HighL and LowL rows
            if all_values:
                mean_value = np.mean(all_values)
            else:
                messagebox.showerror("Error", "No valid numeric data available to plot the distribution.", parent=scrollable_frame)
                plt.close()
                return

            # Plot vertical lines for HighL, LowL, and mean
            for hl in highl_values:
                plt.axvline(x=hl, color='red', linestyle='--', label=f'HighL: {hl}')
            for ll in lowl_values:
                plt.axvline(x=ll, color='darkred', linestyle='--', label=f'LowL: {ll}')
            plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')

            plt.title(histogram_title)
            plt.xlabel('Test Value')
            plt.ylabel('Frequency')
            plt.legend()
            plt.tight_layout()

            # Ensure all lines are visible
            all_x = all_values + highl_values + lowl_values + [mean_value]
            plt.xlim(min(all_x) - 1, max(all_x) + 1)

            # Save the chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Distribution chart with HighL and LowL for a single read point saved successfully at this path: {save_path}", parent=scrollable_frame)
            plt.close()

            # Ask the user if they want to generate a chart excluding HighL and LowL
            user_response = messagebox.askyesno(
                "Generate Focused Distribution",
                "HighL and LowL may distort the chart. Would you like to generate a distribution chart excluding HighL and LowL (showing only the average and actual data)?",
                parent=scrollable_frame
            )
            if user_response:
                plt.figure(figsize=(10, 6))
                plt.hist(all_values, bins=bins, alpha=0.7, label=histogram_title + " (No HighL/LowL)")
                plt.axvline(x=mean_value, color='blue', linestyle='--', label=f'Average: {mean_value:.2f}')
                plt.title(histogram_title + " (No HighL/LowL)")
                plt.xlabel('Test Value')
                plt.ylabel('Frequency')
                plt.legend()
                plt.tight_layout()
                plt.xlim(min(all_values) - 1, max(all_values) + 1)
                save_path_focused = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")],
                    title="Save Focused Distribution Chart As PDF",
                    parent=scrollable_frame
                )
                if save_path_focused:
                    plt.savefig(save_path_focused)
                    messagebox.showinfo("Success", f"Focused distribution chart saved successfully at this path: {save_path_focused}", parent=scrollable_frame)
                plt.close()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot the distribution chart: {str(e)}", parent=scrollable_frame)    
    
    # Function for plot the line graph of a single read point with HighL and LowL for AMAZON
    def plot_line_limit_chart_amazon(app_data):
        try:
            # Check if the Excel file has been uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Check if the Complete Table (with limits) is generated
            if not getattr(app_data, "df_limit_24h_generated", False) or app_data.df_limit_24h is None:
                messagebox.showerror(
                    "Error",
                    "Please first generate the Complete Table that includes the limits to be able to have a line graph with HighL and LowL for a single read point.",
                    parent=scrollable_frame
                )
                return

            # Retrieve all field values (Step 5 + General View)
            test_number_input = line_chart_test_field.get().strip()
            read_point = read_point_field.get().strip()
            color = line_chart_color_field.get().strip()
            legend = line_chart_legend_field.get().strip()
            title = line_chart_title_field.get().strip()
            ylabel = line_chart_ylabel_field.get().strip()
            width_input = line_chart_width_field.get().strip()
            height_input = line_chart_height_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, read_point, color, legend, title, ylabel, width_input, height_input]):
                messagebox.showerror("Error", "Please fill all required entry fields before proceeding.", parent=scrollable_frame)
                return

            # Validate numeric fields
            try:
                test_number = int(test_number_input)
                width = int(width_input)
                height = int(height_input)
            except ValueError:
                messagebox.showerror("Error", "Test number, width, and height must be valid integers.", parent=scrollable_frame)
                return

            if width <= 0 or height <= 0:
                messagebox.showerror("Error", "Please enter a valid positive number to specify the width and height of your chart.", parent=scrollable_frame)
                return

            # Ensure the test number exists in the Complete Table with limits
            df_limit_24h = app_data.df_limit_24h
            if test_number not in df_limit_24h.columns:
                messagebox.showerror(
                    "Error",
                    "Please enter a test number you already used to generate the Complete Table (Step 4).",
                    parent=scrollable_frame
                )
                return

            # Data manipulation (avoid altering original)
            df_single_limit_plot_1 = df_limit_24h.copy()
            # Ensure the key rows are labeled
            df_single_limit_plot_1.at[1, 'Hatrick'] = 'HighL'
            df_single_limit_plot_1.at[2, 'Hatrick'] = 'LowL'

            # Isolate from 'Hatrick' onward, drop row 0 (Test_Name row)
            start_col = df_single_limit_plot_1.columns.get_loc('Hatrick')
            df_single_limit_plot_2 = df_single_limit_plot_1.iloc[1:, start_col:].reset_index(drop=True)

            # Clean HighL / LowL rows
            df_single_limit_plot_2.iloc[0] = df_single_limit_plot_2.iloc[0].replace(' ', np.nan)
            df_single_limit_plot_2.iloc[1] = df_single_limit_plot_2.iloc[1].replace(' ', np.nan)

            # Prepare plotting data (exclude HighL / LowL from X)
            mask = ~df_single_limit_plot_2['Hatrick'].isin(['HighL', 'LowL'])
            x = df_single_limit_plot_2.loc[mask, 'Hatrick']
            y = pd.to_numeric(df_single_limit_plot_2.loc[mask, test_number].replace(' ', np.nan), errors='coerce')

            plt.figure(figsize=(width, height))
            plt.plot(x, y, marker='o', linestyle='-', color=color, label=legend)

            # Plot HighL / LowL horizontal lines
            highl = df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'HighL', test_number].replace(' ', np.nan)
            lowl = df_single_limit_plot_2.loc[df_single_limit_plot_2['Hatrick'] == 'LowL', test_number].replace(' ', np.nan)
            if not highl.empty and pd.notna(highl.values[0]):
                plt.axhline(y=float(highl.values[0]), color='red', linestyle='--', label=f'HighL {test_number}')
            if not lowl.empty and pd.notna(lowl.values[0]):
                plt.axhline(y=float(lowl.values[0]), color='darkred', linestyle='--', label=f'LowL {test_number}')

            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Save chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Line chart with HighL and LowL for a single read point saved successfully at this path: {save_path}.", parent=scrollable_frame)
            plt.close()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    
    
    
    
    
    
    
    
    
    
    
    # Add the "Complete Table" button next to the "Graph Table Preview" button
    complete_table_button = tk.Button(scrollable_frame, text="Complete Table", command=lambda: generate_complete_table(app_data), bg="#1C2541", fg="white", width=13)
    complete_table_button.grid(row=3, column=7, padx=2, pady=10)

    # Add the "5. Line Chart" label
    line_chart_label = tk.Label(scrollable_frame, text="5. Line Chart:", bg="#1C2541", fg="white")
    line_chart_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")

    # Entry field for test number (1st entry field)
    line_chart_test_field = tk.Entry(scrollable_frame, width=30)
    line_chart_test_field.grid(row=4, column=1, padx=10, pady=10)

    # Tooltip for the test number field
    line_chart_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_test_tooltip_label.grid(row=4, column=2, padx=10, pady=10)

    # Function definition for the tooltip of the test number for the Line Chart (Step 5)
    line_chart_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number you wish to plot. Ensure you select the same test number(s) as those chosen in step 4.", scrollable_frame)
    )
    line_chart_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Entry field for line color (2nd entry field)
    line_chart_color_field = tk.Entry(scrollable_frame, width=30)
    line_chart_color_field.grid(row=4, column=3, padx=10, pady=10)

    # Tooltip for the line color field
    line_chart_color_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_color_tooltip_label.grid(row=4, column=4, padx=10, pady=10)

    line_chart_color_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Insert one of these colors: 'darkred', 'olive', 'blue', 'darkorange', 'purple', 'brown', 'black', 'goldenrod', 'teal', 'darkmagenta', 'peru'.", scrollable_frame)
    )
    line_chart_color_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)
    
    # Entry field for legend label to indicate the read point value (3rd entry field)
    line_chart_legend_field = tk.Entry(scrollable_frame, width=30)
    line_chart_legend_field.grid(row=4, column=5, padx=10, pady=10)

    # Tooltip for the legend label field
    line_chart_legend_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_legend_tooltip_label.grid(row=4, column=6, padx=10, pady=10)

    line_chart_legend_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the read point number (e.g., 24H, 48H, 168H, etc.).", scrollable_frame)
    )
    line_chart_legend_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Entry field for chart title (4th entry field)
    line_chart_title_field = tk.Entry(scrollable_frame, width=30)
    line_chart_title_field.grid(row=5, column=1, padx=10, pady=10)

    # Tooltip for the chart title field
    line_chart_title_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_title_tooltip_label.grid(row=5, column=2, padx=10, pady=10)

    line_chart_title_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Select the title for your chart with this naming convention: 'Test Number, Test Name'.", scrollable_frame)
    )
    line_chart_title_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Entry field for y-axis label (5th entry field)
    line_chart_ylabel_field = tk.Entry(scrollable_frame, width=30)
    line_chart_ylabel_field.grid(row=5, column=3, padx=10, pady=10)

    # Tooltip for the y-axis label field
    line_chart_ylabel_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_ylabel_tooltip_label.grid(row=5, column=4, padx=10, pady=10)

    line_chart_ylabel_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number to label the Y-axis (e.g., Test Value #2006).", scrollable_frame)
    )
    line_chart_ylabel_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Entry field for figure width (6th entry field)
    line_chart_width_field = tk.Entry(scrollable_frame, width=30)
    line_chart_width_field.grid(row=5, column=5, padx=10, pady=10)

    # Tooltip for the figure width field
    line_chart_width_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_width_tooltip_label.grid(row=5, column=6, padx=10, pady=10)

    line_chart_width_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the width of your line chart (default: 25).", scrollable_frame)
    )
    line_chart_width_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Entry field for the chart height (7th entry field)
    line_chart_height_field = tk.Entry(scrollable_frame, width=30)
    line_chart_height_field.grid(row=6, column=1, padx=10, pady=10)

    # Tooltip for the figure height field
    line_chart_height_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    line_chart_height_tooltip_label.grid(row=6, column=2, padx=10, pady=10)

    line_chart_height_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the height of your line chart (default: 10).", scrollable_frame)
    )
    line_chart_height_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Function definition for the line graph creation for the single read point
    def plot_line_chart(app_data):
        try:
            # Check if the Excel file has been uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel file first.", parent=scrollable_frame)
                return

            # Ensure the graph_24h DataFrame is available
            if app_data.graph_24h is None or app_data.graph_24h.empty or not hasattr(app_data, "graph_24h_generated") or not app_data.graph_24h_generated:
                raise ValueError(
                    "To plot the line chart for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )

            # Retrieve all field values
            test_number_input = line_chart_test_field.get().strip()
            color = line_chart_color_field.get().strip()
            legend = line_chart_legend_field.get().strip()
            title = line_chart_title_field.get().strip()
            ylabel = line_chart_ylabel_field.get().strip()
            width_input = line_chart_width_field.get().strip()
            height_input = line_chart_height_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, color, legend, title, ylabel, width_input, height_input]):
                raise ValueError("Please fill all required entry fields before proceeding.")

            # Validate integer fields
            try:
                test_number = int(test_number_input)
                width = int(width_input)
                height = int(height_input)
            except ValueError:
                raise ValueError("Test number, width, and height must be valid integers.")

            # Validate width and height as positive numbers
            try:
                width = int(width_input)
                height = int(height_input)
                if width <= 0 or height <= 0:
                    raise ValueError
            except ValueError:
                raise ValueError("Please enter a valid positive number to specify the width and height of your chart.")
            
            # Ensure the test number exists in the dataset
            if test_number not in app_data.graph_24h.columns:
                raise KeyError("Please enter the test number you selected earlier to generate the 'Graph Table' in Step 4.")

            # Replace spaces (' ') with numpy.nan in the selected column
            app_data.graph_24h[test_number] = app_data.graph_24h[test_number].replace(' ', np.nan).astype(float)

            # Plot the line chart
            plt.figure(figsize=(width, height))
            plt.plot(
                app_data.graph_24h['Hatrick'],
                app_data.graph_24h[test_number],
                marker='o',
                linestyle='-',
                color=color,
                label=legend
            )
            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Save the chart
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Line chart graph saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except KeyError as ke:
            messagebox.showerror("Error", str(ke), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    ####### New Buttons for additional data analysis of AMAZON PRODUCTS
    amazon_limit_button = tk.Button(
    scrollable_frame,
    text="Limit",
    command= lambda: limit_with_test_name(app_data),
    bg="#1C2541",
    fg="white",
    width=10
    )
    amazon_limit_button.grid(row=3, column=9,padx=10, pady=10)

    # Plot & Limit (CIENA uses column 4; Clear may need to be shifted to column 5 if currently at 4)
    amazon_plot_limit_button = tk.Button(
        scrollable_frame,
        text="Plot & Limit",
        command=lambda:plot_line_limit_chart_amazon(app_data),
        bg="#1C2541",
        fg="white",
        width=10
    )
    amazon_plot_limit_button.grid(row=6, column=4, padx=10, pady=10)

    # Distribution Plot with Limits (width=25 like CIENA)
    amazon_distribution_limits_button = tk.Button(
        scrollable_frame,
        text="Distribution Plot with Limits",
        command=lambda: dist_amazon_single_read_point(app_data),
        bg="#1C2541",
        fg="white",
        width=25
    )
    amazon_distribution_limits_button.grid(row=8, column=2, padx=10, pady=10)

    # Full Distribution (width=15)
    amazon_full_distribution_button = tk.Button(
        scrollable_frame,
        text="Full Distribution",
        command=lambda:full_distribution_amazon(app_data),
        bg="#1C2541",
        fg="white",
        width=15
    )
    amazon_full_distribution_button.grid(row=8, column=3, padx=10, pady=10)

    # Full Limit (width=15)
    amazon_full_limit_button = tk.Button(
        scrollable_frame,
        text="Full Limit",
        command= lambda: full_limit_amazon(app_data),
        bg="#1C2541",
        fg="white",
        width=15
    )
    amazon_full_limit_button.grid(row=14, column=2, padx=10, pady=10)

    # Full Graph (width=15)
    amazon_full_graph_button = tk.Button(
        scrollable_frame,
        text="Full Graph",
        command= lambda:plot_full_graph_amazon(app_data),
        bg="#1C2541",
        fg="white",
        width=15
    )
    amazon_full_graph_button.grid(row=14, column=3, padx=10, pady=10)

    ########################
    
    # Button creation for line chart
    plot_button = tk.Button(scrollable_frame, text="Plot", command=lambda: plot_line_chart(app_data), bg="#1C2541", fg="white", width=10)
    plot_button.grid(row=6, column=3, padx=10, pady=10)

    # Clear button
    def clear_line_chart_fields(app_data):
        line_chart_test_field.delete(0, tk.END)
        line_chart_color_field.delete(0, tk.END)
        line_chart_legend_field.delete(0, tk.END)
        line_chart_title_field.delete(0, tk.END)
        line_chart_ylabel_field.delete(0, tk.END)
        line_chart_width_field.delete(0, tk.END)
        line_chart_height_field.delete(0, tk.END)

    clear_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_line_chart_fields(app_data), bg="#1C2541", fg="white", width=10)
    clear_button.grid(row=6, column=5, padx=10, pady=10)

    # Add the "6. Distribution Chart" label
    distribution_chart_label = tk.Label(scrollable_frame, text="6. Histogram:", bg="#1C2541", fg="white")
    distribution_chart_label.grid(row=7, column=0, padx=10, pady=10, sticky="w")

    # Add the 1st entry field for test number
    distribution_test_field = tk.Entry(scrollable_frame, width=30)
    distribution_test_field.grid(row=7, column=1, padx=10, pady=10)

    # Add the 1st tooltip for test number
    distribution_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    distribution_test_tooltip_label.grid(row=7, column=2, padx=10, pady=10)

    # Function to show the tooltip for test number
    distribution_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number to analyze its distribution. Use the same test numbers as in step 4. If you wish to plot the distribution for all read points, you need to select a test number as of step 9.", scrollable_frame)
    )
    distribution_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 2nd entry field for bins
    distribution_bins_field = tk.Entry(scrollable_frame, width=30)
    distribution_bins_field.grid(row=7, column=3, padx=10, pady=10)

    # Add the 2nd tooltip for bins
    distribution_bins_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    distribution_bins_tooltip_label.grid(row=7, column=4, padx=10, pady=10)

    distribution_bins_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the number of bins for the histogram. Higher values provide more detail.", scrollable_frame)
    )
    distribution_bins_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 3rd entry field for chart title
    distribution_title_field = tk.Entry(scrollable_frame, width=30)
    distribution_title_field.grid(row=7, column=5, padx=10, pady=10)

    # Add the 3rd tooltip for distribution chart title
    distribution_title_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    distribution_title_tooltip_label.grid(row=7, column=6, padx=10, pady=10)

    # Bind the tooltip functions to mouse enter and leave events for the title
    distribution_title_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Select the title for your Distribution Chart. Use this naming convention: 'Test Number, Test Name'.", scrollable_frame)
    )
    distribution_title_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Function to plot the distribution chart for a single read point
    def plot_distribution_chart(app_data):
        try:
            # Ensure if the Excel file is uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return

            # Ensure the graph_24h DataFrame is available and generated in step 4
            if app_data.graph_24h is None or app_data.graph_24h.empty or not hasattr(app_data, "graph_24h_generated") or not app_data.graph_24h_generated:
                raise ValueError(
                    "To plot the distribution chart (histogram) for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )

            # Retrieve all field values
            test_number_input = distribution_test_field.get().strip()
            bins_input = distribution_bins_field.get().strip()
            title = distribution_title_field.get().strip()

            # Ensure all fields are filled
            if not all([test_number_input, bins_input, title]):
                raise ValueError("Please fill all required entry fields before proceeding.")

            # Validate integer fields
            try:
                test_number = int(test_number_input)
                bins = int(bins_input)
                if bins <= 0:
                    raise ValueError("The value entered for the number of bins and the test number in your distribution chart must be a positive integer. Please correct your input and try again.")
            except ValueError as ve:
                raise ValueError(str(ve))

            # Ensure the test number exists in the dataset
            if test_number not in app_data.graph_24h.columns:
                raise KeyError("Please enter the test number you selected earlier to generate the 'Graph Table' in Step 4.")

            # Replace spaces (' ') with numpy.nan in the selected column
            app_data.graph_24h[test_number] = app_data.graph_24h[test_number].replace(' ', np.nan).astype(float)

            # Plot the distribution chart, skipping NaN values automatically
            plt.figure()
            plt.hist(app_data.graph_24h[test_number].dropna().values, bins=bins, label=f"Test {test_number}", alpha=0.7)
            plt.title(title)
            plt.xlabel("Test Value")
            plt.ylabel("Frequency")
            plt.tight_layout()

            # Save the chart
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], parent=scrollable_frame)
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Distribution chart saved successfully at this path: {save_path}.", parent=scrollable_frame)

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except KeyError as ke:
            messagebox.showerror("Error", str(ke), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot the distribution chart. Please enter the same test numbers used in step 4.", parent=scrollable_frame)

    # Add the "Distribution Plot" button
    distribution_plot_button = tk.Button(scrollable_frame, text="Distribution Plot", command=lambda: plot_distribution_chart(app_data), bg="#1C2541", fg="white", width=15)
    distribution_plot_button.grid(row=8, column=1, columnspan=1, padx=10, pady=10)

    # Function to clear all fields in Label 6
    def clear_distribution_fields(app_data):
        distribution_test_field.delete(0, tk.END)
        distribution_bins_field.delete(0, tk.END)
        distribution_title_field.delete(0, tk.END)

    # Add the "Clear" button for the distribution plot
    clear_distribution_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_distribution_fields(app_data), bg="#1C2541", fg="white", width=10)
    clear_distribution_button.grid(row=8, column=4, padx=10, pady=10)

    # Add the "7. Test Statistics" label
    test_statistics_label = tk.Label(scrollable_frame, text="7. Test Statistics:", bg="#1C2541", fg="white")
    test_statistics_label.grid(row=9, column=0, padx=10, pady=10, sticky="w")

    # Add the entry field for test number
    test_statistics_field = tk.Entry(scrollable_frame, width=30)
    test_statistics_field.grid(row=9, column=1, padx=10, pady=10)

    # Add the tooltip for the test number field
    test_statistics_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    test_statistics_tooltip_label.grid(row=9, column=2, padx=10, pady=10)

    # Function to show the tooltip for test statistics
    test_statistics_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number you previously selected in step 4 to retrieve statistics.", scrollable_frame)
    )
    test_statistics_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Function to generate test statistics and save to a .txt file
    def generate_test_statistics(app_data):
        try:
            # Ensure if the Excel file is uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return

            # Ensure the graph_24h DataFrame is available and generated in step 4
            if app_data.graph_24h is None or app_data.graph_24h.empty or not hasattr(app_data, "graph_24h_generated") or not app_data.graph_24h_generated:
                raise ValueError(
                    "To have test statistics for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )

            # Retrieve the test number from the entry field
            test_number_input = test_statistics_field.get().strip()

            # Ensure the field is filled
            if not test_number_input:
                raise ValueError("Please fill the required entry field before proceeding.")

            # Validate the test number as an integer
            try:
                test_number = int(test_number_input)
            except ValueError:
                raise ValueError("Test number must be a valid number (integer).")

            # Validate that the test number exists in the DataFrame
            if test_number not in app_data.graph_24h.columns:
                raise KeyError("Please enter the test number you selected earlier to generate the 'Graph Table' in Step 4.")

            # Extract the relevant column and calculate statistics
            real_stat_24h = pd.to_numeric(app_data.graph_24h[test_number].copy(), errors='coerce')
            stats = real_stat_24h.describe()

            # Prompt the user to save the statistics to a .txt file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")],
                title="Save Test Statistics As",
                parent=scrollable_frame
            )

            if save_path:
                # Save the statistics to the .txt file
                with open(save_path, 'w') as file:
                    file.write(f"Statistics for Test Number {test_number}:\n")
                    file.write(stats.to_string())

                # Notify the user that the file has been saved
                messagebox.showinfo("Success", f"Test statistics text file (.txt) saved successfully at this path: {save_path}", parent=scrollable_frame)

        except KeyError as e:
            messagebox.showerror("Error", f"{str(e)}", parent=scrollable_frame)
        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Add the "Test Stats" button
    test_statistics_button = tk.Button(scrollable_frame, text="Test Stats", command=lambda: generate_test_statistics(app_data), bg="#1C2541", fg="white", width=10)
    test_statistics_button.grid(row=9, column=3, padx=10, pady=10)

    # Function to clear the Test Statistics entry field
    def clear_test_statistics_field(app_data):
        test_statistics_field.delete(0, tk.END)

    # Add the "Clear" button next to the "Test Stats" button
    clear_test_statistics_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_test_statistics_field(app_data), bg="#1C2541", fg="white", width=10)
    clear_test_statistics_button.grid(row=9, column=4, padx=10, pady=10)

    # Add the "8. Failure Detection" label
    failure_detection_label = tk.Label(scrollable_frame, text="8. Failure Detection:", bg="#1C2541", fg="white")
    failure_detection_label.grid(row=10, column=0, padx=10, pady=10, sticky="w")

    # Add the 1st entry field for the operand of the failure detection section
    failure_operand_field = tk.Entry(scrollable_frame, width=10)
    failure_operand_field.grid(row=10, column=1, padx=10, pady=10)

    # Add the 1st tooltip for the operand
    failure_operand_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    failure_operand_tooltip_label.grid(row=10, column=2, padx=10, pady=10)

    # Function to show the tooltip for the operand
    failure_operand_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter one of these operands: <, <=, >, >=, or =.", scrollable_frame)
    )
    failure_operand_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 2nd entry field for the HBIN number
    failure_hbin_field = tk.Entry(scrollable_frame, width=10)
    failure_hbin_field.grid(row=10, column=3, padx=10, pady=10)

    # Add the 2nd tooltip for the HBIN number
    failure_hbin_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    failure_hbin_tooltip_label.grid(row=10, column=4, padx=10, pady=10)

    # Function to show the tooltip for the HBIN number
    failure_hbin_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the HBIN number you wish to analyze.", scrollable_frame)
    )
    failure_hbin_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Function to generate the failure table
    def generate_failure_table(app_data):
        try:
            # Ensure if the Excel file is uploaded
            if not app_data.file_uploaded:
                messagebox.showerror("Error", "Please upload the Excel File first.", parent=scrollable_frame)
                return 
            
            # Ensure the graph_24h DataFrame is available and generated in step 4
            if app_data.graph_24h is None or app_data.graph_24h.empty or not hasattr(app_data, "graph_24h_generated") or not app_data.graph_24h_generated:
                raise ValueError(
                "To have units with particular HBIN(s) for a single read point, please first generate the 'Graph Table' Excel file in Step 4 by clicking the corresponding button in the application."
                )     
            
            # Ensure the 'HBIN' column exists in the DataFrame
            if 'HBIN' not in app_data.graph_24h.columns:
                raise KeyError("The 'HBIN' column is missing in the STDF dataset.")
            
            # Get the operand and HBIN number from the entry fields
            operand = failure_operand_field.get().strip()
            hbin_number = failure_hbin_field.get().strip()
            
            # Check if both entry fields are empty
            if not operand or not hbin_number:
                raise ValueError("Please ensure both the operand and HBIN number fields are filled before proceeding.")

            # Validate the HBIN number
            if not hbin_number.isdigit():
                raise ValueError("HBIN must be a positive number (integer).")
            hbin_number = int(hbin_number)

            # Validate the operand
            if operand not in ['<', '<=', '>', '>=', '=']:
                raise ValueError("Invalid operand. Please enter one of these operands: <, <=, >, >=, or =.")

            # Convert '=' to '==' for the query
            if operand == '=':
                operand = '=='

            # Dynamically construct the condition based on the operand and HBIN number
            condition = f"HBIN {operand} {hbin_number}"

            # Query the DataFrame
            app_data.failure_24h = app_data.graph_24h.query(condition)

            # Check if the query returned any results
            if app_data.failure_24h.empty:
                raise ValueError(f"No data matches the specified condition: {condition}.")

            # Prompt the user to select a save location
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Failure Table As",
                parent=scrollable_frame
            )

            if save_path:
                # Save the failure table to the specified location
                app_data.failure_24h.to_excel(save_path, sheet_name='Detection', index=False)
                messagebox.showinfo("Success", f"Excel file with specific HBINs saved successfully at this path: {save_path}.", parent=failure_window)

        except ValueError as ve:
            messagebox.showerror("Error", f"{str(ve)}", parent=scrollable_frame)
        except Exception as ve:
            messagebox.showerror("Error", f"{str(ve)}", parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error",  f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)
    
    # Add the "Failure Table Generation" button
    failure_table_button = tk.Button(scrollable_frame, text="Failure Table Generation", command=lambda: generate_failure_table(app_data), bg="#1C2541", fg="white", width=20)
    failure_table_button.grid(row=10, column=5, padx=10, pady=10)

    # Function to clear all entry fields in Label 8 (Failure Detection)
    def clear_failure_detection_fields(app_data):
        failure_operand_field.delete(0, tk.END)
        failure_hbin_field.delete(0, tk.END)

    # Add the "Clear" button next to the "Failure Table Generation" button
    clear_failure_button = tk.Button(scrollable_frame, text="Clear", command=lambda: clear_failure_detection_fields(app_data), bg="#1C2541", fg="white", width=10)
    clear_failure_button.grid(row=10, column=6, padx=10, pady=10)

    # Add the "9. Final Chart" label
    final_chart_label = tk.Label(scrollable_frame, text="9. Final Chart:", bg="#1C2541", fg="white")
    final_chart_label.grid(row=11, column=0, padx=10, pady=10, sticky="w")

    # Add the 1st entry field for test number
    final_chart_test_field = tk.Entry(scrollable_frame, width=30)
    final_chart_test_field.grid(row=11, column=1, padx=10, pady=10)

    # Add the 1st question mark for test number
    final_chart_test_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_test_tooltip_label.grid(row=11, column=2, padx=10, pady=10)

    # Tooltip for the 1st question mark for the test number
    final_chart_test_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter the test number you wish to plot.", scrollable_frame)
    )
    final_chart_test_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 2nd entry field for chart color
    final_chart_color_field = tk.Entry(scrollable_frame, width=30)
    final_chart_color_field.grid(row=11, column=3, padx=10, pady=10)

    # Add the 2nd question mark for chart color
    final_chart_color_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_color_tooltip_label.grid(row=11, column=4, padx=10, pady=10)

    final_chart_color_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the color of your chart. Ensure the number of color names matches the number of read points.", scrollable_frame)
    )
    final_chart_color_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 3rd entry field for read points
    final_chart_read_points_field = tk.Entry(scrollable_frame, width=30)
    final_chart_read_points_field.grid(row=11, column=5, padx=10, pady=10)

    # Add the 3rd question mark for read points
    final_chart_read_points_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_read_points_tooltip_label.grid(row=11, column=6, padx=10, pady=10)

    final_chart_read_points_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Enter all read points you wish to analyze (e.g., 24H, 68H, etc.).", scrollable_frame)
    )
    final_chart_read_points_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 4th entry field for chart title
    final_chart_title_field = tk.Entry(scrollable_frame, width=30)
    final_chart_title_field.grid(row=12, column=1, padx=10, pady=10)

    # Add the 4th question mark for chart title
    final_chart_title_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_title_tooltip_label.grid(row=12, column=2, padx=10, pady=10)

    # Tooltip for the 4th question mark
    final_chart_title_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the title of your chart using the naming convention 'Test Number: Test Name'.", scrollable_frame)
    )
    final_chart_title_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 5th entry field for y-axis label
    final_chart_ylabel_field = tk.Entry(scrollable_frame, width=30)
    final_chart_ylabel_field.grid(row=12, column=3, padx=10, pady=10)

    # Add the 5th question mark for y-axis label
    final_chart_ylabel_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_ylabel_tooltip_label.grid(row=12, column=4, padx=10, pady=10)

    # Tooltip for the 5th question mark
    final_chart_ylabel_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Label the Y-axis using this naming convention: Test Value #2006.", scrollable_frame)
    )
    final_chart_ylabel_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 6th entry field for chart width
    final_chart_width_field = tk.Entry(scrollable_frame, width=30)
    final_chart_width_field.grid(row=12, column=5, padx=10, pady=10)

    # Add the 6th question mark for chart width
    final_chart_width_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_width_tooltip_label.grid(row=12, column=6, padx=10, pady=10)

    # Tooltip for the 6th question mark
    final_chart_width_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the width of your line chart (default: 25).", scrollable_frame)
    )
    final_chart_width_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the 7th entry field for chart height
    final_chart_height_field = tk.Entry(scrollable_frame, width=30)
    final_chart_height_field.grid(row=13, column=1, padx=10, pady=10)

    # Add the 7th question mark for chart height
    final_chart_height_tooltip_label = tk.Label(scrollable_frame, text="?", bg="#1C2541", fg="white", font=("Helvetica", 10, "bold italic"))
    final_chart_height_tooltip_label.grid(row=13, column=2, padx=10, pady=10)

    # Tooltip for the 7th question mark
    final_chart_height_tooltip_label.bind(
        "<Enter>",
        lambda event: show_dynamic_tooltip(event, "Specify the height of your line chart (default: 10).", scrollable_frame)
    )
    final_chart_height_tooltip_label.bind("<Leave>", hide_dynamic_tooltip)

    # Add the "Final Graph" button
    final_graph_button = tk.Button(scrollable_frame, text="Final Graph", bg="#1C2541", fg="white", width=15)
    final_graph_button.grid(row=13, column=4, padx=10, pady=10)

    # Function definition for the final table dataset for all read points
    def generate_final_table(app_data):
        try:
            # Ensure read_point_graphs has enough data
            if not app_data.read_point_graphs or len(app_data.read_point_graphs) < 2:
                raise ValueError("Please generate at least two Read Point Tables or datasets before creating the Final Table Excel file or dataset.")

            # Dynamically select columns for merging based on 'Hatrick' and columns between '150000114' and 'Hatrick'
            merged_table = app_data.read_point_graphs[0][['Hatrick']]
            for graph in app_data.read_point_graphs:
                # Find the positions of "150000114" and "Hatrick" in the columns
                columns = graph.columns.tolist()
                if "150000114" in columns and "Hatrick" in columns:
                    start_index = columns.index("150000114") + 1  # Start after "150000114"
                    end_index = columns.index("Hatrick")  # End before "Hatrick"
                    columns_to_merge = ["Hatrick"] + columns[start_index:end_index]
                else:
                    raise ValueError("Columns '150000114' and 'Hatrick' are not found in the 'read_point_graphs' dataset.")

                # Perform the merge
                merged_table = merged_table.merge(graph[columns_to_merge], on='Hatrick', how='outer')

            # Reset the index for the final table
            merged_table.reset_index(drop=True, inplace=True)

            # Ask the user if they want to save the Excel file
            user_response = messagebox.askyesno(
                "Save Final Table",
                "Do you want to generate the Excel file for the Final Table?",
                parent=scrollable_frame
            )

            if user_response:
                # Prompt the user to save the Excel file
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Final Table As",
                    parent=scrollable_frame
                )
                if not save_path:
                    messagebox.showinfo("Canceled", "The operation was canceled. No Excel file was saved.", parent=scrollable_frame)
                    return
                
                # Save the DataFrame to Excel with the sheet name "Final_Table"
                merged_table.to_excel(save_path, sheet_name="Final_Table", index=False)
                messagebox.showinfo("Success", f"Final Table Excel file with all test numbers and read points saved successfully at this path: {save_path}", parent=scrollable_frame)

            # Cache the final table for later use
            app_data.final_table = merged_table

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Add the "Final Table" button
    final_table_button = tk.Button(scrollable_frame, text="Final Table", bg="#1C2541", fg="white", width=15)
    final_table_button.grid(row=13, column=3, padx=10, pady=10)
    
    # Function definition for the Final Graph of all read points
    def generate_final_graph(app_data):
        try:
            # Ensure the final_table DataFrame is available
            if not hasattr(app_data, 'final_table') or app_data.final_table is None:
                raise ValueError("Unable to create the Final Graph with all read points. Please ensure the Final Table is generated first.")

            # Get user inputs from the entry fields
            test_number = final_chart_test_field.get().strip()
            colors = final_chart_color_field.get().strip()
            labels = final_chart_read_points_field.get().strip()
            title = final_chart_title_field.get().strip()
            ylabel = final_chart_ylabel_field.get().strip()
            width = final_chart_width_field.get().strip()
            height = final_chart_height_field.get().strip()

            # Check if any field is empty
            if not all([test_number, colors, labels, title, ylabel, width, height]):
                raise ValueError("Please ensure all required entry fields are filled before proceeding.")

            # Validate width and height as positive numbers
            try:
                width = int(width)
                height = int(height)
                if width <= 0 or height <= 0:
                    raise ValueError
            except ValueError:
                raise ValueError("Please enter a valid positive number to specify the width and height of your chart.")

            # Convert test_number to string for matching column names
            test_number_str = str(test_number)

            # Filter columns that match the test number
            matching_columns = [col for col in app_data.final_table.columns if col.startswith(test_number_str + '_')]

            # Check if the test number exists in the final_table
            if not matching_columns:
                raise ValueError("Ensure that the test number you entered matches one of the test numbers imported during step 4 of the analysis.")

            # Validate the number of colors matches the number of matching columns
            color_list = colors.split(',')
            if len(color_list) != len(matching_columns):
                raise ValueError("The number of colors you specify in the color entry field must match the number of read points you wish to analyze.")

            # Validate the number of labels matches the number of matching columns
            label_list = labels.split(',')
            if len(label_list) != len(matching_columns):
                raise ValueError(f"Number of labels ({len(label_list)}) does not match the number of read points.")

            # Handle empty cells in the matching columns
            for col in matching_columns:
                app_data.final_table[col] = app_data.final_table[col].replace(' ', np.nan).astype(float)

            # Plot the graph
            plt.figure(figsize=(width, height))
            for col, color, label in zip(matching_columns, color_list, label_list):
                plt.plot(app_data.final_table['Hatrick'], app_data.final_table[col], marker='o', linestyle='-', color=color.strip(), label=label.strip())
            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            # Prompt user to select save location for the PDF file
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Final Graph As PDF",
                parent=scrollable_frame
            )

            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Final line graph with all read points saved successfully at this path: {save_path}", parent=scrollable_frame)

            # Close the plot to avoid displaying it
            plt.close()

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Bind the function to the "Final Graph" button
    final_graph_button.config(command=lambda: generate_final_graph(app_data))
    
    # Bind the function to the "Final Table" button
    final_table_button.config(command=lambda: generate_final_table(app_data))
    
    # Function for excessive units
    def plot_excessive_units(app_data):
        try:
            if not hasattr(app_data, 'final_table') or app_data.final_table is None:
                raise ValueError("Unable to create the Excessive Units Graph. Please ensure the Final Table is generated first.")

            test_number = final_chart_test_field.get().strip()
            colors = final_chart_color_field.get().strip()
            labels = final_chart_read_points_field.get().strip()
            title = final_chart_title_field.get().strip()
            ylabel = final_chart_ylabel_field.get().strip()
            width = final_chart_width_field.get().strip()
            height = final_chart_height_field.get().strip()

            if not all([test_number, colors, labels, title, ylabel, width, height]):
                raise ValueError("Please ensure all required entry fields are filled before proceeding.")

            try:
                width = int(width); height = int(height)
                if width <= 0 or height <= 0:
                    raise ValueError
            except ValueError:
                raise ValueError("Please enter a valid positive number to specify the width and height of your chart.")

            test_number_str = str(test_number)
            matching_columns = [col for col in app_data.final_table.columns if col.startswith(test_number_str + '_')]
            if not matching_columns:
                raise ValueError("Ensure that the test number you entered matches one of the test numbers imported during step 4 of the analysis.")

            color_list = [c.strip() for c in colors.split(',')]
            if len(color_list) != len(matching_columns):
                raise ValueError("The number of colors must match the number of read points.")

            label_list = [l.strip() for l in labels.split(',')]
            if len(label_list) != len(matching_columns):
                raise ValueError("Number of labels does not match the number of read points.")

            for col in matching_columns:
                app_data.final_table[col] = app_data.final_table[col].replace(' ', np.nan).astype(float)

            highest_column = max(matching_columns, key=lambda col: int(col.split('_')[1].replace('h', '')))
            excessive_units = app_data.final_table.copy()
            excessive_units = excessive_units.dropna(subset=[highest_column])

            # Cache dataset + flag for later Full Graph with limits
            app_data.excessive_units = excessive_units[['Hatrick'] + matching_columns]
            app_data.excessive_units_generated = True

            plt.figure(figsize=(width, height))
            for col, color, label in zip(matching_columns, color_list, label_list):
                plt.plot(excessive_units['Hatrick'], excessive_units[col],
                         marker='o', linestyle='-', color=color, label=label)
            plt.title(title)
            plt.xlabel('Hatrick')
            plt.ylabel(ylabel)
            plt.xticks(rotation=45, ha='right')
            plt.grid(True)
            plt.tight_layout()
            plt.legend()

            save_dataset = messagebox.askyesno(
                "Save Dataset",
                "Do you want to save the dataset used for this graph as an Excel file?",
                parent=scrollable_frame
            )
            if save_dataset:
                save_excel_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Dataset As",
                    parent=scrollable_frame
                )
                if save_excel_path:
                    app_data.excessive_units.to_excel(save_excel_path, index=False, sheet_name='Excessive_Units')
                    messagebox.showinfo("Success", f"Dataset saved successfully at this path: {save_excel_path}", parent=scrollable_frame)

            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Excessive Units Graph As PDF",
                parent=scrollable_frame
            )
            if save_path:
                plt.savefig(save_path)
                messagebox.showinfo("Success", f"Excessive Units graph saved successfully at this path: {save_path}", parent=scrollable_frame)
            plt.close()

        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=scrollable_frame)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=scrollable_frame)

    # Add the "Excessive Units" button
    excessive_units_button = tk.Button(scrollable_frame, text="Excessive Units", bg="#1C2541", fg="white", width=15)
    excessive_units_button.grid(row=14, column=1, padx=10, pady=10)
    
    # Bind the function to the relative button
    excessive_units_button.config(command=lambda: plot_excessive_units(app_data))
    
    # Add the "Clear" button for the final chart
    def clear_label_9_fields():
        final_chart_test_field.delete(0, tk.END)
        final_chart_color_field.delete(0, tk.END)
        final_chart_read_points_field.delete(0, tk.END)
        final_chart_title_field.delete(0, tk.END)
        final_chart_ylabel_field.delete(0, tk.END)
        final_chart_width_field.delete(0, tk.END)
        final_chart_height_field.delete(0, tk.END)
        # Reset any functions or variables linked to label 9 buttons if needed

    # Button creation for the 'Clear' button of final chart of all read points
    clear_button_label_9 = tk.Button(scrollable_frame, text="Clear", command=clear_label_9_fields, bg="#1C2541", fg="white", width=10)
    clear_button_label_9.grid(row=13, column=5, padx=10, pady=10)

    # Add the "Clear All" button to reset the application
    def clear_all_fields_and_cache(app_data):
        # Show a warning prompt before clearning
        user_response = messagebox.askyesno(
            "Warning",
            "Are you sure you want to clear all entries? If you click 'YES', all data will be reset and you will need to restart the analysis from the beginning.",
            parent=scrollable_frame,
            icon= "warning"
        )
        if not user_response:
            return  # Do nothing if user selects NO
        
        try:
            # Clear all entry fields in the failure detection window
            entry_field.delete(0, tk.END)
            test_number_field.delete(0, tk.END)
            read_point_field.delete(0, tk.END)
            sheet_name_field.delete(0, tk.END)
            graph_table_test_field.delete(0, tk.END)
            graph_table_sheet_field.delete(0, tk.END)
            line_chart_test_field.delete(0, tk.END)
            line_chart_color_field.delete(0, tk.END)
            line_chart_legend_field.delete(0, tk.END)
            line_chart_title_field.delete(0, tk.END)
            line_chart_ylabel_field.delete(0, tk.END)
            line_chart_width_field.delete(0, tk.END)
            line_chart_height_field.delete(0, tk.END)
            distribution_test_field.delete(0, tk.END)
            distribution_bins_field.delete(0, tk.END)
            distribution_title_field.delete(0, tk.END)
            test_statistics_field.delete(0, tk.END)
            failure_operand_field.delete(0, tk.END)
            failure_hbin_field.delete(0, tk.END)
            final_chart_test_field.delete(0, tk.END)
            final_chart_color_field.delete(0, tk.END)
            final_chart_read_points_field.delete(0, tk.END)
            final_chart_title_field.delete(0, tk.END)
            final_chart_ylabel_field.delete(0, tk.END)
            final_chart_width_field.delete(0, tk.END)
            final_chart_height_field.delete(0, tk.END)

            # Reset all variables and cache
            app_data.file_uploaded = False
            app_data.df_v1_read_point_24h = None
            app_data.df_testing_24h = None
            app_data.df_general_24h = None
            app_data.df_analysis_24h = None
            app_data.df_general_24h_help_1 = None
            app_data.df_general_24h_help_2 = None
            app_data.df_general_24h_help_3 = None
            app_data.graph_24h = None
            app_data.failure_24h = None
            app_data.df_limit_24h = None
            app_data.read_point_graphs= []
            app_data.final_table = None
            app_data.overview_generated = False
            app_data.graph_24h_generated = False
            app_data.excessive_units = None
            app_data.df_limit_24h_generated = False
            app_data.analysis_point_graphs = []
            app_data.full_limit_amazon_merged = None
            app_data.full_limit_amazon_generated = False
            app_data.excessive_units_generated = False
            app_data.plot_df_amazon = None
            app_data.plot_df_amazon_generated = False
            app_data.full_distribution_amazon_df = None
            app_data.full_distribution_amazon_generated = False

            # Close all matplotlib plots in the main thread
            plt.close('all')
        
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while clearing: {str(e)}", parent=scrollable_frame)
    
    # Create clear all button creation
    clear_all_button = tk.Button(scrollable_frame, text="Clear All", command=lambda: clear_all_fields_and_cache(app_data), bg="#1C2541", fg="white", width=10)
    clear_all_button.grid(row=13, column=6, padx=10, pady=10)

     # Construct the path to the Warning Sign Image using the base_path for .exe file application
    warning_logo_path = os.path.join(base_path, "resources", "warning.png")

    # Create a dummy image and label first (will be resized and updated later)
    warning_logo_image = Image.open(warning_logo_path)
    warning_logo_image = warning_logo_image.resize((24, 24), Image.Resampling.LANCZOS)  # Initial size, will be updated
    warning_logo_image_tk = ImageTk.PhotoImage(warning_logo_image)
    warning_logo_label = tk.Label(
        scrollable_frame,
        image=warning_logo_image_tk,
        bg="#1C2541"
    )
    warning_logo_label.image = warning_logo_image_tk  # Keep a reference to avoid garbage collection

    # Place the label initially (will be repositioned and resized by place_warning_logo)
    warning_logo_label.place(x=0, y=0)

    def place_warning_logo():
        button_height = clear_all_button.winfo_height()
        button_width = clear_all_button.winfo_width()
        ratio = 0.9
        img_height = int(button_height * ratio)
        img_width = int(button_height * ratio)
        warning_logo_image = Image.open(warning_logo_path)
        warning_logo_image = warning_logo_image.resize((img_width, img_height), Image.Resampling.LANCZOS)
        warning_logo_image_tk = ImageTk.PhotoImage(warning_logo_image)
        warning_logo_label.config(image=warning_logo_image_tk)
        warning_logo_label.image = warning_logo_image_tk  # Keep a reference
        warning_logo_label.place(
            x=clear_all_button.winfo_x() + button_width + 8,
            y=clear_all_button.winfo_y() + (button_height - img_height) // 2
        )

    scrollable_frame.after(200, place_warning_logo)

    # Construct the path to the ST logo image using the base_path for .exe file
    logo_path = os.path.join(base_path, "resources", "STlogo.gif_min.gif")

    # Load the ST logo image
    logo_image = tk.PhotoImage(file=logo_path)

    # Add the ST logo to the bottom-right corner of the failure detection window for AMAZON
    failure_logo_label = tk.Label(failure_window, image=logo_image, bg="#1C2541")
    failure_logo_label.image = logo_image  # Keep a reference to avoid garbage collection
    failure_logo_label.place(relx=1.0, rely=1.0, anchor='se', x=-20, y=-20)  # Adjust x and y for padding

# Sub-options for Reliability Analysis (Graph)
reliability_sub_button2 = tk.Button(reliability_sub_option_frame, text="Graph", command=open_failure_detection_window, **italic_button_style, width=half_button_width)
reliability_sub_button2.pack(side='top', pady=5)

# Function to toggle sub-options (Garph and Excel Macro) for Reliability Analysis of Makalu
def toggle_reliability_options():
    global reliability_sub_options_visible
    if reliability_sub_options_visible:
        reliability_sub_option_frame.pack_forget()
    else:
        reliability_sub_option_frame.pack(pady=5, after=sub_button2)
    reliability_sub_options_visible = not reliability_sub_options_visible

# Bind the toggle function to the Reliability Analysis button
sub_button2.config(command=toggle_reliability_options)

# Run the application
root.mainloop()