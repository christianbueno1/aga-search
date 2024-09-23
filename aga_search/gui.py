import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from aga_search.config import DOWNLOADS_PATH, DB_FILE_PATH, COLUMN1, COLUMN2, COLUMN3, NEW_COLUMNS_NAME, SHEET_INDEX, SHEET_NAME, FILE_COLUMNS, COLUMN_INDEX
import time
from aga_search.read_spreadsheet import read_file, insert_column_in_excel, create_new_excel_file
import pandas as pd

# Initialize customtkinter
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "dark-blue", "green"

class ExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # width and height of the window
        width = 800
        height = 600
            
        # Set window title and minimum size
        self.title("Excel File Selector")
        self.minsize(width, height)  # Set minimum width and height

        # Make the window responsive by configuring grid rows and columns
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Create a main frame to hold all widgets
        main_frame = ctk.CTkFrame(self, fg_color="gray65")
        main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

        # Configure grid for main_frame
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # Frame for the database file selection
        # frame_db = ctk.CTkFrame(main_frame, fg_color="lightblue")
        frame_db = ctk.CTkFrame(main_frame, fg_color="gray85")
        frame_db.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 10))

        # Configure grid for frame_db
        frame_db.grid_columnconfigure(0, weight=1)

        # General label to show the title
        label_title = ctk.CTkLabel(frame_db, text="Select the Database Excel")
        label_title.grid(row=0, column=0, padx=(10, 10), pady=(10, 5), sticky="w")

        # Label to show the path of the selected file
        self.label_db = ctk.CTkLabel(frame_db, text="No file selected")
        self.label_db.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")

        # Button to select the DB Excel file
        self.button_db = ctk.CTkButton(frame_db, text="Select File", command=self.select_db_file)
        self.button_db.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")

        # Frame for the upload file selection
        # frame_upload = ctk.CTkFrame(main_frame, fg_color="lightgreen")
        frame_upload = ctk.CTkFrame(main_frame, fg_color="gray85")
        frame_upload.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Configure grid for frame_upload
        frame_upload.grid_columnconfigure(0, weight=1)

        # Label for upload file
        label_upload = ctk.CTkLabel(frame_upload, text="Select the Excel File to Upload")
        label_upload.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        # Label to show the path of the selected upload file
        self.label_upload = ctk.CTkLabel(frame_upload, text="No file selected")
        self.label_upload.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")

        # Button to select the upload Excel file
        self.button_upload = ctk.CTkButton(frame_upload, text="Select File", command=self.select_upload_file, state="disabled")
        self.button_upload.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")

        # Frame for action buttons
        frame_actions = ctk.CTkFrame(main_frame, fg_color="gray85")
        frame_actions.grid(row=2, column=0, sticky="nsew", padx=10, pady=(10, 10))

        # Configure grid for frame_actions
        frame_actions.grid_rowconfigure(0, weight=1)
        frame_actions.grid_columnconfigure(0, weight=1)
        frame_actions.grid_columnconfigure(1, weight=1)
        frame_actions.grid_columnconfigure(2, weight=1)
        frame_actions.grid_columnconfigure(3, weight=1)

        # Checkbox to set Downloads as default save directory
        self.use_downloads_var = ctk.BooleanVar(value=True)  # Default checked
        self.checkbox_downloads = ctk.CTkCheckBox(
            frame_actions,
            text="Save to Downloads Directory",
            variable=self.use_downloads_var
        )
        self.checkbox_downloads.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        # Progress bar for file processing
        self.progress_bar = ctk.CTkProgressBar(frame_actions)
        self.progress_bar.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        self.progress_bar.set(0)  # Initialize progress bar to 0
        
        # Button to process the selected files
        self.button_process = ctk.CTkButton(frame_actions, text="Process", command=self.process_files, state="disabled")
        self.button_process.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # Button to close the window
        self.button_close = ctk.CTkButton(frame_actions, text="Close", command=self.close_window)
        self.button_close.grid(row=0, column=2, padx=10, pady=10, sticky="ew")

        # new frame for textbox_output
        frame_textbox_output = ctk.CTkFrame(main_frame, fg_color="gray85")
        frame_textbox_output.grid(row=3, column=0, sticky="nsew", padx=10, pady=(10, 10))

        # CTkTextbox to display positive_new_values_inserted_count and rows_count
        self.textbox_output = ctk.CTkTextbox(
            frame_textbox_output, 
            width=400, 
            height=100,
            state="disabled",
            fg_color="white",
            corner_radius=5
        )
        self.textbox_output.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.check_if_db_exists()
    
    # function to insert the values in the textbox_output
    def insert_values_in_textbox_output(self, positive_new_values_inserted_count, rows_count):
        self.textbox_output.configure(state="normal")
        self.textbox_output.delete("1.0", "end")
        self.textbox_output.insert("end", f"Positive New Values Inserted Count: {positive_new_values_inserted_count}\n")
        self.textbox_output.insert("end", f"Rows Count: {rows_count}\n")
        self.textbox_output.configure(state="disabled")

    def check_if_db_exists(self):
        # Check if the database file exists
        if os.path.exists(DB_FILE_PATH):
            self.label_db.configure(text=DB_FILE_PATH)
            self.button_upload.configure(state="normal")

    def select_db_file(self):
        # Open a file dialog to select the file
        file_path = filedialog.askopenfilename(
            initialdir=DOWNLOADS_PATH,
            title="Select the Database Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        # If the file is selected, show the path in the label
        if file_path:
            self.label_db.configure(text=file_path)
            # Enable the upload file button
            self.button_upload.configure(state="normal")

    def select_upload_file(self):
        # Open a file dialog to select the file
        file_path = filedialog.askopenfilename(
            initialdir=DOWNLOADS_PATH,
            title="Select the Excel File to Upload",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        # If the file is selected, show the path in the label
        if file_path:
            self.label_upload.configure(text=file_path)
            # Enable the process button
            self.button_process.configure(state="normal")

    def process_files(self):
        # Get the paths of the selected files
        db_file = self.label_db.cget("text")
        upload_file = self.label_upload.cget("text")
        # print(f"DB File: {db_file}")
        # print(f"Upload File: {upload_file}")
        # Check if the files are selected
        if db_file == "No file selected" or upload_file == "No file selected":
            messagebox.showerror("Error", "Please select both the database and upload files.")
        else:
            try:
                # Update progress bar to 0
                self.progress_bar.set(0)
                self.update_idletasks()

                # remove to upload_file the extension xlsx
                upload_file_cp = upload_file.replace('.xlsx', '')
                # default new file name
                new_file_name = f"{upload_file_cp}-{time.strftime('%Y-%m-%d-%H-%M-%S')}.xlsx"
                # Process the files
               # crteate the new columns data
                new_columns_data, positive_new_values_inserted_count, rows_count = read_file(file_path=upload_file, sheet_name=SHEET_INDEX, columns=FILE_COLUMNS, new_columns_name=NEW_COLUMNS_NAME, progress_bar=self.progress_bar)
                self.progress_bar.set(0.85)
                self.update_idletasks()
                # time.sleep(0.1)

                # insert the new columns in the excel file
                df: pd.DataFrame =  insert_column_in_excel(upload_file, sheet_name=SHEET_INDEX, col_index=COLUMN_INDEX, new_columns_data=new_columns_data)
                # print(f"df: {df[NEW_COLUMNS_NAME]}")
                self.progress_bar.set(0.90)
                self.update_idletasks()
                # time.sleep(0.1)
                
                # Determine save directory based on checkbox
                if self.use_downloads_var.get():
                    save_dir = DOWNLOADS_PATH  # Use Downloads directory
                else:
                    save_dir = filedialog.askdirectory(
                        title="Select Save Directory",
                        initialdir=DOWNLOADS_PATH  # Optional: Set Downloads as the initial directory
                    )
                    if not save_dir:
                        messagebox.showwarning("Save Directory", "No directory selected. Using Downloads directory.")
                        save_dir = DOWNLOADS_PATH
                new_file_name_path = os.path.join(save_dir, new_file_name)
                
                # Create a new excel file from the DataFrame
                create_new_excel_file(df, new_file_name=new_file_name_path,sheet_name=SHEET_NAME)
                self.progress_bar.set(1)
                self.update_idletasks()
                # time.sleep(0.1)
                # Show success message
                messagebox.showinfo("Success", f"New file created: {new_file_name_path}")

                # insert the values in the textbox_output
                self.insert_values_in_textbox_output(positive_new_values_inserted_count, rows_count)
            
            except Exception as e:
                messagebox.showerror("Processing Error", f"An error occurred while processing the files:\n{e}")

    def close_window(self):
        self.destroy()

    def run(self):
        self.mainloop()

# Entry point of the application
if __name__ == "__main__":
    app = ExcelApp()
    app.run()

