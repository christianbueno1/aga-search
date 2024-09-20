import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from aga_search.config import DOWNLOAD_PATH

# Initialize customtkinter
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "dark-blue", "green"

class ExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Set window title and minimum size
        self.title("Excel File Selector")
        self.minsize(700, 400)  # Set minimum width and height

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
        frame_actions.grid_columnconfigure(0, weight=1)
        frame_actions.grid_columnconfigure(1, weight=1)

        # Button to process the selected files
        self.button_process = ctk.CTkButton(frame_actions, text="Process", command=self.process_files, state="disabled")
        self.button_process.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # Button to close the window
        self.button_close = ctk.CTkButton(frame_actions, text="Close", command=self.close_window)
        self.button_close.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    def select_db_file(self):
        # Open a file dialog to select the file
        file_path = filedialog.askopenfilename(
            initialdir=DOWNLOAD_PATH,
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
                # # Import the file_processor here to avoid circular imports
                # from aga_search.file_processor import process_excel_files

                # # Process the files
                # processed_df = process_excel_files(db_file, upload_file)

                # # Ask user where to save or use default
                # save_dir = filedialog.askdirectory(title="Select Save Directory")
                # if not save_dir:
                #     save_dir = DOWNLOAD_PATH  # Use default download path

                # # Import the file_manager here to avoid circular imports
                # from aga_search.file_manager import save_file

                # # Save the processed DataFrame
                # save_file(processed_df, save_dir)

                # messagebox.showinfo("Success", f"The files have been processed and saved to {save_dir}")
                print("Files processed successfully!")
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
