import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

# Try to import pypff (libpff)
try:
    import pypff
    PYPFF_AVAILABLE = True
except ImportError:
    PYPFF_AVAILABLE = False


# --- App Configuration ---
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class OSTExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window settings
        self.title("OST to EML Extractor - Premium Edition")
        self.geometry("700x500")
        self.minsize(600, 450)

        # Variables
        self.ost_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.is_extracting = False
        
        self._setup_ui()

    def _setup_ui(self):
        # Configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=0)
        self.grid_rowconfigure(3, weight=1)

        # --- Sidebar ---
        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=5, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="OST\nExtractor", font=ctk.CTkFont(size=24, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.info_label = ctk.CTkLabel(self.sidebar_frame, text="Convert Outlook\nOST files directly to\nEML locally.", font=ctk.CTkFont(size=12), text_color="gray")
        self.info_label.grid(row=1, column=0, padx=20, pady=10)
        
        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Dark", "Light", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 20))

        # --- Main View ---

        # 1. Input OST Selection
        self.input_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.input_frame.grid(row=0, column=1, columnspan=3, padx=20, pady=(20, 0), sticky="nsew")
        self.input_frame.grid_columnconfigure(0, weight=1)

        self.input_label = ctk.CTkLabel(self.input_frame, text="1. Select OST File", font=ctk.CTkFont(size=14, weight="bold"))
        self.input_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        self.input_entry = ctk.CTkEntry(self.input_frame, textvariable=self.ost_path_var, placeholder_text="Path to .ost file...")
        self.input_entry.grid(row=1, column=0, sticky="ew", padx=(0, 10))

        self.input_btn = ctk.CTkButton(self.input_frame, text="Browse...", command=self.browse_ost, width=100)
        self.input_btn.grid(row=1, column=1)

        # 2. Output Directory Selection
        self.output_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.output_frame.grid(row=1, column=1, columnspan=3, padx=20, pady=(20, 0), sticky="nsew")
        self.output_frame.grid_columnconfigure(0, weight=1)

        self.output_label = ctk.CTkLabel(self.output_frame, text="2. Select Output Folder", font=ctk.CTkFont(size=14, weight="bold"))
        self.output_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        self.output_entry = ctk.CTkEntry(self.output_frame, textvariable=self.output_dir_var, placeholder_text="Folder to save .eml files...")
        self.output_entry.grid(row=1, column=0, sticky="ew", padx=(0, 10))

        self.output_btn = ctk.CTkButton(self.output_frame, text="Browse...", command=self.browse_output, width=100)
        self.output_btn.grid(row=1, column=1)

        # 3. Status and Run
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.grid(row=2, column=1, columnspan=3, padx=20, pady=(30, 0), sticky="nsew")
        self.action_frame.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(self.action_frame, text="Ready", text_color="green", font=ctk.CTkFont(weight="bold"))
        self.status_label.grid(row=0, column=0, pady=10)

        self.progress_bar = ctk.CTkProgressBar(self.action_frame)
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        self.progress_bar.set(0)

        self.run_btn = ctk.CTkButton(self.action_frame, text="Extract Messages", font=ctk.CTkFont(size=16, weight="bold"), height=40, command=self.start_extraction)
        self.run_btn.grid(row=2, column=0, pady=(10, 20), padx=20, sticky="ew")

        # 4. Log Output
        self.log_textbox = ctk.CTkTextbox(self, wrap="word", font=ctk.CTkFont(size=11))
        self.log_textbox.grid(row=3, column=1, columnspan=3, padx=20, pady=20, sticky="nsew")
        self.log_textbox.insert("0.0", "--- OST Extraction Log ---\n")
        self.log_textbox.configure(state="disabled")
        
        if not PYPFF_AVAILABLE:
            self.log_message("WARNING: 'pypff' library is not installed. Extraction will not work until you install it.")
            self.status_label.configure(text="Dependency missing: pypff", text_color="red")
            self.run_btn.configure(state="disabled")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def browse_ost(self):
        filename = filedialog.askopenfilename(
            title="Select OST File",
            filetypes=(("Outlook Files", "*.ost"), ("All Files", "*.*"))
        )
        if filename:
            self.ost_path_var.set(filename)

    def browse_output(self):
        dirname = filedialog.askdirectory(title="Select Output Folder")
        if dirname:
            self.output_dir_var.set(dirname)

    def log_message(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"{message}\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")
        self.update()

    def start_extraction(self):
        if self.is_extracting:
            return

        ost_path = os.path.normpath(os.path.abspath(os.path.expanduser(self.ost_path_var.get())))
        out_path = os.path.normpath(os.path.abspath(os.path.expanduser(self.output_dir_var.get())))

        if not ost_path or not os.path.exists(ost_path):
            messagebox.showerror("Error", "Please select a valid OST file.")
            return
        
        if not out_path or not os.path.exists(out_path):
            messagebox.showerror("Error", "Please select a valid Output Folder.")
            return

        self.is_extracting = True
        self.run_btn.configure(state="disabled", text="Extracting...")
        self.progress_bar.set(0)
        self.status_label.configure(text="Extracting items...", text_color="yellow")
        self.log_message(f"Starting extraction of {os.path.basename(ost_path)}...")
        
        # Run extraction in a background thread to prevent GUI freezing
        thread = threading.Thread(target=self.extract_ost, args=(ost_path, out_path))
        thread.daemon = True
        thread.start()

    def extract_ost(self, ost_file, out_dir):
        try:
            # Initialize libpff
            pst = pypff.file()
            pst.open(ost_file)
            root = pst.get_root_folder()
            
            # Count total items for progress tracking (simplified recursive count)
            self.log_message("Scanning parsing folder structure...")
            total_items = self._count_items(root)
            self.log_message(f"Found ~{total_items} items.")
            
            self.processed_items = 0
            self._process_folder(root, out_dir, total_items)

            self.status_label.configure(text="Extraction Complete", text_color="green")
            self.log_message("Extraction successfully completed!")
            messagebox.showinfo("Success", f"Extraction finished!\nOutput in: {out_dir}")

        except Exception as e:
            self.status_label.configure(text="Extraction Failed", text_color="red")
            self.log_message(f"ERROR: {str(e)}")
            messagebox.showerror("Error", f"An error occurred during extraction:\n{str(e)}")
        finally:
            if 'pst' in locals() and hasattr(pst, 'close'):
                pst.close()
            self.is_extracting = False
            self.run_btn.configure(state="normal", text="Extract Messages")

    def _count_items(self, folder):
        count = folder.get_number_of_sub_messages()
        for i in range(folder.get_number_of_sub_folders()):
            sub = folder.get_sub_folder(i)
            count += self._count_items(sub)
        return count

    def _process_folder(self, folder, base_dir, total_items):
        folder_name = folder.get_name()
        if not folder_name:
            folder_name = "Root"
            
        # Clean folder name for Windows
        clean_folder = "".join([c for c in folder_name if c.isalpha() or c.isdigit() or c in ' -_']).rstrip()
        current_dir = os.path.join(base_dir, clean_folder)
        
        if not os.path.exists(current_dir):
            os.makedirs(current_dir)

        # Extract Messages
        for i in range(folder.get_number_of_sub_messages()):
            message = folder.get_sub_message(i)
            
            # Save as plain text or try to get headers/body for EML structure
            # Not all items are purely emails. We will dump the transport headers and body.
            try:
                subject = message.get_subject() or "No Subject"
                # Strip invalid chars and truncate to 100 chars to prevent MAX_PATH (>260) errors on Windows
                clean_subject = "".join([c for c in subject if c.isalpha() or c.isdigit() or c in ' -_']).rstrip()[:100]
                # Ensure unique filename
                filename = f"{clean_subject}_{i}.eml"
                filepath = os.path.join(current_dir, filename)

                with open(filepath, 'wb') as f:
                    headers = message.get_transport_headers()
                    if headers:
                        f.write(headers.encode('utf-8', errors='ignore'))
                        f.write(b'\r\n\r\n')
                    
                    try:
                        body = message.get_plain_text_body()
                        if body:
                            f.write(body)
                        else:
                            try:
                                html_body = message.get_html_body()
                                if html_body:
                                    f.write(html_body)
                            except Exception:
                                pass # No HTML body found or corrupted structural size
                    except Exception:
                        pass # No plain text body found

                self.processed_items += 1
                
                # Update progress
                if self.processed_items % 5 == 0 or self.processed_items == total_items:
                    progress = float(self.processed_items) / float(total_items if total_items > 0 else 1)
                    self.progress_bar.set(progress)
                    # Update status in the thread-safe way by scheduling UI updates
                    
            except Exception as e:
                self.log_message(f"Skipped item {i} in {folder_name} due to error: {e}")

        # Process subfolders
        for i in range(folder.get_number_of_sub_folders()):
            subfolder = folder.get_sub_folder(i)
            self._process_folder(subfolder, current_dir, total_items)


if __name__ == "__main__":
    app = OSTExtractorApp()
    app.mainloop()

