import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
import threading
from pub_v1 import * 
from compare_v2 import *
from final_result_v1 import *


class DocumentRevisionGUI:
    """Main GUI window for Document Revision Tool using tkinter."""
    
    def __init__(self, root):
        self.root = root
        self.client_file = ""
        self.home_file = ""
        self.output_file = ""
        self.output_path = ""
        self.is_running = False
        
        # Color scheme - Professional blue
        self.button_color = "#4A90E2"
        self.button_hover = "#357ABD"
        self.button_active = "#2868A6"
        self.bg_color = "#f5f5f5"
        
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface."""
        self.root.title("Document Revision Comparison Tool")
        self.root.geometry("850x700")
        self.root.configure(bg=self.bg_color)
        
        # Make window resizable
        self.root.resizable(True, True)
        
        # Configure grid weights for responsive layout
        self.root.grid_rowconfigure(4, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Title
        title_frame = tk.Frame(self.root, bg=self.bg_color)
        title_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        title_label = tk.Label(
            title_frame,
            text="Document Revision Comparison Tool",
            font=("Arial", 18, "bold"),
            bg=self.bg_color,
            fg="#333333"
        )
        title_label.pack()
        
        # Input files section
        self.create_input_section()
        
        # Output configuration section
        self.create_output_section()
        
        # Action buttons
        self.create_button_section()
        
        # Console output
        self.create_console_section()
        
        # Status bar
        self.create_status_bar()
    
    def create_input_section(self):
        """Create input files selection section."""
        input_frame = tk.LabelFrame(
            self.root,
            text="Input Files",
            font=("Arial", 10, "bold"),
            bg=self.bg_color,
            fg="#333333",
            padx=15,
            pady=10
        )
        input_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        input_frame.grid_columnconfigure(1, weight=1)
        
        # Client file
        tk.Label(
            input_frame,
            text="Client File:",
            font=("Arial", 10),
            bg=self.bg_color,
            width=12,
            anchor="w"
        ).grid(row=0, column=0, sticky="w", pady=5)
        
        self.client_input = tk.Entry(
            input_frame,
            font=("Arial", 9),
            state="readonly",
            readonlybackground="white"
        )
        self.client_input.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        client_btn = tk.Button(
            input_frame,
            text="Browse...",
            command=self.browse_client_file,
            bg=self.button_color,
            fg="white",
            font=("Arial", 9, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            padx=15,
            pady=5
        )
        client_btn.grid(row=0, column=2, pady=5)
        self.bind_button_hover(client_btn)
        
        # Home file
        tk.Label(
            input_frame,
            text="Home File:",
            font=("Arial", 10),
            bg=self.bg_color,
            width=12,
            anchor="w"
        ).grid(row=1, column=0, sticky="w", pady=5)
        
        self.home_input = tk.Entry(
            input_frame,
            font=("Arial", 9),
            state="readonly",
            readonlybackground="white"
        )
        self.home_input.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        home_btn = tk.Button(
            input_frame,
            text="Browse...",
            command=self.browse_home_file,
            bg=self.button_color,
            fg="white",
            font=("Arial", 9, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            padx=15,
            pady=5
        )
        home_btn.grid(row=1, column=2, pady=5)
        self.bind_button_hover(home_btn)
    
    def create_output_section(self):
        """Create output configuration section."""
        output_frame = tk.LabelFrame(
            self.root,
            text="Output Configuration",
            font=("Arial", 10, "bold"),
            bg=self.bg_color,
            fg="#333333",
            padx=15,
            pady=10
        )
        output_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        output_frame.grid_columnconfigure(1, weight=1)
        
        # Output filename
        tk.Label(
            output_frame,
            text="Output Name:",
            font=("Arial", 10),
            bg=self.bg_color,
            width=12,
            anchor="w"
        ).grid(row=0, column=0, sticky="w", pady=5)
        
        self.output_name_input = tk.Entry(
            output_frame,
            font=("Arial", 9)
        )
        self.output_name_input.insert(0, "result_one.xlsx")
        self.output_name_input.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        # Output path
        tk.Label(
            output_frame,
            text="Output Path:",
            font=("Arial", 10),
            bg=self.bg_color,
            width=12,
            anchor="w"
        ).grid(row=1, column=0, sticky="w", pady=5)
        
        self.output_path_input = tk.Entry(
            output_frame,
            font=("Arial", 9),
            state="readonly",
            readonlybackground="white"
        )
        self.output_path_input.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        path_btn = tk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output_path,
            bg=self.button_color,
            fg="white",
            font=("Arial", 9, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            padx=15,
            pady=5
        )
        path_btn.grid(row=1, column=2, pady=5)
        self.bind_button_hover(path_btn)
        
        # Help text
        help_label = tk.Label(
            output_frame,
            text="(If no path is selected, output will be saved in the current directory)",
            font=("Arial", 8, "italic"),
            bg=self.bg_color,
            fg="#666666"
        )
        help_label.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 5))
    
    def create_button_section(self):
        """Create action buttons section."""
        button_frame = tk.Frame(self.root, bg=self.bg_color)
        button_frame.grid(row=3, column=0, pady=20)
        
        # Execute button
        self.execute_btn = tk.Button(
            button_frame,
            text="Execute",
            command=self.execute_comparison,
            bg=self.button_color,
            fg="white",
            font=("Arial", 11, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            width=15,
            height=2
        )
        self.execute_btn.pack(side=tk.LEFT, padx=10)
        self.bind_button_hover(self.execute_btn)
        
        # Reset button
        self.reset_btn = tk.Button(
            button_frame,
            text="Reset",
            command=self.reset_form,
            bg=self.button_color,
            fg="white",
            font=("Arial", 11, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            width=15,
            height=2
        )
        self.reset_btn.pack(side=tk.LEFT, padx=10)
        self.bind_button_hover(self.reset_btn)
    
    def create_console_section(self):
        """Create console output section."""
        console_frame = tk.LabelFrame(
            self.root,
            text="Console Output",
            font=("Arial", 10, "bold"),
            bg=self.bg_color,
            fg="#333333",
            padx=10,
            pady=10
        )
        console_frame.grid(row=4, column=0, sticky="nsew", padx=20, pady=10)
        console_frame.grid_rowconfigure(0, weight=1)
        console_frame.grid_columnconfigure(0, weight=1)
        
        # Create scrolled text widget
        self.console = scrolledtext.ScrolledText(
            console_frame,
            font=("Courier New", 9),
            wrap=tk.WORD,
            state="disabled",
            bg="white",
            height=15
        )
        self.console.grid(row=0, column=0, sticky="nsew")
        
        # Configure tags for colored text
        self.console.tag_config("error", foreground="red")
        self.console.tag_config("success", foreground="green")
        self.console.tag_config("info", foreground="blue")
    
    def create_status_bar(self):
        """Create status bar at the bottom."""
        status_frame = tk.Frame(self.root, bg="#e0e0e0", relief=tk.SUNKEN)
        status_frame.grid(row=5, column=0, sticky="ew")
        
        self.status_label = tk.Label(
            status_frame,
            text="Ready",
            font=("Arial", 9),
            bg="#e0e0e0",
            anchor="w",
            padx=10
        )
        self.status_label.pack(fill=tk.X)
    
    def bind_button_hover(self, button):
        """Bind hover effects to button."""
        def on_enter(e):
            if button['state'] != 'disabled':
                button['bg'] = self.button_hover
        
        def on_leave(e):
            if button['state'] != 'disabled':
                button['bg'] = self.button_color
        
        def on_press(e):
            if button['state'] != 'disabled':
                button['bg'] = self.button_active
        
        def on_release(e):
            if button['state'] != 'disabled':
                button['bg'] = self.button_hover
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
        button.bind("<ButtonPress-1>", on_press)
        button.bind("<ButtonRelease-1>", on_release)
    
    def browse_client_file(self):
        """Open file dialog to select client file."""
        filename = filedialog.askopenfilename(
            title="Select Client File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.client_file = filename
            self.client_input.config(state="normal")
            self.client_input.delete(0, tk.END)
            self.client_input.insert(0, filename)
            self.client_input.config(state="readonly")
            self.log_message(f"Client file selected: {filename}")
            self.update_status("Client file selected")
    
    def browse_home_file(self):
        """Open file dialog to select home file."""
        filename = filedialog.askopenfilename(
            title="Select Home File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.home_file = filename
            self.home_input.config(state="normal")
            self.home_input.delete(0, tk.END)
            self.home_input.insert(0, filename)
            self.home_input.config(state="readonly")
            self.log_message(f"Home file selected: {filename}")
            self.update_status("Home file selected")
    
    def browse_output_path(self):
        """Open directory dialog to select output path."""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_path = directory
            self.output_path_input.config(state="normal")
            self.output_path_input.delete(0, tk.END)
            self.output_path_input.insert(0, directory)
            self.output_path_input.config(state="readonly")
            self.log_message(f"Output path selected: {directory}")
            self.update_status("Output path selected")
    
    def validate_inputs(self):
        """Validate that all required inputs are provided."""
        if not self.client_file:
            messagebox.showwarning(
                "Missing Input",
                "Please select a Client file."
            )
            return False
        
        if not self.home_file:
            messagebox.showwarning(
                "Missing Input",
                "Please select a Home file."
            )
            return False
        
        # Check if files exist
        if not Path(self.client_file).exists():
            messagebox.showerror(
                "File Not Found",
                f"Client file not found:\n{self.client_file}"
            )
            return False
        
        if not Path(self.home_file).exists():
            messagebox.showerror(
                "File Not Found",
                f"Home file not found:\n{self.home_file}"
            )
            return False
        
        return True
    
    def execute_comparison(self):
        """Execute the document comparison process."""
        if self.is_running:
            messagebox.showinfo(
                "Process Running",
                "A comparison process is already running. Please wait for it to complete."
            )
            return
        
        if not self.validate_inputs():
            return
        
        # Prepare output file path
        output_name = self.output_name_input.get().strip()
        if not output_name:
            output_name = "result_one.xlsx"
        
        if not output_name.endswith('.xlsx'):
            output_name += '.xlsx'
        
        if self.output_path:
            self.output_file = str(Path(self.output_path) / output_name)
        else:
            self.output_file = output_name
        
        # Clear console
        self.console.config(state="normal")
        self.console.delete(1.0, tk.END)
        self.console.config(state="disabled")
        
        self.log_message("Starting comparison process...", "info")
        self.log_message(f"Client file: {self.client_file}")
        self.log_message(f"Home file: {self.home_file}")
        self.log_message(f"Output file: {self.output_file}\n")
        
        # Disable buttons during execution
        self.is_running = True
        self.execute_btn.config(state="disabled", bg="#cccccc", cursor="")
        self.reset_btn.config(state="disabled", bg="#cccccc", cursor="")
        self.update_status("Running comparison...")
        
        # Run in separate thread
        thread = threading.Thread(target=self.run_comparison, daemon=True)
        thread.start()
    
    def run_comparison(self):
        """Run the comparison process in a separate thread."""
        try:
           
            
            self.log_message("="*50)
            self.log_message("Document Revision Comparison Tool")
            self.log_message("="*50 + "\n")
            
            # Step 1: Load data
            self.log_message("Step 1: Loading data files...")
            loader = DataLoader(self.client_file, self.home_file)
            client_df, home_df = loader.load_files()
            
            # Step 2: Format client file
            self.log_message("Step 2: Formatting client file...")
            formatter = ClientFormatter(client_df)
            client_df = formatter.process()
            formatter.save_formatted_file('client_formatted.csv')
            
            # Step 3: Process home file
            self.log_message("Step 3: Processing home file...")
            home_processor = HomeProcessor(home_df)
            home_df = home_processor.remove_duplicates()
            
            # Step 4: Compare documents
            self.log_message("Step 4: Comparing documents...")
            comparator = RevisionComparator(client_df, home_df)
            result_df = comparator.process_comparisons()
            
            # Step 5: Save results
            self.log_message("Step 5: Saving results...")
            result_gen = ResultGenerator(result_df, self.output_file)
            result_gen.save_results()
            
            # Step 6: Apply formatting
            self.log_message("Step 6: Applying cell colors...")
            excel_formatter = ExcelFormatter(self.output_file)
            excel_formatter.apply_colors()
            
            # Step 7: Generate summary
            self.log_message("Step 7: Generating summary...")
            result_gen.generate_summary()
            
            self.log_message("\n✓ Process completed successfully!", "success")
            self.root.after(0, self.on_success)
            
        except Exception as e:
            error_msg = f"Error during execution: {str(e)}"
            self.log_message(f"\n❌ {error_msg}", "error")
            self.root.after(0, lambda: self.on_error(error_msg))
    
    def on_success(self):
        """Handle successful completion."""
        self.is_running = False
        self.execute_btn.config(state="normal", bg=self.button_color, cursor="hand2")
        self.reset_btn.config(state="normal", bg=self.button_color, cursor="hand2")
        self.update_status("Process completed successfully")
        
        messagebox.showinfo(
            "Success",
            f"Process completed successfully!\n\nOutput file: {self.output_file}"
        )
    
    def on_error(self, error_msg):
        """Handle errors."""
        self.is_running = False
        self.execute_btn.config(state="normal", bg=self.button_color, cursor="hand2")
        self.reset_btn.config(state="normal", bg=self.button_color, cursor="hand2")
        self.update_status("Error occurred")
        
        messagebox.showerror("Execution Error", error_msg)
    
    def log_message(self, message, tag=None):
        """Add message to console output."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {message}\n"
        
        self.console.config(state="normal")
        if tag:
            self.console.insert(tk.END, formatted_msg, tag)
        else:
            self.console.insert(tk.END, formatted_msg)
        self.console.see(tk.END)
        self.console.config(state="disabled")
    
    def update_status(self, message):
        """Update status bar message."""
        self.status_label.config(text=message)
    
    def reset_form(self):
        """Reset all form fields."""
        if self.is_running:
            messagebox.showinfo(
                "Process Running",
                "Cannot reset while a process is running."
            )
            return
        
        reply = messagebox.askyesno(
            "Reset Form",
            "Are you sure you want to reset all fields?"
        )
        
        if reply:
            self.client_file = ""
            self.home_file = ""
            self.output_path = ""
            
            self.client_input.config(state="normal")
            self.client_input.delete(0, tk.END)
            self.client_input.config(state="readonly")
            
            self.home_input.config(state="normal")
            self.home_input.delete(0, tk.END)
            self.home_input.config(state="readonly")
            
            self.output_path_input.config(state="normal")
            self.output_path_input.delete(0, tk.END)
            self.output_path_input.config(state="readonly")
            
            self.output_name_input.delete(0, tk.END)
            self.output_name_input.insert(0, "result_one.xlsx")
            
            self.console.config(state="normal")
            self.console.delete(1.0, tk.END)
            self.console.config(state="disabled")
            
            self.update_status("Ready")
            self.log_message("Form reset successfully", "info")


def main():
    """Main entry point for GUI application."""
    root = tk.Tk()
    app = DocumentRevisionGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()