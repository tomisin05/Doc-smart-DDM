import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox
import os

class WordControllerGUI:
    def __init__(self):
        self.word = None
        self.doc = None
        self.current_file = None
        
        self.root = tk.Tk()
        self.root.title("Word Document Controller")
        self.root.geometry("400x200")
        
        # File path display
        self.file_label = tk.Label(self.root, text="No file selected", wraplength=350)
        self.file_label.pack(pady=10)
        
        # Buttons
        tk.Button(self.root, text="Select Document", command=self.select_file, width=20).pack(pady=5)
        tk.Button(self.root, text="Open Document", command=self.open_document, width=20).pack(pady=5)
        tk.Button(self.root, text="Close Document", command=self.close_document, width=20).pack(pady=5)
        tk.Button(self.root, text="Close Document (No Save)", command=self.close_no_save, width=20).pack(pady=5)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx *.doc"), ("All Files", "*.*")]
        )
        if file_path:
            self.current_file = file_path
            self.file_label.config(text=f"Selected: {os.path.basename(file_path)}")
    
    def open_document(self):
        if not self.current_file:
            messagebox.showwarning("Warning", "Please select a document first")
            return
        
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = True
            self.doc = self.word.Documents.Open(self.current_file)
            messagebox.showinfo("Success", f"Opened: {os.path.basename(self.current_file)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open document: {e}")
    
    def close_document(self):
        if not self.doc:
            messagebox.showwarning("Warning", "No document is open")
            return
            
        # Ask user about saving
        save_choice = messagebox.askyesnocancel("Save Document?", "Do you want to save changes before closing?")
        if save_choice is None:  # User clicked Cancel
            return
            
        try:
            self.doc.Close(SaveChanges=save_choice)
            self.doc = None
            
            # Check if any documents are still open
            if self.word and self.word.Documents.Count == 0:
                # No documents left, quit Word to avoid black window
                self.word.Quit()
                self.word = None
                messagebox.showinfo("Success", "Document closed and Word application closed")
            else:
                messagebox.showinfo("Success", "Document closed")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to close: {e}")
    
    def close_no_save(self):
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
                self.doc = None
                
            if self.word:
                self.word.Quit()
                self.word = None
                
            messagebox.showinfo("Success", "Document closed without saving")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to close: {e}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = WordControllerGUI()
    app.run()