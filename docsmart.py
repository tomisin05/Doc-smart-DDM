#!/usr/bin/env python3
"""
Doc-smart: Desktop application for managing Word documents for debate preparation
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import os
import subprocess
import platform
from datetime import datetime
from pathlib import Path
import webbrowser
from typing import Dict, List, Optional, Any

# Try to import Windows COM for Word automation
try:
    import win32com.client
    WORD_COM_AVAILABLE = True
except ImportError:
    WORD_COM_AVAILABLE = False

class DocEntry:
    def __init__(self, id: str, name: str, source_type: str, url: str = None, 
                 file_path: str = None, tags: List[str] = None, team_id: str = None,
                 favorite: bool = False, is_open: bool = False, 
                 last_opened_at: float = None, created_at: float = None):
        self.id = id
        self.name = name
        self.source_type = source_type  # "url" or "file"
        self.url = url
        self.file_path = file_path
        self.tags = tags or []
        self.team_id = team_id
        self.favorite = favorite
        self.is_open = is_open
        self.last_opened_at = last_opened_at
        self.created_at = created_at or datetime.now().timestamp()

    def to_dict(self) -> Dict[str, Any]:
        return {
            'id': self.id,
            'name': self.name,
            'source_type': self.source_type,
            'url': self.url,
            'file_path': self.file_path,
            'tags': self.tags,
            'team_id': self.team_id,
            'favorite': self.favorite,
            'is_open': self.is_open,
            'last_opened_at': self.last_opened_at,
            'created_at': self.created_at
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DocEntry':
        return cls(**data)

class Team:
    def __init__(self, id: str, name: str, created_at: float = None):
        self.id = id
        self.name = name
        self.created_at = created_at or datetime.now().timestamp()

    def to_dict(self) -> Dict[str, Any]:
        return {
            'id': self.id,
            'name': self.name,
            'created_at': self.created_at
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Team':
        return cls(**data)

class DocSmartApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Doc-smart - Debate Document Manager")
        self.root.geometry("1200x800")
        
        # Data storage
        self.docs: Dict[str, DocEntry] = {}
        self.teams: Dict[str, Team] = {}
        self.selected_team_id: Optional[str] = None
        self.search_text = tk.StringVar()
        self.favorite_only = tk.BooleanVar()
        
        # Load data
        self.data_file = Path.home() / ".docsmart" / "data.json"
        self.load_data()
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(header_frame, text="Doc-smart", font=("Arial", 20, "bold")).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="Fast access to your debate prep Word documents", 
                 font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 0))
        
        # Buttons
        button_frame = ttk.Frame(header_frame)
        button_frame.pack(side=tk.RIGHT)
        
        ttk.Button(button_frame, text="Add Document", command=self.add_document).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Add Team", command=self.add_team).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Import Folder", command=self.import_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Export Data", command=self.export_data).pack(side=tk.LEFT, padx=2)
        ttk.Separator(button_frame, orient='vertical').pack(side=tk.LEFT, padx=5, fill=tk.Y)
        ttk.Button(button_frame, text="Open Selected", command=self.open_selected_documents).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Close Selected", command=self.close_selected_documents).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Close All Open", command=self.close_all_documents).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Open Team", command=self.open_team_documents).pack(side=tk.LEFT, padx=2)
        
        # Content frame
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left sidebar
        sidebar_frame = ttk.Frame(content_frame)
        sidebar_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # Teams section
        teams_label = ttk.Label(sidebar_frame, text="Teams", font=("Arial", 12, "bold"))
        teams_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.teams_listbox = tk.Listbox(sidebar_frame, width=25, height=10)
        self.teams_listbox.pack(fill=tk.Y, expand=True)
        self.teams_listbox.bind('<<ListboxSelect>>', self.on_team_select)
        self.teams_listbox.bind('<Button-3>', self.show_team_context_menu)  # Right-click
        
        # Team context menu
        self.team_context_menu = tk.Menu(self.root, tearoff=0)
        self.team_context_menu.add_command(label="Rename Team", command=self.rename_selected_team)
        self.team_context_menu.add_command(label="Delete Team", command=self.delete_selected_team)
        
        # Search and filters
        search_frame = ttk.LabelFrame(sidebar_frame, text="Search & Filters", padding=10)
        search_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(search_frame, text="Search:").pack(anchor=tk.W)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_text)
        search_entry.pack(fill=tk.X, pady=(0, 5))
        search_entry.bind('<KeyRelease>', self.on_search_change)
        
        ttk.Checkbutton(search_frame, text="Favorites only", 
                       variable=self.favorite_only, command=self.refresh_documents).pack(anchor=tk.W)
        
        # Main document area
        docs_frame = ttk.Frame(content_frame)
        docs_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Document list
        columns = ('Name', 'Team', 'Tags', 'Status', 'Last Opened')
        self.docs_tree = ttk.Treeview(docs_frame, columns=columns, show='headings', height=20, selectmode='extended')
        
        for col in columns:
            self.docs_tree.heading(col, text=col)
            self.docs_tree.column(col, width=150)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(docs_frame, orient=tk.VERTICAL, command=self.docs_tree.yview)
        self.docs_tree.configure(yscrollcommand=scrollbar.set)
        
        self.docs_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Context menu for documents
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Open in Word", command=self.open_selected_documents)
        self.context_menu.add_command(label="Close in Word", command=self.close_selected_documents)
        self.context_menu.add_command(label="Mark as Favorite", command=self.toggle_favorite_selected)
        self.context_menu.add_command(label="Edit", command=self.edit_selected_document)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Remove Selected", command=self.remove_selected_documents)
        
        self.docs_tree.bind("<Button-3>", self.show_context_menu)
        self.docs_tree.bind("<Double-1>", self.open_selected_documents)
        
        # Initial load
        self.refresh_teams()
        self.refresh_documents()
    
    def generate_id(self, prefix: str = "id") -> str:
        import random
        import string
        return f"{prefix}_{''.join(random.choices(string.ascii_lowercase + string.digits, k=8))}"
    
    def save_data(self):
        """Save data to JSON file"""
        self.data_file.parent.mkdir(exist_ok=True)
        
        data = {
            'docs': {id: doc.to_dict() for id, doc in self.docs.items()},
            'teams': {id: team.to_dict() for id, team in self.teams.items()},
            'selected_team_id': self.selected_team_id
        }
        
        with open(self.data_file, 'w') as f:
            json.dump(data, f, indent=2)
    
    def load_data(self):
        """Load data from JSON file"""
        if not self.data_file.exists():
            return
            
        try:
            with open(self.data_file, 'r') as f:
                data = json.load(f)
            
            self.docs = {id: DocEntry.from_dict(doc_data) 
                        for id, doc_data in data.get('docs', {}).items()}
            self.teams = {id: Team.from_dict(team_data) 
                         for id, team_data in data.get('teams', {}).items()}
            self.selected_team_id = data.get('selected_team_id')
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")
    
    def open_in_word(self, doc: DocEntry):
        """Open document in Microsoft Word"""
        try:
            if doc.source_type == "url":
                # Try Word protocol first, fallback to browser
                word_url = f"ms-word:ofe|u|{doc.url}"
                try:
                    if platform.system() == "Windows":
                        os.startfile(word_url)
                    else:
                        webbrowser.open(doc.url)
                except:
                    webbrowser.open(doc.url)
            else:
                # Open local file
                if not doc.file_path or not os.path.exists(doc.file_path):
                    messagebox.showerror("Error", "File not found. Please check the file path.")
                    return
                
                if platform.system() == "Windows":
                    os.startfile(doc.file_path)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", doc.file_path])
                else:  # Linux
                    subprocess.run(["xdg-open", doc.file_path])
            
            # Mark as opened
            doc.is_open = True
            doc.last_opened_at = datetime.now().timestamp()
            self.save_data()
            self.refresh_documents()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open document: {e}")
    
    def add_document(self):
        """Add new document dialog"""
        dialog = DocumentDialog(self.root, self.teams)
        if dialog.result:
            doc_data = dialog.result
            doc_id = self.generate_id("doc")
            
            doc = DocEntry(
                id=doc_id,
                name=doc_data['name'],
                source_type=doc_data['source_type'],
                url=doc_data.get('url'),
                file_path=doc_data.get('file_path'),
                tags=doc_data.get('tags', []),
                team_id=doc_data.get('team_id')
            )
            
            self.docs[doc_id] = doc
            self.save_data()
            self.refresh_documents()
            messagebox.showinfo("Success", f"Document '{doc.name}' added successfully!")
    
    def add_team(self):
        """Add new team"""
        name = simpledialog.askstring("Add Team", "Enter team name:")
        if name and name.strip():
            team_id = self.generate_id("team")
            team = Team(id=team_id, name=name.strip())
            self.teams[team_id] = team
            self.save_data()
            self.refresh_teams()
            messagebox.showinfo("Success", f"Team '{name}' added successfully!")
    
    def import_folder(self):
        """Import Word documents from a folder"""
        folder_path = filedialog.askdirectory(title="Select folder containing Word documents")
        if not folder_path:
            return
        
        word_extensions = ['.docx', '.doc']
        imported_count = 0
        
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if any(file.lower().endswith(ext) for ext in word_extensions):
                    file_path = os.path.join(root, file)
                    doc_id = self.generate_id("doc")
                    
                    doc = DocEntry(
                        id=doc_id,
                        name=file,
                        source_type="file",
                        file_path=file_path
                    )
                    
                    self.docs[doc_id] = doc
                    imported_count += 1
        
        if imported_count > 0:
            self.save_data()
            self.refresh_documents()
            messagebox.showinfo("Success", f"Imported {imported_count} documents!")
        else:
            messagebox.showinfo("Info", "No Word documents found in the selected folder.")
    
    def export_data(self):
        """Export data to JSON file"""
        file_path = filedialog.asksaveasfilename(
            title="Export Data",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        
        if file_path:
            try:
                data = {
                    'docs': {id: doc.to_dict() for id, doc in self.docs.items()},
                    'teams': {id: team.to_dict() for id, team in self.teams.items()}
                }
                
                with open(file_path, 'w') as f:
                    json.dump(data, f, indent=2)
                
                messagebox.showinfo("Success", "Data exported successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {e}")
    
    def refresh_teams(self):
        """Refresh teams listbox"""
        self.teams_listbox.delete(0, tk.END)
        self.teams_listbox.insert(0, "All Documents")
        self.teams_listbox.insert(1, "Ungrouped")
        
        for team in sorted(self.teams.values(), key=lambda t: t.name):
            self.teams_listbox.insert(tk.END, team.name)
    
    def refresh_documents(self):
        """Refresh documents treeview"""
        # Clear existing items
        for item in self.docs_tree.get_children():
            self.docs_tree.delete(item)
        
        # Filter documents
        filtered_docs = []
        search_term = self.search_text.get().lower()
        
        for doc in self.docs.values():
            # Team filter
            if self.selected_team_id == "ungrouped" and doc.team_id:
                continue
            elif self.selected_team_id and self.selected_team_id != "ungrouped" and doc.team_id != self.selected_team_id:
                continue
            
            # Search filter
            if search_term:
                if (search_term not in doc.name.lower() and 
                    not any(search_term in tag.lower() for tag in doc.tags)):
                    continue
            
            # Favorite filter
            if self.favorite_only.get() and not doc.favorite:
                continue
            
            filtered_docs.append(doc)
        
        # Sort documents (favorites first, then by last opened, then by name)
        filtered_docs.sort(key=lambda d: (
            not d.favorite,
            -(d.last_opened_at or 0),
            d.name.lower()
        ))
        
        # Add to treeview
        for doc in filtered_docs:
            team_name = self.teams[doc.team_id].name if doc.team_id and doc.team_id in self.teams else "—"
            tags_str = ", ".join(doc.tags) if doc.tags else "—"
            status = "Open" if doc.is_open else "Closed"
            last_opened = datetime.fromtimestamp(doc.last_opened_at).strftime("%Y-%m-%d %H:%M") if doc.last_opened_at else "—"
            
            # Add star for favorites
            name_display = f"★ {doc.name}" if doc.favorite else doc.name
            
            self.docs_tree.insert('', tk.END, values=(name_display, team_name, tags_str, status, last_opened))
    
    def on_team_select(self, event):
        """Handle team selection"""
        selection = self.teams_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        if index == 0:  # All Documents
            self.selected_team_id = None
        elif index == 1:  # Ungrouped
            self.selected_team_id = "ungrouped"
        else:
            team_name = self.teams_listbox.get(index)
            for team_id, team in self.teams.items():
                if team.name == team_name:
                    self.selected_team_id = team_id
                    break
        
        self.refresh_documents()
    
    def on_search_change(self, event):
        """Handle search text change"""
        self.refresh_documents()
    
    def show_context_menu(self, event):
        """Show context menu for documents"""
        item = self.docs_tree.identify_row(event.y)
        if item:
            self.docs_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def get_selected_document(self) -> Optional[DocEntry]:
        """Get currently selected document (first one if multiple selected)"""
        selection = self.docs_tree.selection()
        if not selection:
            return None
        
        item = selection[0]
        values = self.docs_tree.item(item, 'values')
        doc_name = values[0].replace("★ ", "")  # Remove star if present
        
        for doc in self.docs.values():
            if doc.name == doc_name:
                return doc
        return None
    
    def get_selected_documents(self) -> List[DocEntry]:
        """Get all currently selected documents"""
        selection = self.docs_tree.selection()
        if not selection:
            return []
        
        selected_docs = []
        for item in selection:
            values = self.docs_tree.item(item, 'values')
            doc_name = values[0].replace("★ ", "")  # Remove star if present
            
            for doc in self.docs.values():
                if doc.name == doc_name:
                    selected_docs.append(doc)
                    break
        
        return selected_docs
    
    def open_selected_documents(self, event=None):
        """Open selected documents in Word"""
        docs = self.get_selected_documents()
        if not docs:
            return
        
        opened_count = 0
        for doc in docs:
            try:
                self.open_in_word(doc)
                opened_count += 1
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open '{doc.name}': {e}")
        
        if opened_count > 0:
            messagebox.showinfo("Success", f"Opened {opened_count} document(s)!")
    
    def toggle_favorite_selected(self):
        """Toggle favorite status of selected documents"""
        docs = self.get_selected_documents()
        if not docs:
            return
        
        # If any selected doc is not favorite, make all favorites
        # If all are favorites, remove favorite from all
        all_favorites = all(doc.favorite for doc in docs)
        new_favorite_status = not all_favorites
        
        for doc in docs:
            doc.favorite = new_favorite_status
        
        self.save_data()
        self.refresh_documents()
        
        action = "Added to" if new_favorite_status else "Removed from"
        messagebox.showinfo("Success", f"{action} favorites: {len(docs)} document(s)!")
    
    def close_selected_documents(self):
        """Close selected documents in Word"""
        docs = self.get_selected_documents()
        open_docs = [doc for doc in docs if doc.is_open]
        
        if not open_docs:
            messagebox.showinfo("Info", "No open documents selected.")
            return
        
        if messagebox.askyesno("Confirm", f"Actually close {len(open_docs)} Word documents?"):
            closed_count = 0
            
            for doc in open_docs:
                if self.actually_close_word_document(doc):
                    doc.is_open = False
                    closed_count += 1
            
            self.save_data()
            self.refresh_documents()
            
            if closed_count > 0:
                messagebox.showinfo("Success", f"Closed {closed_count} Word documents!")
    
    def edit_selected_document(self):
        """Edit selected document"""
        doc = self.get_selected_document()
        if doc:
            dialog = DocumentDialog(self.root, self.teams, doc)
            if dialog.result:
                doc_data = dialog.result
                doc.name = doc_data['name']
                doc.source_type = doc_data['source_type']
                doc.url = doc_data.get('url')
                doc.file_path = doc_data.get('file_path')
                doc.tags = doc_data.get('tags', [])
                doc.team_id = doc_data.get('team_id')
                
                self.save_data()
                self.refresh_documents()
                messagebox.showinfo("Success", "Document updated successfully!")
    
    def remove_selected_documents(self):
        """Remove selected documents"""
        docs = self.get_selected_documents()
        if not docs:
            return
        
        if len(docs) == 1:
            if messagebox.askyesno("Confirm", f"Remove document '{docs[0].name}'?"):
                del self.docs[docs[0].id]
                self.save_data()
                self.refresh_documents()
        else:
            if messagebox.askyesno("Confirm", f"Remove {len(docs)} selected documents?"):
                for doc in docs:
                    del self.docs[doc.id]
                self.save_data()
                self.refresh_documents()
                messagebox.showinfo("Success", f"Removed {len(docs)} documents!")
    
    def show_team_context_menu(self, event):
        """Show context menu for teams"""
        index = self.teams_listbox.nearest(event.y)
        if index >= 2:  # Skip "All Documents" and "Ungrouped"
            self.teams_listbox.selection_clear(0, tk.END)
            self.teams_listbox.selection_set(index)
            self.team_context_menu.post(event.x_root, event.y_root)
    
    def get_selected_team(self) -> Optional[Team]:
        """Get currently selected team"""
        selection = self.teams_listbox.curselection()
        if not selection or selection[0] < 2:
            return None
        
        team_name = self.teams_listbox.get(selection[0])
        for team in self.teams.values():
            if team.name == team_name:
                return team
        return None
    
    def rename_selected_team(self):
        """Rename selected team"""
        team = self.get_selected_team()
        if team:
            new_name = simpledialog.askstring("Rename Team", "Enter new team name:", initialvalue=team.name)
            if new_name and new_name.strip() and new_name.strip() != team.name:
                team.name = new_name.strip()
                self.save_data()
                self.refresh_teams()
                self.refresh_documents()
                messagebox.showinfo("Success", f"Team renamed to '{new_name}'!")
    
    def delete_selected_team(self):
        """Delete selected team"""
        team = self.get_selected_team()
        if team:
            if messagebox.askyesno("Confirm", f"Delete team '{team.name}'? Documents will be ungrouped."):
                # Remove team from all documents
                for doc in self.docs.values():
                    if doc.team_id == team.id:
                        doc.team_id = None
                
                # Delete team
                del self.teams[team.id]
                
                # Reset selection if this team was selected
                if self.selected_team_id == team.id:
                    self.selected_team_id = None
                
                self.save_data()
                self.refresh_teams()
                self.refresh_documents()
                messagebox.showinfo("Success", f"Team '{team.name}' deleted!")
    
    def close_all_documents(self):
        """Actually close all currently open Word documents"""
        open_docs = [doc for doc in self.docs.values() if doc.is_open]
        if not open_docs:
            messagebox.showinfo("Info", "No documents are currently open.")
            return
        
        if messagebox.askyesno("Confirm", f"Actually close {len(open_docs)} Word documents?"):
            closed_count = 0
            
            for doc in open_docs:
                if self.actually_close_word_document(doc):
                    doc.is_open = False
                    closed_count += 1
            
            self.save_data()
            self.refresh_documents()
            messagebox.showinfo("Success", f"Closed {closed_count} Word documents!")
    
    def open_team_documents(self):
        """Open all documents in the selected team"""
        if not self.selected_team_id or self.selected_team_id == "ungrouped":
            messagebox.showinfo("Info", "Please select a specific team first.")
            return
        
        team_docs = [doc for doc in self.docs.values() if doc.team_id == self.selected_team_id]
        if not team_docs:
            messagebox.showinfo("Info", "No documents found in the selected team.")
            return
        
        if messagebox.askyesno("Confirm", f"Open all {len(team_docs)} documents in this team?"):
            opened_count = 0
            for doc in team_docs:
                try:
                    self.open_in_word(doc)
                    opened_count += 1
                except Exception as e:
                    print(f"Failed to open '{doc.name}': {e}")
            
            if opened_count > 0:
                messagebox.showinfo("Success", f"Opened {opened_count} team documents!")
    
    # def actually_close_word_document(self, doc: DocEntry) -> bool:
    #     """Actually close a Word document using COM automation or process killing"""
    #     try:
    #         if platform.system() == "Windows" and WORD_COM_AVAILABLE:
    #             # Try COM automation first (more precise)
    #             try:
    #                 word_app = win32com.client.Dispatch("Word.Application")
                    
    #                 # Find and close the specific document
    #                 for word_doc in word_app.Documents:
    #                     doc_path = word_doc.FullName.lower()
    #                     if doc.file_path and doc.file_path.lower() in doc_path:
    #                         word_doc.Close(SaveChanges=-1)  # -1 = save changes
    #                         return True
    #                     elif doc.name.lower() in doc_path:
    #                         word_doc.Close(SaveChanges=-1)  # -1 = save changes
    #                         return True
                    
    #                 return False
                    
    #             except Exception as e:
    #                 print(f"COM automation failed: {e}")
    #                 # Fall back to process method
    #                 pass
            
    #         # Fallback: Kill all Word processes (nuclear option)
    #         if platform.system() == "Windows":
    #             result = subprocess.run(["taskkill", "/f", "/im", "winword.exe"], 
    #                                   capture_output=True, text=True)
    #             return result.returncode == 0
    #         elif platform.system() == "Darwin":  # macOS
    #             result = subprocess.run(["pkill", "-f", "Microsoft Word"], 
    #                                   capture_output=True)
    #             return result.returncode == 0
    #         else:  # Linux
    #             result = subprocess.run(["pkill", "-f", "libreoffice"], 
    #                                   capture_output=True)
    #             return result.returncode == 0
                
    #     except Exception as e:
    #         print(f"Failed to close Word document: {e}")
    #         return False

    def actually_close_word_document(self, doc: DocEntry) -> bool:
        """Actually close a Word document using COM automation"""
        try:
            if platform.system() == "Windows" and WORD_COM_AVAILABLE:
                try:
                    word_app = win32com.client.Dispatch("Word.Application")
                    
                    # Find and close the specific document
                    for word_doc in word_app.Documents:
                        doc_path = word_doc.FullName.lower()
                        if doc.file_path and doc.file_path.lower() in doc_path:
                            word_doc.Close(SaveChanges=-1)  # -1 = save changes
                            # Check if any documents are still open
                            if word_app.Documents.Count == 0:
                                word_app.Quit()
                            return True
                        elif doc.name.lower() in doc_path:
                            word_doc.Close(SaveChanges=-1)  # -1 = save changes
                            # Check if any documents are still open
                            if word_app.Documents.Count == 0:
                                word_app.Quit()
                            return True
                    
                    return False
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Could not close document '{doc.name}': {e}")
                    return False
            
            # For non-Windows or when COM is not available, show warning
            messagebox.showwarning("Warning", 
                f"Cannot automatically close '{doc.name}'. Please close it manually in Word.")
            return False
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to close Word document: {e}")
            return False

    
    def run(self):
        """Start the application"""
        self.root.mainloop()

class DocumentDialog:
    def __init__(self, parent, teams: Dict[str, Team], doc: DocEntry = None):
        self.result = None
        self.teams = teams
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Add Document" if doc is None else "Edit Document")
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variables
        self.name_var = tk.StringVar(value=doc.name if doc else "")
        self.source_type_var = tk.StringVar(value=doc.source_type if doc else "file")
        self.url_var = tk.StringVar(value=doc.url if doc and doc.url else "")
        self.file_path_var = tk.StringVar(value=doc.file_path if doc and doc.file_path else "")
        self.tags_var = tk.StringVar(value=", ".join(doc.tags) if doc and doc.tags else "")
        self.team_var = tk.StringVar()
        
        if doc and doc.team_id and doc.team_id in teams:
            self.team_var.set(teams[doc.team_id].name)
        
        self.setup_dialog()
        
        # Center dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
        
        self.dialog.wait_window()
    
    def setup_dialog(self):
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Name
        ttk.Label(main_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.name_var, width=50).grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=5)
        
        # Source type
        ttk.Label(main_frame, text="Source:").grid(row=1, column=0, sticky=tk.W, pady=5)
        source_frame = ttk.Frame(main_frame)
        source_frame.grid(row=1, column=1, columnspan=2, sticky=tk.EW, pady=5)
        
        ttk.Radiobutton(source_frame, text="File", variable=self.source_type_var, 
                       value="file", command=self.on_source_change).pack(side=tk.LEFT)
        ttk.Radiobutton(source_frame, text="URL", variable=self.source_type_var, 
                       value="url", command=self.on_source_change).pack(side=tk.LEFT, padx=(10, 0))
        
        # File path
        ttk.Label(main_frame, text="File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.file_entry = ttk.Entry(main_frame, textvariable=self.file_path_var, width=40)
        self.file_entry.grid(row=2, column=1, sticky=tk.EW, pady=5)
        self.browse_button = ttk.Button(main_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=2, column=2, padx=(5, 0), pady=5)
        
        # URL
        ttk.Label(main_frame, text="URL:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.url_entry = ttk.Entry(main_frame, textvariable=self.url_var, width=50)
        self.url_entry.grid(row=3, column=1, columnspan=2, sticky=tk.EW, pady=5)
        
        # Team
        ttk.Label(main_frame, text="Team:").grid(row=4, column=0, sticky=tk.W, pady=5)
        team_combo = ttk.Combobox(main_frame, textvariable=self.team_var, width=47)
        team_combo['values'] = [""] + [team.name for team in self.teams.values()]
        team_combo.grid(row=4, column=1, columnspan=2, sticky=tk.EW, pady=5)
        
        # Tags
        ttk.Label(main_frame, text="Tags:").grid(row=5, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.tags_var, width=50).grid(row=5, column=1, columnspan=2, sticky=tk.EW, pady=5)
        ttk.Label(main_frame, text="(comma-separated)", font=("Arial", 8)).grid(row=6, column=1, sticky=tk.W)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        
        # Initial state
        self.on_source_change()
    
    def on_source_change(self):
        """Handle source type change"""
        if self.source_type_var.get() == "file":
            self.file_entry.config(state=tk.NORMAL)
            self.browse_button.config(state=tk.NORMAL)
            self.url_entry.config(state=tk.DISABLED)
        else:
            self.file_entry.config(state=tk.DISABLED)
            self.browse_button.config(state=tk.DISABLED)
            self.url_entry.config(state=tk.NORMAL)
    
    def browse_file(self):
        """Browse for file"""
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[
                ("Word Documents", "*.docx *.doc"),
                ("All Files", "*.*")
            ]
        )
        if file_path:
            self.file_path_var.set(file_path)
            if not self.name_var.get():
                self.name_var.set(os.path.basename(file_path))
    
    def save(self):
        """Save document"""
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Error", "Name is required!")
            return
        
        source_type = self.source_type_var.get()
        
        if source_type == "file":
            file_path = self.file_path_var.get().strip()
            if not file_path:
                messagebox.showerror("Error", "File path is required!")
                return
            if not os.path.exists(file_path):
                messagebox.showerror("Error", "File does not exist!")
                return
        else:
            url = self.url_var.get().strip()
            if not url:
                messagebox.showerror("Error", "URL is required!")
                return
        
        # Get team ID
        team_id = None
        team_name = self.team_var.get().strip()
        if team_name:
            for tid, team in self.teams.items():
                if team.name == team_name:
                    team_id = tid
                    break
        
        # Parse tags
        tags = [tag.strip() for tag in self.tags_var.get().split(",") if tag.strip()]
        
        self.result = {
            'name': name,
            'source_type': source_type,
            'url': url if source_type == "url" else None,
            'file_path': self.file_path_var.get().strip() if source_type == "file" else None,
            'team_id': team_id,
            'tags': tags
        }
        
        self.dialog.destroy()
    
    def cancel(self):
        """Cancel dialog"""
        self.dialog.destroy()

if __name__ == "__main__":
    app = DocSmartApp()
    app.run()