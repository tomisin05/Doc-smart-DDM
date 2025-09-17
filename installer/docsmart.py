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

class Folder:
    def __init__(self, id: str, name: str, parent_id: str = None, created_at: float = None):
        self.id = id
        self.name = name
        self.parent_id = parent_id
        self.created_at = created_at or datetime.now().timestamp()

    def to_dict(self) -> Dict[str, Any]:
        return {
            'id': self.id,
            'name': self.name,
            'parent_id': self.parent_id,
            'created_at': self.created_at
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Folder':
        return cls(**data)

class DocEntry:
    def __init__(self, id: str, name: str, source_type: str, url: str = None, 
                 file_path: str = None, tags: List[str] = None, team_id: str = None,
                 folder_id: str = None, favorite: bool = False, is_open: bool = False, 
                 last_opened_at: float = None, created_at: float = None):
        self.id = id
        self.name = name
        self.source_type = source_type  # "url" or "file"
        self.url = url
        self.file_path = file_path
        self.tags = tags or []
        self.team_id = team_id
        self.folder_id = folder_id
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
            'folder_id': self.folder_id,
            'favorite': self.favorite,
            'is_open': self.is_open,
            'last_opened_at': self.last_opened_at,
            'created_at': self.created_at
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DocEntry':
        return cls(**data)

class Team:
    def __init__(self, id: str, name: str, folder_id: str = None, created_at: float = None):
        self.id = id
        self.name = name
        self.folder_id = folder_id
        self.created_at = created_at or datetime.now().timestamp()

    def to_dict(self) -> Dict[str, Any]:
        return {
            'id': self.id,
            'name': self.name,
            'folder_id': self.folder_id,
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
        self.folders: Dict[str, Folder] = {}
        self.selected_team_id: Optional[str] = None
        self.selected_folder_id: Optional[str] = None
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
        ttk.Button(button_frame, text="Add Folder", command=self.add_folder).pack(side=tk.LEFT, padx=2)
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
        
        # Teams/Folders section
        teams_label = ttk.Label(sidebar_frame, text="Teams & Folders (📁=Folder, 👥=Team)", font=("Arial", 12, "bold"))
        teams_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Folder buttons
        folder_btn_frame = ttk.Frame(sidebar_frame)
        folder_btn_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(folder_btn_frame, text="Add Folder", command=self.add_folder, width=12).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(folder_btn_frame, text="Add Subfolder", command=self.add_subfolder, width=12).pack(side=tk.LEFT)
        
        # Tree view for folders and teams
        self.tree_frame = ttk.Frame(sidebar_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.folder_tree = ttk.Treeview(self.tree_frame, height=15)
        self.folder_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add a test folder on startup for debugging
        print("DEBUG: Setting up folder tree widget")
        
        tree_scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.folder_tree.yview)
        self.folder_tree.configure(yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.folder_tree.bind('<<TreeviewSelect>>', self.on_folder_tree_select)
        self.folder_tree.bind('<Button-3>', self.show_folder_context_menu)
        
        # Drag and drop bindings
        self.folder_tree.bind('<Button-1>', self.on_drag_start)
        self.folder_tree.bind('<B1-Motion>', self.on_drag_motion)
        self.folder_tree.bind('<ButtonRelease-1>', self.on_drag_end)
        
        # Drag state variables
        self.drag_item = None
        self.drag_data = None
        
        # Context menu
        self.folder_context_menu = tk.Menu(self.root, tearoff=0)
        self.folder_context_menu.add_command(label="Add Subfolder", command=self.add_subfolder)
        self.folder_context_menu.add_command(label="Add Team", command=self.add_team)
        self.folder_context_menu.add_separator()
        self.folder_context_menu.add_command(label="Rename", command=self.rename_selected_item)
        self.folder_context_menu.add_command(label="Delete", command=self.delete_selected_item)
        
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
        self.refresh_folder_tree()
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
            'folders': {id: folder.to_dict() for id, folder in self.folders.items()},
            'selected_team_id': self.selected_team_id,
            'selected_folder_id': self.selected_folder_id
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
            self.folders = {id: Folder.from_dict(folder_data) 
                           for id, folder_data in data.get('folders', {}).items()}
            self.selected_team_id = data.get('selected_team_id')
            self.selected_folder_id = data.get('selected_folder_id')
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")
    
    def add_folder(self):
        """Add new folder"""
        name = simpledialog.askstring("Add Folder", "Enter folder name:")
        if name and name.strip():
            folder_id = self.generate_id("folder")
            folder = Folder(id=folder_id, name=name.strip())
            self.folders[folder_id] = folder
            print(f"DEBUG: Added folder '{name}' with ID {folder_id}")
            print(f"DEBUG: Total folders now: {len(self.folders)}")
            self.save_data()
            self.refresh_folder_tree()
            messagebox.showinfo("Success", f"Folder '{name}' added successfully!\nCheck the tree view below 'All Documents' and 'Ungrouped'")
    
    def add_subfolder(self):
        """Add new subfolder"""
        if not self.selected_folder_id:
            messagebox.showwarning("Warning", "Please select a folder first")
            return
        
        name = simpledialog.askstring("Add Subfolder", "Enter subfolder name:")
        if name and name.strip():
            folder_id = self.generate_id("folder")
            folder = Folder(id=folder_id, name=name.strip(), parent_id=self.selected_folder_id)
            self.folders[folder_id] = folder
            self.save_data()
            self.refresh_folder_tree()
            messagebox.showinfo("Success", f"Subfolder '{name}' added successfully!")

    def refresh_folder_tree(self):
        """Refresh folder tree"""
        print(f"DEBUG: Refreshing folder tree with {len(self.folders)} folders")
        # Clear existing items
        for item in self.folder_tree.get_children():
            self.folder_tree.delete(item)
        
        # Add root items
        self.folder_tree.insert('', 'end', text='All Documents', tags=('all',))
        self.folder_tree.insert('', 'end', text='Ungrouped', tags=('ungrouped',))
        
        # Add folders and their contents
        self._add_folder_items('')
        
        # Debug: Print what's in the tree
        print("DEBUG: Tree contents:")
        for item in self.folder_tree.get_children():
            text = self.folder_tree.item(item, 'text')
            print(f"  - {text}")
    
    def _add_folder_items(self, parent_folder_id: str, parent_tree_id: str = ''):
        """Recursively add folder items to tree"""
        # Add folders - handle both None and empty string for root folders
        if parent_folder_id == '':
            folders = [f for f in self.folders.values() if f.parent_id is None or f.parent_id == '']
        else:
            folders = [f for f in self.folders.values() if f.parent_id == parent_folder_id]
        folders.sort(key=lambda f: f.name)
        
        print(f"DEBUG: Adding {len(folders)} folders with parent_id='{parent_folder_id}'")
        
        for folder in folders:
            folder_tree_id = self.folder_tree.insert(parent_tree_id, 'end', 
                                                    text=f"📁 {folder.name}", 
                                                    tags=('folder', folder.id))
            print(f"DEBUG: Added folder '{folder.name}' to tree")
            # Recursively add subfolders
            self._add_folder_items(folder.id, folder_tree_id)
            
            # Add teams in this folder
            teams = [t for t in self.teams.values() if t.folder_id == folder.id]
            teams.sort(key=lambda t: t.name)
            for team in teams:
                self.folder_tree.insert(folder_tree_id, 'end', 
                                      text=f"👥 {team.name}", 
                                      tags=('team', team.id))
        
        # Add teams without folder (only at root level)
        if not parent_folder_id:
            teams = [t for t in self.teams.values() if not t.folder_id]
            teams.sort(key=lambda t: t.name)
            print(f"DEBUG: Adding {len(teams)} ungrouped teams")
            for team in teams:
                self.folder_tree.insert(parent_tree_id, 'end', 
                                      text=f"👥 {team.name}", 
                                      tags=('team', team.id))
                print(f"DEBUG: Added team '{team.name}' to tree")

    def on_folder_tree_select(self, event):
        """Handle folder tree selection"""
        selection = self.folder_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        tags = self.folder_tree.item(item, 'tags')
        
        if 'all' in tags:
            self.selected_team_id = None
            self.selected_folder_id = None
        elif 'ungrouped' in tags:
            self.selected_team_id = "ungrouped"
            self.selected_folder_id = None
        elif 'team' in tags:
            self.selected_team_id = tags[1]
            self.selected_folder_id = None
        elif 'folder' in tags:
            self.selected_team_id = None
            self.selected_folder_id = tags[1]
        
        self.refresh_documents()

    def show_folder_context_menu(self, event):
        """Show context menu for folders/teams"""
        item = self.folder_tree.identify_row(event.y)
        if item:
            tags = self.folder_tree.item(item, 'tags')
            if tags and tags[0] not in ('all', 'ungrouped'):
                self.folder_tree.selection_set(item)
                self.folder_context_menu.post(event.x_root, event.y_root)

    def get_selected_item(self):
        """Get currently selected item from tree"""
        selection = self.folder_tree.selection()
        if not selection:
            return None, None
        
        item = selection[0]
        tags = self.folder_tree.item(item, 'tags')
        
        if 'team' in tags:
            return 'team', self.teams.get(tags[1])
        elif 'folder' in tags:
            return 'folder', self.folders.get(tags[1])
        return None, None

    def rename_selected_item(self):
        """Rename selected item"""
        item_type, item = self.get_selected_item()
        if not item:
            return
        
        item_name = "Team" if item_type == 'team' else "Folder"
        new_name = simpledialog.askstring(f"Rename {item_name}", f"Enter new {item_name.lower()} name:", initialvalue=item.name)
        
        if new_name and new_name.strip() and new_name.strip() != item.name:
            item.name = new_name.strip()
            self.save_data()
            self.refresh_folder_tree()
            self.refresh_documents()
            messagebox.showinfo("Success", f"{item_name} renamed to '{new_name}'!")

    def delete_selected_item(self):
        """Delete selected item"""
        item_type, item = self.get_selected_item()
        if not item:
            return
        
        item_name = "Team" if item_type == 'team' else "Folder"
        
        if item_type == 'team':
            if messagebox.askyesno("Confirm", f"Delete team '{item.name}'? Documents will be ungrouped."):
                for doc in self.docs.values():
                    if doc.team_id == item.id:
                        doc.team_id = None
                del self.teams[item.id]
                if self.selected_team_id == item.id:
                    self.selected_team_id = None
        
        elif item_type == 'folder':
            if messagebox.askyesno("Confirm", f"Delete folder '{item.name}'? All contents will be moved to parent."):
                for folder in self.folders.values():
                    if folder.parent_id == item.id:
                        folder.parent_id = item.parent_id
                for team in self.teams.values():
                    if team.folder_id == item.id:
                        team.folder_id = item.parent_id
                for doc in self.docs.values():
                    if doc.folder_id == item.id:
                        doc.folder_id = item.parent_id
                del self.folders[item.id]
                if self.selected_folder_id == item.id:
                    self.selected_folder_id = None
        
        self.save_data()
        self.refresh_folder_tree()
        self.refresh_documents()
        messagebox.showinfo("Success", f"{item_name} '{item.name}' deleted!")
    
    def on_drag_start(self, event):
        """Handle drag start"""
        item = self.folder_tree.identify_row(event.y)
        if item:
            tags = self.folder_tree.item(item, 'tags')
            # Only allow dragging teams and folders (not All Documents/Ungrouped)
            if tags and tags[0] in ('team', 'folder'):
                self.drag_item = item
                self.drag_data = {
                    'type': tags[0],
                    'id': tags[1],
                    'text': self.folder_tree.item(item, 'text')
                }
    
    def on_drag_motion(self, event):
        """Handle drag motion - visual feedback"""
        if self.drag_item:
            # Get item under cursor
            target_item = self.folder_tree.identify_row(event.y)
            if target_item and target_item != self.drag_item:
                # Highlight potential drop target
                self.folder_tree.selection_set(target_item)
    
    def on_drag_end(self, event):
        """Handle drag end - perform the move"""
        if not self.drag_item:
            return
        
        target_item = self.folder_tree.identify_row(event.y)
        if target_item and target_item != self.drag_item:
            target_tags = self.folder_tree.item(target_item, 'tags')
            
            # Determine valid drop targets
            if self.drag_data['type'] == 'team':
                # Teams can be dropped on folders or root (for ungrouping)
                if 'folder' in target_tags:
                    self.move_team_to_folder(self.drag_data['id'], target_tags[1])
                elif target_tags[0] in ('all', 'ungrouped'):
                    self.move_team_to_folder(self.drag_data['id'], None)
            
            elif self.drag_data['type'] == 'folder':
                # Folders can be dropped on other folders (to become subfolders) or root
                if 'folder' in target_tags:
                    # Check for circular reference
                    if not self.would_create_circular_reference(self.drag_data['id'], target_tags[1]):
                        self.move_folder_to_parent(self.drag_data['id'], target_tags[1])
                    else:
                        messagebox.showwarning("Invalid Move", "Cannot move folder into its own subfolder")
                elif target_tags[0] in ('all', 'ungrouped'):
                    self.move_folder_to_parent(self.drag_data['id'], None)
        
        # Reset drag state
        self.drag_item = None
        self.drag_data = None
    
    def move_team_to_folder(self, team_id: str, folder_id: str):
        """Move team to specified folder"""
        team = self.teams.get(team_id)
        if team:
            old_folder = team.folder_id
            team.folder_id = folder_id
            self.save_data()
            self.refresh_folder_tree()
            
            folder_name = self.folders[folder_id].name if folder_id else "root"
            messagebox.showinfo("Success", f"Moved team '{team.name}' to {folder_name}")
    
    def move_folder_to_parent(self, folder_id: str, parent_id: str):
        """Move folder to new parent"""
        folder = self.folders.get(folder_id)
        if folder:
            old_parent = folder.parent_id
            folder.parent_id = parent_id
            self.save_data()
            self.refresh_folder_tree()
            
            parent_name = self.folders[parent_id].name if parent_id else "root"
            messagebox.showinfo("Success", f"Moved folder '{folder.name}' to {parent_name}")
    
    def would_create_circular_reference(self, folder_id: str, target_parent_id: str) -> bool:
        """Check if moving folder would create circular reference"""
        current_id = target_parent_id
        while current_id:
            if current_id == folder_id:
                return True
            folder = self.folders.get(current_id)
            current_id = folder.parent_id if folder else None
        return False
    
    def add_team(self):
        """Add new team"""
        name = simpledialog.askstring("Add Team", "Enter team name:")
        if name and name.strip():
            team_id = self.generate_id("team")
            team = Team(id=team_id, name=name.strip(), folder_id=self.selected_folder_id)
            self.teams[team_id] = team
            self.save_data()
            self.refresh_folder_tree()
            messagebox.showinfo("Success", f"Team '{name}' added successfully!")
    
    def refresh_documents(self):
        """Refresh documents treeview"""
        # Clear existing items
        for item in self.docs_tree.get_children():
            self.docs_tree.delete(item)
        
        # Filter documents
        filtered_docs = []
        search_term = self.search_text.get().lower()
        
        for doc in self.docs.values():
            # Team/Folder filter
            if self.selected_team_id == "ungrouped" and (doc.team_id or doc.folder_id):
                continue
            elif self.selected_team_id and self.selected_team_id != "ungrouped" and doc.team_id != self.selected_team_id:
                continue
            elif self.selected_folder_id and not self._is_doc_in_folder_hierarchy(doc):
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
    
    def _is_doc_in_folder_hierarchy(self, doc: DocEntry) -> bool:
        """Check if document is in selected folder hierarchy"""
        if doc.team_id:
            team = self.teams.get(doc.team_id)
            if team and team.folder_id:
                return self._is_folder_in_hierarchy(team.folder_id)
        elif doc.folder_id:
            return self._is_folder_in_hierarchy(doc.folder_id)
        return False
    
    def _is_folder_in_hierarchy(self, folder_id: str) -> bool:
        """Check if folder is in selected folder hierarchy"""
        current_folder_id = folder_id
        while current_folder_id:
            if current_folder_id == self.selected_folder_id:
                return True
            folder = self.folders.get(current_folder_id)
            current_folder_id = folder.parent_id if folder else None
        return False
    
    def on_search_change(self, event):
        """Handle search text change"""
        self.refresh_documents()
    
    def show_context_menu(self, event):
        """Show context menu for documents"""
        item = self.docs_tree.identify_row(event.y)
        if item:
            self.docs_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
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
        
        for doc in docs:
            self.open_in_word(doc)
    
    def open_in_word(self, doc: DocEntry):
        """Open document in Microsoft Word"""
        try:
            if doc.source_type == "url":
                word_url = f"ms-word:ofe|u|{doc.url}"
                try:
                    if platform.system() == "Windows":
                        os.startfile(word_url)
                    else:
                        webbrowser.open(doc.url)
                except:
                    webbrowser.open(doc.url)
            else:
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
    
    def toggle_favorite_selected(self):
        """Toggle favorite status of selected documents"""
        docs = self.get_selected_documents()
        if not docs:
            return
        
        all_favorites = all(doc.favorite for doc in docs)
        new_favorite_status = not all_favorites
        
        for doc in docs:
            doc.favorite = new_favorite_status
        
        self.save_data()
        self.refresh_documents()
    
    def close_selected_documents(self):
        """Close selected documents in Word"""
        docs = self.get_selected_documents()
        open_docs = [doc for doc in docs if doc.is_open]
        
        if not open_docs:
            messagebox.showinfo("Info", "No open documents selected.")
            return
        
        if messagebox.askyesno("Confirm", f"Close {len(open_docs)} Word documents?"):
            for doc in open_docs:
                if self.actually_close_word_document(doc):
                    doc.is_open = False
            
            self.save_data()
            self.refresh_documents()
    
    def actually_close_word_document(self, doc: DocEntry) -> bool:
        """Actually close a Word document using COM automation"""
        try:
            if platform.system() == "Windows" and WORD_COM_AVAILABLE:
                try:
                    word_app = win32com.client.Dispatch("Word.Application")
                    
                    for word_doc in word_app.Documents:
                        doc_path = word_doc.FullName.lower()
                        if doc.file_path and doc.file_path.lower() in doc_path:
                            word_doc.Close(SaveChanges=-1)
                            if word_app.Documents.Count == 0:
                                word_app.Quit()
                            return True
                        elif doc.name.lower() in doc_path:
                            word_doc.Close(SaveChanges=-1)
                            if word_app.Documents.Count == 0:
                                word_app.Quit()
                            return True
                    
                    return False
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Could not close document '{doc.name}': {e}")
                    return False
            
            messagebox.showwarning("Warning", 
                f"Cannot automatically close '{doc.name}'. Please close it manually in Word.")
            return False
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to close Word document: {e}")
            return False
    
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
                team_id=doc_data.get('team_id'),
                folder_id=self.selected_folder_id
            )
            
            self.docs[doc_id] = doc
            self.save_data()
            self.refresh_documents()
            messagebox.showinfo("Success", f"Document '{doc.name}' added successfully!")
    
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
                        file_path=file_path,
                        folder_id=self.selected_folder_id
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
                    'teams': {id: team.to_dict() for id, team in self.teams.items()},
                    'folders': {id: folder.to_dict() for id, folder in self.folders.items()}
                }
                
                with open(file_path, 'w') as f:
                    json.dump(data, f, indent=2)
                
                messagebox.showinfo("Success", "Data exported successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {e}")
    
    def edit_selected_document(self):
        """Edit selected document"""
        docs = self.get_selected_documents()
        if docs:
            doc = docs[0]
            dialog = DocumentDialog(self.root, self.teams, doc)
            if dialog.result:
                doc_data = dialog.result
                doc.name = doc_data['name']
                doc.source_type = doc_data['source_type']
                doc.url = doc_data.get('url')
                doc.file_path = doc_data.get('file_path')
                doc.tags = doc_data.get('tags', [])
                doc.team_id = doc_data.get('team_id')
                doc.folder_id = doc_data.get('folder_id')
                
                self.save_data()
                self.refresh_documents()
                messagebox.showinfo("Success", "Document updated successfully!")
    
    def remove_selected_documents(self):
        """Remove selected documents"""
        docs = self.get_selected_documents()
        if not docs:
            return
        
        if messagebox.askyesno("Confirm", f"Remove {len(docs)} selected documents?"):
            for doc in docs:
                del self.docs[doc.id]
            self.save_data()
            self.refresh_documents()
            messagebox.showinfo("Success", f"Removed {len(docs)} documents!")
    
    def close_all_documents(self):
        """Close all currently open Word documents"""
        open_docs = [doc for doc in self.docs.values() if doc.is_open]
        if not open_docs:
            messagebox.showinfo("Info", "No documents are currently open.")
            return
        
        if messagebox.askyesno("Confirm", f"Close {len(open_docs)} Word documents?"):
            for doc in open_docs:
                if self.actually_close_word_document(doc):
                    doc.is_open = False
            
            self.save_data()
            self.refresh_documents()
    
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
            for doc in team_docs:
                self.open_in_word(doc)
    
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
            'folder_id': None,
            'tags': tags
        }
        
        self.dialog.destroy()
    
    def cancel(self):
        """Cancel dialog"""
        self.dialog.destroy()

if __name__ == "__main__":
    app = DocSmartApp()
    app.run()