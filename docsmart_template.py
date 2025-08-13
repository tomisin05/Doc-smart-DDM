#!/usr/bin/env python3
"""
Doc-smart Template: Desktop application for managing Word documents for debate preparation

This is a learning template with TODOs and helpful comments to guide you through
building the complete application from scratch.

LEARNING OBJECTIVES:
- Build a desktop GUI application with tkinter
- Manage data persistence with JSON
- Integrate with Microsoft Word
- Handle file operations and cross-platform compatibility
- Implement search, filtering, and organization features
"""

# TODO 1: Import all necessary libraries
# HINT: You'll need tkinter for GUI, json for data storage, os/subprocess for file operations
# HINT: Also import platform, datetime, pathlib, webbrowser, typing
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
# TODO: Add remaining imports here

# TODO 2: Try importing Windows COM for Word automation (optional advanced feature)
# HINT: Use try/except block to handle ImportError gracefully
try:
    # TODO: Import win32com.client for Word automation
    pass
    WORD_COM_AVAILABLE = True
except ImportError:
    WORD_COM_AVAILABLE = False

class DocEntry:
    """
    Data class representing a single document entry
    
    LEARNING NOTE: This class stores all information about a document:
    - Basic info (id, name)
    - Source type (URL or local file)
    - Organization (tags, team)
    - Status tracking (favorite, open/closed, timestamps)
    """
    
    def __init__(self, id: str, name: str, source_type: str, url: str = None, 
                 file_path: str = None, tags: list = None, team_id: str = None,
                 favorite: bool = False, is_open: bool = False, 
                 last_opened_at: float = None, created_at: float = None):
        # TODO 3: Initialize all instance variables
        # HINT: Set default values for optional parameters
        # HINT: Use datetime.now().timestamp() for created_at if None
        pass
    
    def to_dict(self):
        """Convert object to dictionary for JSON serialization"""
        # TODO 4: Return dictionary with all instance variables
        # HINT: Use self.__dict__ or manually create dict
        pass
    
    @classmethod
    def from_dict(cls, data: dict):
        """Create DocEntry object from dictionary"""
        # TODO 5: Create and return new DocEntry instance from dict data
        # HINT: Use **data to unpack dictionary as keyword arguments
        pass

class Team:
    """
    Data class representing a team/group for organizing documents
    
    LEARNING NOTE: Teams help organize documents by opponent or topic
    """
    
    def __init__(self, id: str, name: str, created_at: float = None):
        # TODO 6: Initialize team properties
        pass
    
    def to_dict(self):
        """Convert to dictionary for JSON serialization"""
        # TODO 7: Return dictionary representation
        pass
    
    @classmethod
    def from_dict(cls, data: dict):
        """Create Team object from dictionary"""
        # TODO 8: Create Team instance from dict data
        pass

class DocSmartApp:
    """
    Main application class - handles GUI and all functionality
    
    LEARNING NOTE: This is the core of the application. It manages:
    - GUI creation and layout
    - Data storage and retrieval
    - User interactions
    - Document operations
    """
    
    def __init__(self):
        # TODO 9: Initialize the main window
        # HINT: Create tk.Tk(), set title and geometry
        self.root = None  # Replace with tk.Tk()
        
        # TODO 10: Initialize data storage
        # HINT: Create empty dictionaries for docs and teams
        self.docs = {}  # Dict[str, DocEntry]
        self.teams = {}  # Dict[str, Team]
        self.selected_team_id = None
        
        # TODO 11: Create tkinter variables for UI state
        # HINT: Use tk.StringVar() and tk.BooleanVar()
        self.search_text = None  # tk.StringVar()
        self.favorite_only = None  # tk.BooleanVar()
        
        # TODO 12: Set up data file path
        # HINT: Use Path.home() / ".docsmart" / "data.json"
        self.data_file = None
        
        # TODO 13: Load existing data and setup UI
        # self.load_data()
        # self.setup_ui()
    
    def setup_ui(self):
        """
        Create the main user interface
        
        LEARNING NOTE: This method builds the entire GUI layout:
        - Header with title and buttons
        - Sidebar with teams and filters
        - Main area with document list
        - Context menus and event bindings
        """
        
        # TODO 14: Create main frame container
        # HINT: Use ttk.Frame(self.root) and pack with fill and expand
        
        # TODO 15: Create header section
        # HINT: Include app title, description, and action buttons
        # BUTTONS NEEDED: Add Document, Add Team, Import Folder, Export Data
        # ADVANCED: Open Selected, Close Selected, Close All, Open Team
        
        # TODO 16: Create content area with sidebar and main section
        # HINT: Use two frames side by side
        
        # TODO 17: Build sidebar with teams list and filters
        # COMPONENTS NEEDED:
        # - Teams listbox (with "All Documents" and "Ungrouped" options)
        # - Search entry field
        # - Favorites only checkbox
        # - Tag filters (advanced)
        
        # TODO 18: Build main document area
        # COMPONENTS NEEDED:
        # - Treeview with columns: Name, Team, Tags, Status, Last Opened
        # - Scrollbar for the treeview
        # - Enable multi-select (selectmode='extended')
        
        # TODO 19: Create context menus
        # MENU ITEMS NEEDED:
        # - Open in Word
        # - Close in Word (advanced)
        # - Mark as Favorite
        # - Edit
        # - Remove
        
        # TODO 20: Bind events
        # EVENTS NEEDED:
        # - Listbox selection for teams
        # - Right-click for context menu
        # - Double-click to open documents
        # - Search text changes
        
        # TODO 21: Initial data load
        # self.refresh_teams()
        # self.refresh_documents()
        
        pass
    
    def generate_id(self, prefix: str = "id") -> str:
        """Generate unique ID for documents and teams"""
        # TODO 22: Generate random ID string
        # HINT: Use random.choices with string.ascii_lowercase + string.digits
        # HINT: Format as f"{prefix}_{random_string}"
        pass
    
    def save_data(self):
        """Save all data to JSON file"""
        # TODO 23: Create data structure and save to file
        # STEPS:
        # 1. Create data dict with docs, teams, selected_team_id
        # 2. Convert objects to dicts using to_dict() methods
        # 3. Create parent directory if needed
        # 4. Write JSON to file with indent=2
        pass
    
    def load_data(self):
        """Load data from JSON file"""
        # TODO 24: Load and parse JSON data
        # STEPS:
        # 1. Check if file exists
        # 2. Read and parse JSON
        # 3. Convert dicts back to objects using from_dict() methods
        # 4. Handle exceptions gracefully
        pass
    
    def open_in_word(self, doc):
        """Open document in Microsoft Word"""
        # TODO 25: Implement document opening logic
        # STEPS:
        # 1. Check document source type (URL vs file)
        # 2. For URLs: try ms-word protocol, fallback to browser
        # 3. For files: use os.startfile (Windows) or subprocess (Mac/Linux)
        # 4. Update document status (is_open=True, last_opened_at)
        # 5. Save data and refresh display
        # 6. Handle errors with messagebox
        pass
    
    def actually_close_word_document(self, doc):
        """Actually close Word document (ADVANCED FEATURE)"""
        # TODO 26: Implement actual document closing
        # APPROACHES:
        # 1. COM automation (Windows) - find and close specific document
        # 2. Process killing (fallback) - kill Word processes
        # 3. Cross-platform support (taskkill, pkill)
        # RETURN: True if successful, False otherwise
        pass
    
    def add_document(self):
        """Show dialog to add new document"""
        # TODO 27: Create and show document dialog
        # STEPS:
        # 1. Create DocumentDialog instance
        # 2. Check if user provided data (dialog.result)
        # 3. Create new DocEntry with generated ID
        # 4. Add to self.docs dictionary
        # 5. Save data and refresh display
        pass
    
    def add_team(self):
        """Add new team via simple dialog"""
        # TODO 28: Get team name and create team
        # STEPS:
        # 1. Use simpledialog.askstring for team name
        # 2. Validate input (not empty)
        # 3. Create Team object with generated ID
        # 4. Add to self.teams dictionary
        # 5. Save and refresh
        pass
    
    def import_folder(self):
        """Import all Word documents from a folder"""
        # TODO 29: Bulk import Word documents
        # STEPS:
        # 1. Use filedialog.askdirectory to select folder
        # 2. Walk through folder recursively (os.walk)
        # 3. Find files with .docx/.doc extensions
        # 4. Create DocEntry for each file
        # 5. Show success message with count
        pass
    
    def export_data(self):
        """Export data to JSON file"""
        # TODO 30: Export data for backup
        # STEPS:
        # 1. Use filedialog.asksaveasfilename
        # 2. Create data structure (similar to save_data)
        # 3. Write to selected file
        # 4. Show success/error message
        pass
    
    def refresh_teams(self):
        """Update teams listbox display"""
        # TODO 31: Populate teams listbox
        # STEPS:
        # 1. Clear existing items
        # 2. Add "All Documents" and "Ungrouped" options
        # 3. Add all teams sorted by name
        pass
    
    def refresh_documents(self):
        """Update documents treeview display"""
        # TODO 32: Filter and display documents
        # STEPS:
        # 1. Clear existing treeview items
        # 2. Apply filters (team, search, favorites)
        # 3. Sort documents (favorites first, then by last opened)
        # 4. Add to treeview with proper formatting
        # 5. Handle star display for favorites
        pass
    
    def get_selected_documents(self):
        """Get list of currently selected documents"""
        # TODO 33: Extract selected documents from treeview
        # STEPS:
        # 1. Get selection from treeview
        # 2. Extract document names from treeview items
        # 3. Find corresponding DocEntry objects
        # 4. Return list of DocEntry objects
        pass
    
    def open_selected_documents(self, event=None):
        """Open all selected documents"""
        # TODO 34: Open multiple documents at once
        # STEPS:
        # 1. Get selected documents
        # 2. Loop through and call open_in_word for each
        # 3. Count successes and show result message
        # 4. Handle errors gracefully
        pass
    
    def close_selected_documents(self):
        """Close selected documents in Word (ADVANCED)"""
        # TODO 35: Close multiple documents
        # Similar to open_selected_documents but call close method
        pass
    
    def toggle_favorite_selected(self):
        """Toggle favorite status for selected documents"""
        # TODO 36: Bulk favorite toggle
        # LOGIC: If any selected doc is not favorite, make all favorites
        #        If all are favorites, remove favorite from all
        pass
    
    def remove_selected_documents(self):
        """Remove selected documents from library"""
        # TODO 37: Bulk document removal
        # STEPS:
        # 1. Get selected documents
        # 2. Show confirmation dialog
        # 3. Remove from self.docs dictionary
        # 4. Save and refresh
        pass
    
    def on_team_select(self, event):
        """Handle team selection in listbox"""
        # TODO 38: Update selected team and refresh documents
        # STEPS:
        # 1. Get selected index
        # 2. Map to team ID (handle "All" and "Ungrouped" special cases)
        # 3. Update self.selected_team_id
        # 4. Refresh documents display
        pass
    
    def on_search_change(self, event):
        """Handle search text changes"""
        # TODO 39: Trigger document refresh when search changes
        pass
    
    def show_context_menu(self, event):
        """Show right-click context menu"""
        # TODO 40: Display context menu at cursor position
        # STEPS:
        # 1. Identify clicked item
        # 2. Select the item
        # 3. Show context menu at event coordinates
        pass
    
    def run(self):
        """Start the application"""
        # TODO 41: Start the tkinter main loop
        # self.root.mainloop()
        pass

class DocumentDialog:
    """
    Dialog window for adding/editing documents
    
    LEARNING NOTE: This creates a popup window with form fields:
    - Document name
    - Source type (URL or File)
    - URL field or file picker
    - Team selection
    - Tags input
    """
    
    def __init__(self, parent, teams, doc=None):
        self.result = None
        self.teams = teams
        
        # TODO 42: Create dialog window
        # STEPS:
        # 1. Create tk.Toplevel window
        # 2. Set title, geometry, and modal properties
        # 3. Create form variables (tk.StringVar for each field)
        # 4. If editing existing doc, populate with current values
        # 5. Call self.setup_dialog()
        # 6. Center the dialog and wait for result
        pass
    
    def setup_dialog(self):
        """Create dialog form layout"""
        # TODO 43: Build form interface
        # FORM FIELDS NEEDED:
        # - Name entry
        # - Source type radio buttons (URL/File)
        # - URL entry (conditional)
        # - File path entry with browse button (conditional)
        # - Team dropdown
        # - Tags entry
        # - Save/Cancel buttons
        pass
    
    def on_source_change(self):
        """Handle source type radio button changes"""
        # TODO 44: Show/hide URL vs File fields based on selection
        pass
    
    def browse_file(self):
        """Open file picker dialog"""
        # TODO 45: Use filedialog.askopenfilename for Word documents
        # HINT: Set filetypes filter for .docx and .doc files
        pass
    
    def save(self):
        """Validate and save document data"""
        # TODO 46: Validate form and create result dictionary
        # VALIDATION:
        # - Name is required
        # - URL is valid if source type is URL
        # - File exists if source type is File
        # RESULT: Dictionary with all form data
        pass
    
    def cancel(self):
        """Cancel dialog"""
        # TODO 47: Close dialog without saving
        pass

# TODO 48: Add main execution block
if __name__ == "__main__":
    # TODO: Create and run DocSmartApp instance
    pass

"""
LEARNING ROADMAP - Complete these TODOs in order:

PHASE 1 - Basic Structure (TODOs 1-13)
- Set up imports and basic classes
- Create main window and data structures

PHASE 2 - Data Management (TODOs 22-24)
- Implement data persistence with JSON
- Create ID generation and file operations

PHASE 3 - Basic GUI (TODOs 14-21)
- Build main interface layout
- Create teams sidebar and document list

PHASE 4 - Core Functionality (TODOs 25, 27-32)
- Document opening and team management
- Display and filtering logic

PHASE 5 - User Interactions (TODOs 33-41)
- Selection handling and bulk operations
- Event handling and context menus

PHASE 6 - Dialog System (TODOs 42-47)
- Document add/edit dialog
- Form validation and file picking

PHASE 7 - Advanced Features (TODOs 26, 35)
- Word document closing with COM automation
- Advanced bulk operations

TESTING TIPS:
- Test each phase before moving to the next
- Start with simple features and add complexity
- Use print statements for debugging
- Test with sample data

HELPFUL RESOURCES:
- tkinter documentation: https://docs.python.org/3/library/tkinter.html
- JSON handling: https://docs.python.org/3/library/json.html
- File operations: https://docs.python.org/3/library/pathlib.html
- Windows COM: https://pypi.org/project/pywin32/
"""