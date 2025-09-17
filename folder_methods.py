    def add_folder(self):
        """Add new folder"""
        name = simpledialog.askstring("Add Folder", "Enter folder name:")
        if name and name.strip():
            folder_id = self.generate_id("folder")
            folder = Folder(id=folder_id, name=name.strip())
            self.folders[folder_id] = folder
            self.save_data()
            self.refresh_folder_tree()
            messagebox.showinfo("Success", f"Folder '{name}' added successfully!")
    
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
        # Clear existing items
        for item in self.folder_tree.get_children():
            self.folder_tree.delete(item)
        
        # Add root items
        all_docs_id = self.folder_tree.insert('', 'end', text='All Documents', tags=('all',))
        ungrouped_id = self.folder_tree.insert('', 'end', text='Ungrouped', tags=('ungrouped',))
        
        # Add folders and their contents
        self._add_folder_items('')
    
    def _add_folder_items(self, parent_folder_id: str, parent_tree_id: str = ''):
        """Recursively add folder items to tree"""
        # Add folders
        folders = [f for f in self.folders.values() if f.parent_id == parent_folder_id]
        folders.sort(key=lambda f: f.name)
        
        for folder in folders:
            folder_tree_id = self.folder_tree.insert(parent_tree_id, 'end', 
                                                    text=f"📁 {folder.name}", 
                                                    tags=('folder', folder.id))
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
            for team in teams:
                self.folder_tree.insert(parent_tree_id, 'end', 
                                      text=f"👥 {team.name}", 
                                      tags=('team', team.id))

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
                # Remove team from all documents
                for doc in self.docs.values():
                    if doc.team_id == item.id:
                        doc.team_id = None
                
                del self.teams[item.id]
                
                if self.selected_team_id == item.id:
                    self.selected_team_id = None
        
        elif item_type == 'folder':
            if messagebox.askyesno("Confirm", f"Delete folder '{item.name}'? All contents will be moved to parent."):
                # Move subfolders to parent
                for folder in self.folders.values():
                    if folder.parent_id == item.id:
                        folder.parent_id = item.parent_id
                
                # Move teams to parent folder
                for team in self.teams.values():
                    if team.folder_id == item.id:
                        team.folder_id = item.parent_id
                
                # Move documents to parent folder
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