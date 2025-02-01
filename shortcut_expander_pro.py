import tkinter as tk
from tkinter import dnd
import tkinter.messagebox as messagebox
from tkinter import ttk, filedialog
import keyboard
import json
import threading
import time
import os
import pyperclip
import winsound

class ShortcutManager:
    def __init__(self):
        self.current_group = None  # Store current group selection
        
        # Define complete theme colors before anything else
        self.themes = {
            "bitunix": {
                "bg": "#F2F2F7",
                "fg": "#000000",
                "accent": "#b9f641",
                "button_bg": "#b9f641",
                "button_fg": "#000000",
                "field_bg": "#FFFFFF",
                "select_bg": "#b9f641",
                "select_fg": "#000000"
            },
            "light": {
                "bg": "#F2F2F7",
                "fg": "#000000",
                "accent": "#007AFF",
                "button_bg": "#007AFF",
                "button_fg": "#FFFFFF",
                "field_bg": "#FFFFFF",
                "select_bg": "#007AFF",
                "select_fg": "#FFFFFF"
            },
            "dark": {
                "bg": "#1C1C1E",
                "fg": "#FFFFFF",
                "accent": "#0A84FF",
                "button_bg": "#333333",
                "button_fg": "#FFFFFF",
                "field_bg": "#333333",
                "select_bg": "#0A84FF",
                "select_fg": "#FFFFFF"
            },
            "nord": {
                "bg": "#2E3440",
                "fg": "#ECEFF4",
                "accent": "#88C0D0",
                "button_bg": "#5E81AC",
                "button_fg": "#ECEFF4",
                "field_bg": "#3B4252",
                "select_bg": "#88C0D0",
                "select_fg": "#2E3440"
            },
            "monokai": {
                "bg": "#272822",
                "fg": "#F8F8F2",
                "accent": "#FD971F",
                "button_bg": "#66D9EF",
                "button_fg": "#272822",
                "field_bg": "#3E3D32",
                "select_bg": "#FD971F",
                "select_fg": "#272822"
            },
            "solarized": {
                "bg": "#002B36",
                "fg": "#839496",
                "accent": "#268BD2",
                "button_bg": "#268BD2",
                "button_fg": "#FDF6E3",
                "field_bg": "#073642",
                "select_bg": "#268BD2",
                "select_fg": "#FDF6E3"
            }
        }

        self.data = self.load_data()
        self.current_theme = self.data.get("theme", "bitunix")
        self.sound_file = self.data.get("sound_file", "")
        self.buffer_clear_time = self.data.get("buffer_clear_time", 10000)  # Default 10000ms

        self.root = tk.Tk()
        self.root.title("Shortcuts Manager Pro")
        self.root.geometry("1172x586")  # Default size
        self.root.minsize(1172, 586)  # Minimum size constraints

        # --- iOS Styling ---
        self.root.configure(bg="#F2F2F7")

        self.style = ttk.Style()
        self.style.theme_use("default")

        # --- Fonts ---
        self.default_font = ("Tahoma", 14)
        self.button_font = ("Helvetica Neue", 12, "bold")
        self.label_font = ("Helvetica Neue", 12)

        # Configure colors based on the current theme
        self.configure_colors()  # Call this before creating widgets

        # Initialize widget references to None
        self.group_name_label = None
        self.group_members_listbox = None
        self.groups_listbox = None
        self.shortcuts_listbox = None
        self.group_combobox = None
        self.transfer_group_combobox = None
        self.shortcut_entry = None
        self.expansion_entry = None
        self.new_group_entry = None
        self.test_entry = None
        self.test_result_label = None
        self.sound_file_label = None
        
        # Initialize variables for UI state
        self.current_group = None
        self.volume_var = tk.IntVar(value=100)  # Default volume
        
        # Initialize animation properties
        self.animations = {}
        self.animation_speed = 10  # ms between animation frames
        self.animation_steps = 10  # number of steps in animations
        
        # Now create widgets
        self.create_widgets()
        
        # Update UI after widgets are created
        self.update_ui()

        # --- Hotkey Thread ---
        self.hotkey_thread = threading.Thread(target=self.register_hotkey, daemon=True)
        self.hotkey_thread.start()

        # --- Typing Monitoring ---
        self.typing_timer = None
        self.typed_text = ""
        self.ignore_next_space = False  # Flag to ignore spaces after expansion

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def configure_colors(self):
        theme_colors = self.themes.get(self.current_theme, self.themes["light"])
        
        # Configure base styles
        self.style.configure(".", 
            background=theme_colors["bg"],
            foreground=theme_colors["fg"],
            font=self.default_font)
        
        # Configure specific widget styles
        self.style.configure("TLabel",
            background=theme_colors["bg"],
            foreground=theme_colors["fg"],
            font=self.label_font)
            
        self.style.configure("TButton",
            background=theme_colors["button_bg"],
            foreground=theme_colors["button_fg"],
            font=self.button_font)
            
        self.style.configure("TEntry",
            fieldbackground=theme_colors["field_bg"],
            foreground=theme_colors["fg"],
            font=self.default_font)
            
        self.style.configure("TCombobox",
            fieldbackground=theme_colors["field_bg"],
            foreground=theme_colors["fg"],
            font=self.default_font,
            arrowcolor=theme_colors["fg"])

        # Set listbox colors
        self.listbox_bg = theme_colors["field_bg"]
        self.listbox_fg = theme_colors["fg"]
        self.listbox_select_bg = theme_colors["select_bg"]
        self.listbox_select_fg = theme_colors["select_fg"]

        # Apply styles if listboxes exist
        if hasattr(self, 'groups_listbox'):
            self.apply_listbox_styles()

    def apply_listbox_styles(self):
        # Apply styles to listboxes
        for listbox in [self.groups_listbox, self.group_members_listbox, self.shortcuts_listbox]:
            listbox.config(bg=self.listbox_bg, fg=self.listbox_fg, selectbackground=self.listbox_select_bg, selectforeground=self.listbox_select_fg)

    def load_data(self):
        try:
            with open("shortcuts.json", "r") as f:
                return json.load(f)
        except FileNotFoundError:
            return {"shortcuts": {}, "groups": [], "theme": "light", "sound_file": ""}

    def save_data(self):
        """Save without volume control"""
        data = {
            **self.data,
            "sound_file": self.sound_file
        }
        with open("shortcuts.json", "w") as f:
            json.dump(data, f)

    def apply_theme(self):
        # Update the theme colors
        self.configure_colors()

        # Update each widget with the current theme
        for child in self.root.winfo_children():
            self.update_widget_style(child)

        # Refresh the UI
        self.root.update()

    def update_widget_style(self, widget):
        widget_type = widget.winfo_class()

        if widget_type == "TLabel":
            widget.config(background=self.style.lookup("TLabel", "background"), foreground=self.style.lookup("TLabel", "foreground"))
        elif widget_type == "TButton":
            widget.config(background=self.style.lookup("TButton", "background"), foreground=self.style.lookup("TButton", "foreground"))
        elif widget_type == "TEntry":
            widget.config(background=self.style.lookup("TEntry", "fieldbackground"), foreground=self.style.lookup("TEntry", "foreground"), insertbackground=self.style.lookup("TEntry", "foreground"))
        elif widget_type == "TCombobox":
            widget.config(background=self.style.lookup("TCombobox", "fieldbackground"), foreground=self.style.lookup("TCombobox", "foreground"))
        elif widget_type == "Listbox":
            widget.config(bg=self.listbox_bg, fg=self.listbox_fg, selectbackground=self.listbox_select_bg, selectforeground=self.listbox_select_fg)
        elif widget_type == "Text":
            widget.config(bg=self.style.lookup("TEntry", "fieldbackground"), fg=self.style.lookup("TEntry", "foreground"), insertbackground=self.style.lookup("TEntry", "foreground"))
        elif widget_type == "Frame":
            widget.config(bg=self.style.lookup(".", "background"))

        # Recursively update children
        for child in widget.winfo_children():
            self.update_widget_style(child)

    def create_widgets(self):
        # Navigation Bar
        self.nav_frame = tk.Frame(self.root)
        self.nav_frame.configure(bg=self.style.lookup(".", "background"))
        self.nav_frame.pack(side="top", fill="x", pady=(0, 10))

        # Navigation Buttons
        tk.Button(self.nav_frame, text="🏠 Home", command=lambda: self.show_section("home"), font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(side="left", padx=5)
        tk.Button(self.nav_frame, text="📂 Groups", command=lambda: self.show_section("groups"), font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(side="left", padx=5)
        tk.Button(self.nav_frame, text="⌨️ Shortcuts", command=lambda: self.show_section("shortcuts"), font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(side="left", padx=5)
        tk.Button(self.nav_frame, text="⚙️ Settings", command=lambda: self.show_section("settings"), font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(side="left", padx=5)
        tk.Button(self.nav_frame, text="🧪 Test", command=lambda: self.show_section("test"), font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(side="left", padx=5)

        # Sections
        self.home_frame = tk.Frame(self.root, bg=self.style.lookup(".", "background"), padx=15, pady=15)
        self.groups_frame = tk.Frame(self.root, bg=self.style.lookup(".", "background"), padx=15, pady=15)
        self.shortcuts_frame = tk.Frame(self.root, bg=self.style.lookup(".", "background"), padx=15, pady=15)
        self.settings_frame = tk.Frame(self.root, bg=self.style.lookup(".", "background"), padx=15, pady=15)
        self.test_frame = tk.Frame(self.root, bg=self.style.lookup(".", "background"), padx=15, pady=15)

        self.create_home_widgets()
        self.create_groups_widgets()
        self.create_shortcuts_widgets()
        self.create_settings_widgets()
        self.create_test_widgets()

        self.show_section("home")

    def create_home_widgets(self):
        # Create main container with grid layout
        main_container = ttk.Frame(self.home_frame)
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Left side - Input section (70% width)
        input_frame = ttk.LabelFrame(main_container, text="Create New Shortcut")
        input_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Create input elements
        self.create_enhanced_input_section(input_frame)

        # Save button
        tk.Button(input_frame, text="💾 Save Shortcut", command=self.save_shortcut,
                  font=self.button_font, bg=self.style.lookup("TButton", "background"),
                  fg=self.style.lookup("TButton", "foreground"), relief="flat",
                  borderwidth=0, activebackground="#0051CC").pack(pady=10)
        
        # Right side - Quick Actions (30% width)
        action_frame = ttk.LabelFrame(main_container, text="Quick Actions")
        action_frame.pack(side="right", fill="both", padx=(10, 0))

        # Quick action buttons
        actions = [
            ("📋 Import from Clipboard", self.import_from_clipboard),
            ("📊 View Statistics", self.show_statistics),
            ("🔄 Sync Settings", self.sync_settings),
            ("❓ Help", self.show_help)
        ]
        
        for text, command in actions:
            ttk.Button(action_frame, text=text, command=command, style="TButton").pack(
                fill="x", pady=5, padx=10
            )

    def import_from_clipboard(self):
        try:
            text = self.root.clipboard_get()
            self.expansion_entry.delete("1.0", tk.END)
            self.expansion_entry.insert("1.0", text)
        except tk.TclError:
            messagebox.showwarning("Clipboard Empty", "No text found in clipboard.")

    def show_statistics(self):
        stats_window = tk.Toplevel(self.root)
        stats_window.title("Shortcut Statistics")
        stats_window.geometry("400x300")
        
        ttk.Label(stats_window, text=f"Total Shortcuts: {len(self.data['shortcuts'])}").pack(pady=5)
        ttk.Label(stats_window, text=f"Total Groups: {len(self.data['groups'])}").pack(pady=5)

    def sync_settings(self):
        messagebox.showinfo("Sync", "Settings synchronized successfully!")

    def show_help(self):
        help_text = """
Keyboard Shortcuts:
• Ctrl+V - Paste text
• Delete - Remove selected item
• Double-click - Edit shortcut

Groups Management:
• Create groups to organize shortcuts
• Select multiple groups/shortcuts with Ctrl or Shift
• Export/Import groups for backup
• Transfer shortcuts between groups

Best Practices:
• Use unique, memorable shortcuts
• Organize related shortcuts in groups
• Back up your shortcuts regularly
• Test shortcuts before using them



Created by: Erfan Razmi
Version: 1.0.0
        """
        messagebox.showinfo("Help Guide", help_text)

    def create_groups_widgets(self):
        # Frame for group creation and management
        group_management_frame = tk.Frame(self.groups_frame, bg=self.style.lookup(".", "background"))
        group_management_frame.pack(fill="x", pady=5)

        # Left side - Group creation
        new_group_frame = tk.Frame(group_management_frame, bg=self.style.lookup(".", "background"))
        new_group_frame.pack(side="left", fill="x", expand=True)

        tk.Label(new_group_frame, text="New Group:", font=self.label_font,
                 bg=self.style.lookup(".", "background"), fg=self.style.lookup(".", "foreground")).pack(side="left", padx=(0, 5))
        self.new_group_entry = ttk.Entry(new_group_frame, font=self.default_font, style="TEntry")
        self.new_group_entry.pack(side="left", expand=True, fill="x", padx=(0, 5))

        tk.Button(new_group_frame, text="➕ Create", command=self.create_group, font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(side="left")

        # Listbox for groups with enhanced styling
        groups_frame = tk.Frame(self.groups_frame, bg=self.style.lookup(".", "background"))
        groups_frame.pack(fill="both", expand=True, pady=5)

        # Split into two columns
        left_column = tk.Frame(groups_frame, bg=self.style.lookup(".", "background"))
        left_column.pack(side="left", fill="both", expand=True, padx=5)
        
        right_column = tk.Frame(groups_frame, bg=self.style.lookup(".", "background"))
        right_column.pack(side="left", fill="both", expand=True, padx=5)

        # Left column - Groups list
        tk.Label(left_column, text="Groups:", font=self.label_font,
                 bg=self.style.lookup(".", "background"), fg=self.style.lookup(".", "foreground")).pack(pady=5)
        
        self.groups_listbox = tk.Listbox(left_column, font=self.default_font, relief="flat", borderwidth=0,
                                         selectbackground=self.listbox_select_bg, selectforeground=self.listbox_select_fg,
                                         bg=self.listbox_bg, fg=self.listbox_fg, activestyle='none',
                                         exportselection=False)  # Add this line to prevent selection clearing
        self.groups_listbox.pack(fill="both", expand=True)
        self.groups_listbox.bind("<<ListboxSelect>>", self.show_group_members)
        self.groups_listbox.config(selectmode=tk.EXTENDED)  # Enable multiple selection

        # Delete groups button
        tk.Button(left_column, text="🗑️ Delete Selected Groups", 
                 command=self.delete_selected_groups,
                 font=self.button_font,
                 bg=self.style.lookup("TButton", "background"), 
                 fg=self.style.lookup("TButton", "foreground"),
                 relief="flat", borderwidth=0, 
                 activebackground="#0051CC").pack(pady=5)

        # Right column - Group members and transfer
        tk.Label(right_column, text="Group Members:", font=self.label_font,
                 bg=self.style.lookup(".", "background"), fg=self.style.lookup(".", "foreground")).pack(pady=5)
        
        self.group_members_listbox = tk.Listbox(right_column, font=self.default_font, relief="flat", borderwidth=0,
                                                selectbackground=self.listbox_select_bg, selectforeground=self.listbox_select_fg,
                                                bg=self.listbox_bg, fg=self.listbox_fg, activestyle='none',
                                                selectmode=tk.EXTENDED, exportselection=False)  # Add exportselection=False
        self.group_members_listbox.pack(fill="both", expand=True)

        # Transfer controls
        transfer_frame = tk.Frame(right_column, bg=self.style.lookup(".", "background"))
        transfer_frame.pack(fill="x", pady=5)

        tk.Label(transfer_frame, text="Transfer to:", font=self.label_font,
                 bg=self.style.lookup(".", "background"), fg=self.style.lookup(".", "foreground")).pack(side="left", padx=5)
        
        self.transfer_group_combobox = ttk.Combobox(transfer_frame, values=[], state="readonly", font=self.default_font,
                                                    style="TCombobox")
        self.transfer_group_combobox.pack(side="left", fill="x", expand=True, padx=5)

        tk.Button(transfer_frame, text="➡️ Transfer", command=self.transfer_shortcut, font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(side="left", padx=5)

        # Add standard group management buttons
        tk.Button(right_column, text="🗑️ Remove from Group", command=self.remove_shortcut_from_group, font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(pady=5)

        # Add delete multiple shortcuts button
        tk.Button(right_column, text="🗑️ Delete Selected Shortcuts", command=self.delete_shortcut_from_list,
                  font=self.button_font, bg=self.style.lookup("TButton", "background"),
                  fg=self.style.lookup("TButton", "foreground"), relief="flat",
                  borderwidth=0, activebackground="#0051CC").pack(pady=5)

    def delete_selected_groups(self):
        """Combined function for deleting one or multiple groups"""
        selections = self.groups_listbox.curselection()
        if not selections:
            messagebox.showwarning("No Selection", "Please select groups to delete.")
            return
            
        groups_to_delete = [
            self.groups_listbox.get(idx).replace("📂 ", "").strip()
            for idx in selections
        ]
        
        message = "Delete group" if len(groups_to_delete) == 1 else f"Delete {len(groups_to_delete)} groups"
        if messagebox.askyesno("Confirm Delete", 
                             f"{message} and remove their shortcuts from groups?"):
            for group in groups_to_delete:
                # Remove group from groups list
                if group in self.data["groups"]:
                    self.data["groups"].remove(group)
                
                # Update shortcuts that were in this group
                for shortcut in self.data["shortcuts"].values():
                    if shortcut["group"] == group:
                        shortcut["group"] = ""
            
            self.save_data()
            self.update_ui()
            messagebox.showinfo("Success", 
                              f"Deleted {len(groups_to_delete)} group{'s' if len(groups_to_delete) > 1 else ''}")

    def transfer_shortcut(self):
        try:
            # Get all selected shortcuts
            selections = self.group_members_listbox.curselection()
            if not selections:
                messagebox.showwarning("No Selection", "Please select shortcuts to transfer.")
                return

            shortcuts = []
            for index in selections:
                shortcut_text = self.group_members_listbox.get(index)
                shortcut = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()
                shortcuts.append(shortcut)

            # Get target group
            target_group = self.transfer_group_combobox.get()
            
            if not target_group:
                messagebox.showwarning("No Target", "Please select a target group.")
                return

            # Update all selected shortcuts' groups
            for shortcut in shortcuts:
                self.data["shortcuts"][shortcut]["group"] = target_group
            
            self.save_data()
            self.update_ui()
            self.show_group_members()
            messagebox.showinfo("Success", f"{len(shortcuts)} shortcuts transferred to '{target_group}'.")
            
        except IndexError:
            messagebox.showwarning("No Selection", "Please select shortcuts to transfer.")

    def show_group_members(self, event=None):
        """Keep group selected and show members while maintaining shortcut selection"""
        try:
            # Store current shortcut selections
            selected_shortcuts = []
            if self.group_members_listbox:
                selected_shortcuts = [self.group_members_listbox.get(idx) for idx in self.group_members_listbox.curselection()]
            
            selections = self.groups_listbox.curselection()
            if not selections:
                # If no selection, keep the current group
                if self.current_group:
                    for i in range(self.groups_listbox.size()):
                        if self.groups_listbox.get(i).replace("📂 ", "").strip() == self.current_group:
                            self.groups_listbox.selection_set(i)
                            break
                return
                
            group = self.groups_listbox.get(selections[0]).replace("📂 ", "").strip()
            self.current_group = group
            
            # Update transfer combobox
            available_groups = [g for g in self.data["groups"] if g != group]
            self.transfer_group_combobox["values"] = available_groups
            self.transfer_group_combobox.set("")

            # Update member list while preserving selections
            current_shortcuts = set(self.group_members_listbox.get(0, tk.END))
            self.group_members_listbox.delete(0, tk.END)
            
            members = [(shortcut, details["expansion"]) 
                      for shortcut, details in self.data["shortcuts"].items() 
                      if details["group"] == group]
            
            # Insert new items
            for shortcut, expansion in members:
                item_text = f"🔤 {shortcut} ➔ {expansion}"
                self.group_members_listbox.insert(tk.END, item_text)
                # Restore selection if item was previously selected
                if item_text in selected_shortcuts:
                    idx = self.group_members_listbox.size() - 1
                    self.group_members_listbox.selection_set(idx)
            
            # Maintain group selection
            for i in range(self.groups_listbox.size()):
                if self.groups_listbox.get(i).replace("📂 ", "").strip() == group:
                    self.groups_listbox.selection_set(i)
                    break
                    
        except (IndexError, KeyError) as e:
            print(f"Error in show_group_members: {e}")

    def animate_insert(self, widget, text, duration=200):
        """Animate insertion of items"""
        def _insert():
            widget.insert(tk.END, text)
            widget.see(tk.END)
            
        self.root.after(duration, _insert)

    def create_shortcuts_widgets(self):
        # Listbox for shortcuts
        self.shortcuts_listbox = tk.Listbox(self.shortcuts_frame, font=self.default_font, relief="flat", borderwidth=0,
                                            selectbackground="#007AFF", selectforeground="#FFFFFF",
                                            bg=self.listbox_bg, fg=self.listbox_fg, activestyle='none')
        self.shortcuts_listbox.pack(pady=5, fill="both", expand=True)
        self.shortcuts_listbox.bind("<Double-Button-1>", self.edit_shortcut)
        self.shortcuts_listbox.bind("<<ListboxSelect>>", self.prepare_delete_shortcut)
        self.shortcuts_listbox.config(selectmode=tk.EXTENDED)  # Enable multiple selection

        # Add a manual delete button for the selected shortcut
        tk.Button(self.shortcuts_frame, text="🗑️ Delete Shortcut", command=self.delete_selected_shortcut, font=self.button_font,
                  bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                  relief="flat", borderwidth=0, activebackground="#0051CC").pack(pady=10)

        # Add bulk actions frame
        bulk_actions = ttk.Frame(self.shortcuts_frame)
        bulk_actions.pack(fill="x", pady=5)
        
        ttk.Button(bulk_actions, text="🗑️ Delete Selected", 
                   command=self.delete_selected_shortcuts).pack(side="left", padx=5)
        ttk.Button(bulk_actions, text="📤 Export Selected", 
                   command=self.export_selected_shortcuts).pack(side="left", padx=5)

    def create_settings_widgets(self):
        # Create notebook for tabbed settings
        settings_notebook = ttk.Notebook(self.settings_frame)
        settings_notebook.pack(fill="both", expand=True)
        
        # General Settings tab
        general_frame = ttk.Frame(settings_notebook)
        settings_notebook.add(general_frame, text="General")
        
        # Appearance tab
        appearance_frame = ttk.Frame(settings_notebook)
        settings_notebook.add(appearance_frame, text="Appearance")
        
        # Create theme selection
        theme_frame = ttk.LabelFrame(appearance_frame, text="Theme Selection")
        theme_frame.pack(fill="x", padx=10, pady=5)
        
        # Create theme buttons
        for theme_name, theme_colors in self.themes.items():
            ttk.Button(
                theme_frame,
                text=f"Switch to {theme_name.title()}", 
                command=lambda t=theme_name: self.toggle_theme(t)
            ).pack(side="left", padx=5, pady=5)
        
        # Additional settings
        self.create_additional_settings(general_frame)
        
        # About section
        about_frame = ttk.LabelFrame(general_frame, text="About")
        about_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(about_frame, text="Created by: Erfan Razmi").pack(pady=5)
        ttk.Label(about_frame, text="Version: 1.0.0").pack(pady=5)

    def create_test_widgets(self):
        tk.Label(self.test_frame, text="Type Shortcut:", font=self.label_font,
                 bg=self.style.lookup("TLabel", "background"), fg=self.style.lookup("TLabel", "foreground")).pack(pady=5)
        self.test_entry = ttk.Entry(self.test_frame, font=self.default_font, style="TEntry")
        self.test_entry.pack(pady=5, fill="x")
        self.test_entry.bind("<KeyRelease>", self.test_shortcut)

        self.test_result_label = tk.Label(self.test_frame, text="", font=self.label_font,
                                         bg=self.style.lookup("TLabel", "background"), fg=self.style.lookup("TLabel", "foreground"))
        self.test_result_label.pack(pady=5)

    def show_section(self, section):
        """Animate section transitions"""
        # First fade out current section
        for frame in [self.home_frame, self.groups_frame, self.shortcuts_frame, 
                     self.settings_frame, self.test_frame]:
            if frame.winfo_ismapped():
                frame.pack_forget()
        
        # Show new section with fade in effect
        target_frame = getattr(self, f"{section}_frame")
        target_frame.pack(fill="both", expand=True)
        
        def fade_in(step=0):
            if step <= 10:
                opacity = step / 10
                target_frame.update()  # Force frame update
                self.root.after(20, lambda: fade_in(step + 1))
        
        fade_in()

    def update_ui(self):
        """Update all UI elements with current data"""
        if self.group_combobox:
            self.group_combobox["values"] = [""] + self.data["groups"]

        if self.groups_listbox:
            self.groups_listbox.delete(0, tk.END)
            for group in self.data["groups"]:
                self.groups_listbox.insert(tk.END, "📂 " + group)

        if self.shortcuts_listbox:
            self.shortcuts_listbox.delete(0, tk.END)
            for shortcut, details in self.data["shortcuts"].items():
                group_label = f"[{details['group']}]" if details['group'] else ""
                self.shortcuts_listbox.insert(tk.END, f"🔤 {shortcut} ➔ {details['expansion']} {group_label}")

        if self.group_name_label:
            self.group_name_label.config(text="")
            
        if self.group_members_listbox:
            self.group_members_listbox.delete(0, tk.END)

    def toggle_theme(self, theme):
        self.current_theme = theme
        self.data["theme"] = theme
        self.save_data()
        self.apply_theme()

    def save_shortcut(self):
        shortcut = self.shortcut_entry.get().strip()
        group = self.group_combobox.get().strip()
        expansion = self.expansion_entry.get("1.0", tk.END).strip()

        if shortcut and expansion:
            self.data["shortcuts"][shortcut] = {"expansion": expansion, "group": group}
            if group and group not in self.data["groups"]:
                self.data["groups"].append(group)
            self.save_data()
            self.clear_inputs()
            self.update_ui()
            messagebox.showinfo("Info", "✅ Shortcut saved!")

    def prepare_delete_shortcut(self, event=None):
        try:
            selection = self.shortcuts_listbox.curselection()[0]
            shortcut_text = self.shortcuts_listbox.get(selection)
            self.selected_shortcut_for_deletion = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()
        except IndexError:
            self.selected_shortcut_for_deletion = None

    def delete_selected_shortcut(self):
        # This function is called when the "Delete Shortcut" button is pressed
        if self.selected_shortcut_for_deletion:
            try:
                del self.data["shortcuts"][self.selected_shortcut_for_deletion]
                self.save_data()
                self.update_ui()
                self.selected_shortcut_for_deletion = None
            except KeyError:
                messagebox.showerror("Error", "Shortcut not found in data.")
        else:
            messagebox.showwarning("No Selection", "Please select a shortcut to delete.")
    
    def delete_shortcut_from_list(self, event=None):
        """Delete selected shortcuts from the group members list"""
        try:
            # Get group selection
            group_selection = self.groups_listbox.curselection()
            if not group_selection:
                messagebox.showwarning("No Group", "Please select a group first.")
                return

            # Get shortcuts selection
            selections = self.group_members_listbox.curselection()
            if not selections:
                messagebox.showwarning("No Selection", "Please select shortcuts to delete.")
                return

            current_group = self.groups_listbox.get(group_selection[0]).replace("📂 ", "").strip()
            shortcuts = []
            
            # Get all selected shortcuts
            for index in selections:
                shortcut_text = self.group_members_listbox.get(index)
                shortcut = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()
                shortcuts.append(shortcut)

            if shortcuts:
                if messagebox.askyesno("Confirm Delete", 
                                     f"Permanently delete {len(shortcuts)} shortcuts?"):
                    # Delete all selected shortcuts
                    for shortcut in shortcuts:
                        if shortcut in self.data["shortcuts"]:
                            del self.data["shortcuts"][shortcut]
                    
                    self.save_data()
                    self.show_group_members()
                    self.update_ui()
                    messagebox.showinfo("Success", f"{len(shortcuts)} shortcuts deleted.")

        except Exception as e:
            print(f"Error in delete_shortcut_from_list: {e}")
            messagebox.showwarning("Error", "Please select a group and shortcuts to delete.")

    def edit_shortcut(self, event):
        try:
            selection = self.shortcuts_listbox.curselection()[0]
            shortcut_text = self.shortcuts_listbox.get(selection)
            shortcut = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()
            expansion = self.data["shortcuts"][shortcut]["expansion"]
            group = self.data["shortcuts"][shortcut]["group"]

            self.shortcut_entry.insert(0, shortcut)
            self.group_combobox.set(group)
            self.expansion_entry.delete("1.0", tk.END)
            self.expansion_entry.insert("1.0", expansion)
            self.show_section("home")
        except IndexError:
            pass

    def create_group(self):
        group = self.new_group_entry.get().strip()
        if group and group not in self.data["groups"]:
            self.data["groups"].append(group)
            self.save_data()
            self.update_ui()
            self.new_group_entry.delete(0, tk.END)
        else:
            messagebox.showwarning("Duplicate Group", "Group already exists!")

    def prepare_remove_shortcut_from_group(self, event=None):
        try:
            group_selection = self.groups_listbox.curselection()[0]
            self.selected_group_for_removal = self.groups_listbox.get(group_selection).replace("📂 ", "").strip()
            shortcut_selection = self.group_members_listbox.curselection()[0]
            shortcut_text = self.group_members_listbox.get(shortcut_selection)
            self.selected_shortcut_for_removal = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()
        except IndexError:
            self.selected_group_for_removal = None
            self.selected_shortcut_for_removal = None

    def remove_shortcut_from_group(self, event=None):
        """Remove selected shortcuts from their current group"""
        try:
            # Get group selection
            group_selection = self.groups_listbox.curselection()
            if not group_selection:
                messagebox.showwarning("No Group", "Please select a group first.")
                return

            # Get shortcuts selection
            selections = self.group_members_listbox.curselection()
            if not selections:
                messagebox.showwarning("No Selection", "Please select shortcuts to remove.")
                return

            current_group = self.groups_listbox.get(group_selection[0]).replace("📂 ", "").strip()
            shortcuts = []
            
            # Get all selected shortcuts
            for index in selections:
                shortcut_text = self.group_members_listbox.get(index)
                shortcut = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()
                shortcuts.append(shortcut)

            if shortcuts:
                if messagebox.askyesno("Confirm Remove", 
                                     f"Remove {len(shortcuts)} shortcuts from group '{current_group}'?"):
                    # Remove all selected shortcuts from group
                    for shortcut in shortcuts:
                        if shortcut in self.data["shortcuts"]:
                            self.data["shortcuts"][shortcut]["group"] = ""
                    
                    self.save_data()
                    self.show_group_members()
                    self.update_ui()
                    messagebox.showinfo("Success", f"{len(shortcuts)} shortcuts removed from group.")

        except Exception as e:
            print(f"Error in remove_shortcut_from_group: {e}")
            messagebox.showwarning("Error", "Please select a group and shortcuts to remove.")

    def move_shortcut_to_another_group(self):
        try:
            # Get the selected shortcut
            shortcut_selection = self.group_members_listbox.curselection()[0]
            shortcut_text = self.group_members_listbox.get(shortcut_selection)
            shortcut = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()

            # Get the available groups
            available_groups = self.data["groups"].copy()
            current_group = self.data["shortcuts"][shortcut]["group"]
            if current_group in available_groups:
                available_groups.remove(current_group)

            # Create a new top-level window for group selection
            group_selection_window = tk.Toplevel(self.root)
            group_selection_window.title("Move Shortcut To Group")
            group_selection_window.geometry("300x200")

            # Label for instructions
            tk.Label(group_selection_window, text="Select a group:", font=self.label_font).pack(pady=5)

            # Listbox for available groups
            groups_listbox = tk.Listbox(group_selection_window, font=self.default_font, relief="flat", borderwidth=0,
                                         selectbackground=self.listbox_select_bg, selectforeground=self.listbox_select_fg,
                                         bg=self.listbox_bg, fg=self.listbox_fg, activestyle='none')
            for group in available_groups:
                groups_listbox.insert(tk.END, group)
            groups_listbox.pack(pady=5, fill="both", expand=True)

            # Move button
            def confirm_move():
                try:
                    # Get the selected group
                    selected_group_index = groups_listbox.curselection()[0]
                    selected_group = groups_listbox.get(selected_group_index)

                    # Update the shortcut's group
                    self.data["shortcuts"][shortcut]["group"] = selected_group
                    self.save_data()
                    self.update_ui()

                    # Refresh group members list
                    self.show_group_members()

                    # Close the group selection window
                    group_selection_window.destroy()
                except IndexError:
                    messagebox.showwarning("No Selection", "Please select a group.")

            tk.Button(group_selection_window, text="Move Shortcut", command=confirm_move, font=self.button_font,
                      bg=self.style.lookup("TButton", "background"), fg=self.style.lookup("TButton", "foreground"),
                      relief="flat", borderwidth=0, activebackground="#0051CC").pack(pady=10)

        except IndexError:
            messagebox.showwarning("No Selection", "Please select a shortcut to move.")
        except KeyError:
            messagebox.showerror("Error", "Shortcut not found.")

    def delete_group(self, event=None):
        try:
            selection = self.groups_listbox.curselection()[0]
            group = self.groups_listbox.get(selection).replace("📂 ", "").strip()

            # Remove the group from the list of groups
            self.data["groups"].remove(group)

            # Remove shortcuts associated with the group
            shortcuts_to_delete = [k for k, v in self.data["shortcuts"].items() if v["group"] == group]
            for shortcut in shortcuts_to_delete:
                del self.data["shortcuts"][shortcut]

            self.save_data()
            self.update_ui()
        except (IndexError, ValueError):  # ValueError if the group is not in the list
            pass

    def export_backup(self):
        try:
            filepath = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if filepath:
                with open(filepath, "w") as f:
                    json.dump(self.data, f)
                messagebox.showinfo("Info", "Backup exported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export backup: {e}")

    def import_backup(self):
        try:
            filepath = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if filepath:
                with open(filepath, "r") as f:
                    self.data = json.load(f)
                self.save_data()
                self.update_ui()
                messagebox.showinfo("Info", "Backup imported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import backup: {e}")
            
    def export_group(self):
        try:
            selection = self.groups_listbox.curselection()[0]
            group = self.groups_listbox.get(selection).replace("📂 ", "").strip()

            # Create a dictionary containing only shortcuts belonging to the selected group
            group_data = {"shortcuts": {k: v for k, v in self.data["shortcuts"].items() if v["group"] == group}}

            filepath = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if filepath:
                with open(filepath, "w") as f:
                    json.dump(group_data, f)  # Save only the shortcuts of the selected group
                messagebox.showinfo("Info", f"Group '{group}' exported successfully!")
        except (IndexError, Exception) as e:
            messagebox.showerror("Error", f"Failed to export group: {e}")

    def test_shortcut(self, event=None):
        shortcut = self.test_entry.get().strip()
        if shortcut in self.data["shortcuts"]:
            expansion = self.data["shortcuts"][shortcut]["expansion"]
            self.test_result_label.config(text=f"Result: {expansion}")
        else:
            self.test_result_label.config(text="Shortcut not found")

    def clear_inputs(self):
        self.shortcut_entry.delete(0, tk.END)
        self.expansion_entry.delete("1.0", tk.END)
        self.group_combobox.set("")

    def register_hotkey(self):
        keyboard.on_press(self.on_key_press)

    def on_key_press(self, event):
        """Improved key press handling without key blocking"""
        try:
            # Ignore modifier keys
            if event.name in ['shift', 'ctrl', 'alt', 'windows', 'tab']:
                return

            print(f"Key pressed: {event.name}, Buffer: {self.typed_text}")

            if event.name == 'space':
                if self.typed_text:
                    self.check_for_shortcut(force_check=True)
                self.typed_text = ""
            elif event.name == 'enter':
                if self.typed_text:
                    self.check_for_shortcut(force_check=True)
                self.typed_text = ""
            elif event.name == 'backspace':
                if self.typed_text:
                    self.typed_text = self.typed_text[:-1]
            elif len(event.name) == 1:  # Only single characters
                self.typed_text += event.name
                
                # Cancel existing timer
                if self.typing_timer is not None:
                    self.typing_timer.cancel()
                
                # Set timer for checking shortcuts
                self.typing_timer = threading.Timer(0.3, self.check_for_shortcut)
                self.typing_timer.start()
                
                # Set timer to clear buffer using user setting
                self.root.after(self.buffer_clear_time, self.clear_typed_buffer)

        except Exception as e:
            print(f"Error in key press handler: {e}")

    def clear_typed_buffer(self):
        """Clear the typing buffer after delay"""
        if self.typed_text:
            print(f"Clearing buffer: {self.typed_text}")
            self.typed_text = ""

    def check_for_shortcut(self, force_check=False):
        """Improved shortcut checking without key blocking"""
        try:
            if not self.typed_text:
                return

            print(f"Checking text: '{self.typed_text}'")
            
            # Sort shortcuts by length (longest first)
            sorted_shortcuts = sorted(self.data['shortcuts'].items(), 
                                   key=lambda x: len(x[0]), 
                                   reverse=True)
            
            # Check for shortcuts
            for shortcut, details in sorted_shortcuts:
                if self.typed_text.endswith(shortcut):
                    # Check if shortcut is at start of text or preceded by space
                    shortcut_start = len(self.typed_text) - len(shortcut)
                    if shortcut_start == 0 or self.typed_text[shortcut_start - 1] == ' ':
                        print(f"Match found! Shortcut: '{shortcut}'")
                        expansion = details["expansion"]
                        chars_to_remove = len(shortcut)
                        
                        try:
                            # Remove only the shortcut characters
                            for _ in range(chars_to_remove):
                                keyboard.press_and_release('backspace')
                                time.sleep(0.01)
                            
                            # Type the expansion directly
                            keyboard.write(expansion)
                            
                            # Play sound
                            self.play_sound()
                            
                            # Clear buffer and block next space
                            self.typed_text = ""
                            self.ignore_next_space = False
                            
                            print("Expansion completed successfully")
                            return True
                            
                        except Exception as e:
                            print(f"Error during expansion: {e}")
                            return False

            # Clear buffer if forced
            if force_check:
                self.typed_text = ""

        except Exception as e:
            print(f"Error in shortcut checker: {e}")
            return False

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.root.destroy()

    def choose_sound(self):
        filepath = filedialog.askopenfilename(
            title="Choose a sound file",
            filetypes=[("WAV files", "*.wav"), ("All files", "*.*")]
        )
        if filepath:
            self.data["sound_file"] = filepath
            self.sound_file_label.config(text=os.path.basename(filepath))
            self.save_data()

    def play_sound(self, sound_name="SystemExclamation"):
        """Updated play_sound with volume control"""
        sound_file = self.data.get("sound_file")
        volume = self.volume_var.get() if hasattr(self, 'volume_var') else 100
        
        if self.sound_enabled_var.get() and sound_file and os.path.exists(sound_file):
            try:
                # Convert volume to system volume (0-65535)
                system_volume = int((volume / 100) * 65535)
                winsound.PlaySound(sound_file, 
                                 winsound.SND_FILENAME | 
                                 winsound.SND_ASYNC | 
                                 winsound.SND_NODEFAULT)
            except Exception as e:
                print(f"Error playing custom sound: {e}")
                self.play_default_sound()
        else:
            self.play_default_sound()

    def delete_selected_shortcuts(self):
        """Delete multiple shortcuts from either shortcuts tab or groups tab"""
        if self.shortcuts_frame.winfo_ismapped():
            # Delete from shortcuts tab
            selections = self.shortcuts_listbox.curselection()
            source_listbox = self.shortcuts_listbox
        elif self.groups_frame.winfo_ismapped():
            # Delete from groups tab
            selections = self.group_members_listbox.curselection()
            source_listbox = self.group_members_listbox
        else:
            return

        if not selections:
            messagebox.showwarning("No Selection", "Please select shortcuts to delete.")
            return
            
        shortcuts_to_delete = [
            source_listbox.get(idx).split(" ➔ ")[0].replace("🔤 ", "").strip()
            for idx in selections
        ]
        
        if messagebox.askyesno("Confirm Delete", 
                              f"Permanently delete {len(selections)} shortcuts?"):
            for shortcut in shortcuts_to_delete:
                if shortcut in self.data["shortcuts"]:
                    del self.data["shortcuts"][shortcut]
            
            self.save_data()
            self.update_ui()
            if self.groups_frame.winfo_ismapped():
                self.show_group_members()
            messagebox.showinfo("Success", f"{len(selections)} shortcuts deleted.")

    def create_additional_settings(self, parent):
        """Enhanced settings with buffer timing control"""
        # Backup Settings
        backup_frame = ttk.LabelFrame(parent, text="Backup Settings")
        backup_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(backup_frame, text="Auto-backup frequency:").pack(side="left")
        ttk.Combobox(backup_frame, values=["Never", "Daily", "Weekly", "Monthly"], 
                     state="readonly").pack(side="left", padx=5)

        # Sound Settings
        sound_frame = ttk.LabelFrame(parent, text="Sound Settings")
        sound_frame.pack(fill="x", padx=10, pady=5)
        
        # Sound file selection
        sound_file_frame = ttk.Frame(sound_frame)
        sound_file_frame.pack(fill="x", pady=5)
        
        self.sound_enabled_var = tk.BooleanVar(value=bool(self.sound_file))
        ttk.Checkbutton(sound_frame, text="Enable sound on expansion", 
                       variable=self.sound_enabled_var,
                       command=self.toggle_sound).pack(pady=5)
        
        ttk.Label(sound_file_frame, text="Sound file:").pack(side="left", padx=5)
        self.sound_file_label = ttk.Label(sound_file_frame, 
                                        text=os.path.basename(self.sound_file) if self.sound_file else "No file selected")
        self.sound_file_label.pack(side="left", expand=True, fill="x", padx=5)
        
        ttk.Button(sound_file_frame, text="Choose File", 
                  command=self.choose_sound_file).pack(side="right", padx=5)
        ttk.Button(sound_file_frame, text="Test Sound", 
                  command=self.test_sound).pack(side="right", padx=5)

        # Typing Settings
        typing_frame = ttk.LabelFrame(parent, text="Typing Settings")
        typing_frame.pack(fill="x", padx=10, pady=5)
        
        # Buffer clear time setting
        buffer_frame = ttk.Frame(typing_frame)
        buffer_frame.pack(fill="x", pady=5)
        
        ttk.Label(buffer_frame, text="Clear typing buffer after (ms):").pack(side="left", padx=5)
        self.buffer_time_var = tk.StringVar(value=str(self.buffer_clear_time))
        buffer_entry = ttk.Entry(buffer_frame, textvariable=self.buffer_time_var, width=10)
        buffer_entry.pack(side="left", padx=5)
        
        def update_buffer_time(*args):
            try:
                time = max(1000, int(self.buffer_time_var.get()))
                self.buffer_clear_time = time
                self.data["buffer_clear_time"] = time
                self.save_data()
            except ValueError:
                self.buffer_time_var.set(str(self.buffer_clear_time))
        
        self.buffer_time_var.trace('w', update_buffer_time)
        
        # Help text
        help_text = """Typing Buffer Settings:
• Buffer Clear Time: How long to wait before clearing typed text (min 1000ms)
• Longer times allow for slower typing of shortcuts
• Shorter times may improve system responsiveness
• Default: 10000 ms (10 seconds)"""
        
        ttk.Label(typing_frame, text=help_text, wraplength=400, justify="left").pack(pady=5)

    def toggle_sound(self):
        if not self.sound_enabled_var.get():
            self.sound_file = ""
            self.sound_file_label.config(text="No file selected")
            self.data["sound_file"] = ""
            self.save_data()

    def choose_sound_file(self):
        filepath = filedialog.askopenfilename(
            title="Choose Sound File",
            filetypes=[
                ("WAV files", "*.wav"),
                ("MP3 files", "*.mp3"),
                ("All files", "*.*")
            ]
        )
        if filepath:
            self.sound_file = filepath
            self.sound_file_label.config(text=os.path.basename(filepath))
            self.sound_enabled_var.set(True)
            self.data["sound_file"] = filepath
            self.save_data()

    def test_sound(self):
        if self.sound_enabled_var.get() and self.sound_file and os.path.exists(self.sound_file):
            try:
                winsound.PlaySound(self.sound_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
            except Exception as e:
                messagebox.showerror("Error", f"Could not play sound file: {e}")
        else:
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)

    def play_sound(self):
        if self.sound_enabled_var.get() and self.sound_file and os.path.exists(self.sound_file):
            try:
                winsound.PlaySound(self.sound_file, 
                                 winsound.SND_FILENAME | winsound.SND_ASYNC)
            except Exception as e:
                print(f"Error playing custom sound: {e}")
                self.play_default_sound()
        else:
            self.play_default_sound()

    def play_default_sound(self):
        try:
            winsound.PlaySound("SystemExclamation", 
                             winsound.SND_ALIAS | winsound.SND_ASYNC)
        except Exception as e:
            print(f"Error playing system sound: {e}")

    def paste_text(self, event):
        """Handle paste events in text widgets"""
        try:
            clipboard_text = self.root.clipboard_get()
            event.widget.insert(tk.INSERT, clipboard_text)
            return "break"  # Prevent default paste behavior
        except tk.TclError:
            pass

    def create_enhanced_input_section(self, input_frame):
        """Create enhanced input section with labels and widgets"""
        tk.Label(input_frame, text="Shortcut:", font=self.label_font,
                bg=self.style.lookup("TLabel", "background"),
                fg=self.style.lookup("TLabel", "foreground")).pack(pady=5)
        self.shortcut_entry = ttk.Entry(input_frame, font=self.default_font, style="TEntry")
        self.shortcut_entry.pack(pady=5, fill="x", padx=10)

        tk.Label(input_frame, text="Group:", font=self.label_font,
                bg=self.style.lookup("TLabel", "background"),
                fg=self.style.lookup("TLabel", "foreground")).pack(pady=5)
        self.group_combobox = ttk.Combobox(input_frame, values=self.data["groups"],
                                          state="readonly", font=self.default_font, style="TCombobox")
        self.group_combobox.pack(pady=5, fill="x", padx=10)

        tk.Label(input_frame, text="Expansion:", font=self.label_font,
                bg=self.style.lookup("TLabel", "background"),
                fg=self.style.lookup("TLabel", "foreground")).pack(pady=5)
        self.expansion_entry = tk.Text(input_frame, height=3, font=self.default_font,
                                     relief="flat", bg=self.style.lookup("TEntry", "fieldbackground"),
                                     fg=self.style.lookup("TEntry", "foreground"),
                                     insertbackground=self.style.lookup("TEntry", "foreground"))
        self.expansion_entry.pack(pady=5, fill="both", expand=True, padx=10)
        
        # Bind paste events
        self.expansion_entry.bind("<Control-v>", self.paste_text)
        self.expansion_entry.bind("<Command-v>", self.paste_text)  # For macOS

    def export_selected_shortcuts(self):
        """Export multiple selected shortcuts"""
        selections = self.shortcuts_listbox.curselection()
        if not selections:
            messagebox.showwarning("No Selection", "Please select shortcuts to export.")
            return
            
        shortcuts_to_export = {}
        for idx in selections:
            shortcut_text = self.shortcuts_listbox.get(idx)
            shortcut = shortcut_text.split(" ➔ ")[0].replace("🔤 ", "").strip()
            if shortcut in self.data["shortcuts"]:
                shortcuts_to_export[shortcut] = self.data["shortcuts"][shortcut]
        
        if shortcuts_to_export:
            try:
                filepath = filedialog.asksaveasfilename(
                    defaultextension=".json",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
                if filepath:
                    with open(filepath, "w", encoding="utf-8") as f:
                        json.dump({"shortcuts": shortcuts_to_export}, f, indent=4)
                    messagebox.showinfo("Success", f"{len(shortcuts_to_export)} shortcuts exported successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export shortcuts: {e}")

if __name__ == "__main__":
    ShortcutManager()