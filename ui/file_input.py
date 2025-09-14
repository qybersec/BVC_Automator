"""
File Input Component
Handles file selection and display with drag-and-drop support.
"""

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolledtext
from .styles import COLORS, SPACING, FONTS, DIMENSIONS

class FileInputComponent:
    """File input section with drag-and-drop support"""
    
    def __init__(self, parent, browse_callback, drag_drop_setup_callback):
        self.parent = parent
        self.browse_callback = browse_callback
        self.drag_drop_setup_callback = drag_drop_setup_callback
        self.file_section = None
        self.file_display = None
        self.file_display_frame = None
        
    def create(self):
        """Create the file input UI section"""
        self.file_section = tk.Frame(self.parent, **COLORS.get('frame_default', {'bg': COLORS['BACKGROUND_GRAY']}))
        self.file_section.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5), padx=3)
        self.file_section.columnconfigure(0, weight=1)
        
        # Section header
        header_style = {'background': COLORS['BACKGROUND_GRAY']}
        ttk.Label(
            self.file_section, 
            text="📁 Input File", 
            style='Header.TLabel', 
            **header_style
        ).grid(row=0, column=0, pady=(10, 5))
        
        file_frame = tk.Frame(self.file_section, bg=COLORS['BACKGROUND_GRAY'])
        file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10), padx=SPACING['LARGE'])
        file_frame.columnconfigure(0, weight=1)
        
        # File display with clean styling and drag-drop support - expandable width
        self.file_display_frame = tk.Frame(file_frame, bg=COLORS['BACKGROUND_WHITE'], relief='flat', bd=1)
        self.file_display_frame.grid(row=0, column=0, padx=(0, SPACING['LARGE']), sticky=(tk.W, tk.E))
        
        # Create scrollable text widget for multiple file names
        self.file_display = scrolledtext.ScrolledText(
            self.file_display_frame,
            height=DIMENSIONS['FILE_DISPLAY_HEIGHT'],
            width=DIMENSIONS['FILE_DISPLAY_WIDTH'],
            font=FONTS['SMALL'],
            fg=COLORS['TEXT_PRIMARY'],
            bg=COLORS['BACKGROUND_WHITE'],
            wrap=tk.WORD,
            state='disabled',
            borderwidth=0,
            highlightthickness=0
        )
        self.file_display.pack(fill='both', expand=True, padx=6, pady=4)
        
        # Initialize with placeholder text
        self.file_display.config(state='normal')
        self.file_display.insert('1.0', "No files selected")
        self.file_display.config(state='disabled', fg=COLORS['TEXT_MUTED'])
        
        self.file_display_frame.grid_columnconfigure(0, weight=1)
        
        # Enable drag and drop
        if self.drag_drop_setup_callback:
            self.drag_drop_setup_callback(self.file_display_frame)
        
        # Browse button
        browse_button = ttk.Button(
            file_frame, 
            text="📂 Browse", 
            command=self.browse_callback, 
            style='Browse.TButton'
        )
        browse_button.grid(row=0, column=1)
        
        return self.file_section
    
    def update_file_display(self, text, is_placeholder=False):
        """Update the file display text"""
        if self.file_display:
            self.file_display.config(state='normal')
            self.file_display.delete('1.0', tk.END)
            self.file_display.insert('1.0', text)
            
            if is_placeholder:
                self.file_display.config(state='disabled', fg=COLORS['TEXT_MUTED'])
            else:
                self.file_display.config(state='disabled', fg=COLORS['TEXT_PRIMARY'])
    
    def show(self):
        """Show the file input section"""
        if self.file_section:
            self.file_section.grid()
    
    def hide(self):
        """Hide the file input section"""
        if self.file_section:
            self.file_section.grid_remove()
    
    def get_display_frame(self):
        """Get the file display frame for drag-and-drop setup"""
        return self.file_display_frame