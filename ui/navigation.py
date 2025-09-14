"""
Navigation Component
Handles the navigation bar with active state management.
"""

import tkinter as tk
from .styles import COLORS, SPACING, BUTTON_STYLES

class NavigationComponent:
    """Navigation bar component with active state management"""
    
    def __init__(self, parent, select_callback):
        self.parent = parent
        self.select_callback = select_callback
        self.nav_buttons = {}
        self.nav_bar = None
        
    def create(self):
        """Create the navigation bar"""
        self.nav_bar = tk.Frame(self.parent, bg=COLORS['BACKGROUND_BORDER'])
        self.nav_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=0, pady=0)
        self.nav_bar.columnconfigure(0, weight=1)
        
        nav_container = tk.Frame(self.nav_bar, bg=COLORS['BACKGROUND_BORDER'])
        nav_container.pack(pady=SPACING['MEDIUM'], padx=SPACING['LARGE'])
        
        # Store button references for active state management
        self.nav_buttons = {}
        
        # Primary: Basic Report (larger, more prominent)
        self.nav_buttons['basic'] = tk.Button(
            nav_container, 
            text="📊 Basic", 
            font=BUTTON_STYLES['nav_primary_active']['font'],
            bg=BUTTON_STYLES['nav_primary_active']['bg'], 
            fg=BUTTON_STYLES['nav_primary_active']['fg'], 
            relief='flat', 
            bd=0,
            cursor='hand2', 
            command=lambda: self.select_callback('basic'),
            activebackground=COLORS['PRIMARY_BLUE_HOVER'], 
            padx=SPACING['LARGE'], 
            pady=6
        )
        self.nav_buttons['basic'].pack(side='left', padx=(0, SPACING['MEDIUM']))
        
        # Secondary: Other options (smaller)
        self.nav_buttons['detailed'] = tk.Button(
            nav_container, 
            text="📈 Detailed", 
            font=BUTTON_STYLES['nav_secondary_inactive']['font'],
            bg=BUTTON_STYLES['nav_secondary_inactive']['bg'], 
            fg=BUTTON_STYLES['nav_secondary_inactive']['fg'], 
            relief='flat', 
            bd=1,
            cursor='hand2', 
            command=lambda: self.select_callback('detailed'),
            activebackground=COLORS['BACKGROUND_LIGHT'], 
            padx=10, 
            pady=4
        )
        self.nav_buttons['detailed'].pack(side='left', padx=(0, 5))
        
        self.nav_buttons['template'] = tk.Button(
            nav_container, 
            text="📋 Template", 
            font=BUTTON_STYLES['nav_secondary_inactive']['font'],
            bg=BUTTON_STYLES['nav_secondary_inactive']['bg'], 
            fg=BUTTON_STYLES['nav_secondary_inactive']['fg'], 
            relief='flat', 
            bd=1,
            cursor='hand2', 
            command=lambda: self.select_callback('template'),
            activebackground=COLORS['BACKGROUND_LIGHT'], 
            padx=10, 
            pady=4
        )
        self.nav_buttons['template'].pack(side='left')
        
        return self.nav_bar
    
    def update_active_state(self, active_card):
        """Update navigation button visual states"""
        if not self.nav_buttons:
            return
            
        # Reset all buttons to inactive state
        for card_name, button in self.nav_buttons.items():
            if card_name == 'basic':
                if card_name == active_card:
                    # Active basic button
                    style = BUTTON_STYLES['nav_primary_active']
                else:
                    # Inactive basic button  
                    style = BUTTON_STYLES['nav_primary_inactive']
            else:
                if card_name == active_card:
                    # Active secondary button
                    style = BUTTON_STYLES['nav_secondary_active']
                else:
                    # Inactive secondary button
                    style = BUTTON_STYLES['nav_secondary_inactive']
            
            button.configure(
                bg=style['bg'], 
                fg=style['fg'], 
                font=style['font']
            )
    
    def show(self):
        """Show the navigation bar"""
        if self.nav_bar:
            self.nav_bar.grid()
    
    def hide(self):
        """Hide the navigation bar"""
        if self.nav_bar:
            self.nav_bar.grid_remove()