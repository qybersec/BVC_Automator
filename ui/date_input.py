"""
Date Input Component
Handles date range selection for template generation with calendar widgets.
"""

import tkinter as tk
import tkinter.ttk as ttk
from datetime import datetime, timedelta
from .styles import COLORS, SPACING, FONTS, DIMENSIONS

# Check for calendar availability
try:
    from tkcalendar import Calendar
    CALENDAR_AVAILABLE = True
except ImportError:
    CALENDAR_AVAILABLE = False

class DateInputComponent:
    """Date input section for template generation"""
    
    def __init__(self, parent, date_change_callbacks=None):
        self.parent = parent
        self.date_change_callbacks = date_change_callbacks or {}
        self.date_section = None
        self.start_calendar = None
        self.end_calendar = None
        self.date_range_entry = None
        self.date_entry = None  # Fallback entry
        
    def create(self):
        """Create the date input UI section for template generation"""
        self.date_section = tk.Frame(self.parent, bg=COLORS['BACKGROUND_GRAY'])
        self.date_section.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5), padx=3)
        self.date_section.columnconfigure(0, weight=1)
        
        # Section header
        header_style = {'background': COLORS['BACKGROUND_GRAY']}
        ttk.Label(
            self.date_section, 
            text="📅 Date Range for Template", 
            style='Header.TLabel', 
            **header_style
        ).grid(row=0, column=0, pady=(5, 3))
        
        date_frame = tk.Frame(self.date_section, bg=COLORS['BACKGROUND_GRAY'])
        date_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5), padx=SPACING['MEDIUM'])
        date_frame.columnconfigure(0, weight=1)
        
        # Date selection with calendar widgets
        if CALENDAR_AVAILABLE:
            self._create_calendar_widgets(date_frame)
        else:
            self._create_fallback_date_entry(date_frame)
        
        # Initially hide the date section
        self.date_section.grid_remove()
        
        return self.date_section
    
    def _create_calendar_widgets(self, parent_frame):
        """Create compact horizontal calendar layout"""
        # Configure parent frame for better alignment
        parent_frame.grid_columnconfigure(0, weight=1)
        parent_frame.grid_rowconfigure(0, weight=1)
        
        # Main horizontal container with improved layout
        main_container = tk.Frame(parent_frame, bg=COLORS['BACKGROUND_WHITE'])
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=0, pady=0)
        main_container.grid_columnconfigure(0, weight=5)  # Calendars get more space
        main_container.grid_columnconfigure(1, weight=3)  # Controls get proportional space
        
        # Left side: Calendar container with enhanced layout
        calendar_section = tk.Frame(main_container, bg=COLORS['BACKGROUND_GRAY'])
        calendar_section.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, SPACING['LARGE']), pady=0)
        calendar_section.grid_columnconfigure(0, weight=1)
        calendar_section.grid_columnconfigure(1, weight=1)
        calendar_section.grid_rowconfigure(0, weight=1)
        
        # Set default dates first
        today = datetime.now()
        default_start = today
        default_end = today + timedelta(days=4)
        
        # FROM Calendar (Left side)
        from_shadow = tk.Frame(calendar_section, bg=COLORS['BACKGROUND_BORDER'])
        from_shadow.grid(row=0, column=0, padx=(SPACING['MEDIUM'], SPACING['SMALL']), pady=SPACING['MEDIUM'], sticky='nsew')
        
        from_frame = tk.Frame(from_shadow, bg=COLORS['BACKGROUND_WHITE'], relief='flat', bd=0)
        from_frame.pack(padx=0, pady=0, fill='both', expand=True)
        
        from_header = tk.Frame(from_frame, bg=COLORS['PRIMARY_BLUE'], height=DIMENSIONS['CALENDAR_HEADER_HEIGHT'])
        from_header.pack(fill='both', expand=True, pady=0)
        from_header.pack_propagate(False)
        
        tk.Label(
            from_header, 
            text="✨ FROM DATE", 
            font=FONTS['HEADER'],
            fg=COLORS['BACKGROUND_WHITE'], 
            bg=COLORS['PRIMARY_BLUE']
        ).pack(pady=SPACING['MEDIUM'])
        
        self.start_calendar = Calendar(
            from_frame,
            selectmode='day',
            year=default_start.year,
            month=default_start.month,
            day=default_start.day,
            background=COLORS['PRIMARY_BLUE'],
            foreground=COLORS['BACKGROUND_WHITE'],
            selectbackground=COLORS['WARNING_YELLOW'],
            selectforeground=COLORS['TEXT_PRIMARY'],
            normalbackground=COLORS['BACKGROUND_WHITE'],
            normalforeground=COLORS['TEXT_PRIMARY'],
            weekendbackground='#ebf8ff',
            weekendforeground='#2b6cb0',
            othermonthforeground=COLORS['TEXT_DISABLED'],
            othermonthbackground=COLORS['BACKGROUND_LIGHT'],
            headersbackground='#bee3f8',
            headersforeground='#1a365d',
            font=FONTS['SMALL'],
            borderwidth=1,
            bordercolor=COLORS['BACKGROUND_BORDER'],
            cursor='hand2'
        )
        self.start_calendar.pack(padx=SPACING['SMALL'], pady=(0, SPACING['SMALL']), fill='both', expand=True)
        
        # Bind calendar events
        if 'start_date_select' in self.date_change_callbacks:
            self.start_calendar.bind('<<CalendarSelected>>', self.date_change_callbacks['start_date_select'])
        
        # TO Calendar (Right side)
        to_shadow = tk.Frame(calendar_section, bg=COLORS['BACKGROUND_BORDER'])
        to_shadow.grid(row=0, column=1, padx=(SPACING['SMALL'], SPACING['MEDIUM']), pady=SPACING['MEDIUM'], sticky='nsew')
        
        to_frame = tk.Frame(to_shadow, bg=COLORS['BACKGROUND_WHITE'], relief='flat', bd=0)
        to_frame.pack(padx=0, pady=0, fill='both', expand=True)
        
        to_header = tk.Frame(to_frame, bg=COLORS['SUCCESS_GREEN'], height=DIMENSIONS['CALENDAR_HEADER_HEIGHT'])
        to_header.pack(fill='both', expand=True, pady=0)
        to_header.pack_propagate(False)
        
        tk.Label(
            to_header, 
            text="🎯 TO DATE", 
            font=FONTS['HEADER'],
            fg=COLORS['BACKGROUND_WHITE'], 
            bg=COLORS['SUCCESS_GREEN']
        ).pack(pady=SPACING['MEDIUM'])
        
        self.end_calendar = Calendar(
            to_frame,
            selectmode='day',
            year=default_end.year,
            month=default_end.month,
            day=default_end.day,
            background=COLORS['SUCCESS_GREEN'],
            foreground=COLORS['BACKGROUND_WHITE'],
            selectbackground=COLORS['WARNING_YELLOW'],
            selectforeground=COLORS['TEXT_PRIMARY'],
            normalbackground=COLORS['BACKGROUND_WHITE'],
            normalforeground=COLORS['TEXT_PRIMARY'],
            weekendbackground='#f0fff4',
            weekendforeground='#22543d',
            othermonthforeground=COLORS['TEXT_DISABLED'],
            othermonthbackground=COLORS['BACKGROUND_LIGHT'],
            headersbackground='#c6f6d5',
            headersforeground='#1a202c',
            font=FONTS['SMALL'],
            borderwidth=1,
            bordercolor=COLORS['BACKGROUND_BORDER'],
            cursor='hand2'
        )
        self.end_calendar.pack(padx=SPACING['SMALL'], pady=(0, SPACING['SMALL']), fill='both', expand=True)
        
        # Bind calendar events
        if 'end_date_select' in self.date_change_callbacks:
            self.end_calendar.bind('<<CalendarSelected>>', self.date_change_callbacks['end_date_select'])
        
        # Right side: Controls and text input
        controls_section = tk.Frame(main_container, bg=COLORS['BACKGROUND_WHITE'])
        controls_section.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 0), pady=0)
        controls_section.grid_columnconfigure(0, weight=1)
        
        # Text input for date range (synced with calendars)
        controls_header = tk.Label(
            controls_section, 
            text="📝 Date Range", 
            font=FONTS['BODY_BOLD'],
            bg=COLORS['BACKGROUND_WHITE'], 
            fg=COLORS['TEXT_PRIMARY']
        )
        controls_header.grid(row=0, column=0, pady=(SPACING['LARGE'], SPACING['SMALL']), sticky='w')
        
        self.date_range_entry = tk.Entry(
            controls_section,
            font=FONTS['BODY'],
            bg=COLORS['BACKGROUND_WHITE'],
            fg=COLORS['TEXT_PRIMARY'],
            bd=1,
            relief='solid',
            width=25
        )
        self.date_range_entry.grid(row=1, column=0, pady=(0, SPACING['MEDIUM']), padx=SPACING['MEDIUM'], sticky=(tk.W, tk.E))
        self.date_range_entry.insert(0, "Select dates")
        
        # Update the entry with default dates
        start_str = default_start.strftime('%m.%d.%y')
        end_str = default_end.strftime('%m.%d.%y')
        self.date_range_entry.delete(0, tk.END)
        self.date_range_entry.insert(0, f"{start_str} - {end_str}")
        
        # Bind entry events
        if 'date_change' in self.date_change_callbacks:
            self.date_range_entry.bind('<KeyRelease>', self.date_change_callbacks['date_change'])
        if 'date_entry_enter' in self.date_change_callbacks:
            self.date_range_entry.bind('<Return>', self.date_change_callbacks['date_entry_enter'])
    
    def _create_fallback_date_entry(self, parent_frame):
        """Create fallback date entry when calendar is not available"""
        fallback_frame = tk.Frame(parent_frame, bg=COLORS['BACKGROUND_WHITE'])
        fallback_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=SPACING['MEDIUM'], pady=SPACING['MEDIUM'])
        fallback_frame.columnconfigure(0, weight=1)
        
        tk.Label(
            fallback_frame, 
            text="📝 Date Range", 
            font=FONTS['BODY_BOLD'],
            bg=COLORS['BACKGROUND_WHITE'], 
            fg=COLORS['TEXT_PRIMARY']
        ).grid(row=0, column=0, pady=(0, SPACING['SMALL']), sticky='w')
        
        self.date_entry = tk.Entry(
            fallback_frame,
            font=FONTS['BODY'],
            bg=COLORS['BACKGROUND_WHITE'],
            fg=COLORS['TEXT_MUTED'],
            bd=1,
            relief='solid',
            width=50
        )
        self.date_entry.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.date_entry.insert(0, "Enter date range (e.g., 08.04.25 - 08.08.25)")
        
        # Bind entry events
        if 'date_change' in self.date_change_callbacks:
            self.date_entry.bind('<KeyRelease>', self.date_change_callbacks['date_change'])
        if 'date_entry_enter' in self.date_change_callbacks:
            self.date_entry.bind('<Return>', self.date_change_callbacks['date_entry_enter'])
    
    def get_date_range_string(self):
        """Get formatted date range string from text box or calendar widgets"""
        if CALENDAR_AVAILABLE and hasattr(self, 'date_range_entry') and self.date_range_entry:
            # Get from the synced text box
            text = self.date_range_entry.get().strip()
            if text and text != "Select dates":
                return text
            # Fallback to calendar if text box is empty
            elif self.start_calendar and self.end_calendar:
                try:
                    start = self.start_calendar.selection_get()
                    end = self.end_calendar.selection_get()
                    start_str = start.strftime('%m.%d.%y')
                    end_str = end.strftime('%m.%d.%y')
                    return f"{start_str} - {end_str}"
                except:
                    return ""
        elif hasattr(self, 'date_entry') and self.date_entry:
            # Fallback to old text entry
            return self.date_entry.get().strip()
        else:
            return ""
    
    def update_date_range_entry(self):
        """Update the date range entry from calendar selections"""
        if CALENDAR_AVAILABLE and self.start_calendar and self.end_calendar and self.date_range_entry:
            try:
                start = self.start_calendar.selection_get()
                end = self.end_calendar.selection_get()
                start_str = start.strftime('%m.%d.%y')
                end_str = end.strftime('%m.%d.%y')
                
                self.date_range_entry.delete(0, tk.END)
                self.date_range_entry.insert(0, f"{start_str} - {end_str}")
            except:
                pass
    
    def show(self):
        """Show the date input section"""
        if self.date_section:
            self.date_section.grid()
    
    def hide(self):
        """Hide the date input section"""
        if self.date_section:
            self.date_section.grid_remove()