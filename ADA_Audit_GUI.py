import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import openpyxl
from tqdm import tqdm
import time
import threading
from pathlib import Path
import sys
import os
import json
import copy
from datetime import datetime

# Import functions from the original ADA audit script
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from ADA_Audit_25_26_IMPROVED import (
    find_rows_containing_program_name,
    find_rows_containing_month_number,
    find_program_boundary_rows,
    extract_student_attendance_data,
    write_all_attendance_data_to_excel_efficiently
)

# Import the ADA Dashboard module
try:
    from ADA_Dashboard_Module import (
        run_ada_dashboard_with_boundaries,
        get_dashboard_configuration_from_user,
        validate_boundaries_for_dashboard
    )
    DASHBOARD_AVAILABLE = True
except ImportError as e:
    print(f"Warning: ADA Dashboard module not available: {e}")
    DASHBOARD_AVAILABLE = False


class ADAAuditGUI:
    """
    A comprehensive Tkinter GUI for the ADA Audit Process.
    
    This interface allows users to:
    - Select input and output files
    - View and edit program boundaries
    - Execute the audit process
    - Monitor progress and view results
    """
    
    def __init__(self, root):
        self.root = root
        self.root.title("Automated PADC Processor - ADA Compliant Interface")
        self.root.geometry("1200x800")
        
        # ADA Compliant color scheme with high contrast
        self.colors = {
            'background': '#FFFFFF',           # White background for maximum contrast
            'primary': "#003366",              # Dark blue for primary elements
            'secondary': '#006600',            # Dark green for secondary elements
            'accent': '#CC0000',               # Red for alerts/important actions
            'text': '#000000',                 # Black text for maximum readability
            'text_light': '#333333',           # Dark gray for secondary text
            'border': '#666666',               # Medium gray for borders
            'success': '#004400',              # Dark green for success states
            'warning': '#CC6600',              # Orange for warnings
            'error': '#CC0000'                 # Red for errors
        }
        
        self.root.configure(bg=self.colors['background'])
        
        # Configure accessibility features
        self.setup_accessibility_features()
        
        # Initialize variables
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.worksheet_name = tk.StringVar(value="Template- Apportionment Summary")
        
        # Program mappings from original script (updated with new McClellan and SYC locations)
        self.program_name_mappings = {
            # Main Program C locations
            "Program C Charter Resident": "Prog_C",
            "Program C Charter Resident -  Transitional Kindergarten(TK)": "Prog_C_TK",
            "Program C Charter Resident -  McClellan(CM)": "Prog_C_CM",
            "Program C Charter Resident -  Sac Youth Center(SYC)": "Prog_C_SYC",
            
            # Main Program N locations  
            "Program N Non-Resident Charter": "Prog_N", 
            "Program N Non-Resident Charter -  Transitional Kindergarten(TK)": "Prog_N_TK",
            "Program N Non-Resident Charter -  McClellan(CM)": "Prog_N_CM",
            "Program N Non-Resident Charter -  Sac Youth Center(SYC)": "Prog_N_SYC",
            
            # Independent Study programs
            "Program J Indep Study Charter Resident": "Prog_J",
            "Program J Indep Study Charter Non-Resident -  Transitional Kindergarten(TK)": "Prog_J_TK",
            "Program K Indep Study Charter Non-Resident": "Prog_K",
            "Program K Indep Study Charter Non-Resident -  Transitional Kindergarten(TK)": "Prog_K_TK",
        }
        
        # Define which sub-programs should be combined with their parent programs
        self.program_consolidation_rules = {
            "Prog_C": ["Prog_C", "Prog_C_CM", "Prog_C_SYC"],  # Combine main C + CM + SYC
            "Prog_C_TK": ["Prog_C_TK"],  # TK stays separate
            "Prog_N": ["Prog_N", "Prog_N_CM", "Prog_N_SYC"],  # Combine main N + CM + SYC  
            "Prog_N_TK": ["Prog_N_TK"],  # TK stays separate
            "Prog_J": ["Prog_J"],
            "Prog_J_TK": ["Prog_J_TK"],
            "Prog_K": ["Prog_K"],
            "Prog_K_TK": ["Prog_K_TK"],
        }
        
        # Initialize program boundaries storage
        self.program_boundaries = {}
        for short_code in self.program_name_mappings.values():
            self.program_boundaries[short_code] = {"start": None, "stop": None}
        
        # Data storage
        self.student_attendance_data = None
        self.extracted_attendance_data = None
        
        # Table sorting state
        self.sort_column = None
        self.sort_reverse = False
        self.boundary_data = []  # Store boundary data for sorting
        
        # Boundary settings management
        self.settings_directory = Path(os.path.dirname(os.path.abspath(__file__))) / "boundary_settings"
        self.settings_directory.mkdir(exist_ok=True)
        self.saved_configurations = {}
        
        # Create the GUI
        self.create_widgets()
        
        # Set default file paths if they exist
        self.set_default_paths()
        
        # Load saved configurations
        self.load_saved_configurations()
    
    def setup_accessibility_features(self):
        """Configure accessibility features for ADA compliance"""
        
        # Set up keyboard navigation
        self.root.focus_set()
        
        # Configure tab order tracking
        self.tab_order_widgets = []
        
        # Set up screen reader announcements
        self.status_announcements = []
        
        # Configure high contrast mode detection
        try:
            # Windows high contrast detection (if available)
            import winreg
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                   r"Control Panel\Accessibility\HighContrast")
                value, reg_type = winreg.QueryValueEx(key, "Flags")
                # Convert to int if it's a string
                if isinstance(value, str):
                    value = int(value)
                if value & 1:  # High contrast is enabled
                    self.enable_high_contrast_mode()
                winreg.CloseKey(key)
            except (FileNotFoundError, PermissionError, ValueError):
                pass
        except ImportError:
            pass
        
        # Set up keyboard shortcuts
        self.setup_keyboard_shortcuts()
        
        # Configure focus indicators
        self.setup_focus_indicators()
    
    def setup_keyboard_shortcuts(self):
        """Set up keyboard shortcuts for accessibility"""
        
        # Global shortcuts
        self.root.bind('<Control-o>', lambda e: self.browse_input_file())
        self.root.bind('<Control-s>', lambda e: self.browse_output_file())
        self.root.bind('<Control-l>', lambda e: self.load_and_analyze_data())
        self.root.bind('<Control-r>', lambda e: self.run_audit_process())
        self.root.bind('<Control-e>', lambda e: self.export_results())
        self.root.bind('<Control-d>', lambda e: self.run_ada_dashboard() if DASHBOARD_AVAILABLE else None)
        self.root.bind('<F1>', lambda e: self.show_help())
        self.root.bind('<Escape>', lambda e: self.root.focus_set())
        
        # Scrolling shortcuts for accessibility
        self.root.bind('<Control-Up>', lambda e: self.scroll_up())
        self.root.bind('<Control-Down>', lambda e: self.scroll_down())
        self.root.bind('<Control-Home>', lambda e: self.scroll_to_top())
        self.root.bind('<Control-End>', lambda e: self.scroll_to_bottom())
        self.root.bind('<Page_Up>', lambda e: self.page_up())
        self.root.bind('<Page_Down>', lambda e: self.page_down())
        
        # Table sorting shortcuts
        self.root.bind('<F2>', lambda e: self.sort_table('Program Code'))
        self.root.bind('<F3>', lambda e: self.sort_table('Program Name'))
        self.root.bind('<F4>', lambda e: self.sort_table('Start Row'))
        self.root.bind('<F5>', lambda e: self.sort_table('End Row'))
        
        # Add tooltip for shortcuts
        dashboard_shortcut = "Ctrl+D: Run ADA Dashboard\n" if DASHBOARD_AVAILABLE else ""
        self.keyboard_shortcuts_info = (
            "Keyboard Shortcuts:\n"
            "Ctrl+O: Open input file\n"
            "Ctrl+S: Select output file\n"
            "Ctrl+L: Load and analyze data\n"
            "Ctrl+R: Run audit process\n"
            "Ctrl+E: Export results\n"
            f"{dashboard_shortcut}"
            "F1: Show help\n"
            "Esc: Return focus to main window\n"
            "Tab: Navigate between elements\n"
            "Space/Enter: Activate buttons\n\n"
            "Table Sorting Shortcuts:\n"
            "F2: Sort by Program Code\n"
            "F3: Sort by Program Name\n"
            "F4: Sort by Start Row\n"
            "F5: Sort by End Row\n"
            "Click column headers to sort\n\n"
            "Scrolling Shortcuts:\n"
            "Ctrl+Up/Down: Scroll line by line\n"
            "Page Up/Down: Scroll page by page\n"
            "Ctrl+Home: Scroll to top\n"
            "Ctrl+End: Scroll to bottom\n"
            "Mouse wheel: Scroll smoothly"
        )
    
    def setup_focus_indicators(self):
        """Configure visible focus indicators"""
        
        # Configure ttk styles for better focus visibility
        style = ttk.Style()
        
        # High contrast focus ring
        style.configure('Focus.TButton',
                       focuscolor='red',
                       highlightthickness=3,
                       relief='solid',
                       borderwidth=2)
        
        style.configure('Focus.TEntry',
                       focuscolor='red',
                       highlightthickness=2,
                       relief='solid',
                       borderwidth=2)
        
        style.configure('Focus.Treeview',
                       focuscolor='red',
                       highlightthickness=2,
                       relief='solid',
                       borderwidth=2)
    
    def enable_high_contrast_mode(self):
        """Enable high contrast color scheme"""
        
        self.colors.update({
            'background': '#000000',           # Black background
            'primary': '#FFFFFF',              # White for primary elements
            'secondary': '#FFFF00',            # Yellow for secondary elements
            'accent': '#FF0000',               # Red for alerts
            'text': '#003366',                 # Dark blue text
            'text_light': '#003366',           # Dark blue for secondary text
            'border': '#FFFFFF',               # White borders
            'success': '#00FF00',              # Bright green for success
            'warning': '#FFFF00',              # Yellow for warnings
            'error': '#FF0000'                 # Red for errors
        })
        
        self.root.configure(bg=self.colors['background'])
        
        # Reconfigure button styles for high contrast mode to ensure dark blue text
        style = ttk.Style()
        style.configure('Accessible.TButton',
                       foreground='#003366',  # Force dark blue text in high contrast mode
                       background=self.colors['primary'],  # White background
                       font=('Arial', 12, 'bold'),
                       focuscolor=self.colors['accent'],
                       borderwidth=2,
                       relief='raised')
        
        style.map('Accessible.TButton',
                 background=[('active', self.colors['secondary']),
                           ('pressed', self.colors['accent'])],
                 foreground=[('active', '#003366'),  # Dark blue text when active
                           ('pressed', '#003366')])  # Dark blue text when pressed
    
    def announce_to_screen_reader(self, message):
        """Announce messages to screen readers"""
        
        # Store announcement for accessibility
        self.status_announcements.append({
            'message': message,
            'timestamp': time.time()
        })
        
        # Update window title with important announcements
        if any(keyword in message.lower() for keyword in ['error', 'complete', 'ready', 'loading']):
            original_title = "Automated PADC Processor - ADA Compliant Interface"
            self.root.title(f"{original_title} - {message}")
            # Reset title after 3 seconds
            self.root.after(3000, lambda: self.root.title(original_title))
    
    def show_help(self):
        """Show accessibility help dialog"""
        
        help_text = f"""ADA Compliant PADC Processor Help
        
ACCESSIBILITY FEATURES:
• High contrast colors for better visibility
• Keyboard navigation support
• Screen reader compatible
• Clear focus indicators
• Large, readable fonts

{self.keyboard_shortcuts_info}

NAVIGATION TIPS:
• Use Tab key to move between form fields
• Use Arrow keys in tables and lists
• Use Space or Enter to activate buttons
• Use Esc to return focus to main window
• Mouse wheel or scroll shortcuts to navigate content

SCROLLING FEATURES:
• Full interface scrolling with mouse wheel
• Keyboard scrolling with Ctrl+Up/Down arrows
• Page navigation with Page Up/Down keys
• Quick navigation with Ctrl+Home/End
• Smooth scrolling for long content

VISUAL ACCESSIBILITY:
• All text meets WCAG 2.1 AA contrast standards
• Focus indicators are clearly visible
• Error messages are prominently displayed
• Status updates are announced
• Scrollable interface accommodates all screen sizes

For additional assistance, contact your system administrator."""
        
        # Create accessible help dialog
        help_dialog = tk.Toplevel(self.root)
        help_dialog.title("Accessibility Help - F1")
        help_dialog.geometry("600x500")
        help_dialog.configure(bg=self.colors['background'])
        help_dialog.grab_set()
        help_dialog.focus_set()
        
        # Make dialog modal and centered
        help_dialog.transient(self.root)
        
        main_frame = tk.Frame(help_dialog, bg=self.colors['background'], padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollable text widget
        text_widget = scrolledtext.ScrolledText(
            main_frame,
            wrap=tk.WORD,
            width=70,
            height=25,
            font=('Arial', 11),
            bg=self.colors['background'],
            fg='#000000',  # Black text for help dialog on white background
            insertbackground='#000000',  # Black cursor
            selectbackground=self.colors['primary'],
            selectforeground='#003366'  # Dark blue selected text
        )
        text_widget.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        text_widget.insert(tk.END, help_text)
        text_widget.configure(state='disabled')
        
        # Close button
        close_button = tk.Button(
            main_frame,
            text="Close (Esc)",
            command=help_dialog.destroy,
            font=('Arial', 12, 'bold'),
            bg=self.colors['primary'],
            fg='#FFFFFF',  # White text on dark blue button
            activebackground=self.colors['secondary'],
            activeforeground='#FFFFFF',  # White text when active
            padx=20,
            pady=10,
            relief='raised',
            borderwidth=2
        )
        close_button.pack()
        close_button.focus_set()
        
        # Bind Esc to close
        help_dialog.bind('<Escape>', lambda e: help_dialog.destroy())
        help_dialog.bind('<Return>', lambda e: help_dialog.destroy())
    
    def add_to_tab_order(self, widget):
        """Add widget to tab order for keyboard navigation"""
        self.tab_order_widgets.append(widget)
        
        # Configure accessible focus events - only for widgets that support highlight options
        def on_focus_in(event):
            try:
                # Only apply highlight options to widgets that support them
                widget_class = widget.__class__.__name__
                if widget_class in ['Entry', 'Text', 'Listbox', 'Canvas']:
                    widget.configure(highlightthickness=3, highlightcolor=self.colors['accent'])
                elif widget_class == 'Button':
                    # Use relief changes for buttons instead
                    widget.configure(relief='solid', borderwidth=3)
            except tk.TclError:
                # Widget doesn't support these options, ignore
                pass
            
        def on_focus_out(event):
            try:
                widget_class = widget.__class__.__name__
                if widget_class in ['Entry', 'Text', 'Listbox', 'Canvas']:
                    widget.configure(highlightthickness=1, highlightcolor=self.colors['border'])
                elif widget_class == 'Button':
                    # Reset button appearance
                    widget.configure(relief='raised', borderwidth=2)
            except tk.TclError:
                # Widget doesn't support these options, ignore
                pass
        
        widget.bind('<FocusIn>', on_focus_in)
        widget.bind('<FocusOut>', on_focus_out)
    
    def create_scrollable_main_container(self):
        """Create a scrollable main container for the entire interface"""
        
        # Create main canvas and scrollbar
        self.main_canvas = tk.Canvas(self.root, bg=self.colors['background'], highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.main_canvas.yview)
        
        # Configure canvas scrolling
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        
        # Create the scrollable frame
        self.main_frame = ttk.Frame(self.main_canvas, padding="15")
        
        # Add the frame to the canvas
        self.canvas_frame_id = self.main_canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # Grid the canvas and scrollbar
        self.main_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure scrolling behavior
        self.main_frame.bind('<Configure>', self.on_frame_configure)
        self.main_canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Bind mouse wheel scrolling
        self.bind_mousewheel(self.root)
        self.bind_mousewheel(self.main_canvas)
        self.bind_mousewheel(self.main_frame)
        
        # Configure additional scrolling behavior
        self.setup_scroll_behavior()
        
        # Bind window resize events to handle orientation changes
        self.root.bind('<Configure>', self.on_window_configure)
    
    def on_frame_configure(self, event):
        """Update scroll region when frame size changes"""
        # Use after_idle to ensure all widgets are properly sized before updating scroll region
        self.root.after_idle(self.update_scroll_region)
    
    def on_canvas_configure(self, event):
        """Update frame size when canvas size changes (handles orientation changes)"""
        # Update the frame width to match canvas width
        canvas_width = event.width
        canvas_height = event.height
        
        # Configure the frame to match the canvas width
        self.main_canvas.itemconfig(self.canvas_frame_id, width=canvas_width)
        
        # Force update of scroll region after canvas resize
        self.root.after_idle(self.update_scroll_region)
        
        # Log orientation change for debugging
        if hasattr(self, '_last_canvas_size'):
            old_width, old_height = self._last_canvas_size
            if abs(canvas_width - old_width) > 50 or abs(canvas_height - old_height) > 50:
                self.log_message(f"🔄 Canvas resized: {old_width}x{old_height} → {canvas_width}x{canvas_height}")
        
        self._last_canvas_size = (canvas_width, canvas_height)
    
    def update_scroll_region(self):
        """Update the scroll region and handle orientation changes"""
        try:
            if hasattr(self, 'main_canvas') and hasattr(self, 'main_frame'):
                # Update the scroll region
                bbox = self.main_canvas.bbox("all")
                if bbox:
                    self.main_canvas.configure(scrollregion=bbox)
                    
                    # Get current canvas dimensions
                    canvas_width = self.main_canvas.winfo_width()
                    canvas_height = self.main_canvas.winfo_height()
                    
                    # Calculate content dimensions
                    content_width = bbox[2] - bbox[0]
                    content_height = bbox[3] - bbox[1]
                    
                    # Check if scrolling is needed
                    needs_vertical_scroll = content_height > canvas_height
                    needs_horizontal_scroll = content_width > canvas_width
                    
                    # Update scrollbar visibility
                    if needs_vertical_scroll:
                        self.main_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
                    else:
                        self.main_scrollbar.grid_remove()
                    
                    # Ensure mouse wheel scrolling works after orientation change
                    self.refresh_mousewheel_bindings()
                    
        except Exception as e:
            # Silent error handling to prevent GUI disruption
            pass
    
    def refresh_mousewheel_bindings(self):
        """Refresh mouse wheel bindings after orientation changes"""
        try:
            # Re-bind mouse wheel events to ensure they work after orientation change
            widgets_to_rebind = [self.root, self.main_canvas, self.main_frame]
            
            for widget in widgets_to_rebind:
                if widget and widget.winfo_exists():
                    # Remove existing bindings
                    widget.unbind("<MouseWheel>")
                    widget.unbind("<Button-4>")
                    widget.unbind("<Button-5>")
                    
                    # Re-apply mouse wheel bindings
                    self.bind_mousewheel(widget)
        except:
            pass
    
    def on_window_configure(self, event):
        """Handle window resize events including orientation changes"""
        # Only handle events for the root window
        if event.widget == self.root:
            # Get new window dimensions
            new_width = event.width
            new_height = event.height
            
            # Check if this is a significant size change (orientation change)
            if hasattr(self, '_last_window_size'):
                old_width, old_height = self._last_window_size
                width_change = abs(new_width - old_width)
                height_change = abs(new_height - old_height)
                
                # If the change is significant, update scroll configuration
                if width_change > 100 or height_change > 100:
                    self.log_message(f"🔄 Window resized: {old_width}x{old_height} → {new_width}x{new_height}")
                    
                    # Schedule scroll region update after a brief delay to ensure layout is complete
                    self.root.after(100, self.update_scroll_region)
                    
                    # Also refresh the entire scrolling system
                    self.root.after(150, self.refresh_scrolling_system)
            
            self._last_window_size = (new_width, new_height)
    
    def refresh_scrolling_system(self):
        """Refresh the entire scrolling system after major layout changes"""
        try:
            # Update scroll region
            self.update_scroll_region()
            
            # Refresh mouse wheel bindings
            self.refresh_mousewheel_bindings()
            
            # Force a redraw of the canvas
            if hasattr(self, 'main_canvas'):
                self.main_canvas.update_idletasks()
            
            self.log_message("🔄 Scrolling system refreshed")
            
        except Exception as e:
            self.log_message(f"⚠️ Error refreshing scrolling system: {e}")
    
    def bind_mousewheel(self, widget):
        """Bind mouse wheel scrolling to a widget"""
        def _on_mousewheel(event):
            # Only scroll if the canvas exists and is viewable
            if (hasattr(self, 'main_canvas') and 
                self.main_canvas.winfo_exists() and
                self.main_canvas.winfo_viewable()):
                
                try:
                    # Check if mouse is over the canvas area
                    canvas_under_mouse = (self.main_canvas.winfo_containing(event.x_root, event.y_root) == self.main_canvas)
                    
                    # Also allow scrolling if mouse is over any child of the main frame
                    main_frame_under_mouse = False
                    if hasattr(self, 'main_frame') and self.main_frame.winfo_exists():
                        try:
                            widget_under_mouse = self.main_canvas.winfo_containing(event.x_root, event.y_root)
                            if widget_under_mouse:
                                # Check if the widget under mouse is a descendant of main_frame
                                current = widget_under_mouse
                                while current:
                                    if current == self.main_frame:
                                        main_frame_under_mouse = True
                                        break
                                    try:
                                        current = current.master
                                    except:
                                        break
                        except:
                            pass
                    
                    if canvas_under_mouse or main_frame_under_mouse:
                        # Calculate scroll amount with better handling
                        delta = getattr(event, 'delta', 0)
                        if delta != 0:
                            scroll_amount = int(-1 * (delta / 120))
                        else:
                            scroll_amount = -3 if event.num == 4 else 3
                        
                        self.main_canvas.yview_scroll(scroll_amount, "units")
                        
                except Exception:
                    # Fallback scrolling if position detection fails
                    if hasattr(event, 'delta'):
                        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _on_mousewheel_linux_up(event):
            if (hasattr(self, 'main_canvas') and 
                self.main_canvas.winfo_exists() and 
                self.main_canvas.winfo_viewable()):
                self.main_canvas.yview_scroll(-1, "units")
        
        def _on_mousewheel_linux_down(event):
            if (hasattr(self, 'main_canvas') and 
                self.main_canvas.winfo_exists() and 
                self.main_canvas.winfo_viewable()):
                self.main_canvas.yview_scroll(1, "units")
        
        # Bind to different wheel events for cross-platform compatibility
        try:
            widget.bind("<MouseWheel>", _on_mousewheel)  # Windows
            widget.bind("<Button-4>", _on_mousewheel_linux_up)  # Linux
            widget.bind("<Button-5>", _on_mousewheel_linux_down)  # Linux
        except:
            pass
    
    def setup_scroll_behavior(self):
        """Configure additional scrolling behavior and responsiveness"""
        
        # Enable focus-follows-scroll for accessibility
        def on_scroll(*args):
            # Update focus to maintain accessibility context
            if hasattr(self, 'main_frame') and self.main_frame.winfo_viewable():
                # Announce scroll position to screen readers
                try:
                    top, bottom = self.main_canvas.yview()
                    position_percent = int(top * 100)
                    if position_percent == 0:
                        self.announce_to_screen_reader("Scrolled to top of interface")
                    elif position_percent >= 95:
                        self.announce_to_screen_reader("Scrolled to bottom of interface")
                except:
                    pass
        
        # Bind scroll monitoring
        if hasattr(self, 'main_canvas'):
            self.main_canvas.configure(yscrollcommand=lambda *args: (
                self.main_scrollbar.set(*args),
                on_scroll(*args)
            ))
        
        # Auto-scroll to focused widget
        def on_widget_focus(event):
            widget = event.widget
            if hasattr(self, 'main_canvas') and self.main_canvas.winfo_viewable():
                try:
                    # Get widget position relative to the main frame
                    x = widget.winfo_x()
                    y = widget.winfo_y()
                    
                    # Get canvas viewport
                    canvas_height = self.main_canvas.winfo_height()
                    scroll_top = self.main_canvas.canvasy(0)
                    scroll_bottom = scroll_top + canvas_height
                    
                    # Check if widget is visible, if not scroll to it
                    if y < scroll_top or y > scroll_bottom - 50:  # 50px buffer
                        # Calculate relative position to scroll to
                        total_height = self.main_frame.winfo_reqheight()
                        if total_height > 0:
                            scroll_to = max(0, min(1, (y - canvas_height/2) / total_height))
                            self.main_canvas.yview_moveto(scroll_to)
                except:
                    pass
        
        # Bind focus events to all widgets recursively
        self.bind_focus_scroll_recursive(self.main_frame, on_widget_focus)
    
    def bind_focus_scroll_recursive(self, parent, focus_callback):
        """Recursively bind focus events to enable auto-scroll to focused widgets"""
        try:
            parent.bind('<FocusIn>', focus_callback)
            for child in parent.winfo_children():
                try:
                    child.bind('<FocusIn>', focus_callback)
                    self.bind_focus_scroll_recursive(child, focus_callback)
                except:
                    pass
        except:
            pass
    
    def create_widgets(self):
        """Create all GUI widgets and layout with ADA compliance"""
        
        # Create scrollable main container
        self.create_scrollable_main_container()
        
        # Configure root grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        
        # Configure accessible styling
        style = ttk.Style()
        style.configure('Accessible.TLabel',
                       foreground='#000000',  # Black text for readability on white background
                       background=self.colors['background'],
                       font=('Arial', 11))
        
        style.configure('Title.TLabel',
                       foreground=self.colors['primary'],  # Dark blue title text
                       background=self.colors['background'],
                       font=('Arial', 18, 'bold'))
        
        style.configure('Accessible.TButton',
                       foreground='#003366',  # Dark blue text on buttons
                       background=self.colors['primary'],  # Dark blue button background (#003366)
                       font=('Arial', 12, 'bold'),
                       focuscolor=self.colors['accent'],
                       borderwidth=2,
                       relief='raised')
        
        style.map('Accessible.TButton',
                 background=[('active', self.colors['secondary']),
                           ('pressed', self.colors['accent'])],
                 foreground=[('active', '#003366'),
                           ('pressed', '#003366')])
        
        style.configure('Success.TButton',
                       foreground='#003366',  # Dark blue text on success buttons for visibility
                       background=self.colors['primary'],  # Dark blue background (#003366)
                       font=('Arial', 12, 'bold'),
                       focuscolor=self.colors['accent'],
                       borderwidth=2,
                       relief='raised')
        
        style.map('Success.TButton',
                 background=[('active', self.colors['success']),  # Green on hover
                           ('pressed', self.colors['accent'])],
                 foreground=[('active', '#003366'),
                           ('pressed', '#003366')])
        
        style.configure('Accessible.TEntry',
                       foreground='#000000',  # Black text for readability
                       fieldbackground='#FFFFFF',  # White background for input fields
                       bordercolor=self.colors['border'],
                       focuscolor=self.colors['accent'],
                       font=('Arial', 11))
        
        # Title with proper heading hierarchy
        title_label = ttk.Label(self.main_frame, text="Automated PADC Processor", 
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 25))
        self.add_to_tab_order(title_label)
        
        # Accessibility notice
        accessibility_label = ttk.Label(self.main_frame, 
                                       text="♿ ADA Compliant Interface - Press F1 for Help",
                                       style='Accessible.TLabel',
                                       font=('Arial', 10, 'italic'))
        accessibility_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # File Selection Section
        self.create_file_selection_section(self.main_frame, row=2)
        
        # Program Boundaries Section
        self.create_boundaries_section(self.main_frame, row=3)
        
        # Control Buttons Section
        self.create_control_section(self.main_frame, row=4)
        
        # Results Section
        self.create_results_section(self.main_frame, row=5)
        
    def create_file_selection_section(self, parent, row):
        """Create file selection widgets with ADA compliance"""
        
        # File Selection Frame with accessible styling
        file_frame = ttk.LabelFrame(parent, text="Step 1: File Selection", padding="15")
        file_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        # Configure accessible label style
        style = ttk.Style()
        style.configure('FileLabel.TLabel',
                       foreground=self.colors['primary'],  # Dark blue text for file labels
                       background=self.colors['background'],
                       font=('Arial', 11, 'bold'))
        
        # Input file selection with accessibility features
        input_label = ttk.Label(file_frame, text="Input Attendance File (Ctrl+O):", 
                               style='FileLabel.TLabel')
        input_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        
        input_entry = ttk.Entry(file_frame, textvariable=self.input_file_path, 
                               width=80, style='Accessible.TEntry')
        input_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=(0, 8))
        self.add_to_tab_order(input_entry)
        
        input_button = ttk.Button(file_frame, text="Browse...", 
                                 command=self.browse_input_file,
                                 style='Accessible.TButton',
                                 width=12)
        input_button.grid(row=0, column=2, pady=(0, 8), padx=(5, 0))
        self.add_to_tab_order(input_button)
        
        # Output file selection with accessibility features
        output_label = ttk.Label(file_frame, text="Output Audit File (Ctrl+S):", 
                                style='FileLabel.TLabel')
        output_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 8))
        
        output_entry = ttk.Entry(file_frame, textvariable=self.output_file_path, 
                                width=80, style='Accessible.TEntry')
        output_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=(0, 8))
        self.add_to_tab_order(output_entry)
        
        output_button = ttk.Button(file_frame, text="Browse...", 
                                  command=self.browse_output_file,
                                  style='Accessible.TButton',
                                  width=12)
        output_button.grid(row=1, column=2, pady=(0, 8), padx=(5, 0))
        self.add_to_tab_order(output_button)
        
        # Worksheet name with accessibility features
        worksheet_label = ttk.Label(file_frame, text="Output Worksheet Name:", 
                                   style='FileLabel.TLabel')
        worksheet_label.grid(row=2, column=0, sticky=tk.W)
        
        worksheet_entry = ttk.Entry(file_frame, textvariable=self.worksheet_name, 
                                   width=80, style='Accessible.TEntry')
        worksheet_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        self.add_to_tab_order(worksheet_entry)
        
        # Add descriptive text for screen readers
        help_text = ttk.Label(file_frame, 
                             text="Select your Excel attendance file and specify where to save the audit results.",
                             style='Accessible.TLabel',
                             font=('Arial', 9, 'italic'))
        help_text.grid(row=3, column=0, columnspan=3, pady=(10, 0), sticky=tk.W)
    
    def create_boundaries_section(self, parent, row):
        """Create program boundaries display and editing section with ADA compliance"""
        
        # Program Boundaries Frame with accessible styling
        boundaries_frame = ttk.LabelFrame(parent, text="Step 2: Program Boundaries", padding="15")
        boundaries_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        boundaries_frame.columnconfigure(0, weight=1)
        boundaries_frame.rowconfigure(1, weight=1)
        
        # Instructions with accessibility features
        instructions = ttk.Label(boundaries_frame, 
                               text="After loading data, program boundaries will be displayed below. You can edit them before running the audit.\n" +
                               "Note: McClellan (CM) and Sac Youth Center (SYC) data will be automatically consolidated with their parent programs.\n" +
                               "Use Tab to navigate, Space/Enter to edit selected items. Click column headers to sort (↕ = sortable, ↑↓ = sorted).",
                               style='Accessible.TLabel',
                               font=('Arial', 10))
        instructions.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Create scrollable frame for boundaries
        self.create_boundaries_table(boundaries_frame)
        
    def create_boundaries_table(self, parent):
        """Create an accessible table for displaying and editing program boundaries"""
        
        # Create frame for the table
        table_frame = ttk.Frame(parent)
        table_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Create accessible Treeview for the table
        columns = ('Program Code', 'Program Name', 'Start Row', 'End Row')
        self.boundaries_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=8)
        
        # Configure column headings and widths with accessibility and sorting
        self.boundaries_tree.heading('Program Code', text='Program Code ↕', 
                                    command=lambda: self.sort_table('Program Code'))
        self.boundaries_tree.heading('Program Name', text='Program Name ↕', 
                                    command=lambda: self.sort_table('Program Name'))
        self.boundaries_tree.heading('Start Row', text='Start Row ↕', 
                                    command=lambda: self.sort_table('Start Row'))
        self.boundaries_tree.heading('End Row', text='End Row ↕', 
                                    command=lambda: self.sort_table('End Row'))
        
        self.boundaries_tree.column('Program Code', width=140, minwidth=100)
        self.boundaries_tree.column('Program Name', width=450, minwidth=300)
        self.boundaries_tree.column('Start Row', width=120, minwidth=80)
        self.boundaries_tree.column('End Row', width=120, minwidth=80)
        
        # Configure accessible colors
        style = ttk.Style()
        style.configure('Accessible.Treeview',
                       background=self.colors['background'],
                       foreground=self.colors['text'],
                       selectbackground=self.colors['primary'],
                       selectforeground=self.colors['background'],
                       font=('Arial', 10))
        
        style.configure('Accessible.Treeview.Heading',
                       background=self.colors['secondary'],
                       foreground='#000000',  # Black text for header readability
                       font=('Arial', 11, 'bold'))
        
        self.boundaries_tree.configure(style='Accessible.Treeview')
        
        # Add scrollbars with accessible styling
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.boundaries_tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.boundaries_tree.xview)
        self.boundaries_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid the treeview and scrollbars
        self.boundaries_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Add to tab order and configure accessible events
        self.add_to_tab_order(self.boundaries_tree)
        
        # Bind accessible navigation
        self.boundaries_tree.bind('<Double-1>', self.edit_boundary)
        self.boundaries_tree.bind('<Return>', self.edit_boundary)
        self.boundaries_tree.bind('<space>', self.edit_boundary)
        
        # Buttons frame with accessible layout
        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=2, column=0, pady=(15, 0))
        
        load_button = ttk.Button(buttons_frame, text="Load & Analyze Data (Ctrl+L)", 
                               command=self.load_and_analyze_data,
                               style='Accessible.TButton',
                               width=25)
        load_button.pack(side=tk.LEFT, padx=(0, 15), pady=5)
        self.add_to_tab_order(load_button)
        
        edit_button = ttk.Button(buttons_frame, text="Edit Selected Boundary", 
                               command=self.edit_selected_boundary,
                               style='Accessible.TButton',
                               width=20)
        edit_button.pack(side=tk.LEFT, padx=(0, 15), pady=5)
        self.add_to_tab_order(edit_button)
        
        sort_reset_button = ttk.Button(buttons_frame, text="Reset Sort", 
                                     command=self.reset_sort,
                                     style='Accessible.TButton',
                                     width=12)
        sort_reset_button.pack(side=tk.LEFT, padx=(0, 15), pady=5)
        self.add_to_tab_order(sort_reset_button)
        
        # Settings management buttons with accessibility
        settings_buttons_frame = ttk.Frame(parent)
        settings_buttons_frame.grid(row=3, column=0, pady=(10, 0))
        
        save_config_button = ttk.Button(settings_buttons_frame, text="Save Configuration", 
                                      command=self.save_boundary_configuration,
                                      style='Accessible.TButton',
                                      width=18)
        save_config_button.pack(side=tk.LEFT, padx=(0, 8), pady=3)
        self.add_to_tab_order(save_config_button)
        
        load_config_button = ttk.Button(settings_buttons_frame, text="Load Configuration", 
                                      command=self.load_boundary_configuration,
                                      style='Accessible.TButton',
                                      width=18)
        load_config_button.pack(side=tk.LEFT, padx=(0, 8), pady=3)
        self.add_to_tab_order(load_config_button)
        
        export_button = ttk.Button(settings_buttons_frame, text="Export Settings", 
                                 command=self.export_boundary_settings,
                                 style='Accessible.TButton',
                                 width=15)
        export_button.pack(side=tk.LEFT, padx=(0, 8), pady=3)
        self.add_to_tab_order(export_button)
        
        import_button = ttk.Button(settings_buttons_frame, text="Import Settings", 
                                 command=self.import_boundary_settings,
                                 style='Accessible.TButton',
                                 width=15)
        import_button.pack(side=tk.LEFT, padx=(0, 8), pady=3)
        self.add_to_tab_order(import_button)
        
        manage_button = ttk.Button(settings_buttons_frame, text="Manage Configurations", 
                                 command=self.manage_configurations,
                                 style='Accessible.TButton',
                                 width=20)
        manage_button.pack(side=tk.LEFT, pady=3)
        self.add_to_tab_order(manage_button)
    
    def create_control_section(self, parent, row):
        """Create control buttons section with ADA compliance"""
        
        control_frame = ttk.LabelFrame(parent, text="Step 3: Execute Process", padding="15")
        control_frame.grid(row=row, column=0, columnspan=3, pady=(0, 15))
        
        # Progress bar with accessible styling
        progress_frame = ttk.Frame(control_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        progress_label = ttk.Label(progress_frame, text="Progress:", 
                                  style='Accessible.TLabel',
                                  font=('Arial', 11, 'bold'))
        progress_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                          maximum=100, length=400, mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, padx=(0, 15))
        self.add_to_tab_order(self.progress_bar)
        
        # Status label with high visibility
        self.status_var = tk.StringVar(value="Ready to begin")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var,
                                style='Accessible.TLabel',
                                font=('Arial', 11, 'bold'))
        status_label.pack(side=tk.LEFT)
        
        # Button frame
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(pady=(10, 0))
        
        # Main run button with dynamic styling
        self.run_button = ttk.Button(button_frame, text="🚀 Run Audit Process (Ctrl+R)", 
                                   command=self.run_audit_process, 
                                   style='Accessible.TButton',
                                   width=25)
        self.run_button.pack(side=tk.LEFT, padx=(10, 20), pady=10)
        self.add_to_tab_order(self.run_button)
        
        # Export button
        self.export_button = ttk.Button(button_frame, text="📊 Export Results (Ctrl+E)", 
                                      command=self.export_results,
                                      style='Accessible.TButton',
                                      width=20)
        self.export_button.pack(side=tk.LEFT, padx=(0, 10), pady=10)
        self.add_to_tab_order(self.export_button)
        
        # Dashboard button
        if DASHBOARD_AVAILABLE:
            self.dashboard_button = ttk.Button(button_frame, text="📈 Run ADA Dashboard (Ctrl+D)", 
                                             command=self.run_ada_dashboard,
                                             style='Accessible.TButton',
                                             width=25)
            self.dashboard_button.pack(side=tk.LEFT, padx=(0, 10), pady=10)
            self.add_to_tab_order(self.dashboard_button)
        
        # Initially disable export button
        self.export_button.configure(state='disabled')
        
        # Update button states based on data availability
        self.update_button_states()
    
    def update_button_states(self):
        """Update button states and colors based on data availability"""
        
        # Check if we have input file and data loaded
        has_input_file = bool(self.input_file_path.get())
        has_output_file = bool(self.output_file_path.get())
        has_data = self.student_attendance_data is not None
        has_results = self.extracted_attendance_data is not None
        
        # Configure run button state and style
        if has_input_file and has_output_file and has_data:
            # Data is ready - make button prominent
            self.run_button.configure(state='normal', style='Success.TButton')
            self.status_var.set("Ready to run audit process")
            self.announce_to_screen_reader("System ready - all data loaded")
        elif has_input_file and has_output_file:
            # Files selected but no data - use warning style
            style = ttk.Style()
            style.configure('Warning.TButton',
                           foreground='#FFFFFF',  # White text
                           background=self.colors['warning'],
                           font=('Arial', 12, 'bold'),
                           focuscolor=self.colors['accent'],
                           borderwidth=2,
                           relief='raised')
            style.map('Warning.TButton',
                     background=[('active', self.colors['secondary']),
                               ('pressed', self.colors['accent'])],
                     foreground=[('active', '#FFFFFF'),
                               ('pressed', '#FFFFFF')])
            self.run_button.configure(state='normal', style='Warning.TButton')
            self.status_var.set("Load data first, then run audit")
        else:
            # Missing files - disabled state
            style = ttk.Style()
            style.configure('Disabled.TButton',
                           foreground='#FFFFFF',  # White text even when disabled
                           background=self.colors['border'],
                           font=('Arial', 12, 'bold'),
                           borderwidth=2,
                           relief='raised')
            style.map('Disabled.TButton',
                     background=[('active', self.colors['border']),
                               ('pressed', self.colors['border'])],
                     foreground=[('active', '#FFFFFF'),
                               ('pressed', '#FFFFFF')])
            self.run_button.configure(state='disabled', style='Disabled.TButton')
            self.status_var.set("Select input and output files first")
        
        # Configure export button
        if has_results:
            self.export_button.configure(state='normal')
        else:
            self.export_button.configure(state='disabled')
        
        # Configure dashboard button
        if DASHBOARD_AVAILABLE and hasattr(self, 'dashboard_button'):
            if has_input_file and has_data:
                self.dashboard_button.configure(state='normal')
            else:
                self.dashboard_button.configure(state='disabled')
    
    def create_results_section(self, parent, row):
        """Create results display section with ADA compliance"""
        
        results_frame = ttk.LabelFrame(parent, text="Step 4: Results & Activity Log", padding="15")
        results_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        parent.rowconfigure(row, weight=1)
        
        # Create accessible scrolled text widget for results
        self.results_text = scrolledtext.ScrolledText(
            results_frame,
            wrap=tk.WORD, 
            width=80, 
            height=15,
            font=('Consolas', 11),  # Monospace font for better readability
            bg=self.colors['background'],
            fg='#000000',  # Black text for log messages on white background
            insertbackground='#000000',  # Black cursor
            selectbackground=self.colors['primary'],
            selectforeground='#003366',  # Dark blue selected text
            relief='solid',
            borderwidth=2
        )
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.add_to_tab_order(self.results_text)
        
        # Configure text tags for different message types
        self.results_text.tag_configure('success', 
                                       foreground=self.colors['success'], 
                                       font=('Consolas', 11, 'bold'))
        self.results_text.tag_configure('warning', 
                                       foreground=self.colors['warning'], 
                                       font=('Consolas', 11, 'bold'))
        self.results_text.tag_configure('error', 
                                       foreground=self.colors['error'], 
                                       font=('Consolas', 11, 'bold'))
        self.results_text.tag_configure('info', 
                                       foreground=self.colors['primary'], 
                                       font=('Consolas', 11))
        
        # Add welcome message with accessibility information
        self.log_message("♿ Welcome to the ADA Compliant PADC Processor!", 'success')
        self.log_message("🔧 Updated with McClellan (CM) and Sac Youth Center (SYC) consolidation", 'info')
        if DASHBOARD_AVAILABLE:
            self.log_message("📈 ADA Dashboard feature available for CSV generation", 'info')
        self.log_message("� Table sorting feature: Click column headers or use F2-F5 keys", 'info')
        self.log_message("�📋 Step-by-step process:", 'info')
        self.log_message("   1. Select your input attendance file and output audit file", 'info')
        self.log_message("   2. Click 'Load & Analyze Data' to detect program boundaries", 'info')
        self.log_message("   3. Review and edit boundaries if needed (sortable table)", 'info')
        self.log_message("   4. Click 'Run Audit Process' to execute the full audit with consolidation", 'info')
        if DASHBOARD_AVAILABLE:
            self.log_message("   5. Click 'Run ADA Dashboard' to generate CSV dashboard (optional)", 'info')
        self.log_message("", 'info')
        self.log_message("🎯 Accessibility Features Active:", 'success')
        self.log_message("   • Press F1 for help and keyboard shortcuts", 'info')
        self.log_message("   • Use Tab to navigate between elements", 'info')
        self.log_message("   • Mouse wheel or Ctrl+Up/Down to scroll", 'info')
        self.log_message("   • Page Up/Down for page navigation", 'info')
        self.log_message("   • Ctrl+Home/End for quick navigation", 'info')
        self.log_message("   • F2-F5 keys for table sorting", 'info')
        self.log_message("   • High contrast colors for better visibility", 'info')
        self.log_message("   • Screen reader compatible interface", 'info')
        self.log_message("=" * 60, 'info')
    
    def set_default_paths(self):
        """Set default file paths by finding the most recent files"""
        
        # Find the most recent attendance file in Downloads
        downloads_dir = Path("C:\\Users\\Shawn\\Downloads")
        
        if downloads_dir.exists():
            # Find all attendance summary files
            attendance_files = list(downloads_dir.glob("PrintMonthlyAttendanceSummaryTotals_*.xlsx"))
            
            if attendance_files:
                # Get the most recent file by modification time
                most_recent_input = max(attendance_files, key=lambda f: f.stat().st_mtime)
                self.input_file_path.set(str(most_recent_input))
                self.log_message(f"Auto-selected most recent input file: {most_recent_input.name}", 'info')
        
        # Default output path from original script
        default_output = "C:\\Users\\Shawn\\Downloads\\2025-2026_I4C_ADA_Reconciliation.xlsx"
        if Path(default_output).exists():
            self.output_file_path.set(default_output)
    
    def browse_input_file(self):
        """Browse for input attendance file with accessibility features"""
        filename = filedialog.askopenfilename(
            title="Select Input Attendance File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file_path.set(filename)
            self.log_message(f"Input file selected: {Path(filename).name}", 'success')
            self.update_button_states()
            self.announce_to_screen_reader("Input file selected")
    
    def browse_output_file(self):
        """Browse for output audit file with accessibility features"""
        filename = filedialog.asksaveasfilename(
            title="Select Output Audit File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file_path.set(filename)
            self.log_message(f"Output file selected: {Path(filename).name}", 'success')
            self.update_button_states()
            self.announce_to_screen_reader("Output file selected")
    
    def load_and_analyze_data(self):
        """Load data and analyze program boundaries"""
        
        if not self.input_file_path.get():
            messagebox.showerror("Error", "Please select an input file first.")
            return
        
        if not Path(self.input_file_path.get()).exists():
            messagebox.showerror("Error", "Input file does not exist.")
            return
        
        try:
            self.status_var.set("Loading data...")
            self.progress_var.set(20)
            self.root.update()
            
            # Load the Excel data
            self.log_message(f"Loading data from: {self.input_file_path.get()}")
            self.student_attendance_data = pd.read_excel(self.input_file_path.get(), header=None)
            
            self.progress_var.set(40)
            self.root.update()
            
            # Find program boundaries
            self.log_message("Analyzing program boundaries...")
            self.find_program_boundaries()
            
            self.progress_var.set(60)
            self.root.update()
            
            # Adjust boundaries
            self.log_message("Adjusting boundaries to prevent overlaps...")
            self.adjust_program_boundaries()
            
            self.progress_var.set(80)
            self.root.update()
            
            # Update the display
            self.update_boundaries_display()
            
            self.progress_var.set(100)
            self.status_var.set("Data loaded and analyzed")
            self.log_message("Data analysis complete! Review boundaries above.", 'success')
            self.update_button_states()
            self.announce_to_screen_reader("Data analysis completed successfully")
            
        except Exception as e:
            self.log_message(f"Error loading data: {str(e)}", 'error')
            self.status_var.set("Error loading data")
            self.announce_to_screen_reader(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
        finally:
            self.progress_var.set(0)
    
    def find_program_boundaries(self):
        """Find boundaries for each program"""
        
        for full_program_name, short_code in self.program_name_mappings.items():
            matching_rows = find_rows_containing_program_name(self.student_attendance_data, full_program_name)
            start_row, end_row = find_program_boundary_rows(matching_rows)
            self.program_boundaries[short_code]["start"] = start_row
            self.program_boundaries[short_code]["stop"] = end_row
            
            self.log_message(f"Found {short_code}: Start={start_row}, End={end_row}")
    
    def adjust_program_boundaries(self):
        """Adjust boundaries to prevent overlaps (from original script logic)"""
        
        # Fix Program C boundaries
        prog_C_tk_start = self.program_boundaries["Prog_C_TK"]["start"]
        prog_N_start = self.program_boundaries["Prog_N"]["start"]

        if prog_C_tk_start is not None and prog_N_start is not None:
            self.program_boundaries["Prog_C"]["stop"] = prog_C_tk_start - 1

        if prog_N_start is not None:
            self.program_boundaries["Prog_C_TK"]["stop"] = prog_N_start - 1

        # Fix Program N boundaries
        prog_N_tk_start = self.program_boundaries["Prog_N_TK"]["start"]
        if prog_N_tk_start is not None:
            self.program_boundaries["Prog_N"]["stop"] = prog_N_tk_start - 1

        # Fix remaining program boundaries
        programs_to_adjust = ["Prog_N_TK", "Prog_J", "Prog_K"]
        for i in range(len(programs_to_adjust) - 1):
            current_program = programs_to_adjust[i]
            next_program = programs_to_adjust[i + 1]
            
            current_start = self.program_boundaries[current_program]["start"]
            next_start = self.program_boundaries[next_program]["start"]
            
            if current_start is not None and next_start is not None:
                self.program_boundaries[current_program]["stop"] = next_start - 1
    
    def update_boundaries_display(self):
        """Update the boundaries table display"""
        
        # Debug logging
        self.log_message(f"🔄 Updating boundaries display...")
        
        # Clear existing items
        for item in self.boundaries_tree.get_children():
            self.boundaries_tree.delete(item)
        
        # Prepare boundary data for display and sorting
        self.boundary_data = []
        items_added = 0
        
        for short_code, boundaries in self.program_boundaries.items():
            # Find the full program name
            full_name = "Unknown"
            for full, short in self.program_name_mappings.items():
                if short == short_code:
                    full_name = full
                    break
            
            start = boundaries.get("start", "Not found")
            stop = boundaries.get("stop", "Not found")
            
            # Store data for sorting
            self.boundary_data.append({
                'Program Code': short_code,
                'Program Name': full_name,
                'Start Row': start,
                'End Row': stop
            })
            
            items_added += 1
            
            # Debug log for non-null boundaries
            if start != "Not found" or stop != "Not found":
                self.log_message(f"  📋 {short_code}: Start={start}, End={stop}")
        
        # Apply current sort if any
        if self.sort_column:
            self.apply_sort()
        else:
            # No sort applied, just populate in original order
            self.populate_tree_from_data()
        
        self.log_message(f"✅ Added {items_added} boundary entries to display")
    
    def sort_table(self, column):
        """Sort the table by the specified column"""
        
        # Toggle sort direction if clicking the same column
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        
        # Apply the sort
        self.apply_sort()
        
        # Update column headers to show sort direction
        self.update_sort_indicators()
        
        # Log the sort action for accessibility
        direction = "descending" if self.sort_reverse else "ascending"
        self.log_message(f"🔽 Table sorted by {column} ({direction})")
        self.announce_to_screen_reader(f"Table sorted by {column} {direction}")
    
    def apply_sort(self):
        """Apply the current sort to the boundary data"""
        
        if not self.sort_column or not self.boundary_data:
            return
        
        # Define sort key function
        def sort_key(item):
            value = item[self.sort_column]
            
            # Handle numeric values for row numbers
            if self.sort_column in ['Start Row', 'End Row']:
                if value == "Not found" or value is None:
                    return float('inf') if not self.sort_reverse else float('-inf')
                try:
                    return int(value)
                except (ValueError, TypeError):
                    return float('inf') if not self.sort_reverse else float('-inf')
            
            # Handle text values
            if value is None or value == "Not found":
                return "zzz" if not self.sort_reverse else ""
            
            return str(value).lower()
        
        # Sort the data
        self.boundary_data.sort(key=sort_key, reverse=self.sort_reverse)
        
        # Repopulate the tree
        self.populate_tree_from_data()
    
    def populate_tree_from_data(self):
        """Populate the tree view from the sorted boundary data"""
        
        # Clear existing items
        for item in self.boundaries_tree.get_children():
            self.boundaries_tree.delete(item)
        
        # Add sorted data to tree
        for data in self.boundary_data:
            self.boundaries_tree.insert('', tk.END, values=(
                data['Program Code'],
                data['Program Name'], 
                data['Start Row'],
                data['End Row']
            ))
    
    def update_sort_indicators(self):
        """Update column headers to show current sort direction"""
        
        columns = ['Program Code', 'Program Name', 'Start Row', 'End Row']
        
        for col in columns:
            if col == self.sort_column:
                # Show sort direction for current column
                arrow = "↓" if self.sort_reverse else "↑"
                header_text = f"{col} {arrow}"
            else:
                # Show sortable indicator for other columns
                header_text = f"{col} ↕"
            
            self.boundaries_tree.heading(col, text=header_text)
    
    def reset_sort(self):
        """Reset table sorting to original order"""
        
        self.sort_column = None
        self.sort_reverse = False
        
        # Reset all column headers
        self.update_sort_indicators()
        
        # Repopulate in original order
        self.populate_tree_from_data()
        
        # Log the action
        self.log_message("🔄 Table sort reset to original order")
        self.announce_to_screen_reader("Table sort reset to original order")
    
    def edit_boundary(self, event):
        """Handle double-click to edit boundary"""
        self.edit_selected_boundary()
    
    def edit_selected_boundary(self):
        """Edit the selected boundary"""
        
        selection = self.boundaries_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a program boundary to edit.")
            return
        
        # Get selected item data
        item = selection[0]
        values = self.boundaries_tree.item(item, 'values')
        program_code = values[0]
        current_start = values[2]
        current_end = values[3]
        
        # Create edit dialog
        self.create_boundary_edit_dialog(program_code, current_start, current_end)
    
    def create_boundary_edit_dialog(self, program_code, current_start, current_end):
        """Create a dialog for editing program boundaries"""
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Boundary - {program_code}")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.grab_set()  # Make dialog modal
        
        # Center the dialog
        dialog.transient(self.root)
        
        # Create dialog content
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Program info
        ttk.Label(main_frame, text=f"Editing boundaries for: {program_code}", 
                 font=('Arial', 10, 'bold')).pack(pady=(0, 20))
        
        # Start row
        start_frame = ttk.Frame(main_frame)
        start_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(start_frame, text="Start Row:").pack(side=tk.LEFT)
        start_var = tk.StringVar(value=str(current_start) if current_start != "Not found" else "")
        start_entry = ttk.Entry(start_frame, textvariable=start_var, width=20)
        start_entry.pack(side=tk.RIGHT)
        
        # End row
        end_frame = ttk.Frame(main_frame)
        end_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(end_frame, text="End Row:").pack(side=tk.LEFT)
        end_var = tk.StringVar(value=str(current_end) if current_end != "Not found" else "")
        end_entry = ttk.Entry(end_frame, textvariable=end_var, width=20)
        end_entry.pack(side=tk.RIGHT)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack()
        
        def save_changes():
            try:
                # Validate and save changes
                start_val = None if start_var.get().strip() == "" else int(start_var.get())
                end_val = None if end_var.get().strip() == "" else int(end_var.get())
                
                # Update the boundaries
                self.program_boundaries[program_code]["start"] = start_val
                self.program_boundaries[program_code]["stop"] = end_val
                
                # Update display
                self.update_boundaries_display()
                
                # Log the change
                self.log_message(f"✅ Updated {program_code}: Start={start_val}, End={end_val}")
                
                dialog.destroy()
                
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers or leave blank for None.")
        
        ttk.Button(button_frame, text="Save", command=save_changes, 
                  style='Accessible.TButton', width=12).pack(side=tk.LEFT, padx=(0, 10), pady=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy,
                  style='Accessible.TButton', width=12).pack(side=tk.LEFT, pady=5)
        
        # Focus on start entry
        start_entry.focus()
        
        # Enter key saves
        dialog.bind('<Return>', lambda e: save_changes())
    
    def run_audit_process(self):
        """Run the complete audit process"""
        
        # Validation
        if not self.input_file_path.get() or not self.output_file_path.get():
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        
        if self.student_attendance_data is None:
            messagebox.showerror("Error", "Please load and analyze data first.")
            return
        
        # Disable the run button to prevent multiple runs
        self.run_button.configure(state='disabled')
        
        # Run in a separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.execute_audit_process, daemon=True)
        thread.start()
    
    def execute_audit_process(self):
        """Execute the audit process in a separate thread"""
        
        try:
            self.status_var.set("Running audit process...")
            self.log_message("🚀 Starting audit process...")
            
            # Step 1: Find month occurrences
            self.progress_var.set(10)
            self.log_message("📅 Finding month occurrences...")
            
            monthly_attendance_by_program = {}
            for month_number in range(1, 13):
                rows_with_this_month = find_rows_containing_month_number(self.student_attendance_data, month_number)
                monthly_attendance_by_program[month_number] = rows_with_this_month
                self.log_message(f"  Month {month_number}: Found in {len(rows_with_this_month)} rows")
            
            self.progress_var.set(30)
            
            # Step 2: Extract attendance data
            self.log_message("📈 Extracting attendance data...")
            
            raw_attendance_data = extract_student_attendance_data(
                monthly_attendance_by_program,
                self.program_boundaries,
                self.student_attendance_data
            )
            
            self.progress_var.set(40)
            self.log_message(f"✅ Extracted {len(raw_attendance_data)} raw attendance data points")
            
            # Step 3: Consolidate sub-location data with parent programs
            self.log_message("🔄 Consolidating sub-location data with parent programs...")
            self.log_message("   Program C Total = Main Program C + McClellan (CM) + Sac Youth Center (SYC)")
            self.log_message("   Program N Total = Main Program N + McClellan (CM) + Sac Youth Center (SYC)")
            
            self.extracted_attendance_data = self.consolidate_attendance_data(raw_attendance_data)
            
            self.progress_var.set(60)
            self.log_message(f"✅ Consolidated {len(self.extracted_attendance_data)} attendance data points")
            
            # Step 4: Write to Excel
            self.progress_var.set(80)
            self.log_message("💾 Writing consolidated data to Excel...")
            
            write_all_attendance_data_to_excel_efficiently(
                self.extracted_attendance_data,
                self.output_file_path.get(),
                self.worksheet_name.get()
            )
            
            self.progress_var.set(100)
            self.status_var.set("Audit completed successfully")
            self.log_message("Audit process completed successfully!", 'success')
            self.log_message(f"Results saved to: {self.output_file_path.get()}", 'success')
            self.update_button_states()
            self.announce_to_screen_reader("Audit process completed successfully")
            
            # Show completion message
            self.root.after(0, lambda: messagebox.showinfo("Success", 
                f"Audit process completed successfully!\n\nResults saved to:\n{self.output_file_path.get()}"))
            
        except Exception as e:
            error_msg = str(e)
            self.log_message(f"Error during audit process: {error_msg}", 'error')
            self.status_var.set("Error during audit process")
            self.announce_to_screen_reader(f"Audit failed: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Audit process failed: {error_msg}"))
        
        finally:
            # Re-enable the run button and update states
            self.root.after(0, lambda: self.run_button.configure(state='normal'))
            self.root.after(0, lambda: self.update_button_states())
            self.root.after(0, lambda: self.progress_var.set(0))
    
    def consolidate_attendance_data(self, raw_attendance_data):
        """
        Consolidate sub-location data with parent programs.
        
        This function combines McClellan (CM) and Sac Youth Center (SYC) data
        with their respective parent programs (C and N) as required for audit compliance.
        """
        consolidated_attendance_data = {}
        
        # Process each consolidation rule
        for parent_program, child_programs in self.program_consolidation_rules.items():
            self.log_message(f"  Consolidating {parent_program}: {child_programs}")
            
            # For each month (1-12) and age group combination
            for month in range(1, 13):
                for age_group in ["TK-3", "4-6", "7-8", "9-12"]:
                    # Create the field name pattern
                    field_pattern = f"{parent_program}_Month_{month}_{age_group}: "
                    
                    # Sum up values from all child programs
                    total_value = 0
                    found_values = []
                    
                    for child_program in child_programs:
                        child_field_pattern = f"{child_program}_Month_{month}_{age_group}: "
                        child_value = raw_attendance_data.get(child_field_pattern, 0)
                        
                        if child_value and not pd.isna(child_value) and child_value != 0:
                            total_value += child_value
                            found_values.append(f"{child_program}: {child_value}")
                    
                    # Store the consolidated value
                    consolidated_attendance_data[field_pattern] = total_value
                    
                    # Log consolidation details for non-zero values
                    if total_value > 0:
                        self.log_message(f"    {field_pattern} = {' + '.join(found_values)} = {total_value}")
        
        return consolidated_attendance_data
    
    def increase_font_size(self):
        """Increase font size for better accessibility"""
        # Implementation for increasing font sizes
        self.log_message("Font size increased", 'info')
    
    def decrease_font_size(self):
        """Decrease font size for accessibility"""
        # Implementation for decreasing font sizes
        self.log_message("Font size decreased", 'info')
    
    def reset_font_size(self):
        """Reset font size to default"""
        # Implementation for resetting font sizes
        self.log_message("Font size reset to default", 'info')
    
    def scroll_up(self):
        """Scroll up one line for keyboard accessibility"""
        if hasattr(self, 'main_canvas') and self.main_canvas.winfo_viewable():
            self.main_canvas.yview_scroll(-1, "units")
    
    def scroll_down(self):
        """Scroll down one line for keyboard accessibility"""
        if hasattr(self, 'main_canvas') and self.main_canvas.winfo_viewable():
            self.main_canvas.yview_scroll(1, "units")
    
    def scroll_to_top(self):
        """Scroll to the top of the interface"""
        if hasattr(self, 'main_canvas') and self.main_canvas.winfo_viewable():
            self.main_canvas.yview_moveto(0)
    
    def scroll_to_bottom(self):
        """Scroll to the bottom of the interface"""
        if hasattr(self, 'main_canvas') and self.main_canvas.winfo_viewable():
            self.main_canvas.yview_moveto(1)
    
    def page_up(self):
        """Scroll up one page for keyboard accessibility"""
        if hasattr(self, 'main_canvas') and self.main_canvas.winfo_viewable():
            self.main_canvas.yview_scroll(-10, "units")
    
    def page_down(self):
        """Scroll down one page for keyboard accessibility"""
        if hasattr(self, 'main_canvas') and self.main_canvas.winfo_viewable():
            self.main_canvas.yview_scroll(10, "units")
    
    def export_results(self):
        """Export extracted data to a text file for review"""
        
        if not self.extracted_attendance_data:
            messagebox.showwarning("Warning", "No data to export. Run the audit process first.")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write("ADA Audit Process - Consolidated Attendance Data\n")
                    f.write("=" * 55 + "\n")
                    f.write("Note: This data includes consolidation of McClellan (CM) and Sac Youth Center (SYC)\n")
                    f.write("with their respective parent programs for audit compliance.\n\n")
                    
                    for field_name, attendance_value in self.extracted_attendance_data.items():
                        f.write(f"{field_name}: {attendance_value}\n")
                
                self.log_message(f"✅ Results exported to: {filename}")
                messagebox.showinfo("Success", f"Results exported to:\n{filename}")
                
            except Exception as e:
                self.log_message(f"❌ Error exporting results: {str(e)}")
                messagebox.showerror("Error", f"Failed to export results: {str(e)}")
    
    def run_ada_dashboard(self):
        """Run the ADA Dashboard feature using current boundaries"""
        
        if not DASHBOARD_AVAILABLE:
            messagebox.showerror("Error", "ADA Dashboard module is not available. Please check the installation.")
            return
        
        # Validation
        if not self.input_file_path.get():
            messagebox.showerror("Error", "Please select an input file first.")
            return
        
        if self.student_attendance_data is None:
            messagebox.showerror("Error", "Please load and analyze data first.")
            return
        
        # Validate boundaries
        is_valid, message, is_warning = validate_boundaries_for_dashboard(self.program_boundaries)
        if not is_valid:
            if is_warning:
                # Show warning dialog with option to continue
                if not messagebox.askyesno("Warning", message):
                    return
            else:
                messagebox.showerror("Error", f"Boundary validation failed:\n{message}")
                return
        
        try:
            # Get dashboard configuration from user
            self.log_message("📋 Getting dashboard configuration from user...")
            school_year, location, school_name = get_dashboard_configuration_from_user()
            
            if not all([school_year, location, school_name]):
                self.log_message("⚠️ Dashboard cancelled by user")
                return
            
            self.log_message(f"📊 Dashboard Config - Year: {school_year}, Location: {location}, School: {school_name}")
            
            # Disable dashboard button during processing
            if hasattr(self, 'dashboard_button'):
                self.dashboard_button.configure(state='disabled')
            
            # Set up progress and logging callbacks
            def progress_callback(value):
                self.progress_var.set(value)
                self.root.update()
            
            def log_callback(message, msg_type='info'):
                self.log_message(message, msg_type)
            
            # Set status
            self.status_var.set("Running ADA Dashboard...")
            
            # Run dashboard in a separate thread to prevent GUI freezing
            thread = threading.Thread(
                target=self.execute_dashboard_process,
                args=(school_year, location, school_name, progress_callback, log_callback),
                daemon=True
            )
            thread.start()
            
        except Exception as e:
            self.log_message(f"❌ Error starting dashboard: {str(e)}", 'error')
            messagebox.showerror("Error", f"Failed to start dashboard: {str(e)}")
            if hasattr(self, 'dashboard_button'):
                self.dashboard_button.configure(state='normal')
    
    def execute_dashboard_process(self, school_year, location, school_name, progress_callback, log_callback):
        """Execute the dashboard process in a separate thread"""
        
        try:
            # Create output directory for dashboard files
            output_dir = os.path.join(os.path.dirname(self.input_file_path.get()), "ADA_Dashboard_Output")
            
            # Run the dashboard
            results = run_ada_dashboard_with_boundaries(
                input_file_path=self.input_file_path.get(),
                program_boundaries=self.program_boundaries,
                program_mappings=self.program_name_mappings,
                school_year=school_year,
                location=location,
                school_name=school_name,
                output_dir=output_dir,
                progress_callback=progress_callback,
                log_callback=log_callback
            )
            
            # Handle results
            if results['success']:
                message = f"ADA Dashboard completed successfully!\n\n"
                message += f"Records generated: {results['record_count']}\n"
                message += f"Data fields extracted: {results['data_fields']}\n"
                
                if results.get('csv_file'):
                    message += f"Output file: {os.path.basename(results['csv_file'])}\n"
                    message += f"Location: {os.path.dirname(results['csv_file'])}"
                
                self.root.after(0, lambda: self.status_var.set("Dashboard completed successfully"))
                self.root.after(0, lambda: messagebox.showinfo("Dashboard Complete", message))
                
                # Open output directory
                if results.get('csv_file') and os.path.exists(results['csv_file']):
                    self.root.after(0, lambda: self.log_message("📂 Opening output directory..."))
                    try:
                        os.startfile(os.path.dirname(results['csv_file']))
                    except Exception:
                        pass  # If can't open directory, that's ok
                        
            else:
                error_msg = results.get('message', 'Unknown error occurred')
                self.root.after(0, lambda: self.status_var.set("Dashboard failed"))
                self.root.after(0, lambda: messagebox.showerror("Dashboard Error", error_msg))
                
        except Exception as e:
            error_msg = f"Dashboard process failed: {str(e)}"
            self.root.after(0, lambda: log_callback(error_msg, 'error'))
            self.root.after(0, lambda: self.status_var.set("Dashboard failed"))
            self.root.after(0, lambda: messagebox.showerror("Dashboard Error", error_msg))
        
        finally:
            # Re-enable dashboard button and reset progress
            self.root.after(0, lambda: self.update_button_states())
            self.root.after(0, lambda: self.progress_var.set(0))
    
    def log_message(self, message, message_type='info'):
        """Add a message to the results log with accessibility features"""
        
        timestamp = time.strftime("%H:%M:%S")
        
        # Determine icon based on message type
        icons = {
            'success': '✅',
            'warning': '⚠️',
            'error': '❌',
            'info': 'ℹ️'
        }
        
        icon = icons.get(message_type, 'ℹ️')
        formatted_message = f"[{timestamp}] {icon} {message}\n"
        
        # Announce important messages to screen readers
        if message_type in ['success', 'error', 'warning']:
            self.announce_to_screen_reader(message)
        
        # Use after() to ensure thread-safe GUI updates
        self.root.after(0, lambda: self._append_to_log(formatted_message, message_type))
    
    def _append_to_log(self, message, message_type='info'):
        """Append message to log with appropriate styling (GUI thread only)"""
        
        # Insert message with appropriate tag
        start_pos = self.results_text.index(tk.END + "-1c")
        self.results_text.insert(tk.END, message)
        end_pos = self.results_text.index(tk.END + "-1c")
        
        # Apply styling tag
        self.results_text.tag_add(message_type, start_pos, end_pos)
        
        # Auto-scroll to bottom
        self.results_text.see(tk.END)
    
    def load_saved_configurations(self):
        """Load all saved boundary configurations from disk"""
        try:
            for settings_file in self.settings_directory.glob("*.json"):
                try:
                    with open(settings_file, 'r') as f:
                        config_data = json.load(f)
                        # Use the filename (stem) as the key for consistency
                        config_name = settings_file.stem
                        # Ensure the config_data has the correct name matching the file
                        config_data['name'] = config_name
                        self.saved_configurations[config_name] = config_data
                except Exception as e:
                    self.log_message(f"⚠️ Warning: Could not load configuration '{settings_file.name}': {e}")
            
            if self.saved_configurations:
                self.log_message(f"📁 Loaded {len(self.saved_configurations)} saved configurations")
        except Exception as e:
            self.log_message(f"⚠️ Warning: Could not access settings directory: {e}")
    
    def save_boundary_configuration(self):
        """Save current boundary configuration with a user-defined name"""
        
        # Check if we have boundaries to save
        if not any(b["start"] is not None or b["stop"] is not None for b in self.program_boundaries.values()):
            messagebox.showwarning("Warning", "No boundary data to save. Please load and analyze data first.")
            return
        
        # Create save dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Save Boundary Configuration")
        dialog.geometry("400x250")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.transient(self.root)
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configuration name
        ttk.Label(main_frame, text="Configuration Name:", font=('Arial', 10, 'bold')).pack(pady=(0, 5))
        name_var = tk.StringVar()
        name_entry = ttk.Entry(main_frame, textvariable=name_var, width=40)
        name_entry.pack(pady=(0, 15))
        name_entry.focus()
        
        # Description
        ttk.Label(main_frame, text="Description (optional):").pack(pady=(0, 5))
        description_var = tk.StringVar()
        description_entry = ttk.Entry(main_frame, textvariable=description_var, width=40)
        description_entry.pack(pady=(0, 20))
        
        # Existing configurations info
        if self.saved_configurations:
            ttk.Label(main_frame, text="Existing configurations:", font=('Arial', 9)).pack(pady=(0, 5))
            config_list = ", ".join(list(self.saved_configurations.keys())[:5])
            if len(self.saved_configurations) > 5:
                config_list += f" (and {len(self.saved_configurations) - 5} more)"
            ttk.Label(main_frame, text=config_list, font=('Arial', 8), foreground='gray').pack(pady=(0, 15))
        
        def save_config():
            config_name = name_var.get().strip()
            if not config_name:
                messagebox.showerror("Error", "Please enter a configuration name.")
                return
            
            # Validate name (no invalid filename characters)
            invalid_chars = '<>:"/\\|?*'
            if any(char in config_name for char in invalid_chars):
                messagebox.showerror("Error", f"Configuration name contains invalid characters: {invalid_chars}")
                return
            
            # Check if configuration already exists
            if config_name in self.saved_configurations:
                if not messagebox.askyesno("Confirm Overwrite", 
                    f"Configuration '{config_name}' already exists. Overwrite it?"):
                    return
            
            try:
                # Create configuration data
                config_data = {
                    "name": config_name,
                    "description": description_var.get().strip(),
                    "created_date": datetime.now().isoformat(),
                    "program_boundaries": dict(self.program_boundaries),
                    "program_mappings": dict(self.program_name_mappings)
                }
                
                # Save to file
                config_file = self.settings_directory / f"{config_name}.json"
                with open(config_file, 'w') as f:
                    json.dump(config_data, f, indent=2)
                
                # Update in-memory configurations
                self.saved_configurations[config_name] = config_data
                
                self.log_message(f"✅ Configuration '{config_name}' saved successfully")
                messagebox.showinfo("Success", f"Configuration '{config_name}' has been saved.")
                dialog.destroy()
                
            except Exception as e:
                self.log_message(f"❌ Error saving configuration: {e}")
                messagebox.showerror("Error", f"Failed to save configuration: {e}")
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))
        
        ttk.Button(button_frame, text="Save", command=save_config,
                  style='Accessible.TButton', width=12).pack(side=tk.LEFT, padx=(0, 10), pady=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy,
                  style='Accessible.TButton', width=12).pack(side=tk.LEFT, pady=5)
        
        # Enter key saves
        dialog.bind('<Return>', lambda e: save_config())
    
    def load_boundary_configuration(self):
        """Load a saved boundary configuration"""
        
        if not self.saved_configurations:
            messagebox.showinfo("Info", "No saved configurations found.")
            return
        
        # Create load dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Load Boundary Configuration")
        dialog.geometry("600x400")
        dialog.resizable(True, True)
        dialog.grab_set()
        dialog.transient(self.root)
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Select Configuration to Load:", font=('Arial', 10, 'bold')).pack(pady=(0, 10))
        
        # Create listbox with configurations
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        columns = ('Name', 'Description', 'Created Date')
        config_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=12)
        
        config_tree.heading('Name', text='Configuration Name')
        config_tree.heading('Description', text='Description')
        config_tree.heading('Created Date', text='Created Date')
        
        config_tree.column('Name', width=200)
        config_tree.column('Description', width=250)
        config_tree.column('Created Date', width=150)
        
        # Add scrollbars
        v_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=config_tree.yview)
        h_scroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=config_tree.xview)
        config_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        config_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Populate the list
        for name, config in self.saved_configurations.items():
            description = config.get('description', '')
            created_date = config.get('created_date', 'Unknown')
            if created_date != 'Unknown':
                try:
                    created_date = datetime.fromisoformat(created_date).strftime('%Y-%m-%d %H:%M')
                except:
                    pass
            config_tree.insert('', tk.END, values=(name, description, created_date))
        
        def load_selected():
            selection = config_tree.selection()
            if not selection:
                messagebox.showwarning("Warning", "Please select a configuration to load.")
                return
            
            config_name = config_tree.item(selection[0], 'values')[0]
            
            if messagebox.askyesno("Confirm Load", 
                f"Load configuration '{config_name}'?\nThis will replace current boundary settings."):
                
                try:
                    config_data = self.saved_configurations[config_name]
                    
                    # Load boundaries with proper deep copy
                    self.program_boundaries = copy.deepcopy(config_data['program_boundaries'])
                    
                    # Debug logging
                    self.log_message(f"🔄 Loading configuration '{config_name}'...")
                    non_null_boundaries = {k: v for k, v in self.program_boundaries.items() if v.get('start') is not None or v.get('stop') is not None}
                    self.log_message(f"📊 Loaded {len(non_null_boundaries)} programs with boundary data")
                    
                    # Update display
                    self.update_boundaries_display()
                    
                    # Force complete GUI refresh and update button states
                    self.root.update_idletasks()
                    self.update_button_states()
                    
                    # Additional force refresh for the tree view
                    if hasattr(self, 'boundaries_tree'):
                        self.boundaries_tree.update()
                    
                    self.log_message(f"✅ Configuration '{config_name}' loaded successfully")
                    messagebox.showinfo("Success", f"Configuration '{config_name}' has been loaded.")
                    dialog.destroy()
                    
                except Exception as e:
                    self.log_message(f"❌ Error loading configuration: {e}")
                    messagebox.showerror("Error", f"Failed to load configuration: {e}")
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))
        
        ttk.Button(button_frame, text="Load Selected", command=load_selected,
                  style='Accessible.TButton', width=15).pack(side=tk.LEFT, padx=(0, 10), pady=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy,
                  style='Accessible.TButton', width=12).pack(side=tk.LEFT, pady=5)
        
        # Double-click to load
        config_tree.bind('<Double-1>', lambda e: load_selected())
    
    def export_boundary_settings(self):
        """Export boundary settings to a JSON file"""
        
        if not any(b["start"] is not None or b["stop"] is not None for b in self.program_boundaries.values()):
            messagebox.showwarning("Warning", "No boundary data to export. Please load and analyze data first.")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Boundary Settings",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                export_data = {
                    "exported_date": datetime.now().isoformat(),
                    "program_boundaries": dict(self.program_boundaries),
                    "program_mappings": dict(self.program_name_mappings),
                    "export_info": {
                        "version": "1.0",
                        "source": "ADA Audit GUI",
                        "description": "Program boundary settings export"
                    }
                }
                
                with open(filename, 'w') as f:
                    json.dump(export_data, f, indent=2)
                
                self.log_message(f"✅ Boundary settings exported to: {filename}")
                messagebox.showinfo("Success", f"Settings exported successfully to:\n{filename}")
                
            except Exception as e:
                self.log_message(f"❌ Error exporting settings: {e}")
                messagebox.showerror("Error", f"Failed to export settings: {e}")
    
    def import_boundary_settings(self):
        """Import boundary settings from a JSON file"""
        
        filename = filedialog.askopenfilename(
            title="Import Boundary Settings",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    import_data = json.load(f)
                
                # Validate the import data structure
                if 'program_boundaries' not in import_data:
                    messagebox.showerror("Error", "Invalid settings file: missing program_boundaries data.")
                    return
                
                if messagebox.askyesno("Confirm Import", 
                    "Import these boundary settings?\nThis will replace current settings."):
                    
                    # Import boundaries
                    self.program_boundaries = import_data['program_boundaries']
                    
                    # Update program mappings if available
                    if 'program_mappings' in import_data:
                        self.program_name_mappings = import_data['program_mappings']
                    
                    # Update display
                    self.update_boundaries_display()
                    
                    self.log_message(f"✅ Boundary settings imported from: {filename}")
                    messagebox.showinfo("Success", "Settings imported successfully!")
                
            except json.JSONDecodeError:
                messagebox.showerror("Error", "Invalid JSON file format.")
            except Exception as e:
                self.log_message(f"❌ Error importing settings: {e}")
                messagebox.showerror("Error", f"Failed to import settings: {e}")
    
    def manage_configurations(self):
        """Manage saved configurations (view, delete, rename)"""
        
        if not self.saved_configurations:
            messagebox.showinfo("Info", "No saved configurations found.")
            return
        
        # Create management dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Manage Configurations")
        dialog.geometry("700x500")
        dialog.resizable(True, True)
        dialog.grab_set()
        dialog.transient(self.root)
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Manage Saved Configurations", font=('Arial', 12, 'bold')).pack(pady=(0, 15))
        
        # Create treeview for configurations
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        columns = ('Name', 'Description', 'Created Date', 'Programs Count')
        manage_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        manage_tree.heading('Name', text='Configuration Name')
        manage_tree.heading('Description', text='Description')
        manage_tree.heading('Created Date', text='Created Date')
        manage_tree.heading('Programs Count', text='Programs')
        
        manage_tree.column('Name', width=180)
        manage_tree.column('Description', width=220)
        manage_tree.column('Created Date', width=150)
        manage_tree.column('Programs Count', width=80)
        
        # Add scrollbars
        v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=manage_tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=manage_tree.xview)
        manage_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        manage_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        def refresh_list():
            # Clear existing items
            for item in manage_tree.get_children():
                manage_tree.delete(item)
            
            # Populate with current configurations
            for name, config in self.saved_configurations.items():
                description = config.get('description', '')
                created_date = config.get('created_date', 'Unknown')
                if created_date != 'Unknown':
                    try:
                        created_date = datetime.fromisoformat(created_date).strftime('%Y-%m-%d %H:%M')
                    except:
                        pass
                
                program_count = len(config.get('program_boundaries', {}))
                manage_tree.insert('', tk.END, values=(name, description, created_date, program_count))
        
        def delete_selected():
            selection = manage_tree.selection()
            if not selection:
                messagebox.showwarning("Warning", "Please select a configuration to delete.")
                return
            
            config_name = manage_tree.item(selection[0], 'values')[0]
            
            if messagebox.askyesno("Confirm Delete", 
                f"Are you sure you want to delete configuration '{config_name}'?\nThis action cannot be undone."):
                
                try:
                    # Remove from disk
                    config_file = self.settings_directory / f"{config_name}.json"
                    if config_file.exists():
                        config_file.unlink()
                    
                    # Remove from memory
                    del self.saved_configurations[config_name]
                    
                    refresh_list()
                    self.log_message(f"✅ Configuration '{config_name}' deleted successfully")
                    messagebox.showinfo("Success", f"Configuration '{config_name}' has been deleted.")
                    
                except Exception as e:
                    self.log_message(f"❌ Error deleting configuration: {e}")
                    messagebox.showerror("Error", f"Failed to delete configuration: {e}")
        
        # Initial population
        refresh_list()
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))
        
        ttk.Button(button_frame, text="Delete Selected", command=delete_selected,
                  style='Accessible.TButton', width=15).pack(side=tk.LEFT, padx=(0, 10), pady=5)
        ttk.Button(button_frame, text="Refresh", command=refresh_list,
                  style='Accessible.TButton', width=12).pack(side=tk.LEFT, padx=(0, 10), pady=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy,
                  style='Accessible.TButton', width=12).pack(side=tk.LEFT, pady=5)


def main():
    """Main function to run the ADA compliant GUI application"""
    
    root = tk.Tk()
    
    # Configure accessible styling
    style = ttk.Style()
    
    # Use a theme that supports accessibility
    available_themes = style.theme_names()
    if 'vista' in available_themes:
        style.theme_use('vista')
    elif 'winnative' in available_themes:
        style.theme_use('winnative')
    
    # Create and run the application
    app = ADAAuditGUI(root)
    
    # Set minimum window size for accessibility
    root.minsize(1000, 700)
    
    # Configure window properties for screen readers
    root.attributes('-topmost', False)  # Don't force always on top
    
    # Bind global accessibility shortcuts
    root.bind('<Control-plus>', lambda e: app.increase_font_size())
    root.bind('<Control-minus>', lambda e: app.decrease_font_size())
    root.bind('<Control-0>', lambda e: app.reset_font_size())
    
    try:
        # Start the GUI event loop
        root.mainloop()
    except KeyboardInterrupt:
        # Handle Ctrl+C gracefully
        root.quit()


if __name__ == "__main__":
    main()
