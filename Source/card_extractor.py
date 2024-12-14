import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import logging
from pathlib import Path
from typing import Dict, Optional, Any, Union, List
from dataclasses import dataclass
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, flowables, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# To do: alternate character card formats; multiple cards; multiple lorebooks?;
# PNG (in metadata chara category, need to decode from base64)

# Configuration
UI_CONFIG = {
    'WINDOW_SIZE': '900x700',
    'WINDOW_BG': '#f0f0f0',
    'BUTTON_PADDING': 10,
    'PREVIEW_FONT': ('Segoe UI', 10),
    'PREVIEW_BG': '#FFFFFF',
    'PREVIEW_FG': '#2C3E50',
    'HEADER_COLOR': '#2980B9',
    'SECTION_COLOR': '#16A085',
    'BOOK_ENTRY_COLOR': '#8E44AD',
    'SEPARATOR_COLOR': '#BDC3C7',
    'DISABLED_BTN_BG': '#ecf0f1',
    'ACTIVE_BTN_BG': '#2980b9',
    'MENU_BG': '#FFFFFF',           # Add this
    'MENU_FG': '#2C3E50',           # Add this
    'MENU_ACTIVE_BG': '#2980b9',    # Add this
    'MENU_ACTIVE_FG': '#FFFFFF',    # Add this
}

@dataclass
class TextStyle:
    """Defines text styling configuration"""
    font: tuple
    foreground: str
    spacing1: int = 0
    spacing2: int = 0
    spacing3: int = 0
    justify: str = 'left'
    lmargin1: int = 0
    lmargin2: int = 0

TEXT_STYLES = {
    'header': TextStyle(
        font=('Segoe UI', 14, 'bold'),
        foreground=UI_CONFIG['HEADER_COLOR'],
        spacing1=10,
        spacing3=10,
        justify='center'
    ),
    'section_title': TextStyle(
        font=('Segoe UI', 12, 'bold'),
        foreground=UI_CONFIG['SECTION_COLOR'],
        spacing1=15,
        spacing3=5
    ),
    'content': TextStyle(
        font=('Segoe UI', 10),
        foreground=UI_CONFIG['PREVIEW_FG'],
        lmargin1=20,
        lmargin2=20
    ),
    'separator': TextStyle(
        font=('Segoe UI', 10),
        foreground=UI_CONFIG['SEPARATOR_COLOR']
    ),
    'book_entry': TextStyle(
        font=('Segoe UI', 10, 'italic'),
        foreground=UI_CONFIG['BOOK_ENTRY_COLOR'],
        lmargin1=25,
        lmargin2=40
    )
}

def detect_file_type(json_data: Dict[str, Any]) -> str:
    """Detect whether the JSON is a character card or lorebook with improved detection"""
    print(f"\nDebug: File structure - spec: {json_data.get('spec', 'none')}")
    
    # Check if it's a lorebook first
    if 'entries' in json_data and isinstance(json_data['entries'], dict):
        first_entry = next(iter(json_data['entries'].values()), {})
        if any(key in first_entry for key in ['key', 'content', 'comment']):
            print("Debug: Detected as lorebook")
            return 'lorebook'
    
    # Enhanced character detection
    def check_char_fields(data: Dict[str, Any]) -> bool:
        """Check if dictionary contains character card fields"""
        char_card_fields = [
            'name', 'first_mes', 'description', 'personality',
            'char_name', 'char_persona', 'char_greeting', 'example_dialogue',
            'avatar', 'chat', 'persona', 'greeting', 'mes_example'
        ]
        return any(field in data for field in char_card_fields)
    
    # Check for character card by spec
    spec = json_data.get('spec', '')
    if isinstance(spec, str) and 'chara_card' in spec.lower():
        print("Debug: Character card detected by spec")
        return 'character'
        
    # Check main data structure
    if check_char_fields(json_data):
        print("Debug: Character card detected in main structure")
        return 'character'
        
    # Check in 'data' subfield
    if 'data' in json_data and isinstance(json_data['data'], dict):
        if check_char_fields(json_data['data']):
            print("Debug: Character card detected in data field")
            return 'character'
    
    print("Debug: Could not determine file type")
    return 'unknown'

def get_plain_text(formatted_text: str) -> str:
    """Convert formatted text to plain text by removing formatting markers"""
    lines = formatted_text.split('\n')
    plain_lines = []
    skip_next = False
    
    for line in lines:
        # Skip separator lines
        if all(c == '=' for c in line.strip()) or all(c == '─' for c in line.strip()):
            continue
            
        # Handle spacing after titles
        if skip_next and not line.strip():
            skip_next = False
            continue
            
        if line.strip():
            # Remove bullet points
            if line.startswith('►'):
                line = line[1:].strip()
            
            plain_lines.append(line.strip())
            
            # Skip extra newline after uppercase titles
            if line.isupper() and len(line) > 3:
                skip_next = True
            
    return '\n\n'.join(plain_lines)

def create_pdf(formatted_text: str, output_path: str):
    """Convert formatted text to PDF with styling that matches the preview"""
    def create_separator():
        """Create a horizontal line separator"""
        return flowables.HRFlowable(
            width="100%",
            thickness=1,
            color=colors.HexColor('#E0E0E0'),
            spaceBefore=5,
            spaceAfter=10
        )    

    def clean_text(text: str) -> str:
        """Remove problematic Unicode characters and replace with safe alternatives"""
        text = text.replace('■', '').replace('□', '').replace('►', '->')
        text = ''.join(c for c in text if ord(c) < 128)
        return text.strip()

    try:
        # Create the PDF document with metadata
        doc = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        # Get filename for title
        title = Path(output_path).stem
        
        # Set PDF metadata
        doc.title = title
        doc.author = "Character Card Extractor"
        doc.subject = "Extracted Character Data"
        
        # Create styles
        styles = getSampleStyleSheet()
        story = []
        
        # Font handling
        try:
            font_paths = {
                'SegoeUI': ('segoeui.ttf', 'Helvetica'),
                'SegoeUI-Bold': ('segoeuib.ttf', 'Helvetica-Bold'),
                'SegoeUI-Italic': ('segoeuii.ttf', 'Helvetica-Oblique')
            }
            
            registered_fonts = {}
            for font_name, (font_file, fallback) in font_paths.items():
                try:
                    pdfmetrics.registerFont(TTFont(font_name, font_file))
                    registered_fonts[font_name] = font_name
                except:
                    registered_fonts[font_name] = fallback
        except:
            registered_fonts = {
                'SegoeUI': 'Helvetica',
                'SegoeUI-Bold': 'Helvetica-Bold',
                'SegoeUI-Italic': 'Helvetica-Oblique'
            }    

        # Custom styles that match the preview
        styles.add(ParagraphStyle(
            name='DocumentTitle',
            parent=styles['Heading1'],
            fontName=registered_fonts['SegoeUI-Bold'],
            fontSize=16,
            spaceBefore=0,
            spaceAfter=20,
            alignment=1,  # Center alignment
            textColor=colors.HexColor('#2980B9')
        ))

        styles.add(ParagraphStyle(
            name='CharacterName',
            parent=styles['Heading2'],
            fontName=registered_fonts['SegoeUI-Bold'],
            fontSize=14,
            spaceBefore=10,
            spaceAfter=10,
            alignment=1,  # Center alignment
            textColor=colors.HexColor('#2980B9')  # Blue color matching preview
        ))

        styles.add(ParagraphStyle(
            name='CharacterSeparator',
            parent=styles['Normal'],
            fontName=registered_fonts['SegoeUI'],
            fontSize=12,
            spaceBefore=5,
            spaceAfter=5,
            alignment=1,  # Center alignment
            textColor=colors.HexColor('#2980B9')  # Match the name color
        ))
        
        styles.add(ParagraphStyle(
            name='SectionTitle',
            parent=styles['Heading2'],
            fontName=registered_fonts['SegoeUI-Bold'],
            fontSize=14,
            spaceBefore=20,
            spaceAfter=0,
            textColor=colors.HexColor('#16A085'),
            keepWithNext=True
        ))

        styles.add(ParagraphStyle(
            name='LorebookTitle',
            parent=styles['Heading2'],
            fontName=registered_fonts['SegoeUI-Bold'],
            fontSize=14,
            spaceBefore=20,
            spaceAfter=0,
            textColor=colors.HexColor('#16A085'),
            keepWithNext=True
        ))
        
        styles.add(ParagraphStyle(
            name='Content',
            parent=styles['Normal'],
            fontName=registered_fonts['SegoeUI'],
            fontSize=11,
            spaceBefore=0,
            spaceAfter=10,
            leftIndent=20,
            rightIndent=20,
            leading=14
        ))
        
        styles.add(ParagraphStyle(
            name='BookEntry',
            parent=styles['Normal'],
            fontName=registered_fonts['SegoeUI-Italic'],
            fontSize=11,
            spaceBefore=5,
            spaceAfter=5,
            leftIndent=20,
            rightIndent=20,
            textColor=colors.HexColor('#8E44AD')
        ))

        # Define main section titles we expect in character cards
        MAIN_SECTIONS = {
            "CHARACTER DEFINITION",
            "PERSONALITY",
            "CHARACTER NOTE",
            "SCENARIO",
            "FIRST MESSAGE",
            "EXAMPLE MESSAGES",
            "ALTERNATE GREETINGS",
            "CHARACTER BOOK"
        }
        
        # Get filename from path without extension
        title = Path(output_path).stem
        story.append(Paragraph(title, styles['DocumentTitle']))
        story.append(create_separator())

        # Process the text
        lines = formatted_text.split('\n')
        current_section = None
        buffer = []
        in_example_messages = False
        found_char_name = False
        is_lorebook = False

        # Check if this is a lorebook by looking for typical lorebook patterns
        for line in lines[:10]:  # Check first few lines
            if 'Keys:' in line or (line.strip().startswith('►') and 'CHARACTER' not in line):
                is_lorebook = True
                break

        i = 0
        while i < len(lines):
            line = lines[i].rstrip()
            
            # Skip empty lines
            if not line:
                i += 1
                continue

            # Clean the line
            clean_line = clean_text(line)
            if not clean_line:
                i += 1
                continue
            # Lorebook handling
            if is_lorebook:
                # Handle lorebook entry titles
                if clean_line and not clean_line.startswith('Keys:'):
                    # Check if this is a title
                    is_title = False
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        is_title = all(c == '─' for c in next_line) or not next_line
                    
                    if is_title:
                        if buffer:
                            # This is the key change - ensure content is added to story
                            content = '<br/>'.join(buffer)
                            if content.strip():  # Only add if there's actual content
                                story.append(Paragraph(content, styles['Content']))
                            buffer = []
                        story.append(Spacer(1, 20))  # Add space before new entry
                        story.append(Paragraph(clean_line, styles['LorebookTitle']))
                        story.append(create_separator())
                        # Skip the separator line
                        i += 1
                    else:
                        if not clean_line.startswith('Keys:'):
                            buffer.append(clean_line)
                # Handle Keys lines
                elif clean_line.startswith('Keys:'):
                    if buffer:
                        # Ensure content is added before the keys
                        content = '<br/>'.join(buffer)
                        if content.strip():  # Only add if there's actual content
                            story.append(Paragraph(content, styles['Content']))
                        buffer = []
                    story.append(Paragraph(clean_line, styles['BookEntry']))
                else:
                    buffer.append(clean_line)
            else:
                # Character card handling
                if line.strip() and all(c == '=' for c in line.strip()):
                    print(f"Debug: Found potential separator line: '{line}'")
                    # First, if there's any content in buffer, process it
                    if buffer:
                        story.append(Paragraph('<br/>'.join(buffer), styles['Content']))
                        buffer = []
                    
                    next_i = i + 1
                    if next_i < len(lines):
                        next_line = lines[next_i].strip()
                        if next_line and not all(c == '=' for c in next_line):  # Make sure next line isn't another separator
                            potential_name = clean_text(next_line)
                            print(f"Debug: Potential name found: '{potential_name}'")
                            next_next_i = next_i + 1
                            if next_next_i < len(lines):
                                next_next_line = lines[next_next_i].strip()
                                if next_next_line and all(c == '=' for c in next_next_line):
                                    print("Debug: Found complete name pattern!")
                                    # If this is not the first character, force a page break
                                    if found_char_name:
                                        story.append(PageBreak())
                                    else:
                                        # For first character, just add some space after title
                                        story.append(Spacer(1, 20))
                                    
                                    # This is a character name with separators
                                    separator = "=" * 50
                                    story.append(Paragraph(separator, styles['CharacterSeparator']))
                                    story.append(Paragraph(potential_name, styles['CharacterName']))
                                    story.append(Paragraph(separator, styles['CharacterSeparator']))
                                    story.append(Spacer(1, 12))
                                    i = next_next_i + 1
                                    found_char_name = True
                                    continue
                                    
                    # Skip any trailing separator lines
                    if i + 1 < len(lines) and all(c == '=' for c in lines[i + 1].strip()):
                        i += 1
                    continue
                                    
                    # Skip any trailing separator lines
                    if i + 1 < len(lines) and all(c == '=' for c in lines[i + 1].strip()):
                        i += 1
                    continue

                # Handle main section titles
                if clean_line in MAIN_SECTIONS:
                    if buffer:
                        story.append(Paragraph('<br/>'.join(buffer), styles['Content']))
                        buffer = []
                    
                    story.append(Paragraph(clean_line, styles['SectionTitle']))
                    if clean_line != "EXAMPLE MESSAGES":
                        story.append(create_separator())
                    
                    current_section = clean_line
                    in_example_messages = (clean_line == "EXAMPLE MESSAGES")
                    i += 1
                    continue
                
                # Handle bullet points and keys
                if (clean_line.startswith('Keys:') or 
                    clean_line.startswith('->') or 
                    clean_line.startswith('► ') or
                    clean_line.startswith('Greeting')):
                    if buffer:
                        story.append(Paragraph('<br/>'.join(buffer), styles['Content']))
                        buffer = []
                    story.append(Paragraph(clean_line, styles['BookEntry']))
                else:
                    buffer.append(clean_line)
            
            i += 1
        
        # Add any remaining buffered content
        if buffer:
            story.append(Paragraph('<br/>'.join(buffer), styles['Content']))        

        # Build the PDF
        doc.build(story)

    except PermissionError as e:
        raise PermissionError(f"Unable to save PDF: {str(e)}")
    except IOError as e:
        raise IOError(f"I/O error while creating PDF: {str(e)}")
    except Exception as e:
        raise RuntimeError(f"Unexpected error creating PDF: {str(e)}")


class FieldSelectorDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("Select Fields to Extract")
        self.selected_fields: Dict[str, tk.BooleanVar] = {}
        
        # Dialog configuration
        self.transient(parent)
        self.grab_set()
        
        # Center dialog
        window_width = 300
        window_height = 400
        x = (parent.winfo_screenwidth() - window_width) // 2
        y = (parent.winfo_screenheight() - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.resizable(True, True)
        self.selected_fields['name'] = tk.BooleanVar(value=True)
        
        # Field definitions
        self.field_options = {
            'description': 'Character Definition',
            'prompt': 'Character Note',
            'personality': 'Personality',
            'mes_example': 'Example Messages',
            'scenario': 'Scenario',
            'first_mes': 'First Message',
            'character_book': 'Character Book Entries',
            'alternate_greetings': 'Alternate Greetings'
        }
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        ttk.Label(main_frame, text="Select fields to extract:").grid(row=0, column=0, pady=(0, 10), sticky='w')
        
        # Scrollable checkbox container
        checkbox_container = ttk.Frame(main_frame)
        checkbox_container.grid(row=1, column=0, sticky='nsew')
        
        canvas = tk.Canvas(checkbox_container)
        scrollbar = ttk.Scrollbar(checkbox_container, orient="vertical", command=canvas.yview)
        checkbox_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create checkboxes
        for field_key, field_name in self.field_options.items():
            var = tk.BooleanVar(value=True)
            self.selected_fields[field_key] = var
            cb = ttk.Checkbutton(checkbox_frame, text=field_name, variable=var)
            cb.pack(anchor='w', pady=2)
        
        # Configure canvas
        canvas.create_window((0, 0), window=checkbox_frame, anchor='nw')
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        checkbox_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(10, 0), sticky='e')
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=(0, 5))
        ttk.Button(button_frame, text="OK", command=self.ok).pack(side=tk.RIGHT)
        
    def ok(self):
        self.destroy()
        
    def cancel(self):
        self.selected_fields = None
        self.destroy()

class LorebookFieldSelectorDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("Select Fields to Extract")
        self.selected_fields: Dict[str, tk.BooleanVar] = {}
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Center dialog
        window_width = 300
        window_height = 200
        x = (parent.winfo_screenwidth() - window_width) // 2
        y = (parent.winfo_screenheight() - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.resizable(True, True)
        
        # Lorebook field options - removed "Entry" prefix
        self.field_options = {
            'label': 'Labels',
            'content': 'Content',
            'key': 'Keys'
        }
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Label
        ttk.Label(main_frame, text="Select fields to extract:").pack(pady=(0, 10))
        
        # Create checkboxes
        for field_key, field_name in self.field_options.items():
            var = tk.BooleanVar(value=True)
            self.selected_fields[field_key] = var
            cb = ttk.Checkbutton(main_frame, text=field_name, variable=var)
            cb.pack(anchor='w', pady=2)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, pady=(10, 0))
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="OK", command=self.ok).pack(side=tk.RIGHT)
        
    def ok(self):
        self.destroy()
        
    def cancel(self):
        self.selected_fields = None
        self.destroy()

class SaveOptionsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("Save Options")
        self.save_type: Optional[str] = None
        
        self.transient(parent)
        self.grab_set()
        
        # Center dialog
        window_width = 300
        window_height = 200
        x = (parent.winfo_screenwidth() - window_width) // 2
        y = (parent.winfo_screenheight() - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Choose save format:").pack(pady=(0, 15))
        
        for save_type, text in [
            ('formatted', "Save with Formatting (TXT)"),
            ('plain', "Save Plain Text (TXT)"),
            ('pdf', "Save as PDF")
        ]:
            ttk.Button(
                main_frame,
                text=text,
                command=lambda t=save_type: self.set_save_type(t)
            ).pack(fill=tk.X, pady=5)
        
    def set_save_type(self, save_type: str):
        self.save_type = save_type
        self.destroy()

class ModernJSONExtractorGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Character Card & Lorebook Extractor")
        self.root.geometry(UI_CONFIG['WINDOW_SIZE'])
        
        # Add this line to track if we've cleaned up
        self._is_cleaned_up = False
        
        # Add cleanup when window is closed
        self.root.protocol("WM_DELETE_WINDOW", self.cleanup)

        # Data storage
        self.json_data: Optional[Dict[str, Any]] = None
        self.extracted_text: str = ""
        self.selected_fields: Optional[Dict[str, bool]] = None
        self.file_type: Optional[str] = None
        
        self.setup_styles()
        self.create_widgets()

    def create_button_frame(self, parent: ttk.Frame):
        """Create top button frame"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create import menu and button
        self.create_import_menu(button_frame)
        self.import_btn.pack(side=tk.LEFT, padx=5)
        
        # Save button remains the same
        self.save_btn = ttk.Button(
            button_frame,
            text="Save as...",
            command=self.save_txt,
            style='Modern.TButton'
        )
        self.save_btn.pack(side=tk.LEFT, padx=5)

    def setup_styles(self):
        """Configure ttk styles"""
        style = ttk.Style()
        style.configure('Modern.TButton', padding=UI_CONFIG['BUTTON_PADDING'])
        style.configure('Modern.TFrame', background=UI_CONFIG['WINDOW_BG'])
        style.configure("Custom.Vertical.TScrollbar",
            background="#E0E0E0",
            troughcolor="#F5F5F5",
            width=12
        )
        
    def create_widgets(self):
        """Create and configure GUI elements"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.create_button_frame(main_frame)
        self.create_preview_frame(main_frame)
        
    def create_import_menu(self, parent: ttk.Frame) -> None:
        """Create import button with dropdown menu"""
        self.import_menu = tk.Menu(
            parent,
            tearoff=0,
            bg=UI_CONFIG['MENU_BG'],
            fg=UI_CONFIG['MENU_FG'],
            activebackground=UI_CONFIG['MENU_ACTIVE_BG'],
            activeforeground=UI_CONFIG['MENU_ACTIVE_FG']
        )
        self.import_menu.add_command(label="Import Lorebook", command=lambda: self.import_json('lorebook'))
        self.import_menu.add_command(label="Import Card", command=lambda: self.import_json('card'))
        self.import_menu.add_command(label="Import Cards (Multiple)", command=lambda: self.import_json('multiple'))
        
        self.import_btn = ttk.Button(
            parent,
            text="Import",
            style='Modern.TButton'
        )
        self.import_btn.bind('<Button-1>', self.show_import_menu)
        
    def show_import_menu(self, event):
        """Show the import dropdown menu"""
        self.import_menu.post(event.widget.winfo_rootx(), 
                            event.widget.winfo_rooty() + event.widget.winfo_height())

        
    def create_preview_frame(self, parent: ttk.Frame):
        """Create preview area with text widget"""
        preview_frame = ttk.LabelFrame(parent, text="Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        self.setup_preview_text(preview_frame)
        self.setup_extract_button(preview_frame)
    #    
    def setup_preview_text(self, parent: ttk.Frame):
        """Configure preview text widget with all styling"""
        # Create a frame to contain both text widget and scrollbar
        preview_container = ttk.Frame(parent)
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        self.preview_text = tk.Text(
            preview_container,  # Changed parent to preview_container
            wrap=tk.WORD,
            font=UI_CONFIG['PREVIEW_FONT'],
            background=UI_CONFIG['PREVIEW_BG'],
            foreground=UI_CONFIG['PREVIEW_FG'],
            padx=15,
            pady=10,
            spacing1=2,
            spacing2=2,
            spacing3=5,
            relief=tk.FLAT,
            height=20
        )
        
        # Configure text tags
        for tag_name, style in TEXT_STYLES.items():
            self.preview_text.tag_configure(tag_name, 
                font=style.font,
                foreground=style.foreground,
                spacing1=style.spacing1,
                spacing2=style.spacing2,
                spacing3=style.spacing3,
                justify=style.justify,
                lmargin1=style.lmargin1,
                lmargin2=style.lmargin2
            )
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(
            preview_container,  # Changed parent to preview_container
            orient=tk.VERTICAL,
            command=self.preview_text.yview,
            style="Custom.Vertical.TScrollbar"
        )
        self.preview_text.configure(yscrollcommand=scrollbar.set)
        
        # Use grid instead of pack for more precise control
        self.preview_text.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        # Configure grid weights
        preview_container.grid_columnconfigure(0, weight=1)
        preview_container.grid_rowconfigure(0, weight=1)
        
    def setup_extract_button(self, parent: ttk.Frame):
        """Create and configure extract button"""
        self.extract_frame = ttk.Frame(parent)
        self.extract_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.extract_btn = tk.Button(
            self.extract_frame,
            text="Extract Fields",
            command=self.extract_fields,
            font=('Segoe UI', 10),
            relief='flat',
            cursor='hand2',
            padx=10,
            pady=5,
            bg=UI_CONFIG['ACTIVE_BTN_BG'],
            fg='white',
            activebackground='#3498db',
            activeforeground='white',
            disabledforeground='#95a5a6'
        )

    def get_field_value(self, field_name: str) -> Optional[Any]:
        """Safely get field value from JSON data with enhanced nested searching"""
        print(f"Debug: Getting field value for {field_name}")
        
        def search_dict(d: Dict[str, Any], field: str) -> Optional[Any]:
            """Recursively search through nested dictionaries for a field"""
            # Direct check
            if field in d:
                print(f"Debug: Found {field} in direct check")
                return d[field]
                
            # Search nested dictionaries
            for key, value in d.items():
                if isinstance(value, dict):
                    print(f"Debug: Searching in nested dict {key}")
                    result = search_dict(value, field)
                    if result is not None:
                        return result
                        
                # Special handling for extensions/depth_prompt/prompt
                elif key == 'extensions' and isinstance(value, dict):
                    if 'depth_prompt' in value and isinstance(value['depth_prompt'], dict):
                        if 'prompt' in value['depth_prompt']:
                            if field == 'character_note':  # This is typically stored in depth_prompt
                                print("Debug: Found character note in depth_prompt")
                                return value['depth_prompt']['prompt']
        
        if not self.json_data:
            print("Debug: No JSON data available")
            return None
        
        # Try direct access first
        value = search_dict(self.json_data, field_name)
        if value is not None:
            return value
            
        # Then check in 'data' field if it exists
        if 'data' in self.json_data and isinstance(self.json_data['data'], dict):
            value = search_dict(self.json_data['data'], field_name)
            if value is not None:
                return value
        
        print(f"Debug: No value found for {field_name}")
        return None

    def import_json(self, import_type: str = 'card'):
        """Handle JSON file import based on type"""
        try:
            self.reset_interface()
            
            if import_type == 'multiple':
                file_paths = filedialog.askopenfilenames(
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
                if not file_paths:
                    return
                    
                # Process multiple files
                valid_cards = []
                invalid_files = []
                
                print(f"Debug: Selected {len(file_paths)} files")
                
                for file_path in file_paths:
                    try:
                        print(f"Debug: Processing file: {Path(file_path).name}")
                        
                        with open(file_path, 'r', encoding='utf-8') as file:
                            json_data = json.load(file)
                            
                        print("Debug: Successfully loaded JSON")
                        file_type = detect_file_type(json_data)
                        print(f"Debug: Detected type: {file_type}")
                        
                        if file_type == 'character':
                            # Normalize the data structure between v2 and v3
                            if 'data' in json_data and isinstance(json_data['data'], dict):
                                # v3 format - extract data from nested structure
                                char_data = json_data['data']
                                print("Debug: Found v3 character card structure")
                            else:
                                # v2 format - use as is
                                char_data = json_data
                                print("Debug: Found v2 character card structure")
                            
                            # Get character name
                            name = char_data.get('name')
                            if not name:
                                name = Path(file_path).stem
                                
                            print(f"Debug: Found character name: {name}")
                            valid_cards.append((name, char_data))
                        else:
                            print(f"Debug: Invalid file type: {file_type}")
                            invalid_files.append(f"{Path(file_path).name} (Not a character card)")
                            
                    except json.JSONDecodeError as e:
                        print(f"Debug: JSON decode error: {str(e)}")
                        invalid_files.append(f"{Path(file_path).name} (Invalid JSON format)")
                    except Exception as e:
                        print(f"Debug: Error processing file: {str(e)}")
                        invalid_files.append(f"{Path(file_path).name} (Error: {str(e)})")
                
                if invalid_files:
                    error_message = "The following files could not be processed:\n\n" + "\n".join(invalid_files)
                    if valid_cards:
                        error_message += "\n\nContinuing with valid cards..."
                    messagebox.showwarning("Import Warning", error_message)
                    
                    if not valid_cards:
                        print("Debug: No valid cards found")
                        self.reset_interface()
                        return
                
                # Sort cards by name
                print(f"Debug: Sorting {len(valid_cards)} valid cards")
                valid_cards.sort(key=lambda x: x[0].lower())
                
                # Store combined data and set type
                self.json_data = [card[1] for card in valid_cards]
                self.file_type = 'character_multiple'
                print("Debug: Successfully processed multiple cards")
                print(f"Debug: Number of cards stored: {len(self.json_data)}")

            else:
                # Single file import
                file_path = filedialog.askopenfilename(
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
                
                if not file_path:
                    return
                    
                with open(file_path, 'r', encoding='utf-8') as file:
                    self.json_data = json.load(file)
                
                self.file_type = detect_file_type(self.json_data)
                if import_type == 'lorebook' and self.file_type != 'lorebook':
                    messagebox.showerror("Error", "Selected file is not a valid lorebook.")
                    self.reset_interface()
                    return
                elif import_type == 'card' and self.file_type != 'character':
                    messagebox.showerror("Error", "Selected file is not a valid character card.")
                    self.reset_interface()
                    return
            
            self.handle_file_import()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import file(s):\n{str(e)}")
            self.reset_interface()

    def handle_file_import(self):
        """Process imported file based on type"""
        try:
            if self.file_type not in ['character', 'lorebook', 'character_multiple']:
                messagebox.showerror("Error", "Unable to determine file type")
                self.reset_interface()
                return

            # Create appropriate dialog based on file type
            if self.file_type == 'character' or self.file_type == 'character_multiple':
                dialog = FieldSelectorDialog(self.root)
            else:  # lorebook
                dialog = LorebookFieldSelectorDialog(self.root)
                
            self.root.wait_window(dialog)

            if dialog.selected_fields:
                self.selected_fields = {k: v.get() for k, v in dialog.selected_fields.items()}
                self.update_preview_with_selected_fields(dialog.field_options)
                self.enable_extract_button()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file: {str(e)}")
            self.reset_interface()

    def cleanup(self):
        """Clean up resources before closing"""
        if self._is_cleaned_up:
            return
            
        try:
            # Clear text widget
            if hasattr(self, 'preview_text'):
                self.preview_text.configure(state='normal')
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.configure(state='disabled')
            
            # Clear stored data
            self.extracted_text = ""
            if hasattr(self, 'json_data'):
                del self.json_data
            
            # Clear selected fields
            if hasattr(self, 'selected_fields'):
                self.selected_fields = None
            
            # Mark as cleaned up
            self._is_cleaned_up = True
            
            # Destroy the root window if it exists
            if hasattr(self, 'root') and self.root:
                self.root.destroy()
                
        except Exception as e:
            # If something goes wrong during cleanup, log it
            print(f"Error during cleanup: {str(e)}")
            # Still try to destroy the window
            try:
                self.root.destroy()
            except:
                pass

    def reset_interface(self):
        """Reset interface to initial state"""
        self.preview_text.configure(state='normal')
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.configure(state='disabled')
        
        # Clear any stored data
        self.extracted_text = ""
        if hasattr(self, 'json_data'):
            del self.json_data
        if hasattr(self, 'extract_btn'):
            self.extract_btn.pack_forget()
        
        self.json_data = None
        self.selected_fields = None
        self.file_type = None
        self.extracted_text = ""

    def update_preview_with_selected_fields(self, field_options: Dict[str, str]):
        """Update preview with selected fields"""
        self.preview_text.configure(state='normal')
        self.preview_text.delete(1.0, tk.END)
        
        selected = [field_options[k] for k, v in self.selected_fields.items() 
                   if v and k != 'name' and k in field_options]
        preview_text = "Fields selected for extraction:\n\n" + "\n".join(f"• {field}" for field in selected)
        self.preview_text.insert(tk.END, preview_text)
        
        self.preview_text.configure(state='disabled')

    def enable_extract_button(self):
        """Show and enable the extract button"""
        self.extract_btn.configure(
            bg=UI_CONFIG['ACTIVE_BTN_BG'],
            state='normal',
            cursor='hand2'
        )
        self.extract_btn.pack(fill=tk.X, padx=5, pady=5)

    def disable_extract_button(self):
        """Disable the extract button after use"""
        self.extract_btn.configure(
            bg=UI_CONFIG['DISABLED_BTN_BG'],
            state='disabled',
            cursor='arrow'
        )

    def extract_character_fields(self):
        """Extract and format character card fields"""
        print("\nDebug: Starting character field extraction")
        self.preview_text.configure(state='normal')
        
        # Add header
        card_name = self.get_field_value('name') or "Unnamed Character"
        print(f"Debug: Extracting card for {card_name}")
        
        header = f"""{'='*50}
        {card_name}
        {'='*50}\n"""
        self.preview_text.insert(tk.END, header, 'header')
        
        # Process main sections with debugging
        sections = [
            ('description', 'CHARACTER DEFINITION'),
            ('personality', 'PERSONALITY'),
            ('prompt', 'CHARACTER NOTE'),
            ('scenario', 'SCENARIO'),
            ('first_mes', 'FIRST MESSAGE'),
            ('mes_example', 'EXAMPLE MESSAGES')
        ]
        
        for field, title in sections:
            if self.selected_fields.get(field):
                print(f"Debug: Processing section {field}")
                content = self.get_field_value(field)
                if content:
                    print(f"Debug: Found content for {field}")
                    separator = '─' * 50
                    self.preview_text.insert(tk.END, f"\n{title}\n", 'section_title')
                    self.preview_text.insert(tk.END, f"{separator}\n", 'separator')
                    self.preview_text.insert(tk.END, f"{content}\n", 'content')
                else:
                    print(f"Debug: No content found for {field}")
        
        # Handle special sections
        if self.selected_fields.get('alternate_greetings'):
            print("Debug: Processing alternate greetings")
            self.handle_alternate_greetings()
        
        if self.selected_fields.get('character_book'):
            print("Debug: Processing character book")
            self.handle_character_book()

    def format_section(self, field: str, title: str):
        """Format a standard section"""
        if not self.selected_fields.get(field):
            return
            
        content = self.get_field_value(field)
        if not content:
            return
            
        separator = '─' * 50
        self.preview_text.insert(tk.END, f"\n{title}\n", 'section_title')
        self.preview_text.insert(tk.END, f"{separator}\n", 'separator')
        self.preview_text.insert(tk.END, f"{content}\n", 'content')

    def handle_alternate_greetings(self):
        """Handle alternate greetings section"""
        if not self.selected_fields.get('alternate_greetings'):
            return
            
        greetings = self.get_field_value('alternate_greetings')
        if not greetings or not isinstance(greetings, list):
            return
            
        self.preview_text.insert(tk.END, "\nALTERNATE GREETINGS\n", 'section_title')
        self.preview_text.insert(tk.END, "─" * 50 + "\n", 'separator')
        
        for i, greeting in enumerate(greetings, 1):
            greeting = greeting.strip()
            self.preview_text.insert(tk.END, f"► Greeting {i}\n", 'book_entry')
            self.preview_text.insert(tk.END, f"{greeting}\n", 'content')

    def handle_character_book(self):
        """Handle character book section"""
        if not self.selected_fields.get('character_book'):
            return
            
        char_book = self.get_field_value('character_book')
        if not char_book or 'entries' not in char_book:
            return
            
        self.preview_text.insert(tk.END, "\nCHARACTER BOOK\n", 'section_title')
        self.preview_text.insert(tk.END, "─" * 50 + "\n", 'separator')
        
        for entry in char_book['entries']:
            title = entry.get('name', entry.get('comment', ''))
            content = entry.get('content', '')
            if title and content:
                self.preview_text.insert(tk.END, f"► {title}\n", 'book_entry')
                self.preview_text.insert(tk.END, f"{content}\n\n", 'content')

    def extract_fields(self):
        """Extract fields based on file type"""
        print("Debug: Starting field extraction")
        print(f"Debug: File type is {self.file_type}")
        
        if not self.json_data or not self.selected_fields:
            messagebox.showwarning("Warning", "Please import file(s) and select fields first!")
            return
            
        try:
            self.preview_text.configure(state='normal')
            self.preview_text.delete(1.0, tk.END)
            
            if self.file_type == 'character_multiple':
                print(f"Debug: Processing multiple characters, count: {len(self.json_data)}")
                for i, card_data in enumerate(self.json_data):
                    print(f"Debug: Processing card {i + 1}")
                    if i > 0:
                        # Add separator between cards
                        self.preview_text.insert(tk.END, "\n\n" + "="*60 + "\n\n", 'separator')
                    
                    # Temporarily set json_data to current card
                    temp_data = self.json_data
                    self.json_data = card_data
                    
                    # Extract current card
                    self.extract_character_fields()
                    
                    # Restore json_data
                    self.json_data = temp_data
                    
                print("Debug: Completed multiple character extraction")
            elif self.file_type == 'character':
                print("Debug: Processing single character")
                self.extract_character_fields()
            elif self.file_type == 'lorebook':
                print("Debug: Processing lorebook")
                self.extract_lorebook_fields()
                    
            # Store extracted text and disable editing
            self.extracted_text = self.preview_text.get(1.0, tk.END)
            self.preview_text.configure(state='disabled')
            self.disable_extract_button()
                
        except Exception as e:
            print(f"Debug: Error during extraction: {str(e)}")
            messagebox.showerror("Error", f"Failed to extract fields:\n{str(e)}")
            self.reset_interface()

    def extract_multiple_character_fields(self):
        """Extract and format multiple character cards"""
        self.preview_text.configure(state='normal')
        self.preview_text.delete(1.0, tk.END)
        
        # Process each card
        for i, card_data in enumerate(self.json_data):
            if i > 0:
                # Add separator between cards
                self.preview_text.insert(tk.END, "\n\n" + "="*60 + "\n\n", 'separator')
            
            # Temporarily set json_data to current card
            temp_data = self.json_data
            self.json_data = card_data
            
            # Extract current card
            self.extract_character_fields()
            
            # Restore json_data
            self.json_data = temp_data
        
        # Store extracted text and disable editing
        self.extracted_text = self.preview_text.get(1.0, tk.END)
        self.preview_text.configure(state='disabled')

    def extract_lorebook_fields(self):
        """Extract and format lorebook fields"""
        self.preview_text.configure(state='normal')
        self.preview_text.delete(1.0, tk.END)
        
        if not self.json_data or 'entries' not in self.json_data:
            messagebox.showerror("Error", "Invalid lorebook format")
            self.reset_interface()
            return
            
        # Process entries
        first_entry = True
        for entry_id, entry in self.json_data['entries'].items():
            if not first_entry:
                self.preview_text.insert(tk.END, "\n\n")
            first_entry = False
            
            if self.selected_fields.get('label'):
                # Try to get label from either 'name' or 'comment' field
                label = entry.get('name', entry.get('comment', ''))
                if label:
                    self.preview_text.insert(tk.END, f"{label}\n", 'section_title')
                    self.preview_text.insert(tk.END, "─" * 50 + "\n", 'separator')
                    
            if self.selected_fields.get('content'):
                content = entry.get('content', '')
                if content:
                    self.preview_text.insert(tk.END, f"{content}\n", 'content')
                    
            if self.selected_fields.get('key'):
                keys = entry.get('key', [])
                if keys:
                    key_text = ', '.join(keys) if isinstance(keys, list) else str(keys)
                    self.preview_text.insert(tk.END, f"Keys: {key_text}\n", 'book_entry')
        
        # Store the text and disable editing
        self.extracted_text = self.preview_text.get(1.0, tk.END)
        self.preview_text.configure(state='disabled')

    def save_file(self, file_path: str, save_type: str):
        """Save the file in the specified format with enhanced error handling"""
        try:
            if save_type == 'pdf':
                try:
                    create_pdf(self.extracted_text, file_path)
                except PermissionError as e:
                    messagebox.showerror("Permission Error", 
                        "Unable to save PDF file. Please check if:\n"
                        "- You have write permission for the selected location\n"
                        "- The file is not currently open in another program")
                    return
                except IOError as e:
                    messagebox.showerror("I/O Error", 
                        "Failed to save PDF file. Please check if:\n"
                        "- You have enough disk space\n"
                        "- The selected location is accessible")
                    return
                except Exception as e:
                    messagebox.showerror("PDF Creation Error", 
                        f"Failed to create PDF file:\n{str(e)}\n\n"
                        "Try saving as text format instead.")
                    return
            else:
                text_to_save = self.extracted_text if save_type == 'formatted' else get_plain_text(self.extracted_text)
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(text_to_save)
                    
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save file:\n{str(e)}")

    def save_txt(self):
        """Handle saving the extracted text"""
        if not self.extracted_text:
            messagebox.showwarning("Warning", "No extracted text to save!")
            return
            
        # Show save options dialog
        dialog = SaveOptionsDialog(self.root)
        self.root.wait_window(dialog)
        
        if not dialog.save_type:
            return
            
        try:
            # Set up file dialog based on save type
            if dialog.save_type == 'pdf':
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
                )
            else:
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".txt",
                    filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
                )
                
            if not file_path:
                return
                
            self.save_file(file_path, dialog.save_type)
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

    def __del__(self):
        """Destructor - called when the object is about to be destroyed"""
        self.cleanup()

if __name__ == "__main__":
    root = tk.Tk()
    root.configure(bg=UI_CONFIG['WINDOW_BG'])
    app = ModernJSONExtractorGUI(root)
    root.mainloop()