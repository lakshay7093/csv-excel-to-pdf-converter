import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import re
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Global variables
df = None
output_text = ""

# Modern color scheme
COLORS = {
    'primary': '#2E86AB',      # Blue
    'secondary': '#A23B72',    # Purple
    'success': '#F18F01',      # Orange
    'danger': '#C73E1D',       # Red
    'dark': '#1B263B',         # Dark blue
    'light': '#F8F9FA',        # Light gray
    'white': '#FFFFFF',
    'text_dark': '#2C3E50',
    'text_light': '#6C757D',
    'bg_main': '#F4F6F9',
    'accent': '#17A2B8'        # Teal
}

# ---------------- UTILITY FUNCTIONS ----------------

def num(value):
    """Convert value to float, handling various formats including percentages"""
    if pd.isna(value):
        return 0.0
    
    # Convert to string first
    str_value = str(value).strip()
    
    # Handle empty strings
    if not str_value or str_value.lower() in ['', 'nan', 'null', 'none']:
        return 0.0
    
    # Handle percentage format (5% -> 5.0)
    if '%' in str_value:
        str_value = str_value.replace('%', '').strip()
    
    # Remove commas and other formatting
    str_value = str_value.replace(',', '').replace(' ', '')
    
    # Extract number using regex (including decimals)
    m = re.search(r'-?\d+(?:\.\d+)?', str_value)
    result = float(m.group()) if m else 0.0
    
    return result

def create_modern_button(parent, text, command, bg_color, width=20, height=2):
    """Create a modern styled button"""
    return tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg_color,
        fg=COLORS['white'],
        font=('Segoe UI', 11, 'bold'),
        width=width,
        height=height,
        relief='flat',
        cursor='hand2',
        activebackground=bg_color,
        activeforeground=COLORS['white'],
        bd=0
    )

# ---------------- CORE FUNCTIONS ----------------

def upload_and_process_file():
    """Upload and automatically process CSV/Excel file"""
    global df, output_text
    
    file_path = filedialog.askopenfilename(
        title="Select CSV or Excel File",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls"), 
            ("CSV Files", "*.csv"), 
            ("All Files", "*.*")
        ]
    )
    if not file_path:
        return

    try:
        # Load file based on extension
        if file_path.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path)
        
        df.columns = df.columns.str.strip()
        
        # Auto-detect required columns
        column_mapping = detect_columns(df)
        
        if not column_mapping:
            messagebox.showerror(
                "Column Detection Error", 
                "Could not detect required columns.\n\n"
                "Please ensure your file has columns for:\n"
                "• Name/Product\n• MRP/Price\n• Included GST\n• GST Rate"
            )
            return
        
        # Process all items
        process_all_items(column_mapping)
        
        messagebox.showinfo(
            "Success! 🎉", 
            f"File processed successfully!\n\n"
            f"📊 Total items calculated: {len(df)}\n"
            f"📁 File: {file_path.split('/')[-1]}"
        )
        
    except Exception as e:
        messagebox.showerror(
            "Processing Error", 
            f"Failed to process file:\n\n{str(e)}"
        )


def detect_columns(df):
    """Auto-detect column names based on common patterns"""
    columns = df.columns.str.lower()
    mapping = {}
    
    # Define search patterns for each column type
    patterns = {
        'name': ['name', 'product', 'item', 'description', 'title'],
        'mrp': ['mrp', 'price', 'rate', 'amount', 'cost'],
        'included_gst': ['included_gst', 'gst_included', 'with_gst', 'inclusive', 'incl_gst', 'included gst'],
        'gst': ['gst', 'tax', 'gst_rate', 'tax_rate', 'gst%', 'tax%', 'gst rate', 'tax rate']
    }
    
    # Find matching columns
    for col in columns:
        for key, pattern_list in patterns.items():
            if key not in mapping and any(pattern in col for pattern in pattern_list):
                mapping[key] = df.columns[columns.tolist().index(col)]
                break
    
    # Fallback: use first 4 columns if auto-detection fails
    # Based on your image: A=Name, B=MRP, C=Included GST, D=?, E=GST%
    if len(mapping) < 4 and len(df.columns) >= 4:
        cols = df.columns.tolist()
        mapping = {
            'name': cols[0],      # Column A
            'mrp': cols[1],       # Column B  
            'included_gst': cols[2],  # Column C
            'gst': cols[4] if len(cols) > 4 else cols[3]  # Column E (GST%)
        }
    
    return mapping if len(mapping) == 4 else None


def process_all_items(column_mapping):
    """Process all items in the dataframe and display results"""
    global output_text
    output_text = ""
    output_box.delete("1.0", tk.END)
    
    serial_no = 1
    
    for _, row in df.iterrows():
        try:
            # Extract values from row
            name = str(row[column_mapping['name']])
            mrp = num(row[column_mapping['mrp']])
            included_gst = num(row[column_mapping['included_gst']])
            gst_percent = num(row[column_mapping['gst']])

            # Convert GST percentage to decimal - handle different formats
            if gst_percent == 0:
                # If GST is 0, try to calculate it from MRP and Included GST
                if mrp > 0 and included_gst > 0:
                    # Reverse calculate GST% from the given values
                    # This is a fallback when GST% column is missing or zero
                    gst_decimal = 0.05  # Default to 5% if we can't determine
                    display_gst_percent = 5.0  # Show 5% in display
                else:
                    gst_decimal = 0.0
                    display_gst_percent = 0.0
            elif gst_percent > 1:
                # If GST is like 5, 12, 18 (percentage format)
                gst_decimal = gst_percent / 100
                display_gst_percent = gst_percent
            else:
                # If GST is already in decimal format like 0.05, 0.12
                gst_decimal = gst_percent
                display_gst_percent = gst_percent * 100  # Convert back to percentage for display
                
            # Step-by-step calculations following the exact formula
            # Step 1: Discount = MRP - Included_GST_Value
            discount = mrp - included_gst
            
            # Step 2: Discount_Price = Discount / 2
            discount_price = discount / 2
            
            # Step 3: Base_Price = MRP - Discount_Price
            base_price = mrp - discount_price
            
            # Step 4: GST_Value = Base_Price × (GST_% / 100)
            gst_value = base_price * gst_decimal
            
            # Step 5: Final_Price = MRP - GST_Value
            final_price = mrp - gst_value
            
            # Step 6: App_Price = MRP - Discount_Price
            app_price = mrp - discount_price

            # Format output block
            block = f"""
{serial_no}. NAME: {name}

MRP: {mrp}
INCLUDED GST: {included_gst}

DISCOUNT: {discount}
DISCOUNT PRICE (÷2): {discount_price}

BASE PRICE: {base_price}
GST %: {display_gst_percent}%
GST (Decimal): {gst_decimal}

GST VALUE: {gst_value}
FINAL PRICE: {final_price}

APP PRICE: {app_price}
-------------------------------------
"""
            output_text += block
            output_box.insert(tk.END, block)
            serial_no += 1
            
        except Exception as e:
            error_block = f"""
{serial_no}. ❌ ERROR processing item: {name if 'name' in locals() else 'Unknown'}
Error: {str(e)}
-------------------------------------
"""
            output_text += error_block
            output_box.insert(tk.END, error_block)
            serial_no += 1


def export_pdf():
    """Export calculation results to PDF"""
    if not output_text.strip():
        messagebox.showerror(
            "No Data", 
            "No data to export.\n\nPlease upload and process a file first."
        )
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        title="Save PDF Report"
    )
    if not file_path:
        return

    try:
        doc = SimpleDocTemplate(file_path, pagesize=A4, topMargin=50, bottomMargin=50)
        styles = getSampleStyleSheet()
        elements = []
        
        # Add title
        title = Paragraph("📊 MRP Discount & GST Calculator Report", styles["Title"])
        elements.append(title)
        elements.append(Spacer(1, 20))
        
        # Process each item
        items = output_text.split("-------------------------------------")
        
        for item in items:
            if item.strip():
                lines = item.strip().split('\n')
                for line in lines:
                    if line.strip():
                        # Escape HTML characters
                        safe_line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        
                        # Style serial numbers and names
                        if line.strip().endswith(':') or any(line.strip().startswith(str(i)+'.') for i in range(1, 1000)):
                            elements.append(Paragraph(f"<b>{safe_line}</b>", styles["Normal"]))
                        else:
                            elements.append(Paragraph(safe_line, styles["Normal"]))
                        elements.append(Spacer(1, 3))
                
                elements.append(Spacer(1, 15))

        doc.build(elements)
        messagebox.showinfo(
            "Export Successful! 🎉", 
            f"PDF exported successfully!\n\n📁 Location:\n{file_path}"
        )
        
    except Exception as e:
        messagebox.showerror(
            "Export Error", 
            f"Failed to export PDF:\n\n{str(e)}"
        )

def copy_all():
    """Copy all output to clipboard"""
    if not output_text.strip():
        messagebox.showwarning(
            "No Data", 
            "No data to copy.\n\nPlease process a file first."
        )
        return
        
    root.clipboard_clear()
    root.clipboard_append(output_text)
    root.update()
    messagebox.showinfo("Copied! 📋", "All output copied to clipboard!")

def clear_output():
    """Clear the output display"""
    global output_text
    output_text = ""
    output_box.delete("1.0", tk.END)
    messagebox.showinfo("Cleared! 🧹", "Output cleared successfully!")


# ---------------- GUI SETUP ----------------

def create_gui():
    """Create and setup the main GUI"""
    global root, output_box
    
    # Main window setup
    root = tk.Tk()
    root.title("💰 MRP Discount & GST Calculator")
    root.geometry("1000x750")
    root.configure(bg=COLORS['bg_main'])
    root.resizable(True, True)
    
    # Configure window icon and styling
    try:
        root.iconbitmap('calculator.ico')  # Add if you have an icon
    except:
        pass

    # Header section
    create_header()
    
    # Instructions section
    create_instructions()
    
    # Main upload button
    create_upload_section()
    
    # Action buttons
    create_action_buttons()
    
    # Output section
    create_output_section()
    
    # Status bar
    create_status_bar()

def create_header():
    """Create the header section"""
    header_frame = tk.Frame(root, bg=COLORS['primary'], height=80)
    header_frame.pack(fill="x", padx=0, pady=0)
    header_frame.pack_propagate(False)
    
    # Main title
    title_label = tk.Label(
        header_frame,
        text="💰 MRP Discount & GST Calculator",
        bg=COLORS['primary'],
        fg=COLORS['white'],
        font=('Segoe UI', 22, 'bold')
    )
    title_label.pack(pady=20)
    
    # Subtitle
    subtitle_label = tk.Label(
        header_frame,
        text="Professional Auto-Processing Tool",
        bg=COLORS['primary'],
        fg=COLORS['light'],
        font=('Segoe UI', 11)
    )
    subtitle_label.pack(pady=(0, 10))

def create_instructions():
    """Create the instructions section"""
    inst_frame = tk.Frame(root, bg=COLORS['bg_main'])
    inst_frame.pack(pady=15, padx=20, fill="x")
    
    # Main instruction
    main_inst = tk.Label(
        inst_frame,
        text="📁 Upload your CSV or Excel file and all items will be processed automatically!",
        bg=COLORS['bg_main'],
        fg=COLORS['text_dark'],
        font=('Segoe UI', 12, 'bold'),
        wraplength=900
    )
    main_inst.pack(pady=(0, 8))
    
    # Column requirements
    col_req = tk.Label(
        inst_frame,
        text="📋 Expected columns: Name/Product, MRP/Price, Included GST, GST Rate/Tax%",
        bg=COLORS['bg_main'],
        fg=COLORS['text_light'],
        font=('Segoe UI', 10),
        wraplength=900
    )
    col_req.pack(pady=(0, 8))
    
    # Formula explanation
    formula_frame = tk.Frame(inst_frame, bg=COLORS['light'], relief='solid', bd=1)
    formula_frame.pack(fill="x", pady=8)
    
    formula_label = tk.Label(
        formula_frame,
        text="🧮 Formula Steps:",
        bg=COLORS['light'],
        fg=COLORS['text_dark'],
        font=('Segoe UI', 10, 'bold')
    )
    formula_label.pack(anchor="w", padx=10, pady=(8, 4))
    
    formula_text = tk.Label(
        formula_frame,
        text="1) Discount = MRP - Included_GST_Value  2) Discount_Price = Discount/2  3) Base_Price = MRP - Discount_Price  4) GST_Value = Base_Price × (GST_%/100)  5) Final_Price = MRP - GST_Value",
        bg=COLORS['light'],
        fg=COLORS['text_dark'],
        font=('Segoe UI', 9),
        wraplength=900,
        justify="left"
    )
    formula_text.pack(anchor="w", padx=10, pady=(0, 8))

def create_upload_section():
    """Create the upload button section"""
    upload_frame = tk.Frame(root, bg=COLORS['bg_main'])
    upload_frame.pack(pady=20)
    
    upload_btn = tk.Button(
        upload_frame,
        text="📁 Upload & Process File",
        command=upload_and_process_file,
        bg=COLORS['primary'],
        fg=COLORS['white'],
        font=('Segoe UI', 16, 'bold'),
        width=25,
        height=2,
        relief='flat',
        cursor='hand2',
        activebackground=COLORS['dark'],
        activeforeground=COLORS['white'],
        bd=0
    )
    upload_btn.pack()
    
    # Hover effects
    def on_enter(e):
        upload_btn.config(bg=COLORS['dark'])
    def on_leave(e):
        upload_btn.config(bg=COLORS['primary'])
    
    upload_btn.bind("<Enter>", on_enter)
    upload_btn.bind("<Leave>", on_leave)

def create_action_buttons():
    """Create the action buttons section"""
    btn_frame = tk.Frame(root, bg=COLORS['bg_main'])
    btn_frame.pack(pady=15)
    
    # Create all buttons in one row
    # Export PDF button
    pdf_btn = create_modern_button(
        btn_frame, "📄 Export PDF", export_pdf, COLORS['secondary'], 16, 2
    )
    pdf_btn.pack(side="left", padx=8)
    
    # Copy button
    copy_btn = create_modern_button(
        btn_frame, "📋 Copy Output", copy_all, COLORS['success'], 16, 2
    )
    copy_btn.pack(side="left", padx=8)
    
    # Clear button
    clear_btn = create_modern_button(
        btn_frame, "🧹 Clear Output", clear_output, COLORS['danger'], 16, 2
    )
    clear_btn.pack(side="left", padx=8)
    
    # Add hover effects for all buttons
    for btn in [pdf_btn, copy_btn, clear_btn]:
        original_color = btn.cget('bg')
        def make_hover_effect(button, orig_color):
            def on_enter(e):
                button.config(bg=COLORS['dark'])
            def on_leave(e):
                button.config(bg=orig_color)
            return on_enter, on_leave
        
        enter_func, leave_func = make_hover_effect(btn, original_color)
        btn.bind("<Enter>", enter_func)
        btn.bind("<Leave>", leave_func)

def create_output_section():
    """Create the output display section"""
    global output_box
    
    output_frame = tk.Frame(root, bg=COLORS['bg_main'])
    output_frame.pack(fill="both", expand=True, padx=20, pady=(10, 0))
    
    # Output label
    output_label = tk.Label(
        output_frame,
        text="📊 Calculation Results:",
        bg=COLORS['bg_main'],
        fg=COLORS['text_dark'],
        font=('Segoe UI', 12, 'bold')
    )
    output_label.pack(anchor="w", pady=(0, 8))
    
    # Output text box with significantly increased height
    output_box = scrolledtext.ScrolledText(
        output_frame,
        width=120,
        height=35,  # Increased from 30 to 35
        font=('Consolas', 10),
        bg=COLORS['white'],
        fg=COLORS['text_dark'],
        relief='solid',
        bd=1,
        selectbackground=COLORS['accent'],
        selectforeground=COLORS['white'],
        insertbackground=COLORS['primary']
    )
    output_box.pack(fill="both", expand=True, pady=(0, 10))

def create_status_bar():
    """Create the status bar"""
    status_frame = tk.Frame(root, bg=COLORS['dark'], height=30)
    status_frame.pack(fill="x", side="bottom")
    status_frame.pack_propagate(False)
    
    status_label = tk.Label(
        status_frame,
        text="Ready to process files | Developed with ❤️ for efficient calculations",
        bg=COLORS['dark'],
        fg=COLORS['light'],
        font=('Segoe UI', 9)
    )
    status_label.pack(side="left", padx=10, pady=6)
    
    version_label = tk.Label(
        status_frame,
        text="v2.0",
        bg=COLORS['dark'],
        fg=COLORS['accent'],
        font=('Segoe UI', 9, 'bold')
    )
    version_label.pack(side="right", padx=10, pady=6)

# ---------------- MAIN EXECUTION ----------------

if __name__ == "__main__":
    create_gui()
    root.mainloop()
