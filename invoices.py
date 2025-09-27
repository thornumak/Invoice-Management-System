#!/usr/bin/env python3

import os
import re
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
import PyPDF2
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import json

class InvoiceManager:
    def __init__(self, db_file="invoices.db"):
        self.db_file = Path(db_file)
        self.config_file = "invoice_config.json"
        self.load_config()
        self.init_database()
    
    def load_config(self):
        try:
            if Path(self.config_file).exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.last_folder = config.get('last_folder', str(Path.home()))
            else:
                self.last_folder = str(Path.home())
        except:
            self.last_folder = str(Path.home())
    
    def save_config(self):
        config = {'last_folder': self.last_folder}
        with open(self.config_file, 'w') as f:
            json.dump(config, f)
            
    def get_currencies(self):
        """Returns a list of common currency codes."""
        return ["EUR", "USD", "GBP", "JPY", "AUD", "CAD", "CHF"]

    def init_database(self):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        # Check if currency column exists in patterns table and add if not
        cursor.execute("PRAGMA table_info(invoice_patterns)")
        pattern_columns = [column[1] for column in cursor.fetchall()]
        if 'currency' not in pattern_columns:
            cursor.execute("ALTER TABLE invoice_patterns ADD COLUMN currency TEXT DEFAULT 'EUR'")

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS invoice_patterns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pattern TEXT UNIQUE NOT NULL,
                company_name TEXT NOT NULL,
                payment_method TEXT NOT NULL,
                payment_terms_days INTEGER NOT NULL,
                currency TEXT DEFAULT 'EUR',
                created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Check if currency column exists in invoices table and add if not
        cursor.execute("PRAGMA table_info(invoices)")
        invoice_columns = [column[1] for column in cursor.fetchall()]
        if 'currency' not in invoice_columns:
             cursor.execute("ALTER TABLE invoices ADD COLUMN currency TEXT DEFAULT 'EUR'")

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT UNIQUE NOT NULL,
                company_name TEXT NOT NULL,
                amount REAL NOT NULL,
                currency TEXT DEFAULT 'EUR',
                invoice_date DATE NOT NULL,
                due_date DATE NOT NULL,
                payment_method TEXT NOT NULL,
                payment_terms_days INTEGER NOT NULL,
                status TEXT DEFAULT 'Sent',
                file_path TEXT,
                notes TEXT,
                created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
        print(f"Database initialized: {self.db_file}")
    
    def extract_invoice_pattern(self, invoice_number):
        if not invoice_number:
            return None
        
        clean_invoice = invoice_number.strip().lower()
        pattern_match = re.match(r'^([a-zA-Z]+)[-_]?\d+', clean_invoice)
        if pattern_match:
            return pattern_match.group(1).lower()
        
        letters_only = re.findall(r'[a-zA-Z]+', clean_invoice)
        if letters_only:
            return letters_only[0].lower()
        
        return None
    
    def get_company_info_by_pattern(self, invoice_number):
        pattern = self.extract_invoice_pattern(invoice_number)
        if not pattern:
            return None
        
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT company_name, payment_method, payment_terms_days, currency 
            FROM invoice_patterns 
            WHERE pattern = ?
        ''', (pattern,))
        result = cursor.fetchone()
        conn.close()
        
        if result:
            return {
                'company_name': result[0],
                'payment_method': result[1],
                'payment_terms_days': result[2],
                'currency': result[3]
            }
        return None
    
    def add_or_update_pattern(self, pattern, company_name, payment_method, payment_terms_days, currency):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT OR REPLACE INTO invoice_patterns 
                (pattern, company_name, payment_method, payment_terms_days, currency)
                VALUES (?, ?, ?, ?, ?)
            ''', (pattern.lower(), company_name, payment_method, payment_terms_days, currency))
            conn.commit()
            conn.close()
            return True, "Pattern saved successfully"
        except Exception as e:
            conn.close()
            return False, f"Error saving pattern: {e}"
    
    def extract_text_from_pdf(self, file_path):
        try:
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text()
                return text
        except Exception as e:
            return ""
    
    def extract_text_from_docx(self, file_path):
        try:
            doc = Document(file_path)
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text
        except Exception as e:
            return ""
    
    def clean_text_for_parsing(self, text):
        cleaned = text
        cleaned = re.sub(r'\s+', ' ', cleaned)
        cleaned = cleaned.replace('\u00a0', ' ')
        cleaned = cleaned.replace('\u2009', ' ')
        cleaned = cleaned.replace('\u200b', '')
        cleaned = re.sub(r'(\d)\s+[-/]\s+(\d)', r'\1-\2', cleaned)
        cleaned = re.sub(r'(\d)\s+(\d)', r'\1\2', cleaned)
        return cleaned
    
    def parse_invoice_data(self, text):
        cleaned_text = self.clean_text_for_parsing(text)
        
        data = {
            'invoice_number': '',
            'amount': 0.0,
            'invoice_date': '',
            'currency': 'EUR'
        }
        
        invoice_patterns = [
            r'invoice\s+no\.?\s*:?\s*([^\s\n]+(?:\s*[-_]\s*\d+)?)',
            r'invoice\s+no\.?\s*:?\s*([^\n]+)',
            r'invoice\s+no\.?\s+([^\s\n]+(?:\s*[-_]\s*\d+)?)'
        ]
        
        amount_patterns = [
            r'amount\s+due\s*:?\s*\n\s*(\d{1,3}(?:[,\s]\d{3})*[.,]\d{2})\s*(?:EUR|USD|GBP|€|\$|£)',
            r'amount\s+due\s*:?\s*\n\s*(\d{1,3}(?:[,\s]\d{3})*)\s*(?:EUR|USD|GBP|€|\$|£)',
            r'amount\s+due\s*:?\s*\n\s*(?:€|\$|£)?\s*(\d{1,3}(?:[,\s]\d{3})*[.,]\d{2})',
            r'amount\s+due\s*:?\s*(?:€|\$|£)?\s*(\d{1,3}(?:[,\s]\d{3})*[.,]\d{2})',
            r'(?:€|\$|£)\s*(\d{1,3}(?:[,\s]\d{3})*[.,]\d{2})',
        ]

        currency_patterns = {
            'EUR': r'EUR|eur|Eur|€',
            'USD': r'USD|usd|Usd|\$',
            'GBP': r'GBP|gbp|Gbp|£'
        }
        
        ## UPDATED: More flexible date patterns ##
        date_patterns = [
            # Matches YYYY-MM-DD, DD-MM-YYYY, MM-DD-YYYY with separators -, /, or space
            r'(\d{4}[-\/\s]\d{1,2}[-\/\s]\d{1,2})',
            r'(\d{1,2}[-\/\s]\d{1,2}[-\/\s]\d{4})',
            # Matches "September 27, 2025" or "27 September 2025" or "Sep 27 2025" etc.
            r'((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}[,\s]+\d{4})',
            r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?[,\s]+\d{4})'
        ]
        
        for i, pattern in enumerate(invoice_patterns):
            match = re.search(pattern, cleaned_text, re.IGNORECASE)
            if match:
                full_match = match.group(1).strip().replace(' ', '')
                if full_match:
                    data['invoice_number'] = full_match
                    break
        
        for i, pattern in enumerate(amount_patterns):
            match = re.search(pattern, cleaned_text, re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    amount_str = match.groups()[-1].replace(' ', '')
                    if ',' in amount_str and '.' in amount_str:
                        last_comma = amount_str.rfind(',')
                        last_dot = amount_str.rfind('.')
                        if last_comma > last_dot:
                            amount_str = amount_str.replace('.', '').replace(',', '.')
                        else:
                            amount_str = amount_str.replace(',', '')
                    elif ',' in amount_str:
                        amount_str = amount_str.replace(',', '.')
                    
                    data['amount'] = float(amount_str)
                    break
                except ValueError:
                    continue
        
        for currency_code, pattern in currency_patterns.items():
            if re.search(pattern, cleaned_text, re.IGNORECASE):
                data['currency'] = currency_code
                break

        ## UPDATED: Loop through a list of date formats to try ##
        for pattern in date_patterns:
            match = re.search(pattern, cleaned_text, re.IGNORECASE)
            if match:
                date_str = match.group(1).replace('/', '-').replace(' ', '-').replace(',', '')
                
                # List of formats to attempt to parse the date string with
                possible_formats = [
                    '%Y-%m-%d',       # 2025-09-27
                    '%d-%m-%Y',       # 27-09-2025
                    '%m-%d-%Y',       # 09-27-2025
                    '%B-%d-%Y',       # September-27-2025
                    '%d-%B-%Y',       # 27-September-2025
                    '%b-%d-%Y',       # Sep-27-2025
                    '%d-%b-%Y',       # 27-Sep-2025
                ]
                
                for fmt in possible_formats:
                    try:
                        parsed_date = datetime.strptime(date_str, fmt)
                        data['invoice_date'] = parsed_date.strftime('%Y-%m-%d')
                        break # Exit the inner loop once a format works
                    except ValueError:
                        continue # Try the next format
            
            if data['invoice_date']:
                break # Exit the outer loop once a date is found

        return data
    
    def add_invoice(self, invoice_number, amount, currency, invoice_date, company_name=None, 
                    payment_method=None, payment_terms_days=None, file_path=None):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT id FROM invoices WHERE invoice_number = ?", (invoice_number,))
            if cursor.fetchone():
                conn.close()
                return False, f"Invoice number '{invoice_number}' already exists"
            
            if not company_name:
                pattern_info = self.get_company_info_by_pattern(invoice_number)
                if pattern_info:
                    company_name = pattern_info['company_name']
                    payment_method = pattern_info['payment_method']
                    payment_terms_days = pattern_info['payment_terms_days']
                    currency = pattern_info.get('currency', currency) # Use parsed currency if available
                else:
                    return False, f"No company pattern found for '{invoice_number}'. Please set up pattern first."
            
            invoice_date_obj = datetime.strptime(invoice_date, '%Y-%m-%d')
            due_date = invoice_date_obj + timedelta(days=payment_terms_days)
            
            cursor.execute('''
                INSERT INTO invoices 
                (invoice_number, company_name, amount, currency, invoice_date, due_date, 
                 payment_method, payment_terms_days, status, file_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                invoice_number, company_name, amount, currency, invoice_date,
                due_date.strftime('%Y-%m-%d'), payment_method, payment_terms_days, 'Sent', file_path
            ))
            
            conn.commit()
            conn.close()
            return True, f"Invoice '{invoice_number}' added successfully. Due date: {due_date.strftime('%Y-%m-%d')}"
            
        except Exception as e:
            conn.close()
            return False, f"Error adding invoice: {e}"
    
    def get_payment_summary(self):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        today = datetime.now().date()
        
        cursor.execute('''
            SELECT 
                invoice_number, company_name, amount, currency, invoice_date, due_date,
                payment_method, status,
                CASE 
                    WHEN due_date < ? AND status != 'Paid' THEN 'OVERDUE'
                    WHEN due_date <= ? AND status != 'Paid' THEN 'DUE_SOON'
                    ELSE status
                END as payment_status,
                (julianday(?) - julianday(due_date)) as days_overdue
            FROM invoices 
            ORDER BY due_date ASC
        ''', (today.isoformat(), (today + timedelta(days=7)).isoformat(), today.isoformat()))
        
        invoices = cursor.fetchall()
        conn.close()
        return invoices
    
    def get_all_patterns(self):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute("SELECT pattern, company_name, payment_method, payment_terms_days, currency FROM invoice_patterns ORDER BY pattern")
        patterns = cursor.fetchall()
        conn.close()
        return patterns
    
    def update_invoice_status(self, invoice_number, status):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute("UPDATE invoices SET status = ? WHERE invoice_number = ?", (status, invoice_number))
        conn.commit()
        conn.close()
    
    def delete_invoice(self, invoice_number):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute("DELETE FROM invoices WHERE invoice_number = ?", (invoice_number,))
            rows_affected = cursor.rowcount
            conn.commit()
            conn.close()
            
            if rows_affected > 0:
                return True, f"Invoice {invoice_number} deleted successfully"
            else:
                return False, f"Invoice {invoice_number} not found"
                
        except Exception as e:
            conn.close()
            return False, f"Error deleting invoice: {e}"
    
    def get_invoice_by_number(self, invoice_number):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, invoice_number, company_name, amount, currency, invoice_date, 
                   payment_method, payment_terms_days, status, due_date
            FROM invoices 
            WHERE invoice_number = ?
        ''', (invoice_number,))
        result = cursor.fetchone()
        conn.close()
        return result
    
    def update_invoice(self, invoice_id, invoice_number, company_name, amount, currency, invoice_date, 
                       payment_method, payment_terms_days, status):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            invoice_date_obj = datetime.strptime(invoice_date, '%Y-%m-%d')
            due_date = invoice_date_obj + timedelta(days=payment_terms_days)
            
            cursor.execute('''
                UPDATE invoices 
                SET invoice_number = ?, company_name = ?, amount = ?, currency = ?, invoice_date = ?, 
                    due_date = ?, payment_method = ?, payment_terms_days = ?, status = ?
                WHERE id = ?
            ''', (
                invoice_number, company_name, amount, currency, invoice_date,
                due_date.strftime('%Y-%m-%d'), payment_method, payment_terms_days, status, invoice_id
            ))
            
            conn.commit()
            conn.close()
            return True, "Invoice updated successfully"
            
        except Exception as e:
            conn.close()
            return False, f"Error updating invoice: {e}"
    
    def process_file(self, file_path):
        file_path = Path(file_path)
        
        if not file_path.exists():
            return False, "File not found", None
        
        text = ""
        if file_path.suffix.lower() == '.pdf':
            text = self.extract_text_from_pdf(file_path)
        elif file_path.suffix.lower() in ['.docx', '.doc']:
            text = self.extract_text_from_docx(file_path)
        else:
            return False, "Unsupported file type", None
        
        if not text:
            return False, "No text extracted", None
        
        parsed_data = self.parse_invoice_data(text)
        
        if not parsed_data['invoice_number']:
            return False, "Could not find invoice number", None
        
        if not parsed_data['amount']:
            return False, "Could not find amount", None
        
        if not parsed_data['invoice_date']:
            file_date = datetime.fromtimestamp(file_path.stat().st_mtime)
            parsed_data['invoice_date'] = file_date.strftime('%Y-%m-%d')
        
        success, message = self.add_invoice(
            parsed_data['invoice_number'],
            parsed_data['amount'],
            parsed_data['currency'],
            parsed_data['invoice_date'],
            file_path=str(file_path)
        )
        
        if success:
            self.last_folder = str(file_path.parent)
            self.save_config()
        
        return success, message, parsed_data
    
    def get_statistics(self):
        """Get comprehensive invoice statistics"""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        today = datetime.now().date()
        current_month_start = today.replace(day=1)
        next_month = (current_month_start + timedelta(days=32)).replace(day=1)
        month_after = (next_month + timedelta(days=32)).replace(day=1)
        
        stats = {}
        
        # Total invoices and amounts
        cursor.execute("SELECT COUNT(*), SUM(amount) FROM invoices")
        total_invoices, total_amount = cursor.fetchone()
        stats['total_invoices'] = total_invoices or 0
        stats['total_amount'] = total_amount or 0
        
        # Unpaid invoices
        cursor.execute("SELECT COUNT(*), SUM(amount) FROM invoices WHERE status != 'Paid'")
        unpaid_count, unpaid_amount = cursor.fetchone()
        stats['unpaid_invoices'] = unpaid_count or 0
        stats['unpaid_amount'] = unpaid_amount or 0
        
        # Overdue invoices
        cursor.execute("""
            SELECT COUNT(*), SUM(amount) 
            FROM invoices 
            WHERE status != 'Paid' AND due_date < ?
        """, (today.isoformat(),))
        overdue_count, overdue_amount = cursor.fetchone()
        stats['overdue_invoices'] = overdue_count or 0
        stats['overdue_amount'] = overdue_amount or 0
        
        # Paid invoices
        cursor.execute("SELECT COUNT(*), SUM(amount) FROM invoices WHERE status = 'Paid'")
        paid_count, paid_amount = cursor.fetchone()
        stats['paid_invoices'] = paid_count or 0
        stats['paid_amount'] = paid_amount or 0
        
        # Expected payments this month
        cursor.execute("""
            SELECT COUNT(*), SUM(amount) 
            FROM invoices 
            WHERE status != 'Paid' AND due_date >= ? AND due_date < ?
        """, (current_month_start.isoformat(), next_month.isoformat()))
        this_month_count, this_month_amount = cursor.fetchone()
        stats['this_month_count'] = this_month_count or 0
        stats['this_month_amount'] = this_month_amount or 0
        
        # Expected payments next month
        cursor.execute("""
            SELECT COUNT(*), SUM(amount) 
            FROM invoices 
            WHERE status != 'Paid' AND due_date >= ? AND due_date < ?
        """, (next_month.isoformat(), month_after.isoformat()))
        next_month_count, next_month_amount = cursor.fetchone()
        stats['next_month_count'] = next_month_count or 0
        stats['next_month_amount'] = next_month_amount or 0
        
        # Top clients by amount owed
        cursor.execute("""
            SELECT company_name, COUNT(*) as invoice_count, SUM(amount) as total_owed
            FROM invoices 
            WHERE status != 'Paid'
            GROUP BY company_name 
            ORDER BY total_owed DESC 
            LIMIT 5
        """)
        stats['top_clients'] = cursor.fetchall()
        
        # Monthly revenue (last 6 months of paid invoices)
        cursor.execute("""
            SELECT 
                strftime('%Y-%m', invoice_date) as month,
                SUM(amount) as revenue
            FROM invoices 
            WHERE status = 'Paid' 
                AND invoice_date >= date('now', '-6 months')
            GROUP BY strftime('%Y-%m', invoice_date)
            ORDER BY month
        """)
        stats['monthly_revenue'] = cursor.fetchall()
        
        # Payment status breakdown
        cursor.execute("""
            SELECT status, COUNT(*) as count, SUM(amount) as amount
            FROM invoices 
            GROUP BY status
        """)
        stats['status_breakdown'] = cursor.fetchall()
        
        # Clients invoiced per month (last 12 months)
        cursor.execute("""
            SELECT 
                strftime('%Y-%m', invoice_date) as month,
                COUNT(DISTINCT company_name) as unique_clients,
                COUNT(*) as total_invoices
            FROM invoices 
            WHERE invoice_date >= date('now', '-12 months')
            GROUP BY strftime('%Y-%m', invoice_date)
            ORDER BY month
        """)
        stats['monthly_clients'] = cursor.fetchall()
        
        conn.close()
        return stats


class InvoiceGUI:
    def __init__(self):
        self.manager = InvoiceManager()
        self.setup_gui()
    
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Invoice Management System")
        self.root.geometry("1000x700")
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.process_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.process_frame, text="Process Files")
        self.setup_process_tab()
        
        self.patterns_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.patterns_frame, text="Manage Patterns")
        self.setup_patterns_tab()
        
        self.manual_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.manual_frame, text="Manual Entry")
        self.setup_manual_tab()
        
        self.tracking_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.tracking_frame, text="Payment Tracking")
        self.setup_tracking_tab()
        
        self.stats_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_frame, text="Statistics")
        self.setup_statistics_tab()
        
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_process_tab(self):
        main_frame = ttk.Frame(self.process_frame, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        title_label = ttk.Label(main_frame, text="Process Invoice Files", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Instructions
        instructions = ttk.Label(main_frame, text="Select PDF or Word files to automatically extract invoice data", 
                                font=("Arial", 10))
        instructions.pack(pady=(0, 20))
        
        # File selection buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        # Single file button
        single_button = ttk.Button(button_frame, text="📂 Browse for Invoice File", 
                                   command=self.select_file, width=25)
        single_button.pack(side="left", padx=5)
        
        # Multiple files button
        multi_button = ttk.Button(button_frame, text="📂 Browse for Multiple Files", 
                                  command=self.select_multiple_files, width=25)
        multi_button.pack(side="left", padx=5)
        
        # Keyboard shortcut info
        shortcut_frame = ttk.Frame(main_frame)
        shortcut_frame.pack(pady=5)
        
        shortcut_label = ttk.Label(shortcut_frame, text="💡 Shortcuts: Ctrl+O (single file) | Ctrl+Shift+O (multiple files)", 
                                  font=("Arial", 8), foreground="blue")
        shortcut_label.pack()
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self.select_file())
        self.root.bind('<Control-Shift-O>', lambda e: self.select_multiple_files())
        
        ttk.Label(main_frame, text="Processing Results:").pack(anchor="w", pady=(20, 5))
        
        results_frame = ttk.Frame(main_frame)
        results_frame.pack(fill="both", expand=True)
        
        self.results_text = tk.Text(results_frame, height=15, wrap=tk.WORD)
        self.results_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        # Add welcome message
        welcome_msg = """Welcome to the Invoice Processing System! 🎉

📋 How to use:
1. Click "Browse for Invoice File" (or press Ctrl+O) for single files
2. Click "Browse for Multiple Files" (or press Ctrl+Shift+O) for batch processing
3. The system will automatically extract invoice data
4. Check the results below

✨ What gets extracted automatically:
• Invoice number (from "Invoice No.")
• Amount (from "Amount Due")
• Date (from "Invoice Date")

🔧 Setup required:
• Go to "Manage Patterns" tab first
• Set up patterns for your clients (e.g., "ans" → "Ansent Corp")

Ready to process your first invoice? Click the browse button above! 📂
"""
        self.results_text.insert(tk.END, welcome_msg)
    

    def setup_patterns_tab(self):
        main_frame = ttk.Frame(self.patterns_frame, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        title_label = ttk.Label(main_frame, text="Invoice Pattern Management", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        form_frame = ttk.LabelFrame(main_frame, text="Add/Edit Pattern", padding="10")
        form_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Label(form_frame, text="Pattern:").grid(row=0, column=0, sticky="w", pady=5)
        self.pattern_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.pattern_var, width=20).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Company Name:").grid(row=1, column=0, sticky="w", pady=5)
        self.company_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.company_var, width=30).grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Payment Method:").grid(row=2, column=0, sticky="w", pady=5)
        self.payment_method_var = tk.StringVar(value="Bank Transfer")
        payment_combo = ttk.Combobox(form_frame, textvariable=self.payment_method_var, 
                                     values=["Bank Transfer", "PayPal", "Other", "Cash"], width=27)
        payment_combo.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Payment Terms (days):").grid(row=3, column=0, sticky="w", pady=5)
        self.terms_var = tk.StringVar(value="30")
        ttk.Entry(form_frame, textvariable=self.terms_var, width=10).grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Currency:").grid(row=4, column=0, sticky="w", pady=5)
        self.pattern_currency_var = tk.StringVar(value="EUR")
        currency_combo = ttk.Combobox(form_frame, textvariable=self.pattern_currency_var, 
                                      values=self.manager.get_currencies(), width=10)
        currency_combo.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        ttk.Button(form_frame, text="Save Pattern", command=self.save_pattern).grid(row=5, column=1, sticky="w", padx=5, pady=10)
        
        list_frame = ttk.LabelFrame(main_frame, text="Existing Patterns", padding="10")
        list_frame.pack(fill="both", expand=True)
        
        columns = ("Pattern", "Company", "Payment Method", "Terms (days)", "Currency")
        self.patterns_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        
        for col in columns:
            self.patterns_tree.heading(col, text=col)
            self.patterns_tree.column(col, width=150)
        
        self.patterns_tree.pack(side="left", fill="both", expand=True)
        
        patterns_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.patterns_tree.yview)
        patterns_scrollbar.pack(side="right", fill="y")
        self.patterns_tree.configure(yscrollcommand=patterns_scrollbar.set)
        
        ttk.Button(list_frame, text="Refresh", command=self.refresh_patterns).pack(pady=10)
        self.refresh_patterns()
    
    def setup_manual_tab(self):
        main_frame = ttk.Frame(self.manual_frame, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        title_label = ttk.Label(main_frame, text="Manual Invoice Entry", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        form_frame = ttk.LabelFrame(main_frame, text="Enter Invoice Manually", padding="10")
        form_frame.pack(fill="x")
        
        ttk.Label(form_frame, text="Invoice Number:").grid(row=0, column=0, sticky="w", pady=5)
        self.manual_invoice_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.manual_invoice_var, width=20).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Amount:").grid(row=1, column=0, sticky="w", pady=5)
        self.manual_amount_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.manual_amount_var, width=20).grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Currency:").grid(row=1, column=2, sticky="w", pady=5, padx=(10,0))
        self.manual_currency_var = tk.StringVar(value="EUR")
        manual_currency_combo = ttk.Combobox(form_frame, textvariable=self.manual_currency_var, 
                                             values=self.manager.get_currencies(), width=8)
        manual_currency_combo.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Invoice Date (YYYY-MM-DD):").grid(row=2, column=0, sticky="w", pady=5)
        self.manual_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(form_frame, textvariable=self.manual_date_var, width=20).grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Company Name:").grid(row=3, column=0, sticky="w", pady=5)
        self.manual_company_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.manual_company_var, width=30).grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Payment Method:").grid(row=4, column=0, sticky="w", pady=5)
        self.manual_payment_var = tk.StringVar()
        ttk.Combobox(form_frame, textvariable=self.manual_payment_var, 
                     values=["Bank Transfer", "PayPal", "Other", "Cash"], width=27).grid(row=4, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(form_frame, text="Payment Terms Days:").grid(row=5, column=0, sticky="w", pady=5)
        self.manual_terms_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.manual_terms_var, width=10).grid(row=5, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Button(form_frame, text="Add Invoice", command=self.add_manual_invoice).grid(row=6, column=1, sticky="w", padx=5, pady=20)
    
    def setup_tracking_tab(self):
        main_frame = ttk.Frame(self.tracking_frame, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        title_label = ttk.Label(main_frame, text="Payment Tracking", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Button(button_frame, text="Refresh", command=self.refresh_tracking).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Mark as Paid", command=self.mark_as_paid).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Edit Invoice", command=self.edit_selected_invoice).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Delete Invoice", command=self.delete_selected_invoice).pack(side="left", padx=5)
        
        columns = ("Invoice #", "Company", "Amount", "Currency", "Invoice Date", "Due Date", "Payment Method", "Status", "Days")
        self.tracking_tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=20)
        
        for col in columns:
            self.tracking_tree.heading(col, text=col)
            if col == "Company":
                self.tracking_tree.column(col, width=200)
            elif col in ["Amount", "Days", "Currency"]:
                self.tracking_tree.column(col, width=80)
            else:
                self.tracking_tree.column(col, width=120)
        
        self.tracking_tree.pack(side="left", fill="both", expand=True)
        
        tracking_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.tracking_tree.yview)
        tracking_scrollbar.pack(side="right", fill="y")
        self.tracking_tree.configure(yscrollcommand=tracking_scrollbar.set)
        
        self.refresh_tracking()
    
    def select_file(self):
        """Select single file using native Ubuntu file dialog"""
        try:
            # Try to use native Ubuntu file dialog through zenity
            result = subprocess.run([
                'zenity', '--file-selection',
                '--title=Select Invoice File',
                '--file-filter=Invoice files (pdf,docx,doc) | *.pdf *.docx *.doc',
                '--file-filter=PDF files | *.pdf',
                '--file-filter=Word documents | *.docx *.doc',
                '--file-filter=All files | *'
            ], capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and result.stdout.strip():
                file_path = result.stdout.strip()
                self.process_file(file_path)
                return
        except (subprocess.TimeoutExpired, FileNotFoundError, subprocess.SubprocessError):
            # Zenity not available, fall back to tkinter dialog
            pass
        
        # Fallback to tkinter dialog
        file_path = filedialog.askopenfilename(
            parent=self.root,
            initialdir=self.manager.last_folder,
            title="Select Invoice File",
            filetypes=[
                ("Invoice files", "*.pdf *.docx *.doc"),
                ("PDF files", "*.pdf"), 
                ("Word documents", "*.docx *.doc"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.process_file(file_path)
    
    def select_multiple_files(self):
        """Select multiple files for batch processing"""
        try:
            # Try native Ubuntu dialog for multiple files
            result = subprocess.run([
                'zenity', '--file-selection', '--multiple', '--separator=|',
                '--title=Select Multiple Invoice Files',
                '--file-filter=Invoice files (pdf,docx,doc) | *.pdf *.docx *.doc',
                '--file-filter=All files | *'
            ], capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and result.stdout.strip():
                file_paths = result.stdout.strip().split('|')
                self._process_multiple_files(file_paths)
                return
        except (subprocess.TimeoutExpired, FileNotFoundError, subprocess.SubprocessError):
            # Zenity not available, fall back to tkinter dialog
            pass
        
        # Fallback to tkinter dialog
        file_paths = filedialog.askopenfilenames(
            parent=self.root,
            initialdir=self.manager.last_folder,
            title="Select Multiple Invoice Files",
            filetypes=[
                ("Invoice files", "*.pdf *.docx *.doc"),
                ("PDF files", "*.pdf"), 
                ("Word documents", "*.docx *.doc"),
                ("All files", "*.*")
            ]
        )
        
        if file_paths:
            self._process_multiple_files(file_paths)
    
    def _process_multiple_files(self, file_paths):
        """Process multiple files and show progress"""
        self.results_text.insert(tk.END, f"\n🔄 Processing {len(file_paths)} files...\n")
        self.results_text.see(tk.END)
        self.root.update()
        
        success_count = 0
        for file_path in file_paths:
            success, message, data = self.manager.process_file(file_path)
            if success:
                success_count += 1
            
            # Show brief result for each file
            status = "✅" if success else "❌"
            filename = Path(file_path).name
            self.results_text.insert(tk.END, f"{status} {filename}\n")
            self.results_text.see(tk.END)
            self.root.update()
        
        # Summary
        self.results_text.insert(tk.END, f"\n📊 Batch complete: {success_count}/{len(file_paths)} files processed successfully\n")
        self.results_text.insert(tk.END, "-" * 60 + "\n")
        self.results_text.see(tk.END)
        
        self.refresh_tracking()
        self.status_var.set(f"✅ Batch processed: {success_count}/{len(file_paths)} files")
    
    def process_file(self, file_path):
        self.status_var.set("Processing...")
        self.root.update()
        
        success, message, data = self.manager.process_file(file_path)
        
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if success:
            result = f"\n[{timestamp}] ✅ SUCCESS: {Path(file_path).name}\n"
            result += f"   📄 Invoice #: {data['invoice_number']}\n"
            result += f"   💰 Amount: {data['amount']:.2f} {data['currency']}\n"
            result += f"   📅 Date: {data['invoice_date']}\n"
            result += f"   ✅ {message}\n"
            result += "-" * 60 + "\n"
            self.status_var.set("✅ Processed successfully!")
        else:
            result = f"\n[{timestamp}] ❌ FAILED: {Path(file_path).name}\n"
            result += f"   ⚠️  Error: {message}\n"
            result += "-" * 60 + "\n"
            self.status_var.set(f"❌ {message}")
        
        self.results_text.insert(tk.END, result)
        self.results_text.see(tk.END)
        
        self.refresh_tracking()
    
    def save_pattern(self):
        pattern = self.pattern_var.get().strip().lower()
        company = self.company_var.get().strip()
        payment_method = self.payment_method_var.get()
        currency = self.pattern_currency_var.get()

        try:
            terms = int(self.terms_var.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for payment terms")
            return
        
        if not pattern or not company:
            messagebox.showerror("Error", "Please fill in pattern and company name")
            return
        
        success, message = self.manager.add_or_update_pattern(pattern, company, payment_method, terms, currency)
        
        if success:
            self.status_var.set("✅ Pattern saved")
            messagebox.showinfo("Success", message)
            self.refresh_patterns()
            self.pattern_var.set("")
            self.company_var.set("")
            self.terms_var.set("30")
        else:
            self.status_var.set("❌ Error saving pattern")
            messagebox.showerror("Error", message)
    
    def refresh_patterns(self):
        for item in self.patterns_tree.get_children():
            self.patterns_tree.delete(item)
        
        patterns = self.manager.get_all_patterns()
        for pattern in patterns:
            self.patterns_tree.insert("", "end", values=pattern)
    
    def add_manual_invoice(self):
        invoice_number = self.manual_invoice_var.get().strip()
        amount_str = self.manual_amount_var.get().strip()
        currency = self.manual_currency_var.get()
        date_str = self.manual_date_var.get().strip()
        
        if not invoice_number or not amount_str or not date_str:
            messagebox.showerror("Error", "Please fill in invoice number, amount, and date")
            return
        
        try:
            amount = float(amount_str)
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid amount")
            return
        
        company = self.manual_company_var.get().strip()
        payment_method = self.manual_payment_var.get().strip()
        terms_str = self.manual_terms_var.get().strip()
        
        if company:
            if not payment_method or not terms_str:
                messagebox.showerror("Error", "For external invoices, please provide payment method and terms")
                return
            try:
                terms = int(terms_str)
                success, message = self.manager.add_invoice(
                    invoice_number, amount, currency, date_str, company, payment_method, terms
                )
            except ValueError:
                messagebox.showerror("Error", "Please enter valid payment terms")
                return
        else:
            success, message = self.manager.add_invoice(invoice_number, amount, currency, date_str)
        
        if success:
            self.status_var.set("✅ Invoice added")
            messagebox.showinfo("Success", message)
            self.manual_invoice_var.set("")
            self.manual_amount_var.set("")
            self.manual_date_var.set(datetime.now().strftime('%Y-%m-%d'))
            self.manual_company_var.set("")
            self.manual_payment_var.set("")
            self.manual_terms_var.set("")
            self.refresh_tracking()
        else:
            self.status_var.set("❌ Error adding invoice")
            messagebox.showerror("Error", message)
    
    def refresh_tracking(self):
        for item in self.tracking_tree.get_children():
            self.tracking_tree.delete(item)
        
        invoices = self.manager.get_payment_summary()
        for invoice in invoices:
            invoice_num, company, amount, currency, inv_date, due_date, payment_method, status, payment_status, days_overdue = invoice
            
            formatted_amount = f"{amount:.2f}"
            
            if payment_status == 'OVERDUE':
                days_text = f"OVERDUE {int(abs(days_overdue))} days"
                tag = "overdue"
            elif payment_status == 'DUE_SOON':
                days_text = f"Due in {int(abs(days_overdue))} days"
                tag = "due_soon"
            elif status == 'Paid':
                days_text = "PAID"
                tag = "paid"
            else:
                days_text = f"Due in {int(abs(days_overdue))} days"
                tag = "normal"
            
            self.tracking_tree.insert("", "end", values=(
                invoice_num, company, formatted_amount, currency, inv_date, due_date, 
                payment_method, status, days_text
            ), tags=(tag,))
        
        self.tracking_tree.tag_configure("overdue", background="#ffcccc")
        self.tracking_tree.tag_configure("due_soon", background="#fff3cd")
        self.tracking_tree.tag_configure("paid", background="#d4edda")
        self.tracking_tree.tag_configure("normal", background="white")
    
    def mark_as_paid(self):
        selected = self.tracking_tree.selection()
        if not selected:
            messagebox.showwarning("Selection", "Please select an invoice to mark as paid")
            return
        
        invoice_number = self.tracking_tree.item(selected[0])['values'][0]
        self.manager.update_invoice_status(invoice_number, 'Paid')
        self.refresh_tracking()
        self.status_var.set(f"✅ Marked {invoice_number} as paid")
    
    def delete_selected_invoice(self):
        selected = self.tracking_tree.selection()
        if not selected:
            messagebox.showwarning("Selection", "Please select an invoice to delete")
            return
        
        invoice_number = self.tracking_tree.item(selected[0])['values'][0]
        
        if messagebox.askyesno("Confirm Deletion", 
                               f"Are you sure you want to delete invoice {invoice_number}?\n\nThis action cannot be undone."):
            success, message = self.manager.delete_invoice(invoice_number)
            
            if success:
                self.refresh_tracking()
                self.status_var.set(f"✅ Deleted {invoice_number}")
                messagebox.showinfo("Success", message)
            else:
                self.status_var.set("❌ Error deleting invoice")
                messagebox.showerror("Error", message)
    
    def edit_selected_invoice(self):
        """Edit selected invoice"""
        selected = self.tracking_tree.selection()
        if not selected:
            messagebox.showwarning("Selection", "Please select an invoice to edit")
            return
        
        invoice_number = self.tracking_tree.item(selected[0])['values'][0]
        invoice_details = self.manager.get_invoice_by_number(invoice_number)
        
        if not invoice_details:
            messagebox.showerror("Error", "Invoice not found")
            return
        
        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Edit Invoice {invoice_number}")
        edit_window.geometry("500x450")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        (invoice_id, current_number, current_company, current_amount, current_currency, current_date,
         current_payment_method, current_terms, current_status, current_due_date) = invoice_details
        
        main_frame = ttk.LabelFrame(edit_window, text="Edit Invoice Details", padding="15")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ttk.Label(main_frame, text="Invoice Number:").grid(row=0, column=0, sticky="w", pady=5)
        edit_invoice_var = tk.StringVar(value=current_number)
        ttk.Entry(main_frame, textvariable=edit_invoice_var, width=30).grid(row=0, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(main_frame, text="Company Name:").grid(row=1, column=0, sticky="w", pady=5)
        edit_company_var = tk.StringVar(value=current_company)
        ttk.Entry(main_frame, textvariable=edit_company_var, width=30).grid(row=1, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(main_frame, text="Amount:").grid(row=2, column=0, sticky="w", pady=5)
        edit_amount_var = tk.StringVar(value=str(current_amount))
        ttk.Entry(main_frame, textvariable=edit_amount_var, width=20).grid(row=2, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="Currency:").grid(row=2, column=2, sticky="w", padx=(10,5), pady=5)
        edit_currency_var = tk.StringVar(value=current_currency)
        edit_currency_combo = ttk.Combobox(main_frame, textvariable=edit_currency_var, 
                                             values=self.manager.get_currencies(), width=8)
        edit_currency_combo.grid(row=2, column=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(main_frame, text="Invoice Date (YYYY-MM-DD):").grid(row=3, column=0, sticky="w", pady=5)
        edit_date_var = tk.StringVar(value=current_date)
        ttk.Entry(main_frame, textvariable=edit_date_var, width=30).grid(row=3, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(main_frame, text="Payment Method:").grid(row=4, column=0, sticky="w", pady=5)
        edit_payment_var = tk.StringVar(value=current_payment_method)
        payment_combo = ttk.Combobox(main_frame, textvariable=edit_payment_var, 
                                     values=["Bank Transfer", "PayPal", "Other", "Cash"], width=27)
        payment_combo.grid(row=4, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(main_frame, text="Payment Terms (days):").grid(row=5, column=0, sticky="w", pady=5)
        edit_terms_var = tk.StringVar(value=str(current_terms))
        ttk.Entry(main_frame, textvariable=edit_terms_var, width=30).grid(row=5, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(main_frame, text="Status:").grid(row=6, column=0, sticky="w", pady=5)
        edit_status_var = tk.StringVar(value=current_status)
        status_combo = ttk.Combobox(main_frame, textvariable=edit_status_var, 
                                    values=["Sent", "Paid", "Overdue", "Cancelled"], width=27)
        status_combo.grid(row=6, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=4, pady=20)
        
        def save_changes():
            try:
                new_amount = float(edit_amount_var.get())
                new_terms = int(edit_terms_var.get())
                new_date = edit_date_var.get()
                new_currency = edit_currency_var.get()
                
                datetime.strptime(new_date, '%Y-%m-%d')
                
                success, message = self.manager.update_invoice(
                    invoice_id,
                    edit_invoice_var.get().strip(),
                    edit_company_var.get().strip(),
                    new_amount,
                    new_currency,
                    new_date,
                    edit_payment_var.get(),
                    new_terms,
                    edit_status_var.get()
                )
                
                if success:
                    edit_window.destroy()
                    self.refresh_tracking()
                    self.status_var.set("✅ Invoice updated successfully")
                    messagebox.showinfo("Success", message)
                else:
                    messagebox.showerror("Error", message)
                    
            except ValueError:
                messagebox.showerror("Error", "Please check your input values (amount should be a number, date should be YYYY-MM-DD)")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving changes: {e}")
        
        def cancel_edit():
            edit_window.destroy()
        
        ttk.Button(button_frame, text="Save Changes", command=save_changes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=cancel_edit).pack(side="left", padx=5)
    
    def setup_statistics_tab(self):
        """Setup statistics and analytics tab"""
        main_frame = ttk.Frame(self.stats_frame, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        title_label = ttk.Label(main_frame, text="Invoice Statistics & Analytics", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Refresh button
        refresh_button = ttk.Button(main_frame, text="🔄 Refresh Statistics", command=self.refresh_statistics)
        refresh_button.pack(pady=(0, 10))
        
        # Create main container that uses full width
        container_frame = ttk.Frame(main_frame)
        container_frame.pack(fill="both", expand=True)
        
        # Create scrollable frame that fills the container
        canvas = tk.Canvas(container_frame)
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar to fill container
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        self.stats_container = scrollable_frame
        self.refresh_statistics()
    
    def refresh_statistics(self):
        """Refresh all statistics displays"""
        # Clear existing widgets
        for widget in self.stats_container.winfo_children():
            widget.destroy()
        
        stats = self.manager.get_statistics()
        
        # Main content frame that expands to full width
        content_frame = ttk.Frame(self.stats_container)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Overview cards - use grid with proper column configuration
        overview_frame = ttk.LabelFrame(content_frame, text="📊 Overview", padding="15")
        overview_frame.pack(fill="x", pady=(0, 15))
        
        # Create 2x2 grid of overview cards with better sizing
        cards = [
            ("Total Invoices", stats['total_invoices'], f"€{stats['total_amount']:.2f}", "blue"),
            ("Unpaid Invoices", stats['unpaid_invoices'], f"€{stats['unpaid_amount']:.2f}", "orange"),
            ("Overdue Invoices", stats['overdue_invoices'], f"€{stats['overdue_amount']:.2f}", "red"),
            ("Paid Invoices", stats['paid_invoices'], f"€{stats['paid_amount']:.2f}", "green")
        ]
        
        for i, (title, count, amount, color) in enumerate(cards):
            row = i // 2
            col = i % 2
            
            card_frame = ttk.Frame(overview_frame, relief="solid", borderwidth=1)
            card_frame.grid(row=row, column=col, padx=10, pady=10, sticky="ew", ipadx=20, ipady=10)
            
            ttk.Label(card_frame, text=title, font=("Arial", 11, "bold")).pack(pady=5)
            ttk.Label(card_frame, text=str(count), font=("Arial", 16, "bold"), foreground=color).pack()
            ttk.Label(card_frame, text=amount, font=("Arial", 11)).pack(pady=5)
        
        # Configure grid columns to expand equally
        overview_frame.columnconfigure(0, weight=1)
        overview_frame.columnconfigure(1, weight=1)
        
        # Create two-column layout for better width usage
        two_column_frame = ttk.Frame(content_frame)
        two_column_frame.pack(fill="x", pady=(0, 15))
        
        # Left column
        left_column = ttk.Frame(two_column_frame)
        left_column.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Right column  
        right_column = ttk.Frame(two_column_frame)
        right_column.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # Cash flow predictions - left column
        cashflow_frame = ttk.LabelFrame(left_column, text="💰 Expected Cash Flow", padding="15")
        cashflow_frame.pack(fill="x", pady=(0, 15))
        
        this_month_frame = ttk.Frame(cashflow_frame)
        this_month_frame.pack(fill="x", pady=5)
        ttk.Label(this_month_frame, text="This Month:", font=("Arial", 11, "bold")).pack(side="left")
        ttk.Label(this_month_frame, text=f"{stats['this_month_count']} invoices", font=("Arial", 10)).pack(side="left", padx=(15, 0))
        ttk.Label(this_month_frame, text=f"€{stats['this_month_amount']:.2f}", font=("Arial", 11, "bold"), foreground="green").pack(side="right")
        
        next_month_frame = ttk.Frame(cashflow_frame)
        next_month_frame.pack(fill="x", pady=5)
        ttk.Label(next_month_frame, text="Next Month:", font=("Arial", 11, "bold")).pack(side="left")
        ttk.Label(next_month_frame, text=f"{stats['next_month_count']} invoices", font=("Arial", 10)).pack(side="left", padx=(15, 0))
        ttk.Label(next_month_frame, text=f"€{stats['next_month_amount']:.2f}", font=("Arial", 11, "bold"), foreground="blue").pack(side="right")
        
        # Status breakdown - right column
        if stats['status_breakdown']:
            status_frame = ttk.LabelFrame(right_column, text="📋 Invoice Status Breakdown", padding="15")
            status_frame.pack(fill="x", pady=(0, 15))
            
            total_invoices = sum(count for _, count, _ in stats['status_breakdown'])
            
            for status, count, amount in stats['status_breakdown']:
                status_item_frame = ttk.Frame(status_frame)
                status_item_frame.pack(fill="x", pady=3)
                
                percentage = (count / total_invoices * 100) if total_invoices > 0 else 0
                
                # Color coding for status
                color = {
                    'Paid': 'green', 
                    'Overdue': 'red',
                    'Cancelled': 'gray'
                }.get(status, 'orange')
                
                ttk.Label(status_item_frame, text=f"{status}:", font=("Arial", 10, "bold")).pack(side="left")
                ttk.Label(status_item_frame, text=f"{count} invoices ({percentage:.1f}%)", font=("Arial", 10)).pack(side="left", padx=(15, 0))
                ttk.Label(status_item_frame, text=f"€{amount:.2f}", font=("Arial", 10, "bold"), foreground=color).pack(side="right")
        
        # Top clients - full width
        if stats['top_clients']:
            clients_frame = ttk.LabelFrame(content_frame, text="🏢 Top Clients (by amount owed)", padding="15")
            clients_frame.pack(fill="x", pady=(0, 15))
            
            for i, (company, invoice_count, total_owed) in enumerate(stats['top_clients'], 1):
                client_frame = ttk.Frame(clients_frame)
                client_frame.pack(fill="x", pady=3)
                
                ttk.Label(client_frame, text=f"{i}. {company}", font=("Arial", 11, "bold")).pack(side="left")
                ttk.Label(client_frame, text=f"({invoice_count} invoices)", font=("Arial", 10), foreground="gray").pack(side="left", padx=(10, 0))
                ttk.Label(client_frame, text=f"€{total_owed:.2f}", font=("Arial", 11, "bold"), foreground="red").pack(side="right")
        
        # Clients invoiced per month - full width
        if stats['monthly_clients']:
            monthly_clients_frame = ttk.LabelFrame(content_frame, text="👥 Clients Invoiced Per Month (Last 12 Months)", padding="15")
            monthly_clients_frame.pack(fill="x", pady=(0, 15))
            
            max_clients = max(unique_clients for _, unique_clients, _ in stats['monthly_clients']) if stats['monthly_clients'] else 1
            
            for month, unique_clients, total_invoices in stats['monthly_clients']:
                month_frame = ttk.Frame(monthly_clients_frame)
                month_frame.pack(fill="x", pady=3)
                
                # Create a bar chart for better visualization
                bar_length = int((unique_clients / max_clients) * 40) if max_clients > 0 else 0
                bar = "█" * bar_length + "░" * (40 - bar_length)
                
                ttk.Label(month_frame, text=month, font=("Arial", 10, "bold"), width=8).pack(side="left")
                ttk.Label(month_frame, text=bar, font=("Courier", 8), foreground="blue").pack(side="left", padx=(15, 0))
                ttk.Label(month_frame, text=f"{unique_clients} clients", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 0))
                ttk.Label(month_frame, text=f"({total_invoices} invoices)", font=("Arial", 9), foreground="gray").pack(side="right")
        
        # Monthly revenue chart - full width
        if stats['monthly_revenue']:
            revenue_frame = ttk.LabelFrame(content_frame, text="📈 Monthly Revenue (Last 6 Months)", padding="15")
            revenue_frame.pack(fill="x", pady=(0, 15))
            
            max_revenue = max(revenue for _, revenue in stats['monthly_revenue']) if stats['monthly_revenue'] else 1
            
            for month, revenue in stats['monthly_revenue']:
                month_frame = ttk.Frame(revenue_frame)
                month_frame.pack(fill="x", pady=3)
                
                # Create a longer bar chart for better visualization
                bar_length = int((revenue / max_revenue) * 50) if max_revenue > 0 else 0
                bar = "█" * bar_length + "░" * (50 - bar_length)
                
                ttk.Label(month_frame, text=month, font=("Arial", 10, "bold"), width=8).pack(side="left")
                ttk.Label(month_frame, text=bar, font=("Courier", 8), foreground="green").pack(side="left", padx=(15, 0))
                ttk.Label(month_frame, text=f"€{revenue:.2f}", font=("Arial", 10, "bold")).pack(side="right")
    
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    print("Starting Invoice Management System...")
    app = InvoiceGUI()
    app.run()
