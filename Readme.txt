# Invoice Management System

A desktop GUI application built with Python and Tkinter to automate invoice processing, tracking, and analysis. This tool helps you extract key information from PDF and Word invoices, supports multiple currencies, and stores all data in a local database.

---

## 📜 Overview

This application is designed for freelancers, small businesses, or anyone who needs to manage multiple invoices efficiently. Instead of manually entering data into a spreadsheet, you can simply point the application to your invoice files, and it will automatically parse the necessary details.

It features multi-currency support and stores all data in an `invoices.db` SQLite file, ensuring your information remains private.

---

## ✨ Features

-   **📄 Automatic Data Extraction**: Parses invoice number, amount, date, and currency from PDF and DOCX files.
-   **💵 Multi-Currency Support**: Select, store, and track invoices in various currencies (EUR, USD, GBP, etc.).
-   **🗂️ Batch Processing**: Process multiple invoice files at once with a single click.
-   **⚙️ Pattern Management**: Create custom "patterns" to automatically associate invoice codes with specific company details, including a default currency.
-   **📊 Payment Tracking Dashboard**: View all invoices in a table. Invoices are color-coded by status (Overdue, Due Soon, Paid) for easy tracking.
-   **✏️ Full CRUD Functionality**: Manually add, edit, or delete any invoice record. Mark invoices as "Paid" to update their status.
-   **📈 Statistics & Analytics**: A dedicated dashboard provides key financial insights, including total revenue, unpaid amounts, top clients, and monthly revenue charts.
-   **💾 Local Database**: All data is stored in a local `invoices.db` SQLite file. No external services or internet connection is required.

---

## 🚀 Getting Started

Follow these steps to set up and run the application on your local machine.

### Prerequisites

-   **Python 3.8 or newer** is required.
-   **(Optional for Linux)** `zenity` for a native file selection dialog. If not installed, the application will use the default dialog.

### Installation & Setup

1.  **Clone or Download**: Download the project files into a dedicated folder.

2.  **Create `requirements.txt`**: In the project folder, create a new file named `requirements.txt` and paste the following lines into it:
    ```
    PyPDF2
    python-docx
    ```

3.  **Install Dependencies**: Open your terminal or command prompt in the project folder and run the following commands to create a virtual environment and install the required libraries.

    ```bash
    # Create and activate a virtual environment
    # On Windows:
    python -m venv venv
    venv\Scripts\activate
    
    # On macOS/Linux:
    python3 -m venv venv
    source venv/bin/activate
    
    # Install the required packages
    pip install -r requirements.txt
    ```

---

## 🏃 How to Run the Application

With your terminal in the project directory and the virtual environment activated, run the main script:

```bash
# Replace 'invoices.py' with the actual name of your Python file
python invoices.py
```

---

### (Optional) Making the Script Executable

For easier access on Linux and macOS, you can make the script directly executable.

1.  **Add Execute Permissions**: In your terminal, run the following command once:
    ```bash
    chmod +x invoices.py
    ```

2.  **Run Directly**: Now, after activating your virtual environment, you can run the application with a shorter command:
    ```bash
    ./invoices.py
    ```


The first time you run the application, it will automatically create two files in the same directory:
-   `invoices.db`: The database where all your data is stored. **Do not delete this file!**
-   `invoice_config.json`: A configuration file to remember the last folder you opened.

---

## 📝 How It Works (Usage Guide)

### 1. First-Time Setup: Manage Patterns

For the best experience, your first step should be to set up patterns for your clients. A pattern links a short text code (extracted from the invoice number) to a full company profile.

-   Go to the **Manage Patterns** tab.
-   Fill in the details:
    -   **Pattern**: A short, lowercase prefix from your invoice numbers (e.g., if invoice numbers are `ansent-001`, the pattern is `ansent`).
    -   **Company Name**: The full name of the client.
    -   **Payment Method & Terms**: The default payment details.
    -   **Currency**: The default currency for this client.
-   Click **Save Pattern**.

### 2. Processing Invoices

-   Go to the **Process Files** tab.
-   Click **Browse for Invoice File** or **Browse for Multiple Files**.
-   When a file is processed, a new invoice is created with the status **"Sent"**. The results of the operation will be displayed in the text box.

### 3. Manual Entry

-   Go to the **Manual Entry** tab.
-   Fill out the form to add an invoice. You can select the **Currency** from the dropdown menu.
-   New manual invoices are automatically given the status **"Sent"**.

### 4. Tracking Payments

-   Go to the **Payment Tracking** tab to see all your invoices.
-   The table now includes a **Currency** column.
-   The table is color-coded:
    -   🔴 **Red**: Overdue
    -   🟡 **Yellow**: Due within the next 7 days
    -   🟢 **Green**: Paid
-   Select an invoice and use the buttons at the top to **Mark as Paid**, **Edit**, or **Delete** it.

### 5. Viewing Statistics

-   Go to the **Statistics** tab.
-   This dashboard gives you a high-level overview of your business finances.
-   Click the **Refresh Statistics** button to ensure you are viewing the most up-to-date information.
-   **Note**: The overview cards currently aggregate all amounts and display them with a Euro (€) symbol, regardless of their original currency.
