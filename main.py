import sys
import os
import functools
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QFileDialog, QComboBox, QTableWidget, QTableWidgetItem, QCheckBox
from PyQt5.QtCore import Qt
import sqlite3
import pandas as pd
from datetime import datetime

class ReviewApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_file_path = None
        self.setWindowTitle('Product Review Filter')
        self.setGeometry(100, 100, 600, 400)
        
        layout = QVBoxLayout()
        
        self.label = QLabel('Select a month to filter products needing review')
        layout.addWidget(self.label)
        
        self.import_button = QPushButton('Import Excel File')
        self.import_button.clicked.connect(self.import_file)
        layout.addWidget(self.import_button)
        
        self.month_combo = QComboBox()
        self.month_combo.addItems(['Select All', 'January', 'February', 'March', 'April', 'May', 'June', 
                                   'July', 'August', 'September', 'October', 'November', 'December'])
        self.month_combo.currentIndexChanged.connect(self.filter_data)
        layout.addWidget(self.month_combo)
        
        self.table = QTableWidget()
        layout.addWidget(self.table)
        
        self.setLayout(layout)

        self.set_current_month()

    def set_current_month(self):
        current_month_number = datetime.now().month
        self.month_combo.setCurrentIndex(current_month_number)

    def import_excel_to_db(self, file_path):
        df = pd.read_excel(file_path, parse_dates=['Review Date'])
        
        # Convert 'Review Date' to datetime format with dayfirst
        df['Review Date'] = pd.to_datetime(df['Review Date'], dayfirst=True)
        
        if 'Verified' not in df.columns:
            df['Verified'] = False
        
        self.excel_file_path = file_path
        
        conn = sqlite3.connect('product_reviews.db')
        cursor = conn.cursor()

        cursor.execute('DELETE FROM Products')
        
        for index, row in df.iterrows():
            review_date = row['Review Date'].strftime('%Y-%m-%d')  # Convert to 'YYYY-MM-DD' for SQLite storage
            
            verified = row['Verified'] if 'Verified' in row else 0

            try:
                cursor.execute('''
                INSERT INTO Products (product_name, reference, review_date, verified)
                VALUES (?, ?, ?, ?)
                ''', (row['Product Name'], row['Reference'], review_date, verified))
                
            except KeyError as e:
                print(f"Column not found: {e}")
            except Exception as e:
                print(f"An error occurred: {e}")

        conn.commit()
        conn.close()

        print(f"Data from {file_path} has been successfully imported.")


    def import_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)", options=options)
        if file_name:
            self.excel_file_path = file_name
            self.import_excel_to_db(file_name)
            self.filter_data()
    
    def filter_data(self):
        selected_month = self.month_combo.currentText()
        self.display_filtered_products(selected_month)

    def display_filtered_products(self, month):
        conn = sqlite3.connect('product_reviews.db')
        cursor = conn.cursor()

        if month == 'Select All':
            cursor.execute('''
            SELECT product_name, reference, review_date, verified
            FROM Products
            ''')
        else:
            month_number = pd.to_datetime(month, format='%B').month
            cursor.execute('''
            SELECT product_name, reference, review_date, verified
            FROM Products
            WHERE strftime('%m', review_date) = ?
            ''', (f'{month_number:02d}',))

        records = cursor.fetchall()
        conn.close()

        self.table.setRowCount(len(records))
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['Product Name', 'Reference', 'Review Date', 'Verified'])

        for row_index, row_data in enumerate(records):
            for col_index, col_data in enumerate(row_data[:-1]):
                self.table.setItem(row_index, col_index, QTableWidgetItem(str(col_data)))

            checkbox = QCheckBox()
            checkbox.setChecked(bool(row_data[3]))
            checkbox.stateChanged.connect(functools.partial(self.update_verification, row_data[1]))
            self.table.setCellWidget(row_index, 3, checkbox)

    def update_verification(self, reference, state):
        verified = state == Qt.Checked

        with sqlite3.connect('product_reviews.db') as conn:
            cursor = conn.cursor()
            cursor.execute('''
            UPDATE Products
            SET verified = ?
            WHERE reference = ?
            ''', (verified, reference))

        self.update_excel_file(reference, verified)
        self.filter_data()

    def update_excel_file(self, reference, verified):
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            print("Excel file path not set or doesn't exist.")
            return

        try:
            df = pd.read_excel(self.excel_file_path)  # Remove dayfirst argument if not using it

            # Convert 'Review Date' to datetime format with dayfirst handling manually
            df['Review Date'] = pd.to_datetime(df['Review Date'], format='%d/%m/%Y', errors='coerce')
            
            # Convert 'Verified' column to integer if it's not already
            if df['Verified'].dtype != 'int64':
                df['Verified'] = df['Verified'].astype(int)

            # Explicitly cast boolean to integer before assignment
            df.loc[df['Reference'] == reference, 'Verified'] = int(verified)

            # Reformat 'Review Date' for saving to Excel
            df['Review Date'] = df['Review Date'].dt.strftime('%d/%m/%Y')

            df.to_excel(self.excel_file_path, index=False)

            print(f"Excel file {self.excel_file_path} has been updated for reference {reference}.")
        
        except PermissionError as e:
            print(f"Permission error: {e}. Ensure the file is not open and has write permissions.")
        except Exception as e:
            print(f"An error occurred while updating the Excel file: {e}")



# Run the application
app = QApplication(sys.argv)
window = ReviewApp()
window.show()
sys.exit(app.exec_())
