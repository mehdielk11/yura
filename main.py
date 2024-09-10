import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QFileDialog, QComboBox, QTableWidget, QTableWidgetItem
import sqlite3
import pandas as pd

class ReviewApp(QWidget):
    def __init__(self):
        super().__init__()
        
        # Set up the window
        self.setWindowTitle('Product Review Filter')
        self.setGeometry(100, 100, 600, 400)
        
        # Layout
        layout = QVBoxLayout()
        
        # Label
        self.label = QLabel('Select a month to filter products needing review')
        layout.addWidget(self.label)
        
        # Button to import Excel
        self.import_button = QPushButton('Import Excel File')
        self.import_button.clicked.connect(self.import_file)
        layout.addWidget(self.import_button)
        
        # Dropdown for month selection
        self.month_combo = QComboBox()
        self.month_combo.addItems(['January', 'February', 'March', 'April', 'May', 'June', 
                                   'July', 'August', 'September', 'October', 'November', 'December'])
        self.month_combo.currentIndexChanged.connect(self.filter_data)
        layout.addWidget(self.month_combo)
        
        # Table to display filtered data
        self.table = QTableWidget()
        layout.addWidget(self.table)
        
        # Set layout
        self.setLayout(layout)
    
    def import_excel_to_db(self, file_path):
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Print column names for debugging
        print("Column names in the Excel file:", df.columns)
        
        # Connect to SQLite database
        conn = sqlite3.connect('product_reviews.db')
        cursor = conn.cursor()

        # Insert each row into the Products table
        for index, row in df.iterrows():
            # Convert the review_date to string if it's a Timestamp
            review_date = row['Review Date']
            if isinstance(review_date, pd.Timestamp):
                review_date = review_date.strftime('%Y-%m-%d')

            try:
                cursor.execute('''
                INSERT INTO Products (product_name, reference, review_date)
                VALUES (?, ?, ?)
                ''', (row['Product Name'], row['Reference'], review_date))
            except KeyError as e:
                print(f"Column not found: {e}")
            except Exception as e:
                print(f"An error occurred: {e}")

        # Commit and close
        conn.commit()
        conn.close()

        print(f"Data from {file_path} has been successfully imported.")

    def import_file(self):
        # File dialog to select Excel file
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)", options=options)
        if file_name:
            # Import the Excel file
            self.import_excel_to_db(file_name)
    
    def filter_data(self):
        selected_month = self.month_combo.currentText()
        self.display_filtered_products(selected_month)
    
    def display_filtered_products(self, month):
        # Map month to number
        month_number = pd.to_datetime(month, format='%B').month
        
        # Fetch data from database for the selected month
        conn = sqlite3.connect('product_reviews.db')
        cursor = conn.cursor()
        cursor.execute('''
        SELECT product_name, reference, review_date 
        FROM Products
        WHERE strftime('%m', review_date) = ?
        ''', (f'{month_number:02d}',))
        
        records = cursor.fetchall()
        conn.close()
        
        # Display records in the table
        self.table.setRowCount(len(records))
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Product Name', 'Reference', 'Review Date'])
        
        for row_index, row_data in enumerate(records):
            for col_index, col_data in enumerate(row_data):
                self.table.setItem(row_index, col_index, QTableWidgetItem(str(col_data)))

# Run the application
app = QApplication(sys.argv)
window = ReviewApp()
window.show()
sys.exit(app.exec_())
