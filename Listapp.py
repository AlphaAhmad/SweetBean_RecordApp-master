##--------------------------- Command for making exe -------------------------------
# pyinstaller --onefile --windowed --hidden-import=PyQt5  --hidden-import=docx  --hidden-import=docx2pdf --hidden-import=comtypes --add-data "logo.jpeg;."  Listapp.py
##----------------------------------------------------------------------------------

import sys
import os
import platform
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap, QFont, QDoubleValidator
from PyQt5.QtCore import Qt, QTimer, QDateTime, QSettings
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5.QtCore import QSizeF 
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from reportlab.pdfbase.pdfmetrics import stringWidth

# TODO: Exe Working properly 1 Now move date Above, Increase some size of the text and add the order now field at the end


# from docx.enum.section import WD_SECTION  # Import for section breaks

# Global variables
headers = ["Product", "Quantity (Kg)", "Price (Rs)", "Actions"]
Policy1 = "Purchase sample before buying the product. Bought product won't be refunded or Exchanged."
Policy2 = "Only transfer online payment to the mentioned account number and payment detail."
Fixed_Policy_List = [Policy1,Policy2]
def resource_path(relative_path):
    """Get the absolute path to resource files (for PyInstaller)"""
    if getattr(sys, 'frozen', False):  # If running as compiled .exe
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sweet Bean")
        self.setGeometry(100, 100, 700, 700)  # Optimal window size

        # Initialize settings and last invoice number
        self.settings = QSettings("Sweet Bean", "InvoiceApp")
        self.last_invoice_number = self.settings.value("last_invoice_number", 0, int)
        if self.last_invoice_number == 0:
            self.last_invoice_number = 1

        # Load user-defined policies from settings (excluding fixed ones)
        saved_policies = self.settings.value("policies", [], list)
        
        # Filter out duplicates (Avoid saving fixed policies)
        self.policies = [p for p in saved_policies if p not in Fixed_Policy_List]

        # Load the logo path from settings
        self.logo_path = self.settings.value("logo_path", "logo.jpeg")  # Default logo path

        # Set up the UI
        self.setup_ui()

        # Start the timer for auto-updating date and time
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

    def setup_ui(self):
        """Set up the main UI components."""

        # Main content widget
        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 40)
        self.main_layout.setSpacing(15)

        # Sections of the UI
        self.main_layout.addLayout(self.create_metadata_layout())   # Logo, Date, Invoice, NTN

        # Add Customer Name and Address fields
        self.main_layout.addLayout(self.create_customer_details_layout())  # Customer Name and Address

        # Table for product list
        self.main_layout.addLayout(self.create_input_layout())      # Input fields + Add button
        self.table = self.create_table()
        self.main_layout.addWidget(self.table)

        # Total sum widget
        self.total_sum_widget = self.create_total_sum_widget()
        self.main_layout.addWidget(self.total_sum_widget)

        # Payment section
        self.main_layout.addLayout(self.create_payment_layout())       # Payment method & account
        self.main_layout.addLayout(self.create_payment_status_layout()) # Payment status radio buttons

        # Policy Notes Section
        self.main_layout.addWidget(self.create_policy_layout())

        # Order Now Number detail
        order_now_layout = QVBoxLayout()
        order_now_label = QLabel("Order Now on:")
        order_now_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.order_now = QLineEdit()
        self.order_now.setPlaceholderText("0312-xxxxxxx")
        self.order_now.setFont(QFont("Arial", 12, QFont.Bold))
        order_now_layout.addWidget(order_now_label)
        order_now_layout.addWidget(self.order_now)     
        self.main_layout.addLayout(order_now_layout)

        # Save button
        self.save_button = QPushButton("Save as PDF")
        self.save_button.clicked.connect(self.save_as_pdf)
        self.main_layout.addWidget(self.save_button)

        # Print button (save and go to print log at the same time)
        self.print_button = QPushButton("Save and print") 
        self.print_button.clicked.connect(self.print_document)
        self.main_layout.addWidget(self.print_button)

        # Scrollable Area for Large Content
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.main_widget)

        # Set the central widget
        self.setCentralWidget(self.scroll_area)

        # Apply dark theme stylesheet
        self.setStyleSheet("""
            QMainWindow {
                background-color:rgb(11, 11, 11);
            }
            QLabel {
                font-size: 14px;
                color:rgb(0, 0, 0);
            }
            QLineEdit, QSpinBox, QComboBox {
                font-size: 14px;
                padding: 8px;
                border: 1px solid #555;
                border-radius: 4px;
                background-color: #3d3d3d;
                color:rgb(236, 220, 220);
            }
            QPushButton {
                font-size: 14px;
                padding: 8px 16px;
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QTableWidget {
                font-size: 14px;
                border: 1px solid #555;
                border-radius: 4px;
                background-color: #3d3d3d;
                color:rgb(244, 240, 240);
            }
            QTableWidget::item {
                padding: 8px;
            }
            QHeaderView::section {
                background-color: #007bff;
                color: white;
                padding: 8px;
                font-size: 14px;
            }
            QRadioButton {
                font-size: 14px;
                color:rgb(5, 1, 1);
            }
            QLineEdit[readOnly="true"] {
                background-color: #555;
                color:rgb(207, 207, 207);
            }
            QScrollArea {
                border: none;
            }
            QGroupBox {
                font-size: 14px;
                color:rgb(59, 39, 39);
                border: 1px solid #555;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            QTextEdit {
                font-size: 14px;
                background-color: #3d3d3d;
                color: #ffffff;
                border: 1px solid #555;
                border-radius: 4px;
            }
            QListWidget {
                font-size: 14px;
                background-color: #3d3d3d;
                color: #ffffff;
                border: 1px solid #555;
                border-radius: 4px;
            }
        """)

    def create_metadata_layout(self):
        """Create the metadata layout (logo, date/time, invoice, NTN, and Reset Button)."""
        layout = QGridLayout()
        layout.setSpacing(15)

        # Reset Invoice Button (Top Left)
        self.reset_invoice_btn = QPushButton("Reset Invoice")
        self.reset_invoice_btn.clicked.connect(self.reset_invoice_number)
        layout.addWidget(self.reset_invoice_btn, 0, 0, 1, 2)  # Positioned at the top left

        # Logo
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.logo_label, 1, 0, 1, 2)  # Logo spans 1 row and 2 columns

        # Load the logo from settings or use the default
        self.logo_path = self.settings.value("logo_path", "logo.jpeg")  # Default logo path
        self.update_logo()

        # Change Logo Button
        self.change_logo_btn = QPushButton("Change Logo")
        self.change_logo_btn.clicked.connect(self.change_logo)
        layout.addWidget(self.change_logo_btn, 2, 0, 1, 2)  # Button spans 1 row and 2 columns

        # Date and Time
        self.date_time_label = QLabel()
        self.date_time_label.setAlignment(Qt.AlignCenter)
        self.date_time_label.setFont(QFont("Arial", 12))
        layout.addWidget(self.date_time_label, 3, 0, 1, 2)  # Date and Time spans 1 row and 2 columns

        # Invoice Number (Separate Row)
        invoice_text_label = QLabel("Invoice Number:")
        invoice_text_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.invoice_label = QLabel(f"{self.last_invoice_number}")
        self.invoice_label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(invoice_text_label, 4, 0)
        layout.addWidget(self.invoice_label, 4, 1)

        # Invoice Quotation (New Field)
        invoice_quotation_label = QLabel("Invoice Quotation:")
        invoice_quotation_label.setFont(QFont("Arial", 12, QFont.Bold)) 
        self.invoice_quotation_input = QLineEdit()
        self.invoice_quotation_input.setPlaceholderText("Enter Invoice Quotation")
        self.invoice_quotation_input.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(invoice_quotation_label, 5, 0)
        layout.addWidget(self.invoice_quotation_input, 5, 1)

        # NTN (Separate Row)
        NTN_label = QLabel("NTN Number:")
        NTN_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.ntn_text_label = QLineEdit()
        self.ntn_text_label.setPlaceholderText("Enter NTN number")
        self.ntn_text_label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(NTN_label, 6, 0)
        layout.addWidget(self.ntn_text_label, 6, 1)

         # Sales Tax Number (New Field)
        sales_tax_label = QLabel("Sales Tax Number:")
        sales_tax_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.sales_tax_input = QLineEdit()
        self.sales_tax_input.setPlaceholderText("Enter Sales Tax Number")
        self.sales_tax_input.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(sales_tax_label, 7, 0)
        layout.addWidget(self.sales_tax_input, 7, 1)

        return layout

    def update_logo(self):
        """Update the logo displayed in the UI."""
        logo = QPixmap(self.logo_path)
        if logo.isNull():
            # If the logo file is not found, use a default placeholder
            logo = QPixmap("logo.jpeg")  # Default logo
        scaled_logo = logo.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.logo_label.setPixmap(scaled_logo)

    def change_logo(self):
        """Open a file dialog to select a new logo."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Logo", "", "Images (*.png *.jpg *.jpeg *.bmp)"
        )
        if file_path:
            self.logo_path = file_path
            self.settings.setValue("logo_path", self.logo_path)  # Save the new logo path
            self.update_logo()  # Update the logo in the UI

    def create_customer_details_layout(self):
        """Create the layout for Customer Name and Address."""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # Customer Name
        self.customer_name_label = QLabel("Customer Name:")
        self.customer_name_input = QLineEdit()
        self.customer_name_input.setPlaceholderText("Enter Customer Name")

        # Customer Address
        self.customer_address_label = QLabel("Customer Address:")
        self.customer_address_input = QLineEdit()
        self.customer_address_input.setPlaceholderText("Enter Customer Address")

        # Add widgets to the layout
        layout.addWidget(self.customer_name_label)
        layout.addWidget(self.customer_name_input)
        layout.addWidget(self.customer_address_label)
        layout.addWidget(self.customer_address_input)

        return layout

    def create_input_layout(self):
        """Create the input layout (product name, quantity, price, add button)."""
        layout = QHBoxLayout()
        layout.setSpacing(10)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Product Name")

        self.quantity_input = QSpinBox()
        self.quantity_input.setRange(0, 100000000)
        self.quantity_input.setPrefix("Kg ")

        self.price_input = QSpinBox()
        self.price_input.setRange(0, 1000000000)
        self.price_input.setPrefix("Rs ")

        add_button = QPushButton("Add Product")
        add_button.clicked.connect(self.add_row)

        layout.addWidget(self.name_input)
        layout.addWidget(self.quantity_input)
        layout.addWidget(self.price_input)
        layout.addWidget(add_button)

        return layout

    def create_table(self):
        """Create and configure the table for product list."""
        self.table_widget = QTableWidget()  # Store the table widget as an attribute
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(headers)
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        self.table_widget.verticalHeader().setVisible(False)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)

        # Set row height to make rows bigger
        self.table_widget.verticalHeader().setDefaultSectionSize(40)  # Adjust row height

        # Set table size to show at least 3-4 rows
        self.table_widget.setMinimumHeight(160)  # 4 rows * 40px height

        # Add GST Input and Button below the table
        gst_layout = QHBoxLayout()
        gst_layout.setSpacing(10)

        # GST Input Field
        self.gst_label = QLabel("GST (%):")
        self.gst_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.gst_input = QLineEdit()
        self.gst_input.setPlaceholderText("Enter GST Percentage")
        self.gst_input.setFont(QFont("Arial", 12, QFont.Bold))

        # Add GST Button
        self.add_gst_button = QPushButton("Add GST")
        self.add_gst_button.clicked.connect(self.add_gst_to_total)

        # Add widgets to the layout
        gst_layout.addWidget(self.gst_label)
        gst_layout.addWidget(self.gst_input)
        gst_layout.addWidget(self.add_gst_button)

        # Create a container widget for the table and GST layout
        container_widget = QWidget()
        container_layout = QVBoxLayout(container_widget)
        container_layout.addWidget(self.table_widget)  # Add the table widget
        container_layout.addLayout(gst_layout)  # Add the GST layout

        return container_widget

    def create_total_sum_widget(self):
        """Create the total sum widget."""
        total_sum_widget = QLineEdit()
        total_sum_widget.setPlaceholderText("Total Sum (Rs)")
        total_sum_widget.setReadOnly(True)
        total_sum_widget.setAlignment(Qt.AlignRight)
        total_sum_widget.setFont(QFont("Arial", 12, QFont.Bold))

        # Initialize GST-related variables
        self.total_sum = 0
        self.gst_amount = 0
        self.total_with_gst = 0

        return total_sum_widget

    def add_gst_to_total(self):
        """Add GST to the total sum."""
        try:
            gst_percentage = float(self.gst_input.text())
            if gst_percentage < 0:
                QMessageBox.warning(self, "Error", "GST percentage cannot be negative.")
                return

            # Calculate GST amount
            self.total_sum = sum(
                int(self.table_widget.item(row, 2).text()) for row in range(self.table_widget.rowCount())
            )
            self.gst_amount = (self.total_sum * gst_percentage) / 100
            self.total_with_gst = self.total_sum + self.gst_amount

            # Update the total sum widget
            self.total_sum_widget.setText(
                f"Total Sum: Rs {self.total_sum:.2f} + {gst_percentage}% GST (Rs {self.gst_amount:.2f}) = Rs {self.total_with_gst:.2f}"
            )
        except ValueError:
            QMessageBox.warning(self, "Error", "Please enter a valid GST percentage.")

    def add_row(self):
        """Add a new row to the table."""
        product_name = self.name_input.text()
        quantity = self.quantity_input.value()
        price = self.price_input.value()

        if not product_name:
            QMessageBox.warning(self, "Error", "Product Name field is empty.")
            return

        if not quantity:
            QMessageBox.warning(self, "Error", "Quantity field is empty.")
            return

        if not price:
            QMessageBox.warning(self, "Error", "Price field is empty.")
            return

        row_index = self.table_widget.rowCount()  # Use self.table_widget
        self.table_widget.insertRow(row_index)

        self.table_widget.setItem(row_index, 0, QTableWidgetItem(product_name))
        self.table_widget.setItem(row_index, 1, QTableWidgetItem(f"{quantity}"))
        self.table_widget.setItem(row_index, 2, QTableWidgetItem(f"{price}"))

        # Create a smaller delete button
        delete_button = QPushButton("Delete")
        delete_button.setFixedSize(90, 25)  # Set button size (width, height)
        delete_button.clicked.connect(self.delete_row)

        # Center the text inside the button using stylesheet
        delete_button.setStyleSheet("""
            QPushButton {
                text-align: center;  /* Center the text horizontally */
                padding: 0px;        /* Remove padding to ensure text is centered */
                margin: 0px;         /* Remove margin to ensure text is centered */
            }
        """)

        # Center the button in its column
        button_widget = QWidget()
        button_layout = QHBoxLayout(button_widget)
        button_layout.addWidget(delete_button)
        button_layout.setAlignment(Qt.AlignCenter)  # Center the button
        button_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins

        self.table_widget.setCellWidget(row_index, 3, button_widget)  # Use self.table_widget

        # Clear input fields
        self.name_input.clear()
        self.quantity_input.setValue(0)
        self.price_input.setValue(0)

        # Update the total sum
        self.update_total_sum()

    def delete_row(self):
        """Delete a row from the table."""
        button = self.sender()
        if button:
            index = self.table_widget.indexAt(button.pos())  # Use self.table_widget
            if index.isValid():
                self.table_widget.removeRow(index.row())  # Use self.table_widget
                self.update_total_sum()

    def update_time(self):
        """Update the date and time label."""
        current_time = QDateTime.currentDateTime().toString("dd-MM-yyyy  hh:mm:ss")
        self.date_time_label.setText(current_time)

    def update_invoice_label(self):
        """Update the invoice label."""
        self.invoice_label.setText(f"Invoice number: {self.last_invoice_number}")

    def increment_Invoice_number(self):
        self.last_invoice_number+=1
        self.update_invoice_label()

    def reset_invoice_number(self):
        """Reset invoice number to 1 and update the UI."""
        self.last_invoice_number = 1
        self.settings.setValue("last_invoice_number", self.last_invoice_number)  # Store in QSettings
        self.invoice_label.setText(str(self.last_invoice_number))  # Update UI
        QMessageBox.information(self, "Invoice Reset", "Invoice number has been reset to 1.")

    def update_total_sum(self):
        """Calculate and update the total sum of prices in the table."""
        self.total_sum = sum(
            int(self.table_widget.item(row, 2).text()) for row in range(self.table_widget.rowCount())
        )
        self.total_sum_widget.setText(f"Total Sum: Rs {self.total_sum:.2f} + 0.0% GST = Rs {self.total_sum:.2f}")

    def create_payment_layout(self):
        """Create the payment method, account number, and account owner layout with two pairs of fields."""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # First Payment Method, Account Number, and Account Owner
        self.payment_method1_label = QLabel("Payment Method 1:")
        self.payment_method_input1 = QLineEdit()
        self.payment_method_input1.setPlaceholderText("e.g., Bank Transfer")

        self.account_number1_label = QLabel("Account Number 1:")
        self.account_number_input1 = QLineEdit()
        self.account_number_input1.setPlaceholderText("e.g., 1234-5678-9012")

        self.account_owner1_label = QLabel("Account Owner 1:")
        self.account_owner_input1 = QLineEdit()
        self.account_owner_input1.setPlaceholderText("Name of Account Number 1 Owner")

        # Second Payment Method, Account Number, and Account Owner
        self.payment_method2_label = QLabel("Payment Method 2:")
        self.payment_method_input2 = QLineEdit()
        self.payment_method_input2.setPlaceholderText("e.g., Easypaisa")

        self.account_number2_label = QLabel("Account Number 2:")
        self.account_number_input2 = QLineEdit()
        self.account_number_input2.setPlaceholderText("e.g., 0312-xxxxxx9")

        self.account_owner2_label = QLabel("Account Owner 2:")
        self.account_owner_input2 = QLineEdit()
        self.account_owner_input2.setPlaceholderText("Name of Account Number 2 Owner")

        # Add widgets to the layout
        layout.addWidget(QLabel("Payment Details"))
        layout.addWidget(self.payment_method1_label)
        layout.addWidget(self.payment_method_input1)
        layout.addWidget(self.account_number1_label)
        layout.addWidget(self.account_number_input1)
        layout.addWidget(self.account_owner1_label)
        layout.addWidget(self.account_owner_input1)
        layout.addWidget(self.payment_method2_label)
        layout.addWidget(self.payment_method_input2)
        layout.addWidget(self.account_number2_label)
        layout.addWidget(self.account_number_input2)
        layout.addWidget(self.account_owner2_label)
        layout.addWidget(self.account_owner_input2)

        return layout

    def create_payment_status_layout(self):
        """Create the payment status radio buttons layout with a remaining payment input field."""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # Payment Status Radio Buttons
        self.paid_radio = QRadioButton("Paid")
        self.pending_radio = QRadioButton("Pending")

        button_group = QButtonGroup(self)
        button_group.addButton(self.paid_radio)
        button_group.addButton(self.pending_radio)

        layout.addWidget(QLabel("Payment Status"))
        layout.addWidget(self.paid_radio)
        layout.addWidget(self.pending_radio)

        # Remaining Payment Input Field
        self.remaining_payment_label = QLabel("Remaining Payment (Rs):")
        self.remaining_payment_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.remaining_payment_input = QLineEdit()
        self.remaining_payment_input.setPlaceholderText("Enter Remaining Payment")
        self.remaining_payment_input.setFont(QFont("Arial", 12, QFont.Bold))

        # Allow only numbers in the input field
        self.remaining_payment_input.setValidator(QDoubleValidator(0, 999999999, 2, self))  # Allow numbers with up to 2 decimal places

        # Add widgets to the layout
        layout.addWidget(self.remaining_payment_label)
        layout.addWidget(self.remaining_payment_input)

        return layout

    def create_policy_layout(self):
        """Create the policy notes section as a QWidget."""
        group_box = QGroupBox("Policies")  # Group box to contain policies
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # Policy list widget
        self.policy_list_widget = QListWidget()

        # Add fixed policies first
        for policy in Fixed_Policy_List:
            self.add_policy_to_list(policy, fixed=True)

        # Add user-defined policies
        for policy in self.policies:
            self.add_policy_to_list(policy)

        # Input field for new policies
        self.new_policy_input = QTextEdit()
        self.new_policy_input.setPlaceholderText("Enter a new policy")
        self.new_policy_input.setFixedHeight(50)

        # Add Policy button
        add_policy_button = QPushButton("Add Policy")
        add_policy_button.clicked.connect(self.add_policy)

        # Add widgets to layout
        layout.addWidget(self.policy_list_widget)
        layout.addWidget(self.new_policy_input)
        layout.addWidget(add_policy_button)

        group_box.setLayout(layout)  # Set layout to the group box
        return group_box  # Return a QWidget (QGroupBox)

    def add_policy(self):
        """Add a new policy to the list."""
        new_policy = self.new_policy_input.toPlainText().strip()  # Use toPlainText() for QTextEdit
        if new_policy: 
            self.policies.append(new_policy)  # Add to the list
            self.add_policy_to_list(new_policy)  # Add to the UI
            self.new_policy_input.clear()  # Clear the input field

            # Save policies to settings
            self.settings.setValue("policies", self.policies)
        else:
            QMessageBox.warning(self, "Error", "Policy field is empty.")

    def delete_policy(self, item):
        """Delete a policy from the list."""
        row = self.policy_list_widget.row(item)
        if row != -1:
            policy_text = self.policies[row-2]  # Get text from the stored list
            del self.policies[row-2]  # Remove from list
            self.policy_list_widget.takeItem(row)  # Remove from UI

            # Save updated policies
            self.settings.setValue("policies", self.policies)

    def add_policy_to_list(self, policy, fixed=False):
        """Add a policy to the list widget with an optional delete button."""
        item = QListWidgetItem()
        self.policy_list_widget.addItem(item)

        # Create a QTextEdit for policy text (to enable word wrapping)
        policy_text = QTextEdit(policy)
        policy_text.setReadOnly(True)
        policy_text.setFrameStyle(QFrame.NoFrame)
        policy_text.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        policy_text.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        policy_text.setFixedHeight(50)

        # Create a widget layout
        widget = QWidget()
        widget_layout = QHBoxLayout(widget)
        widget_layout.addWidget(policy_text)

        # Only add delete button for non-fixed policies
        if not fixed:
            delete_button = QPushButton("Delete")
            delete_button.setFixedSize(60, 25)
            delete_button.clicked.connect(lambda: self.delete_policy(item))
            widget_layout.addWidget(delete_button)

        widget_layout.setContentsMargins(0, 0, 0, 0)
        item.setSizeHint(widget.sizeHint())
        self.policy_list_widget.setItemWidget(item, widget)

    """Save the entire document as a .pdf file."""
    def save_as_pdf(self):
        try:
            #++++++++++++++ Function To add wrapped text in pdf ++++++++++++++
            def draw_wrapped_text(text):
                nonlocal y_position
                words = text.split()
                line = ""
                for word in words:
                    if stringWidth(line + word, "Helvetica", 20) < content_width:
                        line += word + " "
                    else:
                        pdf.drawString(left_margin, y_position, line.strip())
                        y_position -= 0.4 * inch
                        line = word + " "
                pdf.drawString(left_margin, y_position, line.strip())
                y_position -= 0.4 * inch
            
            # Create a PDF document
            from reportlab.lib.pagesizes import letter
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import inch
            from reportlab.platypus import Table, TableStyle
            from reportlab.lib import colors

            file_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
            if not file_path:
                return  # User canceled the save dialog

            if not file_path.endswith(".pdf"):
                file_path += ".pdf"

            # Define page size and margins
            page_width, page_height = 9.5 * inch, 40.0 * inch
            left_margin, right_margin = 0.5 * inch, 0.5 * inch
            top_margin, bottom_margin = 0.5 * inch, 0.5 * inch
            content_width = page_width - left_margin - right_margin
            y_position = page_height - top_margin

            pdf = canvas.Canvas(file_path, pagesize=(page_width, page_height))
            pdf.setFont("Helvetica", 21.5)

            # Add logo if available
            if self.logo_path:
                logo_width = 2.5 * inch
                logo_height = 2.5 * inch
                pdf.drawInlineImage(self.logo_path, (page_width - logo_width) / 2, y_position - logo_height, width=logo_width, height=logo_height)
                y_position -= logo_height + 0.5 * inch

            # Invoice details
            pdf.setFont("Helvetica-Bold", 24)
            pdf.drawString(left_margin, y_position, "Invoice Details")
            pdf.setFont("Helvetica", 21.5)
            y_position -= 0.5 * inch

            pdf.drawString(left_margin, y_position, f"Date and Time: {self.date_time_label.text()}")
            y_position -= 0.4 * inch

            pdf.drawString(left_margin, y_position, f"Invoice Number: {self.last_invoice_number}")
            y_position -= 0.4 * inch

            pdf.setFont("Helvetica-Bold", 21.5)
            pdf.drawString(left_margin, y_position, "Invoice Quotation:")
            y_position -= 0.4 * inch

            pdf.setFont("Helvetica", 21.5)
            draw_wrapped_text(self.invoice_quotation_input.text())  # Use the function for wrapping
            y_position -= 0.4 * inch

            pdf.drawString(left_margin, y_position, f"NTN: {self.ntn_text_label.text()}")
            y_position -= 0.4 * inch

            pdf.drawString(left_margin, y_position, f"Sales Tax Number: {self.sales_tax_input.text()}")
            y_position -= 1 * inch

            # Customer details
            pdf.setFont("Helvetica-Bold", 24)
            pdf.drawString(left_margin, y_position, "Customer Details")
            pdf.setFont("Helvetica", 21.5)
            y_position -= 0.5 * inch
            pdf.drawString(left_margin, y_position, f"Customer Name: {self.customer_name_input.text()}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Customer Address: {self.customer_address_input.text()}")
            y_position -= 1 * inch

                # Function to center text
            def draw_centered_text(text, font="Helvetica-Bold", font_size=24, y_offset=0):
                pdf.setFont(font, font_size)
                text_width = stringWidth(text, font, font_size)
                center_x = (page_width - text_width) / 2  # Calculate centered position
                pdf.drawString(center_x, y_position + y_offset, text)

            # Table for items
            draw_centered_text("Items", "Helvetica-Bold", 24)  # Centered heading
            pdf.setFont("Helvetica", 21.5)
            y_position -= 0.5 * inch

            table_data = [["Item", "Qty", "Price"]]
            for row in range(self.table_widget.rowCount()):
                row_data = []
                for col in range(3):
                    item = self.table_widget.item(row, col)
                    row_data.append(item.text() if item else "")
                table_data.append(row_data)

            table = Table(table_data, colWidths=[content_width/3]*3)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.white),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 21),  # Increased text size inside the table
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 20),
                ('TOPPADDING', (0, 0), (-1, -1), 10),  # Added top padding for spacing
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ]))
            table.wrapOn(pdf, content_width, y_position)
            table.drawOn(pdf, left_margin, y_position - (len(table_data) * 0.4 * inch))
            y_position -= (len(table_data) + 1) * 0.4 * inch

            # Total price
            pdf.drawString(left_margin, y_position, f"Total: {self.total_sum_widget.text()}")
            y_position -= 1 * inch

            # Payment details
            draw_centered_text("Payment Details")  # Centered heading
            pdf.setFont("Helvetica", 21.5)
            y_position -= 0.5 * inch

            pdf.drawString(left_margin, y_position, f"Payment Method 1: {self.payment_method_input1.text()}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Account Number 1: {self.account_number_input1.text()}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Account Owner 1: {self.account_owner_input1.text()}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Payment Method 2: {self.payment_method_input2.text()}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Account Number 2: {self.account_number_input2.text()}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Account Owner 2: {self.account_owner_input2.text()}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Payment Status: {'Paid' if self.paid_radio.isChecked() else 'Pending'}")
            y_position -= 0.4 * inch
            pdf.drawString(left_margin, y_position, f"Remaining Payment: {self.remaining_payment_input.text()}")  
            y_position -= 1 * inch

            # Policies
            draw_centered_text("Policies")  # Centered heading
            pdf.setFont("Helvetica", 21.5)
            y_position -= 0.5 * inch
            draw_wrapped_text(Policy1)
            y_position -= 0.4 * inch
            draw_wrapped_text(Policy2)
            y_position -= 0.4 * inch
            for policy in self.policies:
                draw_wrapped_text(policy)
                y_position -= 0.4 * inch

            # Order Now section (Centered heading and text)
            draw_centered_text("Order Now")  # Centered heading
            pdf.setFont("Helvetica", 21.5)
            y_position -= 0.5 * inch

            # Centering "Order Now on" text
            order_now_text = f"Order Now on: {self.order_now.text()}"
            order_now_width = stringWidth(order_now_text, "Helvetica", 21.5)
            order_now_x = (page_width - order_now_width) / 2
            pdf.drawString(order_now_x, y_position, order_now_text)
            y_position -= 0.4 * inch


            # Save the PDF
            pdf.save()

            QMessageBox.information(self, "Success", "PDF saved successfully!")
            self.increment_Invoice_number()
            return file_path

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")


    def print_document(self):
        """Save the document as a PDF and open the print dialog."""
        try:
            # Save the document as a PDF
            pdf_path = self.save_as_pdf()
            if not pdf_path:
                return  # Exit if saving failed

            print(f"PDF path: {pdf_path}")  # Debug: Print the PDF path

            # Set up the printer
            printer = QPrinter(QPrinter.HighResolution)

            # Set custom paper size for thermal printer (e.g., 80mm width)
            thermal_paper_width = 80  # Width in millimeters (adjust based on your printer)
            thermal_paper_height = 297  # Length in millimeters (A4 length for roll paper)
            printer.setPageSize(QPrinter.Custom)  # Set to custom paper size
            printer.setPaperSize(QSizeF(thermal_paper_width, thermal_paper_height), QPrinter.Millimeter)  # Set size in mm

            # Set orientation to Portrait
            printer.setOrientation(QPrinter.Portrait)

            # Open the print dialog
            print_dialog = QPrintDialog(printer, self)
            if print_dialog.exec_() == QPrintDialog.Accepted:
                try:
                    # Print the PDF (Windows-specific)
                    if platform.system() == "Windows":
                        os.startfile(pdf_path, "print")
                    else:
                        # For non-Windows systems, use QTextDocument to print the PDF
                        from PyQt5.QtGui import QTextDocument
                        text_document = QTextDocument()

                        # Load the PDF content
                        with open(pdf_path, "rb") as pdf_file:
                            text_document.setHtml(pdf_file.read().decode("utf-8", errors="ignore"))

                        # Scale the content to fit the thermal printer's paper width
                        scale_factor = thermal_paper_width / 210  # Scale factor for A4 width (210mm)
                        text_document.setPageSize(QSizeF(thermal_paper_width, thermal_paper_height * scale_factor))

                        # Print the document
                        text_document.print_(printer)
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"An error occurred while printing: {str(e)}")
                    print(f"Printing error: {str(e)}")  # Debug: Print the error

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    
  
    def closeEvent(self, event):
        """Save the last invoice number when the window is closed."""
        self.settings.setValue("last_invoice_number", self.last_invoice_number)
        event.accept()


if __name__ == "__main__":

    app = QApplication(sys.argv)
    # Force light mode or Fusion style to avoid OS interference
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
