import tkinter as tk
from tkinter import ttk


class InvoiceAutomation:
    
    def __init__(self):

        # Create the main application window
        self.root = tk.Tk()
        self.root.title("GDNC Invoice Automation System")

        # Set a consistent padding for widgets
        PAD = 5

        # Create the Info frame
        self.info_frame = ttk.LabelFrame(self.root, text="Info")
        self.info_frame.grid(row=0, column=0, padx=PAD, pady=PAD, sticky="ew")

        ttkwidgets = ["Company", "Title", "Name", "Address", "Postcode", "City"]
        for i, label in enumerate(ttkwidgets):
            ttk.Label(self.info_frame, text=label).grid(row=i, column=0, sticky="w", padx=PAD, pady=PAD)
            ttk.Entry(self.info_frame).grid(row=i, column=1, sticky="ew", padx=PAD, pady=PAD)
        self.info_frame.columnconfigure(1, weight=1)

        # Create the Contact frame
        self.contact_frame = ttk.LabelFrame(self.root, text="Contact")
        self.contact_frame.grid(row=1, column=0, padx=PAD, pady=PAD, sticky="ew")

        ttkwidgets_contact = ["Phone", "Email"]
        for i, label in enumerate(ttkwidgets_contact):
            ttk.Label(self.contact_frame, text=label).grid(row=i, column=0, sticky="w", padx=PAD, pady=PAD)
            ttk.Entry(self.contact_frame).grid(row=i, column=1, sticky="ew", padx=PAD, pady=PAD)
        self.contact_frame.columnconfigure(1, weight=1)

        # Create the Invoice frame
        self.invoice_frame = ttk.LabelFrame(self.root, text="Invoice")
        self.invoice_frame.grid(row=2, column=0, padx=PAD, pady=PAD, sticky="ew")

        ttkwidgets_invoice = ["Number", "Date", "Deadline"]
        for i, label in enumerate(ttkwidgets_invoice):
            ttk.Label(self.invoice_frame, text=label).grid(row=i, column=0, sticky="w", padx=PAD, pady=PAD)
            ttk.Entry(self.invoice_frame).grid(row=i, column=1, sticky="ew", padx=PAD, pady=PAD)
        self.invoice_frame.columnconfigure(1, weight=1)

        # Create the Products frame
        self.products_frame = ttk.LabelFrame(self.root, text="Products")
        self.products_frame.grid(row=3, column=0, padx=PAD, pady=PAD, sticky="ew")

        # Default product row
        self.frame_default = ttk.Frame(self.products_frame)
        self.frame_default.grid(row=0, column=0, sticky="ew", pady=PAD)

        self.default_label = ttk.Label(self.frame_default, text="Default")
        self.default_label.grid(row=0, column=0, sticky="w", padx=PAD)
        self.default_quantity = ttk.Spinbox(self.frame_default, from_=1, to=100, width=5)
        self.default_quantity.grid(row=0, column=1, padx=PAD)
        self.default_product = ttk.Combobox(self.frame_default, values=["Product 1 from the selection"])
        self.default_product.grid(row=0, column=2, padx=PAD, sticky="ew")
        self.default_price = ttk.Entry(self.frame_default, width=10)
        self.default_price.insert(0, "100 CHF")
        self.default_price.grid(row=0, column=3, padx=PAD)
        ttk.Button(self.frame_default, text="+").grid(row=0, column=4, padx=PAD)

        # Custom product row
        self.frame_custom = ttk.Frame(self.products_frame)
        self.frame_custom.grid(row=1, column=0, sticky="ew", pady=PAD)

        self.custom_label = ttk.Label(self.frame_custom, text="Custom")
        self.custom_label.grid(row=0, column=0, sticky="w", padx=PAD)
        self.custom_quantity = ttk.Spinbox(self.frame_custom, from_=1, to=100, width=5)
        self.custom_quantity.grid(row=0, column=1, padx=PAD)
        self.custom_product = ttk.Entry(self.frame_custom)
        self.custom_product.insert(0, "My Custom Product 1")
        self.custom_product.grid(row=0, column=2, padx=PAD, sticky="ew")
        self.custom_price = ttk.Entry(self.frame_custom, width=10)
        self.custom_price.insert(0, "100 CHF")
        self.custom_price.grid(row=0, column=3, padx=PAD)
        ttk.Button(self.frame_custom, text="+").grid(row=0, column=4, padx=PAD)

        self.products_frame.columnconfigure(0, weight=1)

        # Price frame
        self.price_frame = ttk.Frame(self.root)
        self.price_frame.grid(row=4, column=0, padx=PAD, pady=PAD, sticky="ew")

        ttk.Label(self.price_frame, text="Price").grid(row=0, column=0, sticky="w", padx=PAD)
        self.price_entry = ttk.Entry(self.price_frame, width=10)
        self.price_entry.insert(0, "200 CHF")
        self.price_entry.grid(row=0, column=1, sticky="e", padx=PAD)

        # Create Invoice button
        self.create_invoice_button = ttk.Button(self.root, text="Create invoice")
        self.create_invoice_button.grid(row=5, column=0, padx=PAD, pady=PAD, sticky="ew")

        self.root.mainloop()


if __name__ == "__main__":
    app = InvoiceAutomation()
