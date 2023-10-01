import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate
import datetime
from docx2pdf import convert as convert_to_pdf
import os
import win32print
import win32ui
# import win32com.client
from PIL import Image, ImageWin
import win32con
from pdf2image import convert_from_path
from tkinter import PhotoImage
import sys
from pymongo import MongoClient
import subprocess
# from datetime import datetime
# from time import sleep
sys.stderr = open("consoleoutput.log", "w")
client = MongoClient("Mongodb_Link")  # Replace with your MongoDB connection URI
db = client["Royal_Frame_shopee"]
collection1 = db["Invoice"]
collection2=db["Invoice_no"]
def resource_path(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS2
            except Exception:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)
def check_invoice():
    def back_to_home():
        root.destroy()  # Close the current window (invoice or estimation window)
        home_window.deiconify()  # Show the home page window again
    def get_data():
        tree.delete(*tree.get_children())  # Clear previous data in the Treeview
        for doc in collection1.find():
            tree.insert("", "end", values=(doc["invoice"],doc["name"], doc["phone"], doc["address"],doc["gstno"],doc["paytype"],doc["total"],doc["datefull"],doc["invoice_path"]))
# Function to delete selected data from the Treeview and MongoDB
    def delete_data():
        selected_item = tree.selection()
        if selected_item:
            # ino_to_delete = tree.item(selected_item, "values")[0]
            # collection1.delete_one({"invoice": int(ino_to_delete)})
            # os.remove(tree.item(selected_item, "values")[8])
            # status_label.config(text=f"Deleted data for '{ino_to_delete}'")
            # tree.delete(selected_item)  # Remove selected item from Treeview
            path= tree.item(selected_item, "values")[8]
            subprocess.Popen(['start', '', path], shell=True)
        else:
            status_label.config(text="Select a row to delete")
    def search():
        start_month=selected_start_date.get()
        end_month=selected_end_date.get()
        year=year_label_entry.get()
        if(start_month=='' or end_month=='' or year==''):
            status_label.config(text='Enter all values before search')
        elif(start_month.isalpha() or end_month.isalpha() or year.isalpha()):
            status_label.config(text='Enter valid credentials(should be numeric)')
        elif(int(start_month)>int(end_month) or int(year)>datetime.datetime.now().year ):
            status_label.config(text='Start Month cannot be greater than end month and year cannot be in future')
        else:
            status_label.config(text='')
            tree.delete(*tree.get_children())
            query = {"month": {"$gte": int(start_month), "$lte": int(end_month)}, "year": int(year)}
            for doc in collection1.find(query):
                tree.insert("", "end", values=(doc["invoice"],doc["name"], doc["phone"], doc["address"],doc["gstno"],doc["paytype"],doc["total"],doc["datefull"],doc["invoice_path"]))
        


    root =  tk.Toplevel()
    root.title("Check Invoice")
    root.iconbitmap(resource_path('icon.ico'))
    check_frame = ttk.Frame(root)
    check_frame.grid(row=0, column=1, padx=20, pady=10)
    from_date_label = ttk.Label(check_frame, text="Start Month")
    from_date_label.pack()
    month_types = ('1', '2', '3','4','5','6','7','8','9','10','11','12')
    selected_start_date = tk.StringVar()
    payment_combo = ttk.Combobox(check_frame, textvariable=selected_start_date, values=month_types)
    payment_combo.pack()

    to_date_label = ttk.Label(check_frame, text="End Month")
    to_date_label.pack()
    selected_end_date = tk.StringVar()
    payment_combo2 = ttk.Combobox(check_frame, textvariable=selected_end_date, values=month_types)
    payment_combo2.pack()

    year_label = ttk.Label(check_frame, text="Year")
    year_label.pack()
    year_label_entry = ttk.Entry(check_frame)
    year_label_entry.pack()

    search_button = ttk.Button(check_frame, text="Search", command=search)
    search_button.pack(padx=20, pady=20)

    all_button = ttk.Button(check_frame, text="GetAll", command=get_data)
    all_button.pack(padx=20, pady=20)

    tree = ttk.Treeview(check_frame, columns=("Invoice_No","Name", "Phone", "Address","GSTNO","Payment_Type","Total Amount","Date_of_Generation"), show="headings")
    tree.heading("Invoice_No", text="Invoice_No")
    tree.heading("Name", text="Name")
    tree.heading("Phone", text="Phone")
    tree.heading("Address", text="Address")
    tree.heading("GSTNO",text="GSTNO")
    tree.heading("Payment_Type",text="Payment_Type")
    tree.heading("Total Amount",text="Total Amount")
    tree.heading("Date_of_Generation",text="Date_of_Generation")
    tree.pack()
    delete_button = ttk.Button(check_frame, text="Open File", command=delete_data)
    delete_button.pack()
    status_label = ttk.Label(check_frame, text="")
    status_label.pack()
    back_button = ttk.Button(check_frame, text="Back to Home", command=back_to_home)
    back_button.pack(pady=10)
    get_data()
    root.mainloop()

def open_main_window(is_invoice):
    # home_window.destroy()
    def back_to_home():
        window.destroy()  # Close the current window (invoice or estimation window)
        home_window.deiconify()  # Show the home page window again
    
    if is_invoice:

        # Function to clear item fields
        def clear_item():
            qty_spinbox.delete(0, tk.END)
            qty_spinbox.insert(0, "1")
            describe_entry.delete(0, tk.END)
            # height_entry.delete(0, tk.END)
            # width_entry.delete(0, tk.END)
            rate_entry.delete(0, tk.END)
            rate_entry.insert(0, "0.0")

        # Initialize the invoice list and item counter
        invoice_list = []
        global i
        i = 1
        # Function to add an item to the invoice
        def add_item():
            global i
            try:
                qty = int(qty_spinbox.get())
                desc = describe_entry.get()
                # height = float(height_entry.get())
                # width = float(width_entry.get())
                sft_rate = float(rate_entry.get())

                if qty <= 0   or desc == '' or sft_rate<=0 :
                    raise ValueError("Invalid input values")

                # sft = float(height * width)
                total = float(qty * sft_rate)

                invoice_item = [qty, desc, sft_rate, total]
                invoice_var = [i, desc, qty, sft_rate, total]
                tree.insert('', 0, values=invoice_item)
                clear_item()
                invoice_list.append(invoice_var)
                i += 1

            except ValueError:
                messagebox.showerror("Alert", "Enter Correct Credentials (numeric values greater than 0)!")

        # Function to create a new invoice
        def new_invoice():
            global i
            c_name_entry.delete(0, tk.END)
            c_contact_entry.delete(0, tk.END)
            c_address_entry.delete(0, tk.END)
            c_gst_entry.delete(0, tk.END)
            selected_option.set("")
            clear_item()
            tree.delete(*tree.get_children())
            invoice_list.clear()
            i = 1

        # Function to edit an item in the invoice
        def edit_item():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showerror("Edit Item", "Select an item to edit.")
                return

            # Get the selected item's values
            item_values = tree.item(selected_item, 'values')

            # Create a popup window for editing
            edit_window = tk.Toplevel(window)
            edit_window.title("Edit Item")
            edit_window.iconbitmap(resource_path('icon.ico'))

            edit_frame = ttk.Frame(edit_window)
            edit_frame.pack(padx=20, pady=10)

            # Item Details
            qty_label = ttk.Label(edit_frame, text="Qty:")
            qty_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

            qty_entry = ttk.Entry(edit_frame)
            qty_entry.grid(row=0, column=1, padx=5, pady=5)
            qty_entry.insert(0, item_values[0])

            describe_label = ttk.Label(edit_frame, text="Item Description:")
            describe_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

            describe_entry = ttk.Entry(edit_frame)
            describe_entry.grid(row=1, column=1, padx=5, pady=5)
            describe_entry.insert(0, item_values[1])

            rate_label = ttk.Label(edit_frame, text="Rate (per piece):")
            rate_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

            rate_entry = ttk.Spinbox(edit_frame, from_=0.0, to=700, increment=0.5)
            rate_entry.grid(row=2, column=1, padx=5, pady=5)
            rate_entry.insert(0, item_values[2])

            def save_changes():
                edited_qty = int(qty_entry.get())
                edited_desc = describe_entry.get()
                edited_rate = float(rate_entry.get())

                if edited_qty <= 0   or edited_desc == '' or edited_rate<=0:
                    messagebox.showerror("Edit Item", "Invalid input values")
                    return
                edited_total = float(edited_qty  * edited_rate)

                # Update the treeview with the edited values
                edited_item = [edited_qty, edited_desc, edited_rate, edited_total]
                tree.item(selected_item, values=edited_item)

                # Close the edit window
                edit_window.destroy()

            save_button = ttk.Button(edit_frame, text="Save Changes", command=save_changes)
            save_button.grid(row=5, columnspan=2, pady=10)

        # Function to generate and print the invoice
        def generate_invoice():
            def get_printer_names():
                printer_info = win32print.EnumPrinters(2)
                printer_names = [info[2] for info in printer_info]
                return printer_names

            def print_pdf(pdf_file):
                if not pdf_file:
                    return
                pdf_images = convert_from_path(pdf_file, dpi=300)
                selected_printer = printer_var.get()
                printer_dc = win32ui.CreateDC()
                printer_dc.CreatePrinterDC(selected_printer)
                printer_dc.StartDoc("Print Job")
                printer_dc.StartPage()
                desired_width = 16 * 300
                desired_height = 21 * 300
                scale_x = desired_width / pdf_images[0].width
                scale_y = desired_height / pdf_images[0].height
                printer_dc.SetMapMode(win32con.MM_ANISOTROPIC)
                printer_dc.SetViewportExt((int(desired_width), int(desired_height)))
                printer_dc.SetWindowExt((int(desired_width / scale_x), int(desired_height / scale_y)))
                for pdf_image in pdf_images:
                    dib = ImageWin.Dib(pdf_image)
                    dib.draw(printer_dc.GetHandleOutput(), (0, 0, pdf_image.width, pdf_image.height))
                    printer_dc.EndPage()
                printer_dc.EndDoc()
                root.destroy()

            doc = DocxTemplate(resource_path('invoicetemp.docx'))
            name = c_name_entry.get()
            contact = c_contact_entry.get()
            address = c_address_entry.get()
            gst = c_gst_entry.get()
            pay_option = selected_option.get()
            subtotal = sum(item[4] for item in invoice_list)
            tax = 0.01
            tax_amount=(subtotal * tax)
            total = subtotal + (subtotal * tax)
            document = collection2.find_one({"id": "4828"})
            if (name == '' or contact == '' or pay_option == ''):
                messagebox.showinfo("Alert", "Please Fill Customer Name, Contact, Payment Type")
            else:
                doc.render({"ino":document.get("ino")+1,
                            "name": name,
                            "phone": contact,
                            "address": address,
                            "gstno": gst,
                            "paytype": pay_option,
                            "invoicelist": invoice_list,
                            "tax": str(tax * 100) + "%",
                            "tax_amount": tax_amount,
                            "subtotal": subtotal,
                            "total": total,
                            "date": datetime.datetime.now().strftime("%d-%m-%Y")
                            })
                doc_name = 'invoice' + name + datetime.datetime.now().strftime("%d-%m-%y-%H%M%S") + ".docx"
                output_dir = resource_path('output')
                doc.save(os.path.join(output_dir, doc_name))
                file_name_without_extension, old_extension = os.path.splitext(os.path.join(output_dir, doc_name))
                new_extension = ".pdf"
                new_file_name = file_name_without_extension + new_extension
                data = {"invoice":document.get("ino")+1,
                        "name": name,
                        "phone": contact,
                        "address": address,
                        "gstno": gst,
                        "paytype": pay_option,
                        "invoicelist": invoice_list,
                        "tax": str(tax * 100) + "%",
                        "tax_amount": tax_amount,
                        "subtotal": subtotal,
                        "total": total,
                        "datefull": datetime.datetime.now().strftime("%d-%m-%Y"),
                        "date":datetime.datetime.now().day,
                        "month":datetime.datetime.now().month,
                        "year":datetime.datetime.now().year,
                        "invoice_path": new_file_name
                }
                collection2.update_one({"id": "4828"}, {"$set": {"ino": document.get("ino")+1}})
                collection1.insert_one(data)
                try:
                    path=os.path.join(output_dir, doc_name)
                    # messagebox.showinfo("Alert", path)
                    convert_to_pdf(path)
                except Exception as e:
                    # print(f"PDF conversion error: {e}")
                    messagebox.showerror("Alert", e)
                else:
                    # print(f"PDF conversion successful: {doc_name}")
                    # messagebox.showinfo("Success", "PDF file saved successfully")
                    print("pdf saved successful")

                new_invoice()
                os.remove(os.path.join(output_dir, doc_name))


                root = tk.Toplevel(window)
                root.title("PDF Printer")
                root.iconbitmap(resource_path('icon.ico'))

                print_frame=ttk.Frame(root)
                print_frame.pack(padx=20,pady=10)

                printer_label = ttk.Label(print_frame, text="Select Printer:")
                printer_label.grid(row=1,column=1,padx=10,pady=10)
                available_printers = get_printer_names()
                printer_var = tk.StringVar()
                printer_var.set(available_printers[0])

                printer_option_menu = ttk.OptionMenu(print_frame, printer_var, *available_printers)
                printer_option_menu.grid(row=2,column=1,padx=10,pady=10)

                print_button = ttk.Button(print_frame, text="Print PDF", command=lambda: print_pdf(new_file_name))
                print_button.grid(row=3,column=1,padx=20,pady=20)


        # Create the main window
        window = tk.Toplevel()
        window.title("Royal Frames Shopee")
        window.iconbitmap(resource_path('icon.ico'))
        # Create a frame for customer details
        frame = tk.Frame(window)
        frame.grid(row=0, column=0, padx=20, pady=10)

        # Style

        # Customer Details
        c_name = ttk.Label(frame, text="Customer Name:")
        c_name.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        c_name_entry = ttk.Entry(frame)
        c_name_entry.grid(row=0, column=1, padx=5, pady=5)

        c_contact = ttk.Label(frame, text="Customer Contact:")
        c_contact.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        c_contact_entry = ttk.Entry(frame)
        c_contact_entry.grid(row=1, column=1, padx=5, pady=5)

        c_address = ttk.Label(frame, text="Customer Address:")
        c_address.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        c_address_entry = ttk.Entry(frame)
        c_address_entry.grid(row=2, column=1, padx=5, pady=5)

        c_gst = ttk.Label(frame, text="Customer GSTIN:")
        c_gst.grid(row=3, column=0, padx=5, pady=5, sticky="w")

        c_gst_entry = ttk.Entry(frame)
        c_gst_entry.grid(row=3, column=1, padx=5, pady=5)

        c_pay = ttk.Label(frame, text="Payment Type:")
        c_pay.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        # Payment Type Dropdown
        payment_types = ('Cash', 'NetBanking', 'UPI')
        selected_option = tk.StringVar()
        payment_combo = ttk.Combobox(frame, textvariable=selected_option, values=payment_types)
        payment_combo.grid(row=4, column=1, padx=5, pady=5)

        # Item Details
        qty_label = ttk.Label(frame, text="Qty:")
        qty_label.grid(row=5, column=0, padx=5, pady=5, sticky="w")

        qty_spinbox = ttk.Spinbox(frame, from_=1, to=1000)
        qty_spinbox.grid(row=5, column=1, padx=5, pady=5)

        describe_label = ttk.Label(frame, text="Item Description:")
        describe_label.grid(row=6, column=0, padx=5, pady=5, sticky="w")

        describe_entry = ttk.Entry(frame)
        describe_entry.grid(row=6, column=1, padx=5, pady=5)


        rate_label = ttk.Label(frame, text="Rate (per item):")
        rate_label.grid(row=7, column=0, padx=5, pady=5, sticky="w")

        rate_entry = ttk.Spinbox(frame, from_=0.0, to=700, increment=0.5)
        rate_entry.grid(row=7, column=1, padx=5, pady=5)

        # Buttons
        add_item_button = ttk.Button(frame, text="Add Item", command=add_item)
        add_item_button.grid(row=10, column=0, columnspan=2, pady=10)

        new_invoice_button = ttk.Button(frame, text="New Invoice", command=new_invoice)
        new_invoice_button.grid(row=11, column=0, columnspan=2, pady=10)

        # Invoice Preview
        invoice_frame = ttk.Frame(window)
        invoice_frame.grid(row=0, column=1, padx=20, pady=10)

        invoice_label = ttk.Label(invoice_frame, text="Invoice Preview", font=('Arial', 16))
        invoice_label.pack()

        column = ('Qty', 'Description', 'Rate', 'Total')
        tree = ttk.Treeview(invoice_frame, columns=column, show="headings", height=10)
        tree.pack()

        for col in column:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        # Button to edit an item
        edit_item_button = ttk.Button(invoice_frame, text="Edit Item", command=edit_item)
        edit_item_button.pack(padx=20, pady=20)

        generate_invoice_button = ttk.Button(invoice_frame, text="Generate Invoice", command=generate_invoice)
        generate_invoice_button.pack(padx=20, pady=40,ipadx=40,ipady=20)
        # Start the main event loop
        back_button = ttk.Button(invoice_frame, text="Back to Home", command=back_to_home)
        back_button.pack(pady=10)
        window.mainloop()
    else:
        print("Estimate")

home_window = tk.Tk()
home_window.title("Home Page")
home_window.iconbitmap(resource_path('icon.ico'))
# Create labels and buttons on the home page
style = ttk.Style()
style.configure('TLabel', font=('Arial', 12), foreground='green')
style.configure('TButton', font=('Arial', 12), background='green', foreground='green')
home_label = ttk.Label(home_window, text="Welcome Royal Frame Shopee", font=("Arial", 16))
home_label.pack(pady=20)

generate_invoice_button = ttk.Button(home_window, text="Generate Invoice", command=lambda: open_main_window(True))
generate_invoice_button.pack(pady=10)

generate_estimation_button = ttk.Button(home_window, text="Generate Estimation", command=lambda: open_main_window(False))
generate_estimation_button.pack(pady=10)

generate_estimation_button = ttk.Button(home_window, text="Check Invoices Generated", command=lambda: check_invoice())
generate_estimation_button.pack(pady=10)

# Start the main event loop for the home page
home_window.mainloop()