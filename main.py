import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl

def show_inventory():
    filepath = "inventory.xlsx"
    # filepath = "inventory copy.xlsx"
    if os.path.exists(filepath):
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(row)
        
        if data:
            # Display data in a new window or messagebox
            display_inventory_window = tkinter.Toplevel(window)
            display_inventory_window.title("Inventory Data")
            display_inventory_window.geometry("640x480")  # Adjusted window size
            
            # Configure grid columns to expand
            for i in range(len(data[0])):
                display_inventory_window.grid_columnconfigure(i, weight=1)

            for i, row in enumerate(data):
                if i == 0:  # Skip adding buttons for the header row
                    for j, value in enumerate(row):
                        label = tkinter.Label(display_inventory_window, text=value, width=20, anchor="w", padx=5, pady=2)
                        label.grid(row=i, column=j, sticky="ew")
                else:
                    for j, value in enumerate(row):
                        label = tkinter.Label(display_inventory_window, text=value, width=20, anchor="w", padx=5, pady=2)
                        label.grid(row=i, column=j, sticky="ew")

                    # Buttons for increasing, decreasing, and deleting quantity
                    increase_button = tkinter.Button(display_inventory_window, text="Increase", command=lambda i=i: increase_quantity(i, data, display_inventory_window))
                    increase_button.grid(row=i, column=len(row), padx=5, pady=2, sticky="e")
                    
                    decrease_button = tkinter.Button(display_inventory_window, text="Decrease", command=lambda i=i: decrease_quantity(i, data, display_inventory_window))
                    decrease_button.grid(row=i, column=len(row)+1, padx=5, pady=2, sticky="e")

                    delete_button = tkinter.Button(display_inventory_window, text="Delete", command=lambda i=i: delete_item(i, data, display_inventory_window))
                    delete_button.grid(row=i, column=len(row)+2, padx=5, pady=2, sticky="e")
        else:
            tkinter.messagebox.showinfo("Info", "No items found in inventory.")
    else:
        tkinter.messagebox.showerror("Error", "Inventory file not found.")


def increase_quantity(index, data, window):
    # Update quantity and save to Excel
    current_quantity = int(data[index][1])  # Convert quantity to integer
    data[index] = (data[index][0], current_quantity + 1, data[index][2])
    update_inventory(data)
    # Refresh inventory window
    window.destroy()
    show_inventory()

def decrease_quantity(index, data, window):
    # Update quantity if greater than 0 and save to Excel
    current_quantity = int(data[index][1])  # Convert quantity to integer
    if current_quantity > 0:
        data[index] = (data[index][0], current_quantity - 1, data[index][2])
        update_inventory(data)
    else:
        tkinter.messagebox.showinfo("Info", "Quantity cannot be decreased further.")
    # Refresh inventory window
    window.destroy()
    show_inventory()

def delete_item(index, data, window):
    confirm = tkinter.messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this item?")
    if confirm:
        del data[index]  # Remove the item from the data list
        update_inventory(data)  # Update the inventory
        window.destroy()  # Destroy the current window
        show_inventory()  # Refresh inventory window

def update_inventory(data):
    filepath = "inventory.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for item in data:
        sheet.append(item)
    workbook.save(filepath)

def add_new_product_window():
    new_product_window = tkinter.Toplevel(window)
    new_product_window.title("Add New Product")
    new_product_window.geometry("640x480")  # Adjusted window size "640x480"

    # Widgets for adding new product
    add_product_frame = tkinter.Frame(new_product_window)
    add_product_frame.pack(padx=20, pady=10)

    name_label = tkinter.Label(add_product_frame, text="Name:")
    name_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")

    name_entry = tkinter.Entry(add_product_frame)
    name_entry.grid(row=0, column=1, padx=5, pady=5)

    quantity_label = tkinter.Label(add_product_frame, text="Quantity:")
    quantity_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")

    quantity_entry = tkinter.Entry(add_product_frame)
    quantity_entry.grid(row=1, column=1, padx=5, pady=5)

    price_label = tkinter.Label(add_product_frame, text="Price:")
    price_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")

    price_entry = tkinter.Entry(add_product_frame)
    price_entry.grid(row=2, column=1, padx=5, pady=5)

    save_button = tkinter.Button(add_product_frame, text="Save", command=lambda: add_item(name_entry.get(), quantity_entry.get(), price_entry.get(), new_product_window))
    save_button.grid(row=3, columnspan=2, padx=5, pady=5)

def add_item(name, quantity, price, window):
    if name and quantity and price:
        filepath = "inventory.xlsx"
        
        if not os.path.exists(filepath):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            heading = ["Name", "Quantity", "Price"]
            sheet.append(heading)
            workbook.save(filepath)
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        sheet.append([name, quantity, price])
        workbook.save(filepath)
        tkinter.messagebox.showinfo("Success", "Item added to inventory successfully.")
        window.destroy()
    else:
        tkinter.messagebox.showwarning("Error", "Please fill in all fields.")

def load_data():
    path = "inventory.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tkinter.END, values=value_tuple)
        # treeview.insert('', tkinter.END, values=value_tuple, anchor="center")  

def show_data():
    treeview.delete(*treeview.get_children())
    load_data()

def show_gallery_screen():
    options_frame.pack_forget()
    gallery_screen.pack()
    # Load data when gallery screen is shown
    load_data()

# Main window
window = tkinter.Tk()
style = ttk.Style(window)
window.tk.call("source", "forest-light.tcl")
window.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")
window.title("Inventory Management System")
window.geometry("640x480")  # Set window size to 640x480

# Frame for options
options_frame = tkinter.Frame(window)
options_frame.pack(padx=20, pady=10)

# Option buttons
show_inventory_button = tkinter.Button(options_frame, text="Inventory", command=show_inventory)
show_inventory_button.grid(row=0, column=0, padx=10, pady=5)

add_product_button = tkinter.Button(options_frame, text="Stock Manager", command=add_new_product_window)
add_product_button.grid(row=0, column=1, padx=10, pady=5)


gallery_button = tkinter.Button(options_frame, text="Gallery", command=show_gallery_screen)
gallery_button.grid(row=1, column=1, padx=10, pady=5)

gallery_screen = ttk.Frame(window)

frame = ttk.Frame(gallery_screen)
frame.pack()

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Name", "Quantity", "Price")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Name", width=100)
treeview.column("Quantity", width=50)
treeview.column("Price", width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)

back_button = ttk.Button(gallery_screen, text="Back", command=lambda: gallery_screen.pack_forget() or options_frame.pack())
back_button.pack()


window.mainloop()
