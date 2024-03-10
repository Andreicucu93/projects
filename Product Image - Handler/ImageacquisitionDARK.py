#Most up-to-date PRODUCT IMAGE HANDLER 22724
#In this version I've successfully SPLIT the REMB and Cropp functionalities
#Added batch Name% filter, fixed dropped indicator and added filename caution (duplicate) indicator, scrollbar
#3/10/24 Versioning, icon

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import pyperclip
import os
import shutil
import tkinter.filedialog as fd
import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD
import customtkinter
import datetime
from PIL import Image, ImageTk, ImageOps
import sys
import io
import numpy as np
import threading
import tempfile
import webbrowser
import requests
from packaging.version import parse


#To be compileable..
if hasattr(sys, '_MEIPASS'):
    os.environ['TKDND_LIBRARY'] = os.path.join(sys._MEIPASS, 'tkdnd')


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def pack_files_in_folders_and_submit(dropped_files_mapping, target_directory="D:/Images"):
    # Ensure base target directory exists
    if not os.path.exists(target_directory):
        os.makedirs(target_directory, exist_ok=True)

    for new_file_path, product_id in dropped_files_mapping.items():
        base_destination_folder = os.path.join(target_directory, product_id)
        destination_folder = base_destination_folder
        counter = 1

        # Check if the folder exists and create a unique folder name if necessary
        while os.path.exists(destination_folder):
            destination_folder = f"{base_destination_folder}({counter})"
            counter += 1

        # Now, we have a unique destination folder
        os.makedirs(destination_folder, exist_ok=True)

        destination_path = os.path.join(destination_folder, os.path.basename(new_file_path))
        try:
            shutil.copy(new_file_path, destination_path)
            print(f"Moved {os.path.basename(new_file_path)} to {destination_folder}")
        except Exception as e:
            print(f"Error moving {os.path.basename(new_file_path)}: {e}")

    # Optionally clear the mapping and notify the user
    dropped_files_mapping.clear()
    messagebox.showinfo("Pack Files", "All dropped files have been successfully packed into their respective folders.")


def pack_files():
    image_folder_path = fd.askdirectory(title="Select a folder containing images")
    if not image_folder_path:
        return
    for item_id in tree.get_children():
        item = tree.item(item_id, 'values')
        file_name = item[5]
        product_id = item[1]
        matching_files = [file for file in os.listdir(image_folder_path) if
                          file.startswith(file_name.split('.')[0]) and file.lower().endswith(
                              ('.png', '.jpg', '.jpeg', '.bmp'))]
        num_files = len(matching_files)
        if matching_files:
            destination_folder = os.path.join(image_folder_path, product_id)
            os.makedirs(destination_folder, exist_ok=True)
            for file in matching_files:
                source_file_path = os.path.join(image_folder_path, file)
                destination_file_path = os.path.join(destination_folder, file)
                shutil.move(source_file_path, destination_file_path)

            # Update the tree item based on the number of files moved
            emoji = "📦" * num_files
            tree.set(item_id, column='Status',
                     value=f"Packed {emoji}")  # Mark as Packed with corresponding number of emojis
            print(f"Moved {num_files} files to {destination_folder}")
    update_header_count()


def show_temporary_message(message, text_color="black"):  # Default color is black
    msg_label = customtkinter.CTkLabel(root, text=message, text_color=text_color)
    msg_label.grid(row=13, column=0)  # Adjust the grid position as needed
    root.after(3000, msg_label.destroy)  # Message will disappear after 3000 milliseconds (3 seconds)


def normalize_path(path):
    """Normalize a file path to use the correct separators and ensure it's absolute."""
    normalized_path = os.path.normpath(path.strip("{}").strip())
    if not os.path.isabs(normalized_path):
        # Convert to absolute path if not already
        normalized_path = os.path.abspath(normalized_path)
    return normalized_path

dropped_items = {}
def rename_file_drag_drop(suffix):
    def inner(event):
        global new_file_name, dropped_items
        selected_item = tree.focus()

        if not selected_item:
            show_popup_message("No product selected.", 1000)
            return
        item_values = tree.item(selected_item, 'values')
        if not item_values:
            show_popup_message("No data in the selected record.", 1000)
            return

        original_file_name = item_values[5]
        file_path = event.data.strip("{}*")  # Normalize the file path
        _, file_extension = os.path.splitext(file_path)
        new_file_name = f"{original_file_name}{suffix}{file_extension}"
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)

        try:
            if not os.path.exists(new_file_path):
                os.rename(file_path, new_file_path)
                show_popup_message(f"File renamed to {new_file_name}", 1000)
                # Directly apply "Dropped" status
                tree.item(selected_item, tags=('dropped',))
                tree.set(selected_item, column="#1", value="Dropped")
                dropped_items[selected_item] = new_file_path
            else:
                show_popup_message(f"A file named '{new_file_name}' already exists.", 1000)
        except Exception as e:
            show_popup_message(f"Error during renaming: {e}", 1000)
            return False

        # Ensure Treeview has focus and re-bind the delete key
        tree.focus_set()
        tree.bind('<Delete>', on_delete_key)
        reapply_dropped_status()  # Ensure this is called to refresh the "Dropped" status

        return True

    update_header_count()
    filter_and_highlight_duplicates()
    clear_and_highlight_duplicates()
    reapply_dropped_status()
    return inner


def reapply_dropped_status():
    global dropped_items
    # Ensure this iterates over item IDs, which are the keys in `dropped_items`
    for item_id in dropped_items.keys():
        tree.item(item_id, tags=('dropped',))
        tree.set(item_id, column="#1", value="Dropped")
    # Configure the 'dropped' tag as needed
    tree.tag_configure('dropped', background='#90EE90')



def load_from_excel():
    file_path = fd.askopenfilename(filetypes=[("Excel Files", "*.xlsx")], title="Open File")
    if file_path:
        try:
            df = pd.read_excel(file_path)
            clear_records()  # Clear existing records in the treeview

            # Find the column name that starts with 'Image status'
            image_status_col = next(col for col in df.columns if col.startswith('Status'))

            for _, row in df.iterrows():
                # Use the dynamic column name for 'Image status'
                tree.insert('', 'end', values=(row[image_status_col], row['ID'], row['Name'], row['MC_RPL_UPC'],
                                               row['Package Type'], row['File Name']))
            update_header_count()  # Update the count after loading
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")


def update_header_count():
    global count_loaded
    count_loaded = 0
    for child in tree.get_children():
        image_status = tree.set(child, 'Status')
        count_loaded += image_status.count("📦")
    # Correctly display the count of packed items in the "Image status" column header
    tree.heading("Status", text=f"Status ({count_loaded})")

    # Calculate the total number of items for the "Name" column header
    count_total_items = len(tree.get_children())  # Get the total number of items in the tree
    tree.heading("Name", text=f"Name ({count_total_items})")  # Update "Name" column header with total item count

def on_delete_key(event):
    selected_items = tree.selection()
    items_to_remove = []
    for item in selected_items:
        if item in dropped_items:
            items_to_remove.append(item)
        tree.delete(item)
    for item in items_to_remove:
        del dropped_items[item]  # Remove deleted items from dropped_items
    update_header_count()

    filter_and_highlight_duplicates()
    clear_and_highlight_duplicates()
    reapply_dropped_status()

def reapply_dropped_status():
    for item_id in dropped_items:
        tree.item(item_id, tags=('dropped',))
        tree.set(item_id, column="#1", value="Dropped")
    tree.tag_configure('dropped', background='#90EE90')


def clear_and_highlight_duplicates():
    # Step 1: Clear 'duplicate' tag from all items and reset the filename warning
    for item_id in tree.get_children():
        tree.item(item_id, tags=())
        tree.set(item_id, column="#1", value='')

    for item_id in tree.get_children():
        if 'dropped' not in tree.item(item_id, 'tags'):
            tree.set(item_id, column="#1", value='')

    # Dictionary to track file names and associated item IDs
    file_name_to_item_ids = {}
    for item_id in tree.get_children():
        item = tree.item(item_id, 'values')
        file_name = item[5]  # Assuming the file name is in the 6th column
        if file_name not in file_name_to_item_ids:
            file_name_to_item_ids[file_name] = [item_id]
        else:
            file_name_to_item_ids[file_name].append(item_id)

    # Apply the 'duplicate' tag to items with duplicate file names
    for file_name, item_ids in file_name_to_item_ids.items():
        if len(item_ids) > 1:  # If more than one item has this file name, it's a duplicate
            for item_id in item_ids:
                tree.item(item_id, tags=('filename ⚠',))
                tree.set(item_id, column="#1", value='filename ⚠')
        else:
            # This item is not a duplicate anymore; clear the 'filename ⚠' text
            tree.set(item_ids[0], column="#1", value='')

    # Configure the 'duplicate' tag appearance
    tree.tag_configure('filename ⚠', background='gold')



def on_tree_cell_click(event):
    region = tree.identify("region", event.x, event.y)
    column = tree.identify_column(event.x)
    row_id = tree.identify_row(event.y)

    # Check if the clicked region is a cell and the column is the "Image" column
    if region == "cell" and column == "#1" and row_id:
        current_value = tree.set(row_id, column='Status')
    update_header_count()

#sdf
def on_treeview_click(event):
    region = tree.identify("region", event.x, event.y)
    column = tree.identify_column(event.x)
    if region == "cell" and column == "#1":  # Check if the click is on the "Image" column
        row_id = tree.identify_row(event.y)
        if row_id:  # Check if a row is clicked
            current_value = tree.set(row_id, column='Status')
            update_header_count()  # Update the count


def export_to_excel():
    global current_date
    # Count the number of items with "📦"
    count_loaded = sum("📦" in tree.set(child, 'Status') for child in tree.get_children())
    # Total number of items isn't needed in export but ensure the logic here matches what you display

    data = []
    for item_id in tree.get_children():
        item = tree.item(item_id, 'values')
        data.append(item)
    if data:
        # Keep column names consistent, without dynamically changing them based on counts
        df = pd.DataFrame(data, columns=["Status", "ID", "Name", "MC_RPL_UPC", "Package Type", "File Name"])
        # Generate filename with current date
        current_date = datetime.datetime.now().strftime("%Y-%m-%d")
        default_filename = f"Imaging {current_date}.xlsx"

        file_path = fd.asksaveasfilename(defaultextension='.xlsx',
                                         initialfile=default_filename,
                                         filetypes=[("Excel Files", "*.xlsx")],
                                         title="Save the File")
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Data exported successfully")

def rename_selected_file(suffix):
    selected_item = tree.focus()  # Get selected item
    if not selected_item:
        show_temporary_message("No product selected.", fg='black')
        return
    item_values = tree.item(selected_item, 'values')
    if not item_values:
        show_temporary_message("No data in the selected record.", fg='black')
        return
    original_file_name = item_values[5]  # Assuming file name is in the 6th column (index 5)
    file_path = fd.askopenfilename(title="Select an Image", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.webp")])
    if not file_path:
        return  # No file was selected
    _, file_extension = os.path.splitext(file_path)
    new_file_name = f"{original_file_name}{suffix}{file_extension}"
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
    try:
        os.rename(file_path, new_file_path)
    except Exception as e:
        show_temporary_message(f"Error: {e}")


def clear_records():
    for i in tree.get_children():
        tree.delete(i)
    update_header_count()


def show_popup_message(message, duration=2000):
    popup = tk.Toplevel(root)
    popup.transient(root)  # Make the popup a transient window of the main app
    popup.grab_set()  # Grab the focus
    popup.configure(bg='#c1b3b3')
    popup.overrideredirect(True)  # Remove window decorations

    # Window size and position
    width = 250
    height = 50
    x = root.winfo_x() + (root.winfo_width() // 2) - (width // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (height // 2)
    popup.geometry(f"{width}x{height}+{x}+{y}")

    # Add a label with the message
    message_label = customtkinter.CTkLabel(popup, text=message, bg_color='#525c95')
    message_label.pack(expand=True, fill='both', pady=10)

    popup.lift()  # Bring the popup above other windows
    popup.focus_force()  # Force focus on the popup

    # Automatically destroy the popup after 'duration' milliseconds
    popup.after(duration, popup.destroy)



def look_online():
    copy_file_name_to_clipboard()
    selected_item = tree.focus()  # Get selected item
    item_values = tree.item(selected_item, 'values')
    if item_values:
        name = item_values[2]  # Assuming the Name is in the 2nd column
        search_query = f"https://www.google.com/search?tbm=isch&q={name}"
        webbrowser.open(search_query)

def copy_file_name_to_clipboard():
    selected_item = tree.focus()  # Get selected item
    item_values = tree.item(selected_item, 'values')
    if item_values:
        file_name = item_values[5]
        pyperclip.copy(file_name)

def upc_length(upc):
    if len(upc) == 9 or (len(upc) == 10 and not upc[-1].isdigit()):
        return 4
    elif len(upc) == 10 or (len(upc) == 11 and not upc[-1].isdigit()):
        return 5
    elif len(upc) == 11 or (len(upc) == 12 and not upc[-1].isdigit()):
        return 6
    elif len(upc) == 12 or (len(upc) == 13 and not upc[-1].isdigit()):
        return 7
    else:
        return 0  # Default case

def database_connection_result(success, message):
    if success:
        print("Database connected successfully")
        # Update UI or proceed with database operations
    else:
        print(f"Failed to connect to database: {message}")
        # This should ideally be run in a thread-safe manner if updating GUI:
        root.after(0, lambda: messagebox.showerror("Database Connection Error", message))

def load_database_async(callback):
    def run():
        global conn, cursor
        database_path = r'C:/Space/Database/MC_Productlibrary.accdb'

        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={database_path};'
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            # Call the callback function on success
            callback(True, "Database connected successfully.")
        except Exception as e:
            # Call the callback function on failure
            callback(False, str(e))

    # Start the thread
    thread = threading.Thread(target=run)
    thread.start()

def load_database():
    global conn, cursor
    database_path = r'C:/Space/Database/MC_Productlibrary.accdb'

    # Check if the database file exists
    if not os.path.exists(database_path):
        messagebox.showerror("Database Error", "Database file not found."
                                               "\nPlease make sure it is located in C:\Space\Database")
        return

    # Check if the database file is older than two weeks
    last_modified = datetime.datetime.fromtimestamp(os.path.getmtime(database_path))
    if datetime.datetime.now() - last_modified > datetime.timedelta(weeks=2):
        messagebox.showwarning("Database Warning", "Your database file is older than two weeks. \n"
                                                   "Consider updating it.")

    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={database_path};'
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    except Exception as e:
        messagebox.showerror("Database Connection Error", str(e))


def fetch_records(cursor, single_checked, multipack_checked):
    try:
        query = "SELECT TOP 200 ID, Name, Desc20, Desc10 FROM MC_Products WHERE Status <> 'Live'"

        conditions = []
        if single_checked:
            conditions.append("Desc10 = 'Single'")
        if multipack_checked:
            conditions.append("Desc10 = 'Multi'")

        if conditions:
            query += " AND (" + " OR ".join(conditions) + ")"

        query += " ORDER BY LEN(ID) DESC, ID DESC"

        cursor.execute(query)
        records = cursor.fetchall()
        display_records(records, clear_existing=True)
        update_header_count()
    except Exception as e:
        raise e
    update_header_count()
    filter_and_highlight_duplicates()
    clear_and_highlight_duplicates()

name_window_open = False

def name_filter_window(filter_type):
    def on_close_window():
        global name_window_open
        name_window_open = False
        name_filter_root.destroy()
    global name_window_open  # Use the global variable within your function

    # Check if the window is already open
    if name_window_open:
        return  # If the window is open, do not proceed to open another window

    def execute_name_filter():
        global name_window_open, on_close_window

        if not cursor:
            messagebox.showerror("Error", "Database not connected")
            return
        input_text = name_filter_entry.get("1.0", "end-1c")
        if input_text:
            items = [item.strip() for item in input_text.split('\n') if item.strip()]  # Split by newlines and strip
            # Modify the query to use LIKE for partial matches, wrapping each item in wildcards
            query = "SELECT ID, Name, Desc20, Desc10 FROM MC_Products WHERE " + " OR ".join([f"Name LIKE ?" for _ in items])
            # Wrap each item in % for the LIKE wildcard match
            like_items = [f"%{item}%" for item in items]  # Adjusted to wrap items in % for both sides
            try:
                cursor.execute(query, like_items)  # Pass the list of like_items as parameters
                display_records(cursor.fetchall())
                # Call these functions after items are added to ensure accurate processing
                clear_and_highlight_duplicates()
                update_header_count()
                reapply_dropped_status()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to execute query: {e}")
        name_filter_root.destroy()
        name_window_open = False  # Update the flag when the window is closed

    # Before creating the window, set the flag to True
    name_window_open = True

    name_filter_root = tk.Toplevel()
    name_filter_root.iconphoto(False, tk.PhotoImage(file='pih_icon.png'))
    name_filter_root.title(f"Batch Filter by {filter_type}")
    name_filter_entry = tk.Text(name_filter_root, height=15, width=40)
    name_filter_entry.pack()
    customtkinter.CTkButton(name_filter_root, text="Filter", command=execute_name_filter).pack(pady=(5, 5))

    # Bind the window close ('X') button to update the flag
    name_filter_root.protocol("WM_DELETE_WINDOW", on_close_window)




window_open = False

def batch_filter_window(filter_type):
    def on_close_window():
          global window_open
          window_open = False
          batch_filter_root.destroy()
    global window_open
    # Check if the window is already open
    if window_open:
        return  # If the window is open, do not proceed to open another window

    def execute_batch_filter():
        global window_open, on_close_window

        if not cursor:
            messagebox.showerror("Error", "Database not connected.")
            return
        input_text = batch_filter_entry.get("1.0", "end-1c")
        if input_text:
            items = [item.strip() for item in input_text.split('\n') if item.strip()]  # Split by newlines
            if filter_type == "ID":
                query = f"SELECT ID, Name, Desc20, Desc10 FROM MC_Products WHERE ID IN ({','.join('?' * len(items))})"
            else:  # MC RPL UPC
                query = f"SELECT ID, Name, Desc20, Desc10 FROM MC_Products WHERE Desc20 IN ({','.join('?' * len(items))})"
            try:
                cursor.execute(query, items)  # Fixed *items to items
                display_records(cursor.fetchall())
                # Call these functions after items are added to ensure accurate processing
                clear_and_highlight_duplicates()
                update_header_count()
                reapply_dropped_status()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to execute query: {e}")
        batch_filter_root.destroy()
        window_open = False  # Update the flag when the window is closed

    # Before creating the window, set the flag to True
    window_open = True

    batch_filter_root = tk.Toplevel()
    batch_filter_root.iconphoto(False, tk.PhotoImage(file='pih_icon.png'))
    batch_filter_root.title(f"Batch Filter by {filter_type}")
    batch_filter_entry = tk.Text(batch_filter_root, height=15, width=40)
    batch_filter_entry.pack()
    customtkinter.CTkButton(batch_filter_root, text="Filter", command=execute_batch_filter).pack(pady=(5, 5))

    # Bind the window close ('X') button to update the flag
    batch_filter_root.protocol("WM_DELETE_WINDOW", on_close_window)




def treeview_sort_column(tree, col, reverse):
    l = [(tree.set(k, col), k) for k in tree.get_children('')]
    l.sort(reverse=reverse)

    # Rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tree.move(k, '', index)

    # Reverse sort next time
    tree.heading(col, command=lambda: treeview_sort_column(tree, col, not reverse))


def on_tree_select(event):
    selected_item = tree.focus()
    if selected_item:
        print(f"Item selected: {selected_item}, Values: {tree.item(selected_item, 'values')}")



def display_records(records, clear_existing=False):
    if clear_existing:
        for i in tree.get_children():
            tree.delete(i)
    existing_ids = {tree.item(child_id, 'values')[1] for child_id in
                    tree.get_children()}  # Assuming ID is in the second column (index 1)
    for record in records:
        if record[0] not in existing_ids:  # Check if ID is not already in the tree
            prefix_length = upc_length(record[2])
            file_name = record[2][prefix_length:] if prefix_length else record[2]
            # Inserting at index '0' to place new records at the top of the Treeview
            tree.insert('', '0', values=('', record[0], record[1], record[2], record[3], file_name),
                        tags=('',))


def show_drop_status(success, file_name):
    global status_label
    message = f"Success: Renamed {file_name}" if success else f"Failed: Could not rename {file_name}"
    message_color = "green" if success else "red"
    status_label.config(text=message, fg=message_color)
    status_label.grid(row=13, column=0)  # Adjust as per your layout
    root.after(3000, status_label.grid_remove)  # Message disappears after 3 seconds


def filter_and_highlight_duplicates():
    file_name_to_item_ids = {}

    # Iterate through all items in the tree after filtering
    for item_id in tree.get_children():
        item = tree.item(item_id, 'values')
        file_name = item[5]  # Assuming the file name is in the 6th column

        # Track item IDs by file name
        if file_name not in file_name_to_item_ids:
            file_name_to_item_ids[file_name] = [item_id]
        else:
            file_name_to_item_ids[file_name].append(item_id)

    # Configure the tag for duplicates to change the text color to red
    tree.tag_configure('duplicate', background='red')

    # Apply the 'duplicate' tag to items with duplicate file names
    for item_ids in file_name_to_item_ids.values():
        if len(item_ids) > 1:  # If more than one item has this file name
            for item_id in item_ids:
                tree.item(item_id, tags=('duplicate',))

def main():
    global root, tree, package_type_var, single_filter_var, single_search_type, conn, cursor, \
        mc_rpl_upc_check_var, \
        id_check_var, search_type_var, single_check_var, multipack_check_var, context_menu

    root = TkinterDnD.Tk()
    root.attributes("-alpha", 1)
    root.title("Product Image Toolkit")
    root.geometry('920x700')  # Width x Height
    root.iconphoto(False, tk.PhotoImage(file='pih_icon.png'))
 #   root.minsize(width=832, height=800)
 #   root.maxsize(width=832, height=800)
    search_type_var = tk.StringVar(value="MC RPL UPC")
    mc_rpl_upc_check_var = tk.BooleanVar(value=False)
    id_check_var = tk.BooleanVar(value=False)
    style = ttk.Style()
    available_themes = style.theme_names()
    style.theme_use('vista')
    print("Available themes:", available_themes)
    root.configure(bg='#2c2e42')
    style = ttk.Style(root)
    style.configure("Treeview", rowheight=25)  # Adjust the row height if necessary
    style.configure("Centered.Treeview", justify='center')  # Custom style for centered text

    # background_color_one_a = 'ivory2'
    # background_color_one_b = 'azure4'
    # root['bg'] = background_color_one_a
    tree = ttk.Treeview(root, columns=("Status", "ID", "Name", "MC_RPL_UPC", "Package Type", "File Name"),
                        show='headings', style="Centered.Treeview")

    # Database Load Button
    load_database_async(database_connection_result)
    update_header_count()
    root.grid_rowconfigure(11, weight=1)

    current_version = "2.0"

    def display_update_details(details):
        # Corrected references to Toplevel and Scrollbar
        detail_window = tk.Toplevel(root)  # Make sure to use tk.Toplevel
        detail_window.title("Update Details")
        detail_window.geometry("400x300")
        text = tk.Text(detail_window, wrap=tk.WORD)
        text.pack(expand=True, fill=tk.BOTH)
        scroll = tk.Scrollbar(detail_window, command=text.yview)  # Use tk.Scrollbar
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.config(yscrollcommand=scroll.set)
        text.insert(tk.END, details)
        text.config(state=tk.DISABLED)

    def show_custom_update_dialog(latest_version, download_url, update_details):
        dialog = tk.Toplevel(root)
        dialog.iconphoto(False, tk.PhotoImage(file='pih_icon.png'))
        dialog.title("Update Available")
        dialog.geometry("400x150")  # Adjust size as needed
        dialog.attributes("-topmost", True)

        full_message = f"Version {latest_version} is available. Would you like to download?\n\nUpdate Details:\n{update_details}"
        message_label = tk.Label(dialog, text=full_message, wraplength=380, justify="left")
        message_label.pack(pady=10, padx=10, fill=tk.X, expand=True)

        # Frame for buttons
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=(5, 10), padx=10, fill=tk.X)

        # Download button
        download_button = tk.Button(button_frame, text="Download",
                                    command=lambda: [webbrowser.open(download_url), dialog.destroy()])
        download_button.pack(side='left', expand=True)

        # Later button
        later_button = tk.Button(button_frame, text="Later", command=dialog.destroy)
        later_button.pack(side='right', expand=True)

    def check_for_update():
        update_url = "https://raw.githubusercontent.com/Andreicucu93/projects/main/Product%20Image%20-%20Handler/PIH_version.json"

        try:
            response = requests.get(update_url)
            data = response.json()
            latest_version = data['latest_version']
            download_url = data['download_url']
            update_details = data.get('update_details', 'No update details provided.')

            if parse(latest_version) > parse(current_version):
                show_custom_update_dialog(latest_version, download_url, update_details)
        except Exception as e:
            messagebox.showerror("Error", f"Error checking for update: {e}")

    check_for_update()

    tree.focus_set()
    tree.bind("<<TreeviewSelect>>", on_tree_select)
    root.bind('<Delete>', lambda e: on_delete_key(e))

    single_check_var = tk.BooleanVar()
    multipack_check_var = tk.BooleanVar()

    fetch_frame = customtkinter.CTkFrame(root, fg_color="#2c2e42")
    fetch_frame.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

    light_mode_image = Image.open(resource_path("discover_button.png"))

    top_drop = Image.open(resource_path("dropfortop.png"))
    top_drop_tk = ImageTk.PhotoImage(top_drop, size=(100, 30))
    side_drop = Image.open(resource_path("dropforside.png"))
    side_drop_tk = ImageTk.PhotoImage(side_drop, size=(100, 30))
    front_drop = Image.open(resource_path("dropforfront.png"))
    front_drop_tk = ImageTk.PhotoImage(front_drop, size=(100, 30))

    discover_items = customtkinter.CTkImage(light_image=light_mode_image, size=(259, 47))

    fetch_button = customtkinter.CTkButton(fetch_frame, image=discover_items, text="", fg_color="#2c2e42", width=100,

                                           command=lambda: fetch_records(cursor, single_check_var.get(),
                                                                         multipack_check_var.get()))
    fetch_button.grid(row=1, column=2, pady=(5, 0))


    single_filter_var = tk.StringVar()

    batch_filter_frame = customtkinter.CTkFrame(root, height=47, fg_color="#2c2e42")
    batch_filter_frame.grid(row=2, column=0, padx=10, sticky='ew')


    idfilter_button = Image.open(resource_path("idfilter_button.png"))
    idfilterbutton = customtkinter.CTkImage(light_image=idfilter_button, size=(100, 30))
    mcfilter_button = Image.open(resource_path("mcfilter_button.png"))
    mcfilterbutton = customtkinter.CTkImage(light_image=mcfilter_button, size=(115, 30))
    batch_filter_by_id_button = customtkinter.CTkButton(batch_filter_frame, image=idfilterbutton, text="", fg_color="#2c2e42",
                                                 width=100,
                                                 hover_color="#2c2e42", command=lambda: batch_filter_window("ID"))

    batch_filter_by_id_button.grid(row=1, column=1, pady=7, padx=(0, 5))


    batch_filter_by_mc_button = customtkinter.CTkButton(batch_filter_frame, image=mcfilterbutton, text="", fg_color="#2c2e42",
                                                 width=100,
                                                 hover_color="#2c2e42", command=lambda: batch_filter_window("MC RPL UPC"))
    batch_filter_by_mc_button.grid(row=1, column=2, pady=5)

    clear_list_button = Image.open(resource_path("clear_button.png"))
    clearlistbutton = customtkinter.CTkImage(light_image=clear_list_button, size=(100, 30))

    clear_button = customtkinter.CTkButton(batch_filter_frame, image=clearlistbutton, text="",
                    fg_color="#2c2e42", width=100, hover_color="#2c2e42", command=clear_records)
    clear_button.grid(row=1, column=6, padx=10, pady=5)

    namefilter_button = Image.open(resource_path("namefilter_button.png"))
    namefilterbutton = customtkinter.CTkImage(light_image=namefilter_button, size=(110, 30))

    name_filter_button = customtkinter.CTkButton(batch_filter_frame, command=lambda: name_filter_window("Name"),
                                                 image=namefilterbutton, text="", hover_color="#2c2e42", fg_color="#2c2e42")
    name_filter_button.grid(row=1, column=5, pady=5)

    # Treeview
    tree_frame = customtkinter.CTkFrame(root, fg_color='#2c2e42')
    tree_frame.grid(row=7, column=0, padx=10, pady=10, sticky="nsew")

    # Make sure tree_frame can expand to fill space
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)
    tree = ttk.Treeview(tree_frame, columns=("Status", "ID", "Name", "MC_RPL_UPC", "Package Type", "File Name"),
                        show='headings', style="Centered.Treeview")

    # Configure tags for different statuses
    tree.tag_configure('dropped', background='#90EE90')
    tree.grid(row=7, column=0, padx=10, pady=10)
    tree.heading("Status", text="Status", command=lambda: treeview_sort_column(tree, "Status", False))
    tree.heading("MC_RPL_UPC", text="MC_RPL_UPC", command=lambda: treeview_sort_column(tree, "MC_RPL_UPC", False))
    tree.heading("Package Type", text="Package Type", command=lambda: treeview_sort_column(tree, "Package Type", False))
    tree.heading("File Name", text="File Name", command=lambda: treeview_sort_column(tree, "File Name", False))
    tree.heading("ID", text="ID", command=lambda: treeview_sort_column(tree, "ID", False))
    tree.heading("Name", text="Name", command=lambda: treeview_sort_column(tree, "Name", True))
    tree.column("Status", width=70)
    tree.column("ID", width=40)
    tree.column("Name", width=400)
    tree.column("MC_RPL_UPC", width=100)
    tree.column("Package Type", width=85)
    tree.column("File Name", width=70)

    scrollbar = customtkinter.CTkScrollbar(tree_frame, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.grid(row=0, column=0, sticky="nsew")
    # Position the Scrollbar next to the Treeview
    scrollbar.grid(row=0, column=1, sticky="ns")
    tree_frame.grid_columnconfigure(0, weight=1)
    tree_frame.grid_rowconfigure(0, weight=1)

    # Bottom frame primary
    bottom_frame = customtkinter.CTkFrame(root, fg_color="#2c2e42")
    bottom_frame.grid(row=8, column=0)

    import_image = Image.open(resource_path("import_button.png"))
    importimage = customtkinter.CTkImage(light_image=import_image, size=(85, 30))
    export_image = Image.open(resource_path("export_button.png"))
    exportimage = customtkinter.CTkImage(light_image=export_image, size=(85, 30))
    look_online_image = Image.open(resource_path("look_online_button.png"))
    lookonlineimage = customtkinter.CTkImage(light_image=look_online_image, size=(110, 30))
    pack_files_image = Image.open(resource_path("pack_files_button.png"))
    packfilesimage = customtkinter.CTkImage(light_image=pack_files_image, size=(100, 30))
    single_multi_image = Image.open(resource_path("single_multi_button.png"))
    singlemultiimage = customtkinter.CTkImage(light_image=single_multi_image, size=(100, 30))

    look_online_button = customtkinter.CTkButton(bottom_frame, image=lookonlineimage, text="", fg_color="#2c2e42", command=look_online,
                                                 hover_color="#2c2e42", width=50)
    look_online_button.grid(row=0, column=3, padx=(5, 0))


    style = ttk.Style()
    style.configure("Custom.TButton", foreground="gray")
    export_button = customtkinter.CTkButton(bottom_frame, image=exportimage, text="", fg_color="#2c2e42",
                                            command=export_to_excel, width=50, hover_color="#2c2e42")
    export_button.grid(row=0, column=2)

    load_button = customtkinter.CTkButton(bottom_frame, image=importimage, text="", fg_color="#2c2e42",
                                          command=load_from_excel, width=50, hover_color="#2c2e42")
    load_button.grid(row=0, column=1)

    pack_files_button = customtkinter.CTkButton(bottom_frame, image=packfilesimage, text="", fg_color="#2c2e42", command=pack_files,
                                                 hover_color="#2c2e42", width=50)
    pack_files_button.grid(row=0, column=4)

    single_to_multi_button = customtkinter.CTkButton(bottom_frame, image=singlemultiimage, text="", fg_color="#2c2e42", command=make_multipacks,
                                                 hover_color="#2c2e42", width=50)
    single_to_multi_button.grid(row=0, column=5)

    # Rename frame
    #rename_frame = customtkinter.CTkFrame(root, width=500, height=100, fg_color="#2c2e42")
    #rename_frame.grid(row=10, column=0, pady=(10, 0))
    #renaming_label = customtkinter.CTkLabel(rename_frame,
    #                                        text="Rename your images by dragging and dropping them into the designated areas below",
    #                                        text_color="white")
    #renaming_label.grid(row=0, columnspan=3, pady=(0, 10))
    # Second bottom frame
    second_bottom_frame = tk.Frame(root, bg='#2c2e42')
    second_bottom_frame.grid(row=11, column=0, sticky='nsew', padx=10, pady=(20, 0))
    root.columnconfigure(0, weight=1)  # Allow the column to expand
    root.rowconfigure(11, weight=1)  # Allow the row to expand
    root.columnconfigure(0, weight=1)  # Allow the column to expand
    root.rowconfigure(11, weight=1)
    # Example setup of labels (make sure this part already exists in your code)
    front_label = tk.Label(second_bottom_frame, image=front_drop_tk, bg="#2c2e42")
    front_label.grid(row=0, column=0, sticky='nsew')

    side_label = tk.Label(second_bottom_frame, image=side_drop_tk, bg="#2c2e42")
    side_label.grid(row=0, column=1, sticky='nsew')

    top_label = tk.Label(second_bottom_frame, image=top_drop_tk, bg="#2c2e42")
    top_label.grid(row=0, column=2, sticky='nsew')

    for label, suffix in [(front_label, '.1'), (side_label, '.2'), (top_label, '.3')]:
        label.drop_target_register(DND_FILES)
        label.dnd_bind('<<Drop>>', lambda e, s=suffix: rename_file_drag_drop(s)(e))

        second_bottom_frame.columnconfigure((0, 1, 2), weight=1)  # Allow columns to expand
    second_bottom_frame.rowconfigure(0, weight=1)
    status_label = customtkinter.CTkLabel(root, text="", text_color="green")
    status_label.grid(row=12, column=0)
    status_label.grid_remove()  # Initially hide the label

    msg_label = customtkinter.CTkLabel(root, text="", text_color="green")
    msg_label.grid(row=13, column=0)

    root.mainloop()

current_concatenated_image = None
current_image_path = None
base_image = None

global original_image_path  # Declare this at the beginning of your script
original_image_path = None

def save_image():
    global current_concatenated_image, original_image_path
    if current_concatenated_image:
        save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")], initialfile="multipack")
        if save_path:
            current_concatenated_image.save(save_path)
            print(f"Image saved to {save_path}")

def save_image_overwrite():
    global current_concatenated_image, original_image_path
    if current_concatenated_image and original_image_path:
        try:
            current_concatenated_image.save(original_image_path)
            print(f"Image successfully overwritten at {original_image_path}")
        except Exception as e:
            print(f"Error saving image: {e}")
    else:
        print("No image or original path available to overwrite.")

def trigger_display_refresh():
    global num_concatenations_entry, window

    # Fetch the current value from the entry widget
    current_num_cans = num_concatenations_entry.get()

    # Directly call the function to reapply the current settings and refresh the display
    apply_concatenation_and_display(current_num_cans)

# Global dictionary to hold references to PhotoImage objects to prevent garbage collection
image_references = {}

def display_updated_image(window, pil_image):
    global image_references

    # Generate a unique key for this particular image display instance
    unique_key = "image_display_" + str(id(pil_image))

    try:
        # Clear existing images in the window to prevent memory leaks from old references
        for widget in window.winfo_children():
            if isinstance(widget, tk.Label) and hasattr(widget, 'image'):
                widget.destroy()

        # Add a visible border around the image. You can adjust the border width and color as needed.
        # Here, I'm adding a 10-pixel wide black border.
        pil_image_with_border = ImageOps.expand(pil_image, border=10, fill='black')

        # Convert the PIL image (with border) to a PhotoImage
        photo_image = ImageTk.PhotoImage(pil_image_with_border)
        image_references[unique_key] = photo_image  # Store the reference to prevent garbage collection

        # Create a new label for the image and pack it into the window
        image_label = tk.Label(window, image=photo_image, bg='#2c2e42')
        image_label.image = photo_image  # Keep a reference within the label too, as an extra precaution
        image_label.pack(pady=20, fill='both', expand=True)
    except Exception as e:
        show_popup_message_window(f"Error displaying image: {e}", 2000)

def auto_correct_orientation(pil_image, correction_angle=0):
    """
    Automatically adjusts the orientation of a PIL image to be horizontal.

    Args:
    - pil_image (PIL.Image): The image to correct.
    - correction_angle (float): The angle to rotate the image by to correct the skew.

    Returns:
    - PIL.Image: The adjusted image.
    """
    # Rotate the image to correct the orientation
    corrected_image = pil_image.rotate(-correction_angle, expand=True, resample=Image.BICUBIC, fillcolor='white')

    return corrected_image



def process_and_display_image(path, window, fixed_size, force_refresh=False):
    global current_concatenated_image, current_image_path, base_image

    try:
        # Clear existing images in the window
        for widget in window.winfo_children():
            if isinstance(widget, tk.Label) and hasattr(widget, 'image'):
                widget.destroy()

        # Force reprocessing of the image if the same file is dropped again
        if force_refresh or base_image is None or path != current_image_path:
            base_image = Image.open(path)
            current_image_path = path

        num_concat = int(num_concatenations_entry.get()) if num_concatenations_entry.get().isdigit() else 1

        # Always start concatenation from the base image
        new_concatenated_image = Image.new('RGBA', (base_image.width * num_concat, base_image.height), (255, 0, 0, 0))
        for i in range(num_concat):
            new_concatenated_image.paste(base_image, (i * base_image.width, 0))

        # Update the global variable with the new concatenated image
        current_concatenated_image = new_concatenated_image

        # Determine scaling based on the largest side of the concatenated image
        scale_factor = 200 / max(new_concatenated_image.width, new_concatenated_image.height)
        new_width = int(new_concatenated_image.width * scale_factor)
        new_height = int(new_concatenated_image.height * scale_factor)
        resized_image = new_concatenated_image.resize((new_width, new_height), Image.LANCZOS)
        tk_image = ImageTk.PhotoImage(resized_image)

        # Display the updated image
        display_updated_image(window, resized_image)

    except Exception as e:
        show_popup_message_window("The file needs to be processed first.", 2000)
        print(f"Error processing image: {e}")


def apply_final_cropping(image, percent=22):
    """Crop a given percentage from the top and bottom of the image."""
    width, height = image.size
    crop_height = int((percent / 100.0) * height / 2)

    top = crop_height
    bottom = height - crop_height

    cropped_image = image.crop((0, top, width, bottom))

    print(f"Original height: {height}, Cropped height: {bottom - top}, Crop height: {crop_height}")

    # Diagnostic prints
    print(f"Original height: {height}, New height: {height - 2 * crop_height}")
    print(f"Cropping {crop_height}px from top and bottom, resulting in {bottom - top}px height")

    return cropped_image

def apply_additional_cropping(image, percent=1, additional_bottom_percent=1):
    """
    Apply additional cropping from top-bottom and left-right symmetrically.

    Args:
    - image (PIL.Image): The image to be cropped.
    - percent (float): The percentage of the image's height and width to be cropped from each side.

    Returns:
    - PIL.Image: The cropped image.
    """
    width, height = image.size
    # Calculate crop amounts based on the specified percentage
    crop_height = int((1 / 100.0) * height / 2) #modified this from percent
    crop_width = int((6 / 100.0) * width / 2) #modified this from percent

    additional_bottom_crop = int((additional_bottom_percent / 100.0) * height)

    # Calculate new boundaries
    left = crop_width
    upper = crop_height
    right = width - crop_width
    lower = height - crop_height - additional_bottom_crop  # Adjusted additional bottom crop

    # Perform cropping
    cropped_image = image.crop((left, upper, right, lower))
    return cropped_image



def apply_concatenation_and_display(num_cans):
    global base_image, current_concatenated_image, window, fixed_size
    if not base_image:
        print("No base image to work with.")
        return

    # Adjust the number of cans based on the input
    num_concat = int(num_cans) if num_cans.isdigit() else 1

    # Create a new image for concatenation
    new_concatenated_image = Image.new('RGBA', (base_image.width * num_concat, base_image.height), (255, 0, 0, 0))
    for i in range(num_concat):
        new_concatenated_image.paste(base_image, (i * base_image.width, 0))

    # Apply final cropping to the concatenated image
    final_cropped_image = apply_final_cropping(new_concatenated_image, 22)  # Using a large value for visibility

    # Update the global variable with the final image
    current_concatenated_image = final_cropped_image

    # Display the final, cropped image
    display_updated_image(window, final_cropped_image)

from rembg import remove
def remove_background():
    global base_image, current_image_path

    if base_image:
        # Convert base_image to byte array
        img_byte_arr = io.BytesIO()
        base_image.save(img_byte_arr, format='PNG')
        img_byte_arr = img_byte_arr.getvalue()

        # Remove background
        result_img_data = remove(img_byte_arr)
        processed_image = Image.open(io.BytesIO(result_img_data))

        # Update base_image with the processed image
        base_image = processed_image
    update_preview()


def remove_extra_space():
    global base_image, current_concatenated_image, window, num_concatenations_entry
    if base_image:
        for _ in range(1):
            # Crop extra space from all sides (if any additional cropping logic is needed, include it here)
            cropped_image = crop_excess_space(base_image)

            # Apply additional cropping from top-bottom and left-right (if needed)
            final_cropped_image = apply_additional_cropping(cropped_image, percent=1)  # Adjust percentage as needed

            # Update base_image with the final cropped image
            base_image = final_cropped_image

            # Update the global variable for display and saving
            current_concatenated_image = base_image

            # Ensure the display is updated with the correctly oriented and processed image
            update_display_with_consistent_scaling(window, current_concatenated_image)

            # Reset and reapply original number of cans for correct display
            original_num_cans = num_concatenations_entry.get()
            apply_concatenation_and_display(original_num_cans)
    update_preview()

def update_display_with_consistent_scaling(window, image):
    window.update_idletasks()  # Ensure window size is updated
    display_width = window.winfo_width()
    display_height = window.winfo_height() - 20  # Adjust based on your layout needs

    # Calculate scaling factors and resize the image
    scale_factor_width = display_width / image.width
    scale_factor_height = display_height / image.height
    scale_factor = min(scale_factor_width, scale_factor_height)

    new_width = int(image.width * scale_factor)
    new_height = int(image.height * scale_factor)

    resized_image = image.resize((new_width, new_height), Image.LANCZOS)

    # Update the displayed image
    display_updated_image(window, resized_image)


def crop_excess_space(image):
    """Crop excess space from the PIL image."""
    img_array = np.array(image)

    # Find the bounding box of non-transparent pixels
    non_transparent = np.where(img_array[..., 3] != 0)
    if non_transparent[0].size and non_transparent[1].size:
        ymin, ymax = np.min(non_transparent[0]), np.max(non_transparent[0])
        xmin, xmax = np.min(non_transparent[1]), np.max(non_transparent[1])
        cropped_image = image.crop((xmin, ymin, xmax + 1, ymax + 1))
        return cropped_image
    return image


window = None

def convert_and_process_image(path, window, fixed_size):
    try:
        image = Image.open(path)
        # Convert image to PNG in memory if it's not already a PNG
        if image.format != 'PNG':
            # Convert the image to RGBA to maintain transparency
            image = image.convert('RGBA')
        # Use a temporary path if you need to save the converted image temporarily
        temp_path = os.path.join(tempfile.gettempdir(), os.path.basename(path).rsplit('.', 1)[0] + '.png')
        image.save(temp_path, format='PNG')
        process_and_display_image(temp_path, window, fixed_size)
        # Optionally delete the temporary file if you want to clean up immediately
        # os.remove(temp_path)
    except Exception as e:
        print(f"Error processing image: {e}")


def on_drop(event):
    global base_image, current_concatenated_image, current_image_path, original_image_path

    dropped_files = event.data.split('\n')

    for file_path in dropped_files:
        # Normalize the file path
        cleaned_path = file_path.strip().strip('\"').strip('{}').rstrip("\r")
        print(f"Processed path: {cleaned_path}")

        if os.path.isfile(cleaned_path):
            original_image_path = cleaned_path
            # Reset the current image variables to force a refresh
            base_image = None
            current_concatenated_image = None
            current_image_path = None

            file_format = os.path.splitext(cleaned_path)[-1].lower()
            # Process PNG images directly
            if file_format == '.png':
                process_and_display_image(cleaned_path, window, 200)
            # Convert and process other image formats
            else:
                # Ensure conversion maintains transparency for preview
                convert_and_process_image(cleaned_path, window, 200)
        else:
            print("Unsupported file format or file does not exist.")




def make_multipacks():
    global window, drop_target_label, num_concatenations_entry, num_concatenations_var

    def ensure_model_file_exists():
        # Define the path to the model file within your application's resources
        # This path needs to be adjusted based on where you've included the model file in your PyInstaller package
        source_model_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'u2net.onnx')

        # Define the target path in the user's home directory
        target_directory = os.path.join(os.path.expanduser('~'), '.u2net')
        target_model_path = os.path.join(target_directory, 'u2net.onnx')

        # Create the target directory if it doesn't exist
        if not os.path.exists(target_directory):
            os.makedirs(target_directory)

        # Copy the model file if it doesn't already exist in the target location
        if not os.path.isfile(target_model_path):
            shutil.copy(source_model_path, target_model_path)
            print(f"Model file copied to: {target_model_path}")
        else:
            print("Model file already exists in the target location.")

    ensure_model_file_exists()

    if window is not None:
        return
    num_concatenations_var = tk.StringVar(value="1")
    window = tk.Toplevel(root)
    window.iconphoto(False, tk.PhotoImage(file='pih_icon.png'))
    window.geometry('700x350')  # Increased height to accommodate new UI elements
    window.minsize(height=700, width=350)
    window.maxsize(height=700, width=350)
    window.configure(bg='#2c2e42')
    window.title("Product Image editing")

    drop_target_label = customtkinter.CTkLabel(window, text="", font=('Helvetica', 13, 'bold'), text_color='white')
    drop_target_label.pack(pady=20)

    single_replicate_image = Image.open(resource_path("can_replicate.png"))
    singlereplicateimage = ImageTk.PhotoImage(single_replicate_image)

    replicate_instructions = customtkinter.CTkButton(window, image=singlereplicateimage, text="", hover_color='#2c2e42', fg_color='#2c2e42')
    replicate_instructions.pack(pady=(5, 5))

    # Entry for number of concatenations
    num_concat_label = customtkinter.CTkLabel(window, text="Number of cans:", text_color='white')
    num_concat_label.pack()
    num_concatenations_entry = customtkinter.CTkEntry(window, textvariable=num_concatenations_var)
    num_concatenations_entry.pack()
    num_concatenations_var.trace_add("write", lambda name, index, mode, sv=num_concatenations_var: update_preview())

    save_button = customtkinter.CTkButton(window, text="Save Image as", command=save_image)
    save_button.pack(pady=10)

    overwrite_save_button = customtkinter.CTkButton(window, text="Save", command=save_image_overwrite)
    overwrite_save_button.pack(pady=10)

    image_editing_frame = customtkinter.CTkFrame(window, width=320, height=25, fg_color='#2c2e42')
    image_editing_frame.pack(pady=10)

    remove_bg_button = customtkinter.CTkButton(image_editing_frame, text="Remove background", command=remove_background, fg_color='gray17')
    remove_bg_button.grid(row=0, column=0, padx=5)

    remove_space_button = customtkinter.CTkButton(image_editing_frame, text="Remove space", command=remove_extra_space, fg_color='gray17')
    remove_space_button.grid(row=0, column=1, padx=5)

    drag_and_drop_image = Image.open(resource_path("draganddrop.png"))
    dnd_image = ImageTk.PhotoImage(drag_and_drop_image)

    drop_instruction = customtkinter.CTkButton(window, image=dnd_image, text="", hover_color='#2c2e42', fg_color='#2c2e42')
    drop_instruction.pack(pady=5)

    # Make the window a drop target for files
    window.drop_target_register(DND_FILES)
    window.dnd_bind('<<Drop>>', on_drop)

    position_window_next_to_root()

    window.protocol('WM_DELETE_WINDOW', on_window_close)

def update_preview():
    global current_image_path, window
    if current_image_path:
        process_and_display_image(current_image_path, window, 200)


def on_window_close():
    global window
    if window is not None:
        window.destroy()
    window = None

def position_window_next_to_root():
    global window, root
    if window is not None and root is not None:
        x = root.winfo_x()
        y = root.winfo_y()
        width = root.winfo_width()
        window.geometry(f"+{x + width + 10}+{y}")


def show_popup_message_window(message, duration=2000):
    popup = tk.Toplevel(window)
    popup.transient(window)  # Make the popup a transient window of the main app
    popup.grab_set()  # Grab the focus
    popup.configure(bg='#c1b3b3')
    popup.overrideredirect(True)  # Remove window decorations

    # Window size and position
    width = 250
    height = 50
    x = window.winfo_x() + (window.winfo_width() // 2) - (width // 2)
    y = window.winfo_y() + (window.winfo_height() // 2) - (height // 2)
    popup.geometry(f"{width}x{height}+{x}+{y}")

    # Add a label with the message
    message_label = customtkinter.CTkLabel(popup, text=message, bg_color='#525c95')
    message_label.pack(expand=True, fill='both', pady=10)

    popup.lift()  # Bring the popup above other windows
    popup.focus_force()  # Force focus on the popup

    # Automatically destroy the popup after 'duration' milliseconds
    popup.after(duration, popup.destroy)


if __name__ == "__main__":
    conn = None
    cursor = None
    main()
