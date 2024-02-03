import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import pyperclip
import webbrowser
import os
import tkinter.filedialog as fd
import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD
import shutil
import customtkinter
import datetime
from PIL import Image, ImageTk
import sys
import win32com.client as win32
import threading


#To be compileable..
if hasattr(sys, '_MEIPASS'):
    os.environ['TKDND_LIBRARY'] = os.path.join(sys._MEIPASS, 'tkdnd')


#Email submission
def send_email_thread(to, subject, body, attachments=None):
    trigger_outlook_email(to=to, subject=subject, body=body, attachments=attachments)


def trigger_outlook_email(to, subject, body, attachments=None):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # 0 is for a mail item
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    # Attach files, if any
    if attachments:
        for attachment in attachments:
            mail.Attachments.Add(attachment)
    mail.Display(True)  # True to display the item to the user


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def pack_files_in_folders():
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
            tree.set(item_id, column='Image status',
                     value=f"Packed {emoji}")  # Mark as Packed with corresponding number of emojis
            print(f"Moved {num_files} files to {destination_folder}")
    update_header_count()


def show_temporary_message(message, text_color="black"):  # Default color is black
    msg_label = customtkinter.CTkLabel(root, text=message, text_color=text_color)
    msg_label.grid(row=13, column=0)  # Adjust the grid position as needed
    root.after(3000, msg_label.destroy)  # Message will disappear after 3000 milliseconds (3 seconds)


def rename_file_drag_drop(suffix):
    def inner(event):
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
        if os.path.exists(new_file_path):
            show_popup_message(f"A file named '{new_file_name}' already exists.", 1000)
            return False
        try:
            os.rename(file_path, new_file_path)
            show_popup_message(f"File renamed to {new_file_name}", 1000)
            return True
        except Exception as e:
            show_popup_message("Error during renaming.", 1000)
            return False

    return inner


def load_from_excel():
    file_path = fd.askopenfilename(filetypes=[("Excel Files", "*.xlsx")], title="Open File")
    if file_path:
        try:
            df = pd.read_excel(file_path)
            clear_records()  # Clear existing records in the treeview

            # Find the column name that starts with 'Image status'
            image_status_col = next(col for col in df.columns if col.startswith('Image status'))

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
        image_status = tree.set(child, 'Image status')
        count_loaded += image_status.count("📦")
    tree.heading("Image status", text=f"Image status ({count_loaded})")


def on_delete_key(event):
    selected_items = tree.selection()
    if selected_items:
        for item in selected_items:
            tree.delete(item)
        update_header_count()  # Update the count after deletion


def on_tree_cell_click(event):
    region = tree.identify("region", event.x, event.y)
    column = tree.identify_column(event.x)
    row_id = tree.identify_row(event.y)

    # Check if the clicked region is a cell and the column is the "Image" column
    if region == "cell" and column == "#1" and row_id:
        current_value = tree.set(row_id, column='Image status')
    update_header_count()


def on_treeview_click(event):
    region = tree.identify("region", event.x, event.y)
    column = tree.identify_column(event.x)
    if region == "cell" and column == "#1":  # Check if the click is on the "Image" column
        row_id = tree.identify_row(event.y)
        if row_id:  # Check if a row is clicked
            current_value = tree.set(row_id, column='Image status')
            update_header_count()  # Update the count


def export_to_excel():
    global current_date
    # Count the number of loaded items
    count_loaded = sum("📦" in tree.set(child, 'Image status') for child in tree.get_children())

    data = []
    for item_id in tree.get_children():
        item = tree.item(item_id, 'values')
        data.append(item)
    if data:
        # Create the DataFrame with the dynamic column name
        df = pd.DataFrame(data, columns=["Image status (" + str(count_loaded) + " loaded)", "ID",
                                         "Name", "MC_RPL_UPC", "Package Type", "File Name"])
        # Generate filename with current date
        current_date = datetime.datetime.now().strftime("%Y-%m-%d")  # Format the date
        default_filename = f"Imaging {current_date}.xlsx"  # Default filename "Imaging YYYY-MM-DD.xlsx"

        file_path = fd.asksaveasfilename(defaultextension='.xlsx',
                                         initialfile=default_filename,  # Use the generated filename as the default
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
    file_path = fd.askopenfilename(title="Select an Image", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])
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
    message_label = tk.Label(popup, text=message, bg='#525c95')
    message_label.pack(expand=True, fill='both', pady=10)

    popup.lift()  # Bring the popup above other windows
    popup.focus_force()  # Force focus on the popup

    # Automatically destroy the popup after 'duration' milliseconds
    popup.after(duration, popup.destroy)


def select_and_rename_file(record, suffix):
    file_name_column = 5  # Adjust the column index based on your Treeview setup
    original_file_name = record[file_name_column]  # Get the file name from the record
    # Open file dialog to select an image
    file_path = fd.askopenfilename(title="Select an Image", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])
    if not file_path:
        return  # No file was selected
    # Extract the file extension
    _, file_extension = os.path.splitext(file_path)
    # Create the new file name
    new_file_name = f"{original_file_name}{suffix}{file_extension}"
    # Rename and save the file
    new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
    os.rename(file_path, new_file_path)
    messagebox.showinfo("Success", f"File renamed to {new_file_name}")


def look_online():
    copy_file_name_to_clipboard()
    selected_item = tree.focus()  # Get selected item
    item_values = tree.item(selected_item, 'values')
    if item_values:
        name = item_values[2]  # Assuming the Name is in the 2nd column
        search_query = f"https://www.google.com/search?tbm=isch&q={name}"
        webbrowser.open(search_query)


def copy_name_to_clipboard():
    selected_item = tree.focus()  # Get selected item
    item_values = tree.item(selected_item, 'values')
    if item_values:
        name = item_values[2]  # Assuming the Name is in the 2nd column
        pyperclip.copy(name)


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
    except Exception as e:
        raise e
    update_header_count()


def single_search(error_label_entry, error_label_not_found):
    search_input = single_filter_var.get()
    if not search_input.strip():
        error_label_entry.grid()
        error_label_not_found.grid_remove()
        return
    else:
        error_label_entry.grid_remove()

    selected_search_type = search_type_var.get()
    if selected_search_type == "MC RPL UPC":
        cursor.execute("SELECT ID, Name, Desc20, Desc10 FROM MC_Products WHERE Desc20 = ? AND Status <> 'Approved'", search_input)
    elif selected_search_type == "ID":
        cursor.execute("SELECT ID, Name, Desc20, Desc10 FROM MC_Products WHERE ID = ? AND Status <> 'Approved'", search_input)
    else:
        error_label_entry.config(text="Please select a search type")
        error_label_entry.grid()
        return

    records = cursor.fetchall()
    if not records:
        error_label_not_found.grid()
    else:
        error_label_not_found.grid_remove()
        display_records(records)


def batch_filter_window(filter_type):
    def execute_batch_filter():
        global cursor  # Declare cursor as global
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
                cursor.execute(query, *items)
                display_records(cursor.fetchall())
            except Exception as e:
                messagebox.showerror("Error", f"Failed to execute query: {e}")
        batch_filter_root.destroy()

    batch_filter_root = tk.Toplevel()
    batch_filter_root.title(f"Batch Filter by {filter_type}")
    batch_filter_entry = tk.Text(batch_filter_root, height=15, width=40)
    batch_filter_entry.pack()
    customtkinter.CTkButton(batch_filter_root, text="Filter products", command=execute_batch_filter).pack(pady=(5, 5))


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

is_help_window_open = False
help_window = None

def open_help_window():
    global help_window

    if help_window is not None:
        # Help window already open
        return

    help_window = tk.Toplevel(root)
    help_window.geometry('300x825')
    help_window.minsize(height=825, width=300)
    help_window.maxsize(height=825, width=300)
    help_window.configure(bg='#2c2e42')
    help_window.title("Instructions")

    # Load the image (use resource_path to find the image)
    instruction_image = Image.open(resource_path("Capture.PNG"))
    tk_image = ImageTk.PhotoImage(instruction_image)

    # Create a label to display the image
    image_label = tk.Label(help_window, image=tk_image)
    image_label.image = tk_image  # Keep a reference!
    image_label.pack()

    # Position the window (optional)
    x = root.winfo_x()
    y = root.winfo_y()
    width = root.winfo_width()
    help_window.geometry(f"+{x + width + 10}+{y}")

    # Bind the close event
    help_window.protocol("WM_DELETE_WINDOW", on_help_window_close)


def on_help_window_close():
    global help_window
    if help_window is not None:
        help_window.destroy()  # Close the window
    help_window = None  # Reset the global variable


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
            tree.insert('', '0', values=('       ----------', record[0], record[1], record[2], record[3], file_name),
                        tags=('       ----------',))


def show_drop_status(success, file_name):
    global status_label
    message = f"Success: Renamed {file_name}" if success else f"Failed: Could not rename {file_name}"
    message_color = "green" if success else "red"
    status_label.config(text=message, fg=message_color)
    status_label.grid(row=13, column=0)  # Adjust as per your layout
    root.after(3000, status_label.grid_remove)  # Message disappears after 3 seconds


def main():
    global root, tree, package_type_var, single_filter_var, single_search_type, conn, cursor, \
        mc_rpl_upc_check_var, \
        id_check_var, search_type_var, single_check_var, multipack_check_var, context_menu

    root = TkinterDnD.Tk()
    root.attributes("-alpha", 1)
    root.title("Product Image Handler")
    root.geometry('832x825')  # Width x Height
    root.minsize(width=832, height=825)
    root.maxsize(width=832, height=825)
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
    tree = ttk.Treeview(root, columns=("Image status", "ID", "Name", "MC_RPL_UPC", "Package Type", "File Name"),
                        show='headings', style="Centered.Treeview")

    # Database Load Button
    load_database()
    update_header_count()
    root.grid_rowconfigure(11, weight=1)

    tree.focus_set()
    tree.bind("<<TreeviewSelect>>", on_tree_select)
    root.bind('<Delete>', lambda e: on_delete_key(e))

    single_check_var = tk.BooleanVar()
    multipack_check_var = tk.BooleanVar()

    fetch_frame = customtkinter.CTkFrame(root, fg_color="#2c2e42")
    fetch_frame.grid(row=1, column=0, padx=10, pady=10, sticky='ew')


    light_mode_image = Image.open(resource_path("discover_button.png"))
    guide_image = Image.open(resource_path("instructions.png"))
    guide_image_tk = ImageTk.PhotoImage(guide_image)

    top_drop = Image.open(resource_path("dropfortop.png"))
    top_drop_tk = ImageTk.PhotoImage(top_drop)
    side_drop = Image.open(resource_path("dropforside.png"))
    side_drop_tk = ImageTk.PhotoImage(side_drop)
    front_drop = Image.open(resource_path("dropforfront.png"))
    front_drop_tk = ImageTk.PhotoImage(front_drop)

    # Create a CTkImage object with the loaded image
    discover_items = customtkinter.CTkImage(light_image=light_mode_image, size=(259, 47))

    fetch_button = customtkinter.CTkButton(fetch_frame, image=discover_items, text="", fg_color="#2c2e42", width=100,
                                           hover_color="#2c2e42",
                                           command=lambda: fetch_records(cursor, single_check_var.get(),
                                                                         multipack_check_var.get()))
    fetch_button.grid(row=1, column=2, pady=(5, 0))

    help_button = customtkinter.CTkButton(fetch_frame, image=guide_image_tk, text="", command=open_help_window,
                                          height=1, width=1, fg_color='#2c2e42', hover_color="#2c2e42")
    help_button.grid(row=1, column=3, padx=(480, 0))

    single_filter_var = tk.StringVar()

    batch_filter_frame = customtkinter.CTkFrame(root, height=47, fg_color="#2c2e42")
    batch_filter_frame.grid(row=2, column=0, padx=10, sticky='ew')

    batch_filter_by_id_button = customtkinter.CTkButton(batch_filter_frame, text="Batch Filter by ID",
                                                    fg_color='SlateBlue2', hover_color='SlateBlue3',
                                                        command=lambda: batch_filter_window("ID"))
    batch_filter_by_id_button.grid(row=1, column=1, padx=10, pady=5)

    batch_filter_by_mc_button = customtkinter.CTkButton(batch_filter_frame, text="Batch Filter by MC RPL UPC",
               fg_color='SlateBlue2', hover_color='SlateBlue3', command=lambda: batch_filter_window("MC RPL UPC"))
    batch_filter_by_mc_button.grid(row=1, column=2, padx=10, pady=5)

    clear_button = customtkinter.CTkButton(batch_filter_frame, text="Clear all products", fg_color='red2',
                                           hover_color='red1', command=clear_records, width=17)
    clear_button.grid(row=1, column=5, padx=10, pady=10)


    # Treeview
    tree = ttk.Treeview(root, columns=("Image status", "ID", "Name", "MC_RPL_UPC", "Package Type", "File Name"),
                        show='headings', style="Centered.Treeview")

    # Configure tags for different statuses
    tree.tag_configure('       ----------', foreground='gray')
    tree.grid(row=7, column=0, padx=10, pady=10)
    tree.heading("Image status", text="Image status", command=lambda: treeview_sort_column(tree, "Image status", False))
    tree.column("Image status", width=105)
    tree.heading("ID", text="ID")
    tree.heading("Name", text="Name")
    tree.heading("MC_RPL_UPC", text="MC_RPL_UPC")
    tree.heading("Package Type", text="Package Type", command=lambda: treeview_sort_column(tree, "Package Type", False))
    tree.heading("File Name", text="File Name")
    tree.heading("ID", text="ID", command=lambda: treeview_sort_column(tree, "ID", False))
    tree.heading("Name", text="Name", command=lambda: treeview_sort_column(tree, "Name", False))
    tree.column("ID", width=50)
    tree.column("Name", width=400)
    tree.column("MC_RPL_UPC", width=100)
    tree.column("Package Type", width=85)
    tree.column("File Name", width=70)
    tree.grid(row=7, column=0, padx=10, pady=10)
    # Bottom frame primary
    bottom_frame = customtkinter.CTkFrame(root, width=200, height=30, fg_color="#2c2e42")
    bottom_frame.grid(row=8, column=0, pady=(10, 0))
    copy_file_name_button = customtkinter.CTkButton(bottom_frame, text="Copy File Name", command=copy_file_name_to_clipboard,
                                                    fg_color='#1f6aa5', hover_color='RoyalBlue3')
    copy_file_name_button.grid(row=0, column=0)
    look_online_button = customtkinter.CTkButton(bottom_frame, text="Look Online", command=look_online,
                                                 fg_color='#1f6aa5', hover_color='RoyalBlue3')
    look_online_button.grid(row=0, column=1, padx=(10, 0))

    # Bottom frame secondary
    bottom_frame_secondary = customtkinter.CTkFrame(root, width=200, height=30, fg_color="#2c2e42")
    bottom_frame_secondary.grid(row=9, column=0)
    style = ttk.Style()
    style.configure("Custom.TButton", foreground="gray")
    export_button = customtkinter.CTkButton(bottom_frame_secondary, text="Export in xlsx",
                               command=export_to_excel, fg_color='SpringGreen4', hover_color='#04aa56')
    export_button.grid(row=1, column=0)
    load_button = customtkinter.CTkButton(bottom_frame_secondary, text="Load xlsx", command=load_from_excel
                                          , fg_color='SpringGreen4', hover_color='#04aa56')
    load_button.grid(row=1, column=1, pady=10, padx=(10, 0))  # Adjust grid position as needed
    pack_files_button = customtkinter.CTkButton(bottom_frame_secondary, text="Pack files", command=pack_files_in_folders,
                                                fg_color='DarkGoldenrod3', hover_color='DarkGoldenrod2')
    pack_files_button.grid(row=2, column=0)

    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    send_email_button = customtkinter.CTkButton(bottom_frame_secondary, text="Send Email",
                                                fg_color='DarkGoldenrod3', hover_color='DarkGoldenrod2',
                                                command=lambda: threading.Thread(target=send_email_thread, args=(
                                                    "SpaceLab.Support@molsoncoors.com",
                                                    f"Image submission - {current_date}",
                                                    "Here is the imaging report and related files.",
                                                    None,  # Assuming no attachments for simplicity; add as needed
                                                )).start())
    send_email_button.grid(row=2, column=1, padx=(10, 0))

    # Rename frame
    rename_frame = customtkinter.CTkFrame(root, width=500, height=100, fg_color="#2c2e42")
    rename_frame.grid(row=10, column=0, pady=(10, 0))
    renaming_label = customtkinter.CTkLabel(rename_frame,
                                            text="Rename your images by dragging and dropping them into the designated areas below",
                                            text_color="white")
    renaming_label.grid(row=0, columnspan=3, pady=(0, 10))
    # Second bottom frame
    second_bottom_frame = tk.Frame(root, height=200, bg='#22202d')  # Set background color to match customtkinter style
    second_bottom_frame.grid(row=11, column=0, sticky='nsew', padx=10, pady=10)
    root.columnconfigure(0, weight=1)  # Allow the column to expand
    root.rowconfigure(11, weight=1)  # Allow the row to expand
    root.columnconfigure(0, weight=1)  # Allow the column to expand
    root.rowconfigure(11, weight=1)
    front_label = tk.Label(second_bottom_frame, image=front_drop_tk, bg="RosyBrown2", height=40, width=40)
    front_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew') #Drag and drop for front rename

    side_label = tk.Label(second_bottom_frame, image=side_drop_tk, bg="RosyBrown2", height=40, width=40)
    side_label.grid(row=0, column=1, padx=10, pady=10, sticky='nsew') #Drag and drop for side rename

    top_label = tk.Label(second_bottom_frame, image=top_drop_tk, bg="RosyBrown2", height=40, width=40)
    top_label.grid(row=0, column=2, padx=10, pady=10, sticky='nsew') #Drag and drop for top rename

    second_bottom_frame.columnconfigure((0, 1, 2), weight=1)  # Allow columns to expand
    second_bottom_frame.rowconfigure(0, weight=1)
    status_label = customtkinter.CTkLabel(root, text="", text_color="green")
    status_label.grid(row=12, column=0)
    status_label.grid_remove()  # Initially hide the label
    # Bind drag and drop events
    for CTkLabel, suffix in [(front_label, '.1'), (side_label, '.2'), (top_label, '.3')]:
        CTkLabel.drop_target_register(DND_FILES)
        CTkLabel.dnd_bind('<<Drop>>', lambda e, s=suffix: rename_file_drag_drop(s)(e))
    msg_label = customtkinter.CTkLabel(root, text="", text_color="green")
    msg_label.grid(row=13, column=0)

    root.mainloop()


if __name__ == "__main__":
    conn = None
    cursor = None
    main()