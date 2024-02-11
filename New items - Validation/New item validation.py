import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import pyodbc
import pyperclip
import re
from tkinter import messagebox
import json
import urllib.request
import webbrowser
import customtkinter


#versionControl: 02072014
#Theme: config
customtkinter.set_appearance_mode("dark")

config_url = 'https://raw.githubusercontent.com/questionmarkdude/questionmark/main/config.json'


def fetch_config():
    try:
        response = urllib.request.urlopen(config_url)
        if response.getcode() == 200:
            config = json.loads(response.read())
            return config
        else:
            return None
    except Exception as e:
        print("Exception:", e)
        return None


class ButtonedTreeview(ttk.Treeview):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def insert_button(self, item, column, text, command):
        button = tk.Button(self, text=text, command=command, relief=tk.FLAT, padx=0, pady=0)
        button.state = text  # Store the state ('O' or 'X') in the button itself
        self.place_widget(button, item, column)

    def place_widget(self, widget, item, column):
        widget.update_idletasks()
        self.update_idletasks()  # Add this line
        x, y, width, height = self.bbox(item, column)
        widget.place(x=x, y=y, width=width, height=height, anchor='nw')


class App:

    def __init__(self, master):
        self.master = master
        master.title('New Item Checker 020724')

        self.accdb_file_path = None

        self.feedback_entry = tk.Entry(master, width=8)
        self.feedback_entry.bind("<FocusOut>", self.hide_feedback_entry)
        self.feedback_entry.bind("<Return>", self.on_feedback_edit)

        self.results = {}

        self.treeview_frame = tk.Frame(master)  # Create a frame for the treeview and scrollbar
        self.treeview_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.treeview_frame.columnconfigure(0, weight=1)  # Add this line to make the treeview expand horizontally

        self.treeview = ButtonedTreeview(self.treeview_frame,  # Use ButtonedTreeview instead of ttk.Treeview
                                 columns=('check', 'name', 'UPC', '12DigitUPC', 'NameCheck', 'Feedback'),
                                 show='headings', selectmode='browse')

        self.treeview.column('check', width=10, anchor=tk.CENTER)

        self.treeview.heading('Feedback', text='Feedback')
        self.treeview.column('Feedback', width=200, anchor=tk.CENTER)

        self.treeview.heading('name', text='NAME', command=self.sort_by_name)
        self.treeview.column('name', width=400, anchor=tk.W)

        self.treeview.heading('UPC', text='UPC')
        self.treeview.column('UPC', width=5, anchor=tk.CENTER)

        self.upc_dict = {}
        self.treeview.heading('12DigitUPC', text='12DigitUPC')
        self.treeview.column('12DigitUPC', width=5, anchor=tk.CENTER)

        self.treeview.heading('NameCheck', text='NameCheck')
        self.treeview.column('NameCheck', width=5, anchor=tk.CENTER)

        self.treeview.grid(row=0, column=0, sticky='nsew')  # Place the treeview in the frame

        # Create a scrollbar and associate it with the treeview
        self.scrollbar = ttk.Scrollbar(self.treeview_frame, orient='vertical', command=self.treeview.yview)
        self.scrollbar.grid(row=0, column=1, sticky='ns')
        self.treeview.configure(yscrollcommand=self.scrollbar.set)

        self.treeview.bind('<<TreeviewSelect>>', self.on_treeview_select)

        self.treeview.column('check', width=20, anchor=tk.CENTER)
        self.treeview.heading('check', text='Check', command=self.toggle_all_checkboxes)  # Add the command back

        self.treeview.bind('<Button-1>', self.show_combobox)

        self.button_frame = tk.Frame(master, background='#242424')  # Create a frame for the buttons
        self.button_frame.pack(pady=5)

        self.load_accdb_button = customtkinter.CTkButton(self.button_frame, text='Load ACCDB', command=self.load_accdb)
        self.load_csv_button = customtkinter.CTkButton(self.button_frame, text='Load CSV', command=self.load_csv)

        self.load_accdb_button.grid(row=0, column=0, padx=(5, 5), pady=5, sticky='w')
        self.load_csv_button.grid(row=0, column=1, padx=(5, 5), pady=5, sticky='w')

        self.analyze_button = customtkinter.CTkButton(self.button_frame, text='Analyze', command=self.analyze_all)
        self.analyze_button.grid(row=0, column=3, padx=5, pady=5)

        # Change the order of buttons according to your request
        self.copy_button = customtkinter.CTkButton(self.button_frame, text="Clipboard: Full UPC", fg_color='#3b91a6'
                                                   , hover_color='#1a819a', command=self.copy_selected_item_upc)
        self.copy_button.grid(row=0, column=4, padx=5, pady=5)

        self.copy_brand_family_button = customtkinter.CTkButton(self.button_frame, text="Clipboard: Brand Family",
                                  fg_color='#3b91a6', hover_color='#1a819a', command=self.copy_brand_family)
        self.copy_brand_family_button.grid(row=0, column=5, padx=5, pady=5)

        self.copy_name_button = customtkinter.CTkButton(self.button_frame, text="Clipboard: Name", fg_color='#3b91a6',
                                                        hover_color='#1a819a', command=self.copy_name)
        self.copy_name_button.grid(row=0, column=6, padx=5, pady=5)

        self.clipboard_button = customtkinter.CTkButton(self.button_frame, text='Clipboard: UPC IN', fg_color='#3b91a6',
                                                        hover_color='#1a819a', command=self.copy_upc_to_clipboard)
        self.clipboard_button.grid(row=0, column=7, padx=5, pady=5)

        self.clipboard_feedback_button = customtkinter.CTkButton(self.button_frame, text='Clipboard: Feedback', fg_color='#3b91a6',
                                                                 hover_color='#1a819a', command=self.copy_feedback_to_clipboard)
        self.clipboard_feedback_button.grid(row=0, column=8, padx=5, pady=5)

        self.feedback_entry = tk.Entry(master, width=9)
        self.feedback_entry.bind("<FocusOut>", self.hide_feedback_entry)
        self.feedback_entry.bind("<Return>", self.on_feedback_edit)

        self.add_to_database_button = customtkinter.CTkButton(self.button_frame, text='Add to database', command=self.add_to_database)
        self.add_to_database_button.grid(row=0, column=9, padx=5, pady=5)

        self.check_upc_button = customtkinter.CTkButton(self.button_frame, text='GS1 Check', fg_color='#7436d2',
                                                        hover_color='#6d29d2', command=self.check_upc_online, width=5)
        self.check_upc_button.grid(row=0, column=10, padx=5, pady=5)

        self.frame_output_selected = tk.Frame(master)
        self.frame_output_selected.pack(fill=tk.BOTH, expand=True)  # modified to expand
        self.frame_output_selected.grid_rowconfigure(0, weight=1)  # allow row to expand
        self.frame_output_selected.grid_columnconfigure(0, weight=1)  # allow column to expand

        self.output_selected_treeview = ttk.Treeview(self.frame_output_selected,
                                                     columns=('Manufacturer ID', 'Rest of UPC', 'Name', 'Manufacturer',
                                                              'Brand Family', 'Country', 'State'),
                                                     show='headings', height=1)  # Set height to 1

        self.output_selected_treeview.column('Manufacturer ID', width=100, anchor=tk.W)
        self.output_selected_treeview.column('Rest of UPC', width=70, anchor=tk.W)
        self.output_selected_treeview.column('Name', width=400, anchor=tk.W)
        self.output_selected_treeview.column('Manufacturer', width=300, anchor=tk.W)
        self.output_selected_treeview.column('Brand Family', width=250, anchor=tk.W)
        self.output_selected_treeview.column('Country', width=70, anchor=tk.W)
        self.output_selected_treeview.column('State', width=70, anchor=tk.W)
        self.output_selected_treeview.heading('Manufacturer ID', text='Manufacturer ID')
        self.output_selected_treeview.heading('Rest of UPC', text='Rest of UPC')
        self.output_selected_treeview.heading('Name', text='Name')
        self.output_selected_treeview.heading('Manufacturer', text='Manufacturer')
        self.output_selected_treeview.heading('Brand Family', text='Brand Family')
        self.output_selected_treeview.heading('Country', text='Country')
        self.output_selected_treeview.heading('State', text='State')
        self.output_selected_treeview.grid(sticky='nsew')  # modified to expand

        self.frame_output_brand = tk.Frame(master)
        self.frame_output_brand.pack(fill=tk.BOTH, expand=True)  # modified to expand
        self.frame_output_brand.grid_rowconfigure(0, weight=1)  # allow row to expand
        self.frame_output_brand.grid_columnconfigure(0, weight=1)  # allow column to expand

        self.output_brand_treeview = ttk.Treeview(self.frame_output_brand,
                                                  columns=('Manufacturer ID', 'Rest of UPC', 'Name', 'Manufacturer',
                                                           'Brand Family',
                                                           'Country', 'State'),
                                                  show='headings')
        self.output_brand_treeview.column('Manufacturer ID', width=100, anchor=tk.W)
        self.output_brand_treeview.column('Rest of UPC', width=70, anchor=tk.W)
        self.output_brand_treeview.column('Name', width=400, anchor=tk.W)
        self.output_brand_treeview.column('Manufacturer', width=300, anchor=tk.W)
        self.output_brand_treeview.column('Brand Family', width=250, anchor=tk.W)
        self.output_brand_treeview.column('Country', width=70, anchor=tk.W)
        self.output_brand_treeview.column('State', width=70, anchor=tk.W)
        self.output_brand_treeview.heading('Manufacturer ID', text='Manufacturer ID')
        self.output_brand_treeview.heading('Rest of UPC', text='Rest of UPC')
        self.output_brand_treeview.heading('Name', text='Name')
        self.output_brand_treeview.heading('Manufacturer', text='Manufacturer')
        self.output_brand_treeview.heading('Brand Family', text='Brand Family')
        self.output_brand_treeview.heading('Country', text='Country')
        self.output_brand_treeview.heading('State', text='State')
        self.output_brand_treeview.grid(sticky='nsew')  # modified to expand

        self.output = tk.Text(master, wrap=tk.WORD, height=10, width=100)
        self.output.pack()
        self.output.tag_configure('blue', foreground='blue')
        self.output.tag_configure('bold', font=('TkDefaultFont', 10, 'bold'))
        self.output.tag_configure('red', foreground='red')
        self.output.tag_configure('green', foreground='green')

        self.csv_data = None
        self.accdb_data = None
        self.results = {}  # Add a dictionary to store results

        self.treeview.tag_bind('checkbox', '<Button-1>', self.on_mouse_click)  # Bind mouse click event to 'checkbox' tag

    def toggle_button(self, button):
        new_state = 'X' if button['state'] == 'O' else 'O'
        button['state'] = new_state
        button['text'] = new_state

    def sort_by_name(self):
        # Toggle sort order
        self.name_sort_order = not getattr(self, 'name_sort_order', False)  # Default to ascending order on first click

        # Retrieve all items from the treeview
        items = [(self.treeview.item(item_id, "values")[1], item_id) for item_id in self.treeview.get_children('')]

        # Sort items by name, toggle order based on current sort order
        items.sort(reverse=self.name_sort_order)

        # Rearrange items in the treeview based on their sorted order
        for index, (name, item_id) in enumerate(items):
            self.treeview.move(item_id, '', index)


    def on_mouse_click(self, event):
        item = self.treeview.identify_row(event.y)
        col = self.treeview.identify_column(event.x)

        # If the "check" column was clicked
        if col == '#1':
            current_value = self.treeview.item(item, "values")[0]
            new_value = 'X' if current_value == 'O' else 'O'
            self.treeview.set(item, 'check', new_value)

            # Update the checked items count
            self.update_checked_items_count()

    def update_checked_items_count(self):
        checked = sum(1 for item in self.treeview.get_children() if self.treeview.item(item, "values")[0] == 'X')
        total = len(self.treeview.get_children())
        self.treeview.heading('check', text=f'Check ({checked}/{total})')  # Update the "Check" column heading

    def toggle_all_checkboxes(self):
        all_items = self.treeview.get_children()
        if all_items:
            first_value = self.treeview.set(all_items[0], 'check')
            new_value = 'X' if first_value == 'O' else 'O'
            for item in all_items:
                self.treeview.set(item, 'check', new_value)

        self.update_checked_items_count()  # Update the checked items count after toggling all checkboxes

    def calculate_checked_items(self):
        all_items = self.treeview.get_children()
        total_items = len(all_items)
        checked_items = sum(self.treeview.set(item, 'check') == 'X' for item in all_items)
        return checked_items, total_items

    def show_combobox(self, event):
        column = self.treeview.identify_column(event.x)
        row = self.treeview.identify_row(event.y)

        if column == "#1":  # Allow checking individual items in the "Check" column
            current_value = self.treeview.set(row, column)
            new_value = 'X' if current_value == 'O' else 'O'
            self.treeview.set(row, column, new_value)

            # Update the checked items count after changing an item's check status
            self.update_checked_items_count()

        elif column == "#6":  # Show the Entry widget in the "Feedback" column
            x, y, width, height = self.treeview.bbox(row, column)
            self.feedback_entry.place(x=x + self.treeview_frame.winfo_x(), y=y + self.treeview_frame.winfo_y(),
                                      width=width, height=height)
            self.feedback_entry.delete(0, tk.END)
            self.feedback_entry.insert(0, self.treeview.set(row, column))
            self.feedback_entry.focus_set()
            self.current_row = row

    def hide_feedback_entry(self, event=None):
        self.feedback_entry.place_forget()

    def on_feedback_edit(self, event=None):
        self.treeview.set(self.current_row, "Feedback", self.feedback_entry.get())
        self.hide_feedback_entry()

    def check_item_name(self, item_name):
        allowed_symbols = {'/', '.'}
        contains_double_space = False
        contains_disallowed_symbol = False

        for char in item_name:
            if char.isspace() and item_name.count('  ') > 0:
                contains_double_space = True
            if not (char.isalnum() or char.isspace() or char in allowed_symbols):
                contains_disallowed_symbol = True

        if contains_double_space and contains_disallowed_symbol:
            return 'Double Space & Symbol'
        elif contains_double_space:
            return 'Double Space'
        elif contains_disallowed_symbol:
            return 'Symbol'

        return 'YES'

    def calculate_twelve_digit_upc(self, upc):
        if not upc.isnumeric():
            return upc

        if len(upc) >= 12:
            upc = upc[:11]

        weights = [3, 1, 3, 1, 3, 1, 3, 1, 3, 1, 3]
        check_digit = 10 - sum([int(x) * y for x, y in zip(upc.zfill(11), weights)]) % 10

        if check_digit == 10:
            check_digit = 0

        return upc + str(check_digit)

    def check_upc_online(self):
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showerror("Error", "No item selected!")
            return

        item_values = self.treeview.item(selected_item[0], 'values')
        if item_values:
            twelve_digit_upc = item_values[3]  # Assuming the 12-digit UPC is in the 4th column
            url = f"https://www.gs1.org/services/verified-by-gs1/results?gtin={twelve_digit_upc}"
            webbrowser.open(url)
        else:
            messagebox.showerror("Error", "No UPC found for selected item!")


    def copy_brand_family(self):
        selected_item: Tuple = self.treeview.selection()
        if not selected_item:
            return

        selected_item_name = self.treeview.item(selected_item)['values'][1]
        brand_family = self.csv_data.loc[self.csv_data['NAME'] == selected_item_name, 'BRAND FRANCHISE'].values[0]
        pyperclip.copy(f"{brand_family} %")

    def copy_name(self):
        selected_item: Tuple = self.treeview.selection()
        if not selected_item:
            return

        selected_item_name = self.treeview.item(selected_item)['values'][1]
        selected_item_brand = self.csv_data.loc[self.csv_data['NAME'] == selected_item_name, 'BRAND'].values[0]
        selected_item_inner_pack = self.csv_data.loc[self.csv_data['NAME'] == selected_item_name, 'INNER PACK'].values[
            0]

        if selected_item_inner_pack == 1:
            copy_text = f"{selected_item_brand} SINGLE"
        elif selected_item_inner_pack > 1:
            copy_text = f"{selected_item_brand} {selected_item_inner_pack} PACK"
        else:
            copy_text = f"{selected_item_brand}"

        pyperclip.copy(copy_text)

    def copy_feedback_to_clipboard(self):
        items_with_feedback = []
        for child in self.treeview.get_children():
            feedback = self.treeview.set(child, 'Feedback')
            if feedback:
                name = self.treeview.set(child, 'name')
                items_with_feedback.append(f"{name}: {feedback}")

        if items_with_feedback:
            pyperclip.copy('\n'.join(items_with_feedback))

    def copy_upc_to_clipboard(self):
        upc_list = []
        for item in self.treeview.get_children():
            if self.treeview.set(item, 'check') == 'X':
                upc_list.append(self.treeview.set(item, 'UPC'))

        if upc_list:
            upc_str = ', '.join(f"'{upc}'" for upc in upc_list)
            self.master.clipboard_clear()
            self.master.clipboard_append(upc_str)
            self.master.update()
            messagebox.showinfo("Clipboard", "UPC(s) copied to clipboard!")
        else:
            messagebox.showerror("Error", "No UPC selected!")

        upc_in_string = "UPC in (" + ",\n".join(f"'{upc}'" for upc in upc_list) + ")"  # Updated line
        pyperclip.copy(upc_in_string)
        print("UPC list copied to clipboard.")

    def reload_accdb(self):
        if self.accdb_file_path:
            try:
                conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + self.accdb_file_path
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()

                # noinspection PyArgumentList
                tables = cursor.tables(tableType='TABLE')

                table_name = tables.fetchone().table_name

                cursor.execute(f'SELECT * FROM {table_name}')
                columns = [column[0] for column in cursor.description]
                self.accdb_data = pd.DataFrame.from_records(data=cursor.fetchall(), columns=columns)
                self.accdb_data['UPC'] = self.accdb_data['UPC'].astype(str)

                cursor.close()
                conn.close()
                print('ACCDB file reloaded successfully.')
                self.add_to_database_button.configure(text='Items added', fg_color='dark green', hover_color='#045004')
            except Exception as e:
                print('Error reloading ACCDB file:', e)

    def add_to_database(self):
        items_to_add = []
        for item in self.treeview.get_children():
            check_value = self.treeview.set(item, 'check')
            if check_value == "X":
                index = self.treeview.index(item)
                selected_row = self.csv_data.iloc[index]
                items_to_add.append(selected_row)

        if not items_to_add:
            print('Nothing to add.')
            return

        items_to_add = pd.DataFrame(items_to_add)

        if items_to_add.empty:
            print('Nothing to add.')
            return

        try:
            conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + self.accdb_file_path
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            # noinspection PyArgumentList
            tables = cursor.tables(tableType='TABLE')
            table_name = tables.fetchone().table_name

            items_to_add['ID'] = 'JUST ADDED'
            for index, row in self.csv_data.iterrows():
                # Assume you have a unique ID for each row or generate one
                unique_id = row['ID']  # Or use `str(uuid.uuid4())` to generate a UUID
                self.treeview.insert('', 'end', iid=unique_id, values=(...))


            for index, row in items_to_add.iterrows():
                query = f''' 
           INSERT INTO {table_name} (ID, Name, Manufacturer, UPC, Desc4, Desc3, Desc35) 
           VALUES (?, ?, ?, ?, ?, ?, ?) 
           '''
                cursor.execute(query, row['ID'], row['NAME'], row['MANUFACTURER'], row['UPC'], row['BRAND FAMILY'],
                               row['BREWERY LOCATION COUNTRY'], row['BREWERY LOCATION STATE'])
                conn.commit()

            cursor.close()
            conn.close()

            print('Items added successfully.')
            self.load_accdb_button.configure(text='Database Loaded', fg_color='dark green', hover_color='#045004')
            self.add_to_database_button.configure(text='Items added', fg_color='dark green', hover_color='#045004')
            self.load_csv_button.configure(text='Load CSV', fg_color='#1f6aa5', hover_color='#144870')
            self.analyze_button.configure(text='Analyze', fg_color='#1f6aa5', hover_color='#144870')

            self.reload_accdb()

        except Exception as e:
            self.add_to_database_button.configure(text='Fail to add', fg_color='red')
            print('Error adding items to ACCDB file:', e)

    def sort_output_brand_treeview(self):
        items = [(self.output_brand_treeview.set(child, 'Name'), child) for child in
                 self.output_brand_treeview.get_children('')]
        items.sort()

        for index, (value, child) in enumerate(items):
            self.output_brand_treeview.move(child, '', index)

    def analyze_all(self):
        # Clear the output widget
        self.add_to_database_button.configure(text='Add to database', fg_color='#1f6aa5')

        self.output.delete('1.0', tk.END)

        # Clear the output_brand_treeview widget
        self.output_brand_treeview.delete(*self.output_brand_treeview.get_children())
        self.output_selected_treeview.delete(
            *self.output_selected_treeview.get_children())  # Clear old selected item data

        # Clear previous results
        self.results = {}

        for item in self.treeview.get_children():
            check_value = self.treeview.item(item, 'values')[0]
            if check_value == "X":
                self.treeview.selection_set(item)
                upc_results = self.check_upc()
                brand_results = self.check_brand()

                self.results[item] = {'upc': upc_results, 'brand': brand_results}  # Store results using item key

                # Update the UPC result in the treeview
                if upc_results == '12 digit UPC':
                    self.treeview.set(item, "12DigitUPC", 'YES')
                    self.treeview.set(item, "Feedback", 'YES')
                elif upc_results == 'UPC not found':
                    self.treeview.set(item, "12DigitUPC", 'NO')
                    self.treeview.set(item, "Feedback", 'NO')

                # within analyze_all method
                name = self.treeview.item(item, 'values')[1]  # assuming 'name' is the second column
                name_check_result = self.check_item_name(name)

                # Update the Check Name result in the treeview
                self.treeview.set(item, "NameCheck", name_check_result)

        # Update the output Text widget with the results for the currently selected item
        selected_item = self.treeview.selection()
        if selected_item:
            self.on_treeview_select(None)  # Update the Text widgets with the stored results
        self.analyze_button.configure(text='Analyze - Done', fg_color='dark green', hover_color='#045004')

    def toggle_checkbox(self, event):
        column = self.treeview.identify('column', event.x, event.y)
        item = self.treeview.identify('item', event.x, event.y)
        if column == '#1' and item:  # Check if the clicked column is 'Check' column
            current_value = self.treeview.item(item, 'values')[0]
            if current_value == "O":
                self.treeview.set(item, column='check', value="X")
                self.treeview.item(item, tags=('bold',))  # Bold the text
            else:
                self.treeview.set(item, column='check', value="O")
                self.treeview.item(item, tags=('bold',))  # Bold the text

    def on_treeview_select(self, event):
        selected_item = self.treeview.selection()

        # Clear both treeviews
        self.output_brand_treeview.delete(*self.output_brand_treeview.get_children())
        self.output_selected_treeview.delete(*self.output_selected_treeview.get_children())

        if selected_item:
            selected_item = selected_item[0]  # Extract the selected item's key

            if selected_item in self.results:
                # Display previous UPC results if available
                self.output.delete('1.0', tk.END)
                for line in self.results[selected_item]['upc'].split('\n'):
                    if "Exact match:" in line or "Approximated UPC check:" in line:
                        self.output.insert(tk.END, line + '\n')
                    elif "No exact match found" in line or "No match found" in line:
                        self.output.insert(tk.END, line + '\n', 'red')
                    else:
                        self.output.insert(tk.END, line + '\n', 'green')

                # Display selected item's brand details
                selected_item_brand = self.results[selected_item]['brand'][0]
                self.output_selected_treeview.insert('', 'end', values=selected_item_brand)

                # Display previous Check Brand results if available
                for matched_item in self.results[selected_item]['brand'][1]:
                    self.output_brand_treeview.insert('', 'end', values=matched_item)

                # Sort the items in the output_brand_treeview
                self.sort_output_brand_treeview()

            else:  # Handle case when selected item isn't in the results
                self.output.delete(1.0, tk.END)
                self.output.insert(tk.END, 'Please analyze the item before trying to select it.\n', 'red')
        else:
            self.output.delete(1.0, tk.END)
            self.output.insert(tk.END, 'Please select an item from the list.\n')

    def check_upc(self):
        selected_item = self.treeview.selection()
        if selected_item:
            index = self.treeview.index(selected_item[0])
            selected_row = self.csv_data.iloc[index]

            upc_csv = selected_row['UPC']
            upc_csv_no_last_digit = upc_csv[:-1]

            exact_matches = []
            approx_matches = []

            for _, row in self.accdb_data.iterrows():
                if row['UPC'] == upc_csv:
                    exact_matches.append((row['ID'], row['Name']))
                elif row['UPC'] == upc_csv_no_last_digit:
                    approx_matches.append((row['ID'], row['Name']))

            output_text = ""
            if exact_matches:
                matched_items = ', '.join([f"ID {item_id} - {item_name}" for item_id, item_name in exact_matches])
                output_text += f"Exact match:\n {matched_items}\n"
            else:
                output_text += 'Exact match:\nNo exact match found.\n'

            if approx_matches:
                approx_items = ', '.join([f"ID {item_id} - {item_name}" for item_id, item_name in approx_matches])
                output_text += f"Approximated UPC check:\n {approx_items}\n"
            else:
                output_text += 'Approximated UPC check:\n No match found.\n'

            return output_text
        else:
            return 'Please select an item from the list.\n'

    def load_accdb(self):
        accdb_file_path = filedialog.askopenfilename(filetypes=[('Access Database', '*.accdb')])
        if accdb_file_path:
            self.accdb_file_path = accdb_file_path
            try:
                conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + accdb_file_path
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()

                # noinspection PyArgumentList
                tables = cursor.tables(tableType='TABLE')

                table_name = tables.fetchone().table_name

                cursor.execute(f'SELECT * FROM {table_name}')
                columns = [column[0] for column in cursor.description]
                self.accdb_data = pd.DataFrame.from_records(data=cursor.fetchall(), columns=columns)
                self.accdb_data['UPC'] = self.accdb_data['UPC'].astype(str)

                cursor.close()
                conn.close()
                print('ACCDB file loaded successfully.')
                self.load_accdb_button.configure(text='Database Loaded', fg_color='dark green', hover_color='#045004')
            except Exception as e:
                print('Error loading ACCDB file:', e)
                self.load_accdb_button.configure(text='Error loading DB', fg_color='red')
    
    def copy_12_digit_upc_to_clipboard(self, index):
        try:
            twelve_digit_upc = self.upc_dict[index]
            pyperclip.copy(twelve_digit_upc)
            print("12DigitUPC copied to clipboard:", twelve_digit_upc)
        except Exception as e:
            print(f"Error: {e}, Index: {index}")

    def copy_selected_item_upc(self):
        selected_items = self.treeview.selection()
        if not selected_items:
            print("No item is selected.")
            return

        item = selected_items[0]
        index = int(item.split("_")[-1])  # Extract index from the tag (e.g., "row_7" -> 7)
        self.copy_12_digit_upc_to_clipboard(index)

    def load_csv(self):
        csv_file_path = filedialog.askopenfilename(filetypes=[('CSV file', '*.csv')])
        if csv_file_path:
            self.load_csv_button.configure(text='CSV Loaded', fg_color='dark green', hover_color='#045004')
            self.analyze_button.configure(text='Analyze', fg_color='#1f6aa5', hover_color='#144870')
            self.add_to_database_button.configure(text='Add to database', fg_color='#1f6aa5')
            self.csv_data = pd.read_csv(csv_file_path, dtype={'UPC': str})

            self.treeview.delete(*self.treeview.get_children())

            total_rows = len(self.csv_data)

            for index, row in self.csv_data.iterrows():
                name = row['NAME']
                upc = row['UPC']
                twelve_digit_upc = self.calculate_twelve_digit_upc(upc)
                self.upc_dict[index] = twelve_digit_upc
                print(f"Adding index {index} to upc_dict")

                name_check = self.check_item_name(name)

                # Add a tag to each row that includes the index
                tag = f"row_{index}"

                if re.search(r'[^a-zA-Z0-9 .\/]', name) or '  ' in name:
                    item = self.treeview.insert('', 'end', tag, values=('X', name, upc, twelve_digit_upc, name_check),
                                                tags=(tag, 'red'))
                else:
                    item = self.treeview.insert('', 'end', tag, values=('X', name, upc, twelve_digit_upc, name_check),
                                                tags=(tag,))

                # Make sure the item is visible before placing the button
                fraction = index / total_rows
                self.treeview.yview_moveto(fraction)
                self.update_checked_items_count()  # Update the checked items count after loading CSV data

    def check_brand(self):
        selected_item = self.treeview.selection()
        if selected_item:
            index = self.treeview.index(selected_item[0])
            selected_row = self.csv_data.iloc[index]

            brand_csv = selected_row['BRAND FRANCHISE']
            brand_family = selected_row['BRAND FAMILY']  # Get the Brand Family

            matched_items = []
            for _, row in self.accdb_data.iterrows():
                # Use a regular expression to match the brand_csv as a prefix, followed by any characters
                if re.match(f"^{re.escape(brand_csv)}", str(row['Desc4'])):
                    upc = str(row['UPC'])
                    manufacturer_id = upc[:6]
                    upc_part2 = upc[6:]
                    matched_items.append((manufacturer_id, upc_part2, row['Name'], row['Manufacturer'],
                                          row['Desc4'],
                                          row['Desc3'], row['Desc35']))

            # Sort matched_items based on the 'Name' column
            matched_items.sort(key=lambda x: x[2])  # Assuming the 'Name' column is the third one in the tuple

            selected_upc = selected_row['UPC']
            selected_name = selected_row['NAME']
            manufacturer = selected_row['MANUFACTURER']
            manufacturer_id = selected_upc[:6]
            upc_part2 = selected_upc[6:]
            country = selected_row['BREWERY LOCATION COUNTRY']
            state = selected_row['BREWERY LOCATION STATE']

            selected_item_output = (manufacturer_id, upc_part2, selected_name,
                                    manufacturer, brand_family, country, state)

            return selected_item_output, matched_items
        else:
            return None, []


if __name__ == '__main__':
    root = customtkinter.CTk()
    app = App(root)
    root.mainloop()
