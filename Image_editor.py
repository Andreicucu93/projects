import os
import tkinter
from tkinter import Tk, Label, Entry, Button, filedialog, StringVar, Listbox, messagebox, Checkbutton, BooleanVar, \
    Radiobutton
from PIL import Image, ImageFilter
import customtkinter


root = customtkinter.CTk()
root.title("Imaging - file editor (v.012624)")
root.geometry("500x480") #width x height
root.minsize(height=480, width=500)
root.maxsize(height=480, width=500)
customtkinter.set_appearance_mode("Dark")  # or "Light"


extension_type = StringVar()
extension_type.set(".jpg")

def resize_images(directory, size):
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith(('.jpg', '.png')):
                try:
                    with Image.open(os.path.join(root, filename)) as img:
                        max_size = max(img.size)
                        if max_size > 249:
                            if img.size[0] > img.size[1]:
                                width_percent = (size / float(img.size[0]))
                                height = int((float(img.size[1]) * float(width_percent)))
                                img = img.resize((size, height), resample=Image.Resampling.NEAREST)
                            else:
                                height_percent = (size / float(img.size[1]))
                                width = int((float(img.size[0]) * float(height_percent)))
                                img = img.resize((width, size), resample=Image.Resampling.NEAREST)
                            img.save(os.path.join(root, filename))
                except IOError:
                    print(f"{filename} is not an image file.")
    messagebox.showinfo("Information", "The Images have been resized.")


def browse_folder():
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    subdirs = []
    for root, dirs, files in os.walk(filename):
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            for file in os.listdir(dir_path):
                if file.lower().endswith(('.jpg', '.png')):
                    subdirs.append(dir_path)
                    break
    listbox.delete(0, 'end')
    for subdir in subdirs:
        listbox.insert('end', subdir)


folder_path = StringVar()
label = customtkinter.CTkLabel(root, text="Select the location containing images:")
label.pack(pady=(10, 10))

select_files_frame = customtkinter.CTkFrame(root, width=450, bg_color='#242424')
select_files_frame.pack()

entry = customtkinter.CTkEntry(select_files_frame, textvariable=folder_path)
entry.grid(row=0, column=0, padx=10)

browse = customtkinter.CTkButton(select_files_frame, text="Browse", command=browse_folder)
browse.grid(row=0, column=1)

selectedFiles_label = customtkinter.CTkLabel(root, text="Inner paths selected 🡻\n(If applicable)")
selectedFiles_label.pack(pady=(5, 5))

listbox = Listbox(root, width=80)
listbox.pack()

width_label = customtkinter.CTkLabel(root, text="🡻 Enter new size for loaded images (batch resize)\n"
                                                "350 = standard")
width_label.pack(pady=(15, 5))

resize_frame = customtkinter.CTkFrame(root, width=200, height=20, fg_color="#242424")
resize_frame.pack(pady=(10, 10))

width_entry = customtkinter.CTkEntry(resize_frame, width=60)
width_entry.grid(row=0, column=0)

resize_button = customtkinter.CTkButton(resize_frame, text="Resize", fg_color='darkblue', width=50, command=lambda: check_inputs())
resize_button.grid(row=0, column=1, padx=(10, 10))


def check_inputs():
    if folder_path.get() and width_entry.get():
        resize_images(folder_path.get(), int(width_entry.get()))
    elif not folder_path.get():
        messagebox.showerror("Error", "Please select a folder containing the images.")
    else:
        messagebox.showerror("Error", "Please enter a width for the images.")


def remove_extension(directory):
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith(('.jpg', '.png')):
                new_filename = os.path.splitext(filename)[0]
                os.rename(os.path.join(root, filename), os.path.join(root, new_filename))
    messagebox.showinfo("Information", "The file extensions have been removed.")


remove_button = customtkinter.CTkButton(resize_frame, text="Remove extension", width=30, command=lambda: remove_extension(folder_path.get()))
remove_button.grid(row=0, column=2, padx=(0, 20))


def add_extension(directory, extension):
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if not filename.lower().endswith(('.jpg', '.png')):
                new_filename = f"{filename}{extension}"
                os.rename(os.path.join(root, filename), os.path.join(root, new_filename))
    messagebox.showinfo("Information", "The file extensions have been added.")


extension_var = StringVar()
extension_var.set(".jpg")

extension_radiobuttons = customtkinter.CTkFrame(root, width=200, height=20, fg_color="#242424")
extension_radiobuttons.pack(pady=(10, 20))

extension_jpg = customtkinter.CTkRadioButton(extension_radiobuttons, text=".jpg", variable=extension_var, value=".jpg",
                                             width=50)
extension_jpg.grid(row=0, column=0, padx=(10, 10))

extension_png = customtkinter.CTkRadioButton(extension_radiobuttons, text=".png", variable=extension_var, value=".png",
                                             width=50)
extension_png.grid(row=0, column=1, padx=(10, 10))

add_extension_button = customtkinter.CTkButton(extension_radiobuttons, text="Add selected extension",
                              command=lambda: add_extension(folder_path.get(), extension_var.get()))
add_extension_button.grid(row=0, column=2)


root.mainloop()
