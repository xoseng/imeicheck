# coding=utf8
# author: Xose Brais Noya Garcia


#IMPORTS

from PIL import Image
from tkinter import *
from tkinter import filedialog
from PIL import ImageTk, Image
import webbrowser
import threading
from tkinter import ttk

#GLOBAL VARS
version_software='IMEI Check RC V2.3'
author='Xosé Brais Noya García'
url='https://github.com/xoseng/imeicheck'
email='xose.noya.garcia@gmail.com'
github='https://github.com/xoseng'
logo_img='./img/logo.ico'

# GUI FUNCTIONS
def about_program():
    #import Tkinter as tk
    import tkinter as tk
    root = tk.Tk()
    root.title("About")  # Titulo de la ventana
    root.iconbitmap(logo_img)  # Icono de la ventana, en ico o xbm en Linux
    root.resizable(0, 0)
    texto = tk.Text(root)
    texto.pack()
    texto.config(width=38, height=4, padx=5, pady=5)
    texto.insert(tk.END,version_software+"\nAuthor: "+author+"\nEmail: "+email+"\nGitHub: "+github+"\n")
    texto.config(state="disabled")
    root.mainloop()

def warning_path():
    import tkinter as tk
    root = tk.Tk()
    root.title("Notice path")  # Titulo de la ventana
    root.iconbitmap(logo_img)  # Icono de la ventana, en ico o xbm en Linux
    root.resizable(0, 0)
    texto = tk.Text(root)
    texto.pack()
    texto.config(width=31, height=2, padx=5, pady=5)
    texto.insert(tk.END,"Notice!\nRequired file path to works!")
    texto.config(state="disabled")
    root.mainloop()

def warning_file():
    import tkinter as tk
    root = tk.Tk()
    root.title("Notice path")  # Titulo de la ventana
    root.iconbitmap(logo_img)  # Icono de la ventana, en ico o xbm en Linux
    root.resizable(0, 0)
    texto = tk.Text(root)
    texto.pack()
    texto.config(width=38, height=2, padx=5, pady=5)
    texto.insert(tk.END,"Notice!\nYou must select at least one file!")
    texto.config(state="disabled")
    root.mainloop()

def warning_option():
    import tkinter as tk
    root = tk.Tk()
    root.title("Notice path")  # Titulo de la ventana
    root.iconbitmap(logo_img)  # Icono de la ventana, en ico o xbm en Linux
    root.resizable(0, 0)
    texto = tk.Text(root)
    texto.pack()
    texto.config(width=43, height=2, padx=5, pady=5)
    texto.insert(tk.END,"Notice!\nYou must choose at least one sheet option!")
    texto.config(state="disabled")
    root.mainloop()

def license_agreement_gnu():
    #import Tkinter as tk
    import tkinter as tk
    root = tk.Tk()
    root.title("license agreement")  # Titulo de la ventana
    root.iconbitmap(logo_img)  # Icono de la ventana, en ico o xbm en Linux
    root.resizable(0, 0)
    texto = tk.Text(root)
    texto.pack()
    texto.config(width=77, height=8, padx=5, pady=5)
    texto.insert(tk.END,"GNU General Public License v3.0\n\n"
                        "Permissions of this strong copyleft license are conditioned\n"
                        "on making available complete source code of licensed works and modifications,\n"
                        "which include larger works using a licensed work, under the same license.\n"
                        "Copyright and license notices must be preserved.\n"
                        "Contributors provide an express grant of patent rights.\n")
    texto.config(state="disabled")
    root.mainloop()

def exit_program():
    sys.exit()
    root.destroy()

def center(win):
    """
    centers a tkinter window
    :param win: the root or Toplevel window to center
    """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()

def project_website():
    import webbrowser
    webbrowser.open(url)

def get_filepathload():
    import tkinter as tk
    from tkinter import filedialog
    root = Tk()
    root.withdraw()
    filez=root.filename = filedialog.askopenfilenames(initialdir="/", title="Select files",
                                               filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")),defaultextension=".xlsx")
    lst = list(filez)
    root.destroy()
    #return list with path of files to load!
    return lst

def set_filepathsave():
    import tkinter as tk
    from tkinter import filedialog
    root = Tk()
    root.withdraw()
    root.filename = filedialog.asksaveasfile(initialdir="/", title="Save as",
                                             filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")), mode='w',
                                             defaultextension=".xlsx")
    try:
        filepath=str(root.filename).strip()
        filepath = filepath.replace("/", "\\")
        filepath = filepath.split("'")[1].strip()
    except:
        filepath=''
    root.destroy()
    return filepath