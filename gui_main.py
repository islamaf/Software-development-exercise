import os
from tkinter import Tk, ttk, filedialog
import pandas as pd
from win32 import win32api

root = Tk()
root.title('Ahram Exam')
root.resizable(True, True)

root.frame_header = ttk.Frame()
root.geometry("350x250")
root.eval('tk::PlaceWindow . center')

ttk.Label(root.frame_header, text='Browse file to open:', style='Header.TLabel', font=("Arial", 15)).grid(row=1, column=1)

filename = ttk.Button(root.frame_header, text="Browse", command=lambda: open_file()).grid(row=4, column=1)

print_result = ttk.Button(root.frame_header, text="Print result", command=lambda: print_file())
print_result.grid(row=12, column=1)
print_result['state'] = 'disabled'


def open_file():
    file_to_open = filedialog.askopenfilename(initialdir="C:/", title="Select file",
                                              filetypes=(("all files", "*.*"), ("excel files", "*.xls")))
    df = pd.read_excel(file_to_open)
    os.startfile(file_to_open)

    ttk.Label(root.frame_header, text='All averages:', style='Header.TLabel',font=("Arial", 15)).grid(row=6, column=1)
    ttk.Label(root.frame_header, text=df.mean(), style='Header.TLabel', font=("Arial", 15)).grid(row=8, column=1)

    ttk.Label(root.frame_header, text=get_max_mean(df), style='Header.TLabel', font=("Arial", 15)).grid(row=10, column=1)

    f = open('maximum_average.txt', 'w')
    f.write(get_max_mean(df))
    f.close()

    root.geometry("350x350")

    print_result['state'] = 'enabled'


def print_file():
    file_to_print = "maximum_average.txt"
    if file_to_print:
        win32api.ShellExecute(0, "print", file_to_print, None, ".", 0)


def get_max_mean(l):
    max_val = 0
    max_column = ''
    winner = ""
    for i, x in zip(l.columns, l.mean()):
        if x > max_val:
            max_val = x
            max_column = i
            winner = f'{max_column} is the maximum'
    return winner


root.frame_header.pack(pady=10, anchor="center")
root.mainloop()