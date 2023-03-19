import tkinter as tk
from tkinter import filedialog as fd
import os
from win32com.client import Dispatch
import win32com.client

root = tk.Tk()
root.title("LOL Language Changer")
p1 = tk.PhotoImage(file = 'icon.png')
root.iconphoto(False,p1)
root.geometry("400x350")

file_path = ""


def select_path():
    global file_path
    global show_path

    with open("path.txt", 'r+') as path:
        line = path.readline()

        if line.strip() == "":
            file_path = fd.askopenfilename()
            path.write(f"{file_path}")
        else:
            file_path = f"{line}"

        show_path.config(text=f"Path:\n {file_path}")


def reset_path():
    with open("path.txt", 'r+') as path:
        path.truncate(0)
        show_path.config(text="Choose path")


def select_language(option):
    global file_path

    if file_path == "":
        language_choice.config(text="Please select a file path first.")
        return

    language_choice.config(text=f"Created shortcut for {option} language")
    language_names = {
        'Japanese': 'ja_JP',
        'Korean': 'ko_KR',
        'Chinese': 'zh_CN',
        'Taiwanese': 'zh_TW',
        'Spanish (Spain)': 'es_ES',
        'Spanish (Latin America)': 'es_MX',
        'English (United States)': 'en_US',
        'English (United Kingdom)': 'en_GB',
        'English (Australia)': 'en_AU',
        'French': 'fr_FR',
        'German': 'de_DE',
        'Italian': 'it_IT',
        'Polish': 'pl_PL',
        'Romanian': 'ro_RO',
        'Greek': 'el_GR',
        'Portuguese (Brazil)': 'pt_BR',
        'Hungarian': 'hu_HU',
        'Russian': 'ru_RU',
        'Turkish': 'tr_TR'
    }

    option = language_names[option]

    shortcut_path = os.path.join(os.getcwd(), f"League Of Legends.lnk")

    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)

    wsh = win32com.client.Dispatch("WScript.Shell")
    shortcut = wsh.CreateShortcut(shortcut_path)

    shortcut.TargetPath = file_path
    shortcut.Arguments = f"--locale={option}"
    shortcut.WorkingDirectory = os.path.dirname(file_path)
    shortcut.Description = f"League of Legends Shortcut"
    shortcut.IconLocation = file_path
    shortcut.Save()

    open_button = tk.Button(root, text="Open League Of Legends", command=lambda: os.startfile(shortcut_path))
    open_button.pack(pady=10, anchor=tk.N, padx=20)


button_select_path = tk.Button(root, text="Select path", command=select_path)
button_reset_path = tk.Button(root, text="Reset", command=reset_path)

Languages = ["Japanese", "Korean", "Chinese", "Taiwanese", "Spanish (Spain)", "Spanish (Latin America)",
             "English (United States)", "English (United Kingdom)", "English (Australia)", "French", "German",
             "Italian", "Polish", "Romanian", "Greek", "Portuguese", "Hungarian", "Russian", "Turkish"]

selected_language = tk.StringVar(root)
selected_language.set(Languages[0])

language_menu = tk.OptionMenu(root, selected_language, *Languages, command=select_language)

hint = tk.Label(root, text="Select path to LeagueClient.exe in your League Of Legends directory")
show_path = tk.Label(root, text=f"Choose Path")
language_choice = tk.Label(root, text="")

hint.pack(pady=20)
button_select_path.pack(pady=10, anchor=tk.N, padx=20)
button_reset_path.pack(pady=10, anchor=tk.N, padx=20)

language_menu.pack(pady=10, anchor=tk.N, padx=20)
language_choice.pack(pady=10)
show_path.pack(side=tk.BOTTOM, pady=10)

root.mainloop()
