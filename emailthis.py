import argparse
import sys
import sqlite3

import smtplib
import os
import mimetypes
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.multipart import MIMEMultipart

import tkinter as tk
import tkinter.messagebox as mb
from tkinter import Tk, Toplevel, Text, BOTH, X, N, LEFT, RIGHT
from tkinter.ttk import Frame, Label, Entry, Combobox

import emtsconf as cnf

absFilePath = os.path.abspath(__file__)
path, filename = os.path.split(absFilePath)

conn = sqlite3.connect(path+"\emailthis.db")
cursor = conn.cursor()
cursor.execute("""CREATE TABLE IF NOT EXISTS mailto(name text, addr text)""")
conn.commit()

def send_email(addr_to, msg_subj, msg_text, files):
    name_from = cnf.efrom

    msg = MIMEMultipart()
    msg['From'] = name_from
    msg['To'] = addr_to
    msg['Subject'] = msg_subj

    body = msg_text
    msg.attach(MIMEText(body, 'plain'))

    process_attachement(msg, files)

    addr_from = cnf.efrom
    password = cnf.epass

    server = smtplib.SMTP_SSL(cnf.eserver, cnf.eport)
    #server.starttls()
    #server.set_debuglevel(True)
    server.login(addr_from, password)
    server.send_message(msg)
    server.quit()
    print('Письмо отправлено') 

def process_attachement(msg, files):
    for f in files:
        if os.path.isfile(f):
            attach_file(msg,f)
        elif os.path.exists(f):
            dir = os.listdir(f)
            for file in dir:
                attach_file(msg,f+"/"+file)

def attach_file(msg, filepath):
    filename = os.path.basename(filepath)
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    if maintype == 'text':
        with open(filepath, encoding="utf-8") as fp:
            file = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
    elif maintype == 'image':
        with open(filepath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
    elif maintype == 'audio':
        with open(filepath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
    else:
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)
            file.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(file)
    file.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(file)
    print(f'Добавлен файл: {filename}')

def get_addr(name):
    sql = "SELECT addr FROM mailto WHERE name = ?"
    caddr = cursor.execute(sql, (name,))
    faddr = caddr.fetchall()
    addr = faddr[0][0]
    return addr

def change_addr(event):
    sql = "SELECT addr FROM mailto WHERE name = ?"
    ewg = event.widget.get()
    print(ewg)
    caddr = cursor.execute(sql, (ewg,))
    faddr = caddr.fetchall()
    addr = faddr[0][0]
    print(addr)

class settingsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

class addDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.username = tk.StringVar()
        self.usermail = tk.StringVar()

        label_name_m = tk.Label(self,text = 'Имя')
        label_name_m.pack()
        entry_name_m = tk.Entry(self, textvariable=self.username)
        entry_name_m.pack()
        label_addr_m = tk.Label(self,text = 'Адрес')
        label_addr_m.pack()
        entry_addr_m = tk.Entry(self, textvariable=self.usermail)
        entry_addr_m.pack()
        button_add_m = tk.Button(self, text='Добавить', command=self.write_user)
        button_add_m.pack()

    def write_user(self):
        sql_add = "INSERT INTO mailto (name, addr) VALUES (?, ?)"
        cursor.execute(sql_add, (self.username.get(), self.usermail.get()))
        conn.commit()
        if mb.showinfo("Информация", 'Получатель добавлен в базу данных'):
            self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        #self.geometry("640x480+300+300")

        self.frame = Frame()
        self.frame.pack(fill=BOTH, expand=True)

        ######################################################################################
        self.frame1 = Frame(self.frame)
        self.frame1.pack(fill=X)

        #self.cnames = cursor.execute("SELECT name FROM mailto")
        #self.names = self.cnames.fetchall()

        self.label_to = tk.Label(self.frame1, text='Кому:', width=10)
        self.label_to.pack(side=LEFT, padx=5, pady=5)

        self.combo_to = Combobox(
            self.frame1,
            #values = self.names,
            postcommand=self.fill_combo
            )
        self.fill_combo()

        self.combo_to.current(0)
        self.combo_to.bind("<<ComboboxSelected>>", change_addr)
        self.combo_to.pack(side=LEFT, fill=X, padx=5, expand=True)

        self.button_add = tk.Button(self.frame1, text = 'Добавить', command=self.open_add)
        self.button_add.pack(side=LEFT, padx=5, pady=5)
        self.button_del = tk.Button(self.frame1, text = 'Удалить', command=self.del_user)
        self.button_del.pack(side=LEFT, padx=5, pady=5)

        ###########################################################################################
        self.frame2 = Frame(self.frame)
        self.frame2.pack(fill=X)
        self.label_subj = tk.Label(self.frame2, text='Тема:', width=10)
        self.label_subj.pack(side=LEFT, padx=5, pady=5)
        self.entry_subj = tk.Entry(self.frame2)
        self.entry_subj.insert(0, cnf.subj)
        self.entry_subj.pack(fill=X, padx=5, expand=True)     

        ##########################################################################################
        self.frame3 = Frame(self.frame)
        self.frame3.pack(fill=BOTH, anchor=N, expand=True)
        self.label_msg = tk.Label(self.frame3, text='Сообщение:', width=10)
        self.label_msg.pack(side=LEFT, anchor=N, padx=5, pady=5)
        self.text_msg = tk.Text(self.frame3, height = 10, width=25)
        self.text_msg.pack(side=LEFT,fill=X, anchor=N, pady=5, padx=5, expand=True)
        self.text_msg.insert(1.0, cnf.text)

        ###########################################################################################
        self.frame4 = Frame(self.frame)
        self.frame4.pack(fill=BOTH, anchor=N, expand=True)
        self.label_files = tk.Label(self.frame4, text='Файлы:', width=10)
        self.label_files.pack(side=LEFT, anchor=N, padx=5, pady=5)
        self.labels = []
        for i in range(len(args.list)):
            self.File = tk.Label(self.frame4, text = args.list[i])
            self.File.pack(anchor=N)
            self.labels.append(self.File)

        #############################################################################################
        self.frame5 = Frame(self.frame, height = 5)
        self.frame5.pack(side = RIGHT)
        self.button = tk.Button(
            self.frame5, 
            text="Настройки!",
            command = self.open_settings
        )
        self.button.pack( padx=5, pady=5)
        self.button = tk.Button(
            self.frame5, 
            text="Отправить!",
            command = self.send_click
        )
        self.button.pack( padx=5, pady=5)

    def fill_combo(self):
        self.cnames = cursor.execute("SELECT name FROM mailto")
        self.names = self.cnames.fetchall()
        self.combo_to['values'] = self.names
        self.combo_to.current(0)

    def open_add(self):
        add_dialog = addDialog(self)
        add_dialog.transient(app)
        add_dialog.grab_set()
        add_dialog.focus_set()
        add_dialog.wait_window()

    def open_settings(self):
        set_dialog = settingsDialog(self)
        set_dialog.transient(app)
        set_dialog.grab_set()
        set_dialog.focus_set()
        set_dialog.wait_window()

    def del_user(self):
        if mb.askyesno("Вопрос", f"Вы действительно хотите удалить\nпользователя: {self.combo_to.get()} — {get_addr(self.combo_to.get())}?"):
            sql = "DELETE FROM mailto WHERE addr = ?"
            cursor.execute(sql, (get_addr(self.combo_to.get()),))
            conn.commit()


    def send_click(self):
        send_email(get_addr(self.combo_to.get()), self.entry_subj.get(), self.text_msg.get("1.0", tk.END), args.list)
        msg = "Ваше письмо успешно отправлено"
        mb.showinfo("Информация", msg)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog='eMail',
        description='Программа отсылающая файлы по почте',
        epilog='(c) bigvik'
    )
    parser.add_argument ('-l', '--list', nargs='+')
    args = parser.parse_args(sys.argv[1:])
    if  not args.list:
        msg = "Не переданы файлы для отправки:\n--list список_файлов"
        mb.showwarning("Предупреждение", msg)
        quit()
    else:
        app = App()
        app.mainloop()
