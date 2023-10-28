import os, openpyxl, sqlite3
from tkinter import *
from multimoney import *
from comercios import *
from manual_1 import *

os.system('cls')

global segoe, style, basebg, fulldatapack
FONT = 'Segoe UI'
BASEBG = '#111'

def replace(remove, make):
    remove.destroy()
    make()

root = Tk()
root.title('Organizador de documentos')
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
root.config(background='#1E1E1E')
root.state('zoomed')
try: root.iconbitmap('icon.ico')
except: pass
menubar = Menu(root)
root.config(menu=menubar)

def init(parentframe, text1, text2, content):
    parentframe.config(background='#1E1E1E')
    parentframe.grid(row=0, column=0, sticky='new', padx=150, pady=(50,0))
    parentframe.columnconfigure(0, weight=1)
    Label(parentframe, text=text1, background='#111', fg='#555', font=(FONT, 12)).grid(row=0, column=0, sticky='ew', padx=20, ipadx=10, ipady=10)
    Label(parentframe, text=text2, background='#000', fg='#008000', font=(FONT, 15), highlightthickness=5, highlightbackground='#001A00').grid(row=1, column=0, sticky='ew', ipady=10)
    content.grid(row=2, column=0, sticky='ew', padx=20, ipadx=30, pady=(0,50))
    content.columnconfigure(0, weight=1)
    content.config(background='#111')

def buildf_home():
    # Required.
    global on_frame, frame1, content1
    frame1 = Frame(root)
    on_frame = frame1
    content1 = Frame(frame1)
    init(frame1, 'PÁGINA DE INCIO', 'PROCESADOR DE DOCUMENTOS', content1)

    # Personalized / By calling "init" no rows are consumed.
    Label(content1, text='Seleccione una opción para empezar...', background=BASEBG, fg='#FFF').grid(row=0, column=0, sticky='ew', pady=20)
    wrapbtns = Frame(content1, background=BASEBG)
    wrapbtns.grid(row=1, column=0)
    wrapbtns.columnconfigure([0,1,2], weight=1)
    Button(wrapbtns, text='Multimoney', cursor='hand2', background=BASEBG, fg='#0F0', border=0, command=lambda:replace(on_frame, buildf_multimoney)).grid(row=0, column=0, padx=20, pady=20)
    Button(wrapbtns, text='Comercios', cursor='hand2', background=BASEBG, fg='#0F0', border=0, command=lambda:replace(on_frame, buildf_comercios)).grid(row=0, column=1, padx=20, pady=20)
    Button(wrapbtns, text='Beto', cursor='hand2', background=BASEBG, fg='#0F0', border=0, command=lambda:replace(on_frame, buildf_beto)).grid(row=0, column=2, padx=20, pady=20)
    Button(wrapbtns, text='Manual', cursor='hand2', background=BASEBG, fg='#0F0', border=0, command=lambda:replace(on_frame, buildf_manual_1)).grid(row=1, column=0, padx=20, pady=20)
    Button(wrapbtns, text='CIC y análisis', cursor='hand2', background=BASEBG, fg='#0F0', border=0, command=lambda:replace(on_frame, buildf_manual_2)).grid(row=1, column=1, padx=20, pady=20)
    Button(wrapbtns, text='Actualización', cursor='hand2', background=BASEBG, fg='#0F0', border=0, command=lambda:replace(on_frame, buildf_actualizacion)).grid(row=1, column=2, padx=20, pady=20)

buildf_home()

def caps(event):
    for item in range(len(stack_inputs)):
        incaps = stack_inputs[item].get().upper()
        stack_inputs[item].delete(0, END)
        stack_inputs[item].insert(0, incaps)

def cls(stack_inputs):
    for field in range(len(stack_inputs)):
        stack_inputs[field].delete(0, END)
    stack_inputs[0].focus_set()

def buildf_multimoney():
    # Required.
    global on_frame, frame2, content2
    frame2 = Frame(root)
    on_frame = frame2
    content2 = Frame(frame2)
    init(frame2, 'MULTIMONEY', 'DATOS DEL CLIENTE', content2)

    # Personalized / By calling "init" no rows are consumed.
    # Readout/data entry panel.
    global document, serial, snames, slnames
    incolumns_1 = Frame(content2, background=BASEBG)
    incolumns_1.columnconfigure([0,1], weight=1)
    a = StringVar(); b = StringVar(); c = StringVar(); d = StringVar()

    # row=0 << content2
    incolumns_1.grid(row=0, column=0, sticky='ew')

    # row=0 << incolumns_1
    document = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=a)
    document.grid(row=0, column=0, sticky='ew', padx=(150,30), pady=(50,0), ipady=3)
    document.bind('<KeyRelease>', caps)
    serial = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=b)
    serial.grid(row=0, column=1, sticky='ew', padx=(30,150), pady=(50,0), ipady=3)
    serial.bind('<KeyRelease>', caps)

    def clear_document(event):
        global record_document
        if document.get() != '': record_document = document.get()
        document.delete(0, END)
        document.focus_set()

    def undo_clear_document(event):
        document.delete(0, END)
        document.insert(0, record_document)
        document.focus_set()

    def clear_serial(event):
        global record_serial
        if serial.get() != '': record_serial = serial.get()
        serial.delete(0, END)

    def undo_clear_serial(event):
        serial.delete(0, END)
        serial.insert(0, record_serial)

    def clear_name(event):
        global record_snames
        if snames.get() != '': record_snames = snames.get()
        snames.delete(0, END)
        snames.focus_set()

    def undo_clear_snames(event):
        snames.delete(0, END)
        snames.insert(0, record_snames)
        snames.focus_set()

    def clear_sname(event):
        global record_slnames
        if slnames.get() != '': record_slnames = slnames.get()
        slnames.delete(0, END)
        slnames.focus_set()

    def undo_clear_slnames(event):
        slnames.delete(0, END)
        slnames.insert(0, record_slnames)
        slnames.focus_set()

    # row=1 << incolumns_1
    label_document = Label(incolumns_1, text='PAGARÉ', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_document.grid(row=1, column=0, sticky='ew', padx=(150,30), pady=(5,30))
    label_document.bind('<Button-1>', clear_document)
    label_document.bind('<Button-3>', undo_clear_document)
    label_serial = Label(incolumns_1, text='IDENTIFICACIÓN', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_serial.grid(row=1, column=1, sticky='ew', padx=(30,150), pady=(5,30))
    label_serial.bind('<Button-1>', clear_serial)
    label_serial.bind('<Button-3>', undo_clear_serial)

    # row=2 << incolumns_1
    snames = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=c)
    snames.grid(row=2, column=0, sticky='ew', padx=(150,30), ipady=3)
    snames.bind('<KeyRelease>', caps)
    slnames = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=d)
    slnames.grid(row=2, column=1, sticky='ew', padx=(30,150), ipady=3)
    slnames.bind('<KeyRelease>', caps)

    # row=3 << incolumns_1
    label_name = Label(incolumns_1, text='NOMBRE(S)', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_name.grid(row=3, column=0, sticky='ew', padx=(150,30), pady=(5,30))
    label_name.bind('<Button-1>', clear_name)
    label_name.bind('<Button-3>', undo_clear_snames)
    label_sname = Label(incolumns_1, text='APELLIDO(S)', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_sname.grid(row=3, column=1, sticky='ew', padx=(30,150), pady=(5,30))
    label_sname.bind('<Button-1>', clear_sname)
    label_sname.bind('<Button-3>', undo_clear_slnames)

    incolumns_2 = Frame(content2, background=BASEBG)
    incolumns_2.columnconfigure(0, weight=1)

    # row=1 << content2
    incolumns_2.grid(row=1, column=0)

    # row=0 << incolumns_2
    global clear, pull, push, auto, lastfolder
    clear = Button(incolumns_2, text='LIMPIAR CAMPOS', background=BASEBG, fg='#74FF92', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2')
    clear.grid(row=0, column=0, padx=20, pady=(0,30))
    pull = Button(incolumns_2, text='ABRIR', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2')
    pull.grid(row=0, column=1, padx=20, pady=(0,30))
    push = Button(incolumns_2, text='PROCESAR', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    push.grid(row=0, column=2, padx=20, pady=(0,30))
    auto = Button(incolumns_2, text='AUTO', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2')
    auto.grid(row=0, column=3, padx=20, pady=(0,30))
    document.focus_set()

    # row=2 << content2
    lastfolder = Button(content2, text='N/A', background=BASEBG, fg='#74FF92', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    lastfolder.grid(row=2, column=0, pady=(50,0))

    # row=3 << content2
    Label(content2, text='Última carpeta procesada', background=BASEBG, fg='#797979', font=(FONT, 11)).grid(row=3, column=0, pady=(0,50))

    # Workspace data packing.
    global stack_inputs, stack_buttons
    stack_inputs = [document, serial, snames, slnames]
    stack_buttons = [pull, push, auto, clear, lastfolder]

    def mm_pullpush():
        pull.invoke()
        push.invoke()

    # Command center.
    pull.config(command=lambda:mm_pullrequest(stack_inputs, push))
    push.config(command=lambda:mm_pushrequest(stack_inputs, push, lastfolder))
    auto.config(command=lambda:mm_pullpush())
    clear.config(command=lambda:cls(stack_inputs))

def buildf_comercios():
    # Required.
    global on_frame, frame3, content3
    frame3 = Frame(root)
    on_frame = frame3
    content3 = Frame(frame3)
    init(frame3, 'COMERCIOS', 'DATOS DEL CLIENTE', content3)

    # Personalized / By calling "init" no rows are consumed.
    global document, serial, snames, slnames
    incolumns_1 = Frame(content3, background=BASEBG)
    incolumns_1.columnconfigure([0,1], weight=1)
    a = StringVar(); b = StringVar(); c = StringVar(); d = StringVar()

    # row=0 << content3
    incolumns_1.grid(row=0, column=0, sticky='ew')

    # row=0 << incolumns_1
    document = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=a)
    document.grid(row=0, column=0, sticky='ew', padx=(150,30), pady=(50,0), ipady=3)
    document.bind('<KeyRelease>', caps)
    serial = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=b)
    serial.grid(row=0, column=1, sticky='ew', padx=(30,150), pady=(50,0), ipady=3)
    serial.bind('<KeyRelease>', caps)

    def clear_document(event):
        global record_document
        if document.get() != '': record_document = document.get()
        document.delete(0, END)
        document.focus_set()

    def undo_clear_document(event):
        document.delete(0, END)
        document.insert(0, record_document)
        document.focus_set()

    def clear_serial(event):
        global record_serial
        if serial.get() != '': record_serial = serial.get()
        serial.delete(0, END)
        serial.focus_set()

    def undo_clear_serial(event):
        serial.delete(0, END)
        serial.insert(0, record_serial)
        serial.focus_set()

    def clear_name(event):
        global record_snames
        if snames.get() != '': record_snames = snames.get()
        snames.delete(0, END)
        snames.focus_set()

    def undo_clear_snames(event):
        snames.delete(0, END)
        snames.insert(0, record_snames)
        snames.focus_set()

    def clear_sname(event):
        global record_slnames
        if slnames.get() != '': record_slnames = slnames.get()
        slnames.delete(0, END)
        slnames.focus_set()

    def undo_clear_slnames(event):
        slnames.delete(0, END)
        slnames.insert(0, record_slnames)
        slnames.focus_set()

    # row=1 << incolumns_1
    label_document = Label(incolumns_1, text='PAGARÉ', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_document.grid(row=1, column=0, sticky='ew', padx=(150,30), pady=(5,30))
    label_document.bind('<Button-1>', clear_document)
    label_document.bind('<Button-3>', undo_clear_document)
    label_serial = Label(incolumns_1, text='IDENTIFICACIÓN', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_serial.grid(row=1, column=1, sticky='ew', padx=(30,150), pady=(5,30))
    label_serial.bind('<Button-1>', clear_serial)
    label_serial.bind('<Button-3>', undo_clear_serial)

    # row=2 << incolumns_1
    snames = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=c)
    snames.grid(row=2, column=0, sticky='ew', padx=(150,30), ipady=3)
    snames.bind('<KeyRelease>', caps)
    slnames = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=d)
    slnames.grid(row=2, column=1, sticky='ew', padx=(30,150), ipady=3)
    slnames.bind('<KeyRelease>', caps)

    # row=3 << incolumns_1
    label_name = Label(incolumns_1, text='NOMBRE(S)', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_name.grid(row=3, column=0, sticky='ew', padx=(150,30), pady=(5,30))
    label_name.bind('<Button-1>', clear_name)
    label_name.bind('<Button-3>', undo_clear_snames)
    label_sname = Label(incolumns_1, text='APELLIDO(S)', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_sname.grid(row=3, column=1, sticky='ew', padx=(30,150), pady=(5,30))
    label_sname.bind('<Button-1>', clear_sname)
    label_sname.bind('<Button-3>', undo_clear_slnames)

    incolumns_2 = Frame(content3, background=BASEBG)
    incolumns_2.columnconfigure(0, weight=1)

    # row=1 << content3
    incolumns_2.grid(row=1, column=0)

    # row=0 << incolumns_2
    global clear, pull, push, auto, lastfolder
    clear = Button(incolumns_2, text='LIMPIAR CAMPOS', background=BASEBG, fg='#74FF92', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2')
    clear.grid(row=0, column=0, padx=20, pady=(0,30))
    pull = Button(incolumns_2, text='ABRIR', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2')
    pull.grid(row=0, column=1, padx=20, pady=(0,30))
    push = Button(incolumns_2, text='PROCESAR', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    push.grid(row=0, column=2, padx=20, pady=(0,30))
    auto = Button(incolumns_2, text='AUTO', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2')
    auto.grid(row=0, column=3, padx=20, pady=(0,30))
    document.focus_set()

    # row=2 << content3
    lastfolder = Button(content3, text='N/A', background=BASEBG, fg='#74FF92', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    lastfolder.grid(row=2, column=0, pady=(50,0))

    # row=3 << content3
    Label(content3, text='Última carpeta procesada', background=BASEBG, fg='#797979', font=(FONT, 11)).grid(row=3, column=0, pady=(0,50))

    # Workspace data packing.
    global stack_inputs, stack_buttons
    stack_inputs = [document, serial, snames, slnames]

    def com_pullpush():
        pull.invoke()
        push.invoke()

    # Command center.
    pull.config(command=lambda:com_pullrequest(stack_inputs, push))
    push.config(command=lambda:com_pushrequest(stack_inputs, push, lastfolder))
    auto.config(command=com_pullpush)
    clear.config(command=lambda:cls(stack_inputs))

def buildf_beto():
    # Required.
    global on_frame, frame4, content4
    frame4 = Frame(root)
    on_frame = frame4
    content4 = Frame(frame4)
    init(frame4, 'BETO', 'DATOS DEL CLIENTE', content4)

    # Personalized / By calling "init" no rows are consumed.
    Label(content4, text='Capa en proceso de elaboración.', background=BASEBG, fg='#FFF').grid(row=0, column=0, sticky='ew', pady=20)

def buildf_manual_1():
    # Required.
    global on_frame, frame5, content5
    frame5 = Frame(root)
    on_frame = frame5
    content5 = Frame(frame5)
    init(frame5, 'MANUAL', 'DATOS DEL CLIENTE', content5)

    # Personalized / By calling "init" no rows are consumed.
    Label(content5, text='Capa en proceso de elaboración.', background=BASEBG, fg='#FFF').grid(row=0, column=0, sticky='ew', pady=20)
    global document, serial, snames, slnames, kycdated
    incolumns_1 = Frame(content5, background=BASEBG)
    incolumns_1.columnconfigure([0,1,2], weight=1)
    a = StringVar(); b = StringVar(); c = StringVar(); d = StringVar(); e = StringVar()

    # row=0 << content5
    incolumns_1.grid(row=0, column=0, sticky='ew')

    def clear_document(event):
        global record_document
        if document.get() != '': record_document = document.get()
        document.delete(0, END)
        document.focus_set()

    def undo_clear_document(event):
        document.delete(0, END)
        try: document.insert(0, record_document)
        except: pass
        document.focus_set()

    def clear_kycdated(event):
        global record_kycdated
        if kycdated.get() != '': record_kycdated = kycdated.get()
        kycdated.delete(0, END)
        kycdated.focus_set()

    def undo_clear_kycdated(event):
        kycdated.delete(0, END)
        try: kycdated.insert(0, record_kycdated)
        except: pass
        kycdated.focus_set()

    def clear_serial(event):
        global record_serial
        if serial.get() != '': record_serial = serial.get()
        serial.delete(0, END)

    def undo_clear_serial(event):
        serial.delete(0, END)
        try: serial.insert(0, record_serial)
        except: pass
        serial.focus_set()

    def clear_name(event):
        global record_snames
        if snames.get() != '': record_snames = snames.get()
        snames.delete(0, END)
        snames.focus_set()

    def undo_clear_snames(event):
        snames.delete(0, END)
        try: snames.insert(0, record_snames)
        except: pass
        snames.focus_set()

    def clear_sname(event):
        global record_slnames
        if slnames.get() != '': record_slnames = slnames.get()
        slnames.delete(0, END)
        slnames.focus_set()

    def undo_clear_slnames(event):
        slnames.delete(0, END)
        try: slnames.insert(0, record_slnames)
        except: pass
        slnames.focus_set()

    def key13(event):
        x = str(document.get())
        document.delete(0, END)
        x = x.strip()
        document.insert(0, x)
        if event.char == '\r' and event.keycode == 13:
            searchx = document.get()
            stack_inputs[1].delete(0, END)
            stack_inputs[2].delete(0, END)
            stack_inputs[3].delete(0, END)
            stack_inputs[4].delete(0, END)
            dbpath = os.getcwd()
            dbpath = dbpath.replace('\\','/')
            dbpath = f'{dbpath}/lista_de_expedientes.db'
            con = sqlite3.connect(dbpath)
            cur = con.cursor()
            try:
                req = cur.execute(f'SELECT * FROM excel_load WHERE PAGARE="{searchx}"')
                res = req.fetchone()
                if res != None:
                    data = list(res)
                    stack_inputs[2].insert(0, data[1])
                    data = data[2].upper().split(' ')
                    if len(data) == 2:
                        stack_inputs[3].insert(0, data[0])
                        stack_inputs[4].insert(0, data[1])
                    if len(data) == 3:
                        stack_inputs[3].insert(0, data[0])
                        stack_inputs[4].insert(0, f'{data[1]} {data[2]}')
                    elif len(data) == 4:
                        stack_inputs[3].insert(0, f'{data[0]} {data[1]}')
                        stack_inputs[4].insert(0, f'{data[2]} {data[3]}')
                    elif len(data) > 4:
                        stack_inputs[3].insert(0, data[:-2])
                        stack_inputs[4].insert(0, f'{data[-2]} {data[-1]}')
            except: pass
            con.close()

    # row=0 << incolumns_1
    document = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=a)
    document.grid(row=0, column=0, sticky='ew', padx=(150,30), pady=(50,0), ipady=3)
    document.bind('<KeyRelease>', key13)
    kycdated = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=e)
    kycdated.grid(row=0, column=1, sticky='ew', padx=30, pady=(50,0), ipady=3)
    kycdated.bind('<KeyRelease>', caps)

    # row=1 << incolumns_1
    label_document = Label(incolumns_1, text='PAGARÉ', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_document.grid(row=1, column=0, sticky='ew', padx=(150,30), pady=(5,30))
    label_document.bind('<Button-1>', clear_document)
    label_document.bind('<Button-3>', undo_clear_document)

    label_kycdated = Label(incolumns_1, text='FECHA EN KYC #2', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_kycdated.grid(row=1, column=1, sticky='ew', padx=30, pady=(5,30))
    label_kycdated.bind('<Button-1>', clear_kycdated)
    label_kycdated.bind('<Button-3>', undo_clear_kycdated)

    # row=2 << incolumns_1
    serial = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=b)
    serial.grid(row=2, column=0, sticky='ew', padx=(150,30), ipady=3)
    serial.bind('<KeyRelease>', caps)
    snames = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=c)
    snames.grid(row=2, column=1, sticky='ew', padx=30, ipady=3)
    snames.bind('<KeyRelease>', caps)
    slnames = Entry(incolumns_1, fg='#090', font=(FONT, 14), selectbackground='#080', justify="center", textvariable=d)
    slnames.grid(row=2, column=2, sticky='ew', padx=(30,150), ipady=3)
    slnames.bind('<KeyRelease>', caps)

    # row=3 << incolumns_1
    label_serial = Label(incolumns_1, text='IDENTIFICACIÓN', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_serial.grid(row=3, column=0, sticky='ew', padx=(150,30), pady=(5,30))
    label_serial.bind('<Button-1>', clear_serial)
    label_serial.bind('<Button-3>', undo_clear_serial)
    label_name = Label(incolumns_1, text='NOMBRE(S)', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_name.grid(row=3, column=1, sticky='ew', padx=30, pady=(5,30))
    label_name.bind('<Button-1>', clear_name)
    label_name.bind('<Button-3>', undo_clear_snames)
    label_sname = Label(incolumns_1, text='APELLIDO(S)', background=BASEBG, fg='#797979', font=(FONT, 12), cursor='hand2')
    label_sname.grid(row=3, column=2, sticky='ew', padx=(30,150), pady=(5,30))
    label_sname.bind('<Button-1>', clear_sname)
    label_sname.bind('<Button-3>', undo_clear_slnames)

    incolumns_2 = Frame(content5, background=BASEBG)
    incolumns_2.columnconfigure(0, weight=1)

    # row=1 << content5
    incolumns_2.grid(row=1, column=0)

    # row=0 << incolumns_2
    global savein, pull, readf, push, auto, lastfolder
    savein = Button(incolumns_2, text='GUARDAR EN', background=BASEBG, fg='#FFFF00', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2')
    savein.grid(row=0, column=0, padx=20, pady=(0,30))
    pull = Button(incolumns_2, text='BUSCAR EN', background=BASEBG, fg='#FF0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    pull.grid(row=0, column=1, padx=20, pady=(0,30))
    readf = Button(incolumns_2, text='LEER', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    readf.grid(row=0, column=2, padx=20, pady=(0,30))
    push = Button(incolumns_2, text='PROCESAR', background=BASEBG, fg='#0F0', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    push.grid(row=0, column=3, padx=20, pady=(0,30))
    document.focus_set()

    # row=2 << content5
    lastfolder = Button(content5, text='N/A', background=BASEBG, fg='#74FF92', font=(FONT, 11), activebackground=BASEBG, activeforeground='#FFF', border=0, cursor='hand2', state='disabled')
    lastfolder.grid(row=2, column=0, pady=(50,0))

    # row=3 << content5
    Label(content5, text='Última carpeta procesada', background=BASEBG, fg='#797979', font=(FONT, 11)).grid(row=3, column=0, pady=(0,50))

    # Workspace data packing.
    global stack_inputs, stack_buttons
    stack_inputs = [document, kycdated, serial, snames, slnames]
    stack_buttons = [savein, pull, readf, push, lastfolder]

    def pullpush():
        pull.invoke()
        push.invoke()

    # Command center.
    savein.config(command=lambda:set_savingdirectory(stack_buttons))
    pull.config(command=lambda:man_pullrequest(stack_inputs, stack_buttons))
    readf.config(command=lambda:man_readf(stack_inputs, stack_buttons))
    push.config(command=lambda:man_pushrequest(stack_inputs, stack_buttons))

def buildf_manual_2():
    # Required.
    global on_frame, frame6, content6
    frame6 = Frame(root)
    on_frame = frame6
    content6 = Frame(frame6)
    init(frame6, 'ANÁLISIS Y CIC', 'DATOS DEL CLIENTE', content6)

    # Personalized / By calling "init" no rows are consumed.
    Label(content6, text='Capa en proceso de elaboración.', background=BASEBG, fg='#FFF').grid(row=0, column=0, sticky='ew', pady=20)

def buildf_actualizacion():
    # Required.
    global on_frame, frame7, content7
    frame7 = Frame(root)
    on_frame = frame7
    content7 = Frame(frame7)
    init(frame7, 'ACTUALIZACIÓN', 'DATOS DEL CLIENTE', content7)

    # Personalized / By calling "init" no rows are consumed.
    Label(content7, text='Capa en proceso de elaboración.', background=BASEBG, fg='#FFF').grid(row=0, column=0, sticky='ew', pady=20)

def showshorcuts():
    messagebox.showinfo('Accesos directos por teclado', 'F1: Abrir\nF2: Procesar\nF3: Auto\nF4: Última carpeta procesada\nF5: Limpiar campos\n\nF6: Accesos directos de teclado')

def remspaces():
    global rs
    rs = stack_inputs[0].get().upper()
    rs = rs.replace(' ', '').replace('\n', '')
    stack_inputs[0].delete(0, END)
    stack_inputs[0].insert(0, rs)

global slqdata
sqldata = []

def getxcldata():
    remspaces()
    dbpath = os.getcwd()
    dbpath = dbpath.replace('\\','/')
    dbpath = f'{dbpath}/lista_de_expedientes.db'
    verify_bd = os.path.exists(dbpath)
    if verify_bd:
        con = sqlite3.connect(dbpath)
        cur = con.cursor()
        req = cur.execute('SELECT * FROM excel_load')
        res = req.fetchall()
        for r in res:
            r = list(r)
            sqldata.append(list(r))
    con.close()

def cargar_datos_excel():
    excel_doc = filedialog.askopenfile(mode='r', defaultextension=['xlxs'])
    xcpath = excel_doc.name
    wb = openpyxl.load_workbook(xcpath)
    sheet = wb.active
    maxrow = sheet.max_row
    rc = 2
    dbpath = os.getcwd()
    dbpath = dbpath.replace('\\','/')
    dbpath = f'{dbpath}/lista_de_expedientes.db'
    con = sqlite3.connect(dbpath)
    cur = con.cursor()
    try:
        cur.execute('''
            CREATE TABLE excel_load(
                PAGARE VARCHAR(12),
                CEDULA VARCHAR(20),
                NOMBRE VARCHAR(100)
            )
        ''')
    except Exception as e: print(e)
    getxcldata()
    craftedtext = []
    for i in range(maxrow-1):
        craftedtext.append([str(sheet.cell(row=rc, column=1).value), str(sheet.cell(row=rc, column=2).value), sheet.cell(row=rc, column=3).value])
        ctxt = craftedtext[i]
        cur.execute(f'INSERT INTO excel_load VALUES ("{ctxt[0]}", "{ctxt[1]}", "{ctxt[2]}")')
        rc += 1
    excel_doc.close()
    con.commit()
    con.close()
    messagebox.showinfo('Procesador de documentos', f'El archivo\n\n"{xcpath}"\n\nHa sido cargado correctamente.')

def menu():
    global menubar, mfile, mbrand, mdata, mhelp
    mfile = Menu(menubar, tearoff=0)
    mfile.add_command(label='Cerrar sesión', state='disabled')
    mfile.add_command(label='Salir')
    mbrand = Menu(menubar, tearoff=0)
    mbrand.add_command(label='Inicio', command=lambda:replace(on_frame, buildf_home))
    mbrand.add_command(label='Multimoney', command=lambda:replace(on_frame, buildf_multimoney))
    mbrand.add_command(label='Comercios', command=lambda:replace(on_frame, buildf_comercios))
    mbrand.add_command(label='Beto', command=lambda:replace(on_frame, buildf_beto))
    mbrand_sub = Menu(mbrand, tearoff=0)
    mbrand_sub.add_command(label='Ordenar documentos', command=lambda:replace(on_frame, buildf_manual_1))
    mbrand_sub.add_command(label='Análisis y CIC', command=lambda:replace(on_frame, buildf_manual_2))
    mbrand.add_cascade(label='Manual', menu=mbrand_sub)
    mbrand.add_command(label='Actualización', command=lambda:replace(on_frame, buildf_actualizacion))
    mdata = Menu(menubar, tearoff=0)
    mdata.add_command(label='Cargar datos desde Excel', command=cargar_datos_excel)
    mhelp = Menu(menubar, tearoff=0)
    mhelp.add_command(label='Teclado', command=showshorcuts)
    mhelp.add_command(label='Documentación', state='disabled')
    mhelp.add_command(label='GitHub', state='disabled')
    menubar.add_cascade(label='Archivo', menu=mfile, state='disabled')
    menubar.add_cascade(label='Aplicación', menu=mbrand)
    menubar.add_cascade(label='Datos', menu=mdata)
    menubar.add_cascade(label='Ayuda', menu=mhelp)

menu()

def kclisteners(event):
    if event.keycode == 112: pull.invoke()
    elif event.keycode == 113: push.invoke()
    elif event.keycode == 114: auto.invoke()
    elif event.keycode == 115: clear.invoke()
    elif event.keycode == 116: lastfolder.invoke()
    elif event.keycode == 117: showshorcuts()

def kclisteners_2(event):
    if event.keycode == 112: savein.invoke()
    elif event.keycode == 113: push.invoke()
    elif event.keycode == 114: readf.invoke()
    elif event.keycode == 115: pull.invoke()
    elif event.keycode == 116: lastfolder.invoke()
    elif event.keycode == 117: showshorcuts()

def keyboardlistener(event):
    try:
        if on_frame == frame2:kclisteners(event)
    except: pass
    try:
        if on_frame == frame3: kclisteners(event)
    except: pass
    try:
        if on_frame == frame5: kclisteners_2(event)
    except: pass

# replace(frame1, buildf_comercios)

root.bind('<KeyPress>', keyboardlistener)
root.mainloop()