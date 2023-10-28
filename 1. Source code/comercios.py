import os, shutil
from tkinter import filedialog, messagebox, END
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from PIL import Image

def folderscan():
    global scanner
    scanner = os.listdir(workingdirectory)

def capsfilenames():
    for f in scanner:
        _f_ = f.upper().replace('.PDF', '.pdf').replace('.PNG', '.png').replace('.JPG', '.jpg').replace('.JPEG', '.jpeg')
        os.rename(f'{workingdirectory}{f}', f'{workingdirectory}{_f_}')
    folderscan()

def namesplit(r_fname):
    global r_name, r_lname
    r_name = ''
    r_lname = ''
    r_fname = r_fname.split(' ')
    normalized_data = []
    for rf in r_fname:
        if rf != '' and rf != ' ' and rf != 'y': normalized_data.append(rf)
    r_fname = normalized_data
    normalized_data = []
    if len(r_fname) == 3:
        try:
            r_name = r_fname[0]
            r_lname = f'{r_fname[1]} {r_fname[2]}'
        except Exception as e: print(e, 'len(r_fname) == 3')
    elif len(r_fname) == 4:
        try:
            r_name = f'{r_fname[0]} {r_fname[1]}'
            r_lname = f'{r_fname[2]} {r_fname[3]}'
        except Exception as e: print(e, 'len(r_fname) == 4')
    elif len(r_fname) > 4:
        try:
            r_lname = f'{r_fname[-2]} {r_fname[-1]}'
            r_name.pop()
            r_name.pop()
            r_name = ' '.join(r_fname)
        except Exception as e: print(e, 'len(r_fname) > 4')

def onlypdfs():
    for s in scanner:
        if s.__contains__('.png') or s.__contains__('.jpg') or s.__contains__('.jpeg'):
            img = Image.open(f'{workingdirectory}{s}')
            img = img.convert('RGB')
            onlyfilename = s.replace('.png', '').replace('.jpg', '').replace('.jpeg', '')
            img.save(f'{workingdirectory}{onlyfilename}.pdf')
            if s.__contains__('ID'): pass
            else: os.remove(f'{workingdirectory}{s}')

def createpdf(subfolder, titleit, pages):
    folderscan()
    titleit = f'{titleit} {render}'
    global aff
    for s in scanner:
        if s.__contains__('AFF'): aff = s
    source = f'{workingdirectory}{aff}'
    pdf = PdfReader(source)
    _writer = PdfWriter()
    for p in pages:
        _writer.add_page(pdf.pages[p])
    with open(f'{workingdirectory}{subfolder}{titleit}.pdf', 'wb') as f:
        _writer.write(f)
        f.close()

def _folder0():
    global folder0
    folder0 = '0. OTROS DOCUMENTOS/'
    os.makedirs(f'{workingdirectory}{folder0}')
    folderscan()
    moving = f'{workingdirectory}{folder0}'
    for s in scanner:
        if s.__contains__('AFF'): shutil.move(f'{workingdirectory}{s}', f'{moving}KIT {render}.pdf')
        elif s.__contains__('FIRMA'): shutil.move(f'{workingdirectory}{s}', f'{moving}KIT FIRMA REP LEGAL {render}.pdf')
        elif s.__contains__('ID1.pdf') or s.__contains__('ID2.pdf'):
            os.remove(f'{workingdirectory}{s}')
        elif s.__contains__('ID1'):
            ss = s.split('.')
            shutil.move(f'{workingdirectory}{s}', f'{moving}ID CARA 1 {render}.{ss[1]}')
        elif s.__contains__('ID2'):
            ss = s.split('.')
            shutil.move(f'{workingdirectory}{s}', f'{moving}ID CARA 2 {render}.{ss[1]}')

def _folder1():
    global folder1
    folder1 = '1. INFORMACIÓN GENERAL/'
    os.makedirs(f'{workingdirectory}{folder1}')
    createpdf(folder1, 'AUTORIZACIÓN CIC', [24,26,27])
    createpdf(folder1, 'CICAC', [25,26,27])
    createpdf(folder1, 'CONSENTIMIENTO INFORMADO', [22,26,27])
    createpdf(folder1, 'KYC', [20,21,26,27])
    _merge = PdfMerger()
    for s in scanner:
        if s.__contains__('.pdf'):
            if s.__contains__('ID1') or s.__contains__('ID2'):
                _merge.append(f'{workingdirectory}{s}')
    outputidname = f'{workingdirectory}{folder1}ID {render}.pdf'
    with open(outputidname, 'wb') as f:
        _merge.write(f)
        _merge.close()
        f.close()
    try:
        os.remove(f'{workingdirectory}ID1.pdf')
        os.remove(f'{workingdirectory}ID2.pdf')
    except Exception as e: print(e)

def _folder2():
    global folder2
    folder2 = '2. APROBACIONES CREDITICIAS/'
    os.makedirs(f'{workingdirectory}{folder2}')
    createpdf(folder2, 'CONTRATO', [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,26,27])
    createpdf(folder2, 'PAGARÉ', [19,26,27])

def _folder3():
    global folder3
    folder3 = '3. INFORMACIÓN PARA ANÁLISIS CAPACIDAD/'
    os.makedirs(f'{workingdirectory}{folder3}')
    createpdf(folder3, 'DECLARACIÓN JURADA', [23,26,27])
    finaltitle = ''
    for s in scanner:
        _s_ = s.upper().replace('.PDF', '').replace('.PNG', '').replace('.JPG', '').replace('.JPEG', '')
        if s.__contains__('pdf') and s.__contains__('ORDEN'):
            if s.__contains__('n.pdf') or s.__contains__('N.pdf'): finaltitle = f'ORDEN PATRONAL 1 {render}'
            else: nn = _s_.split('ORDEN'); finaltitle = f'ORDEN PATRONAL {nn[1]} {render}'
            _from = f'{workingdirectory}{s}'
            _to = f'{workingdirectory}{folder3}{finaltitle}.pdf'
            shutil.move(_from, _to)
        elif s.__contains__('pdf') and s.__contains__('ORIGEN'):
            if s.__contains__('n.pdf') or s.__contains__('N.pdf'): finaltitle = f'ORIGEN DE INGRESOS 1 {render}'
            else: nn = _s_.split('ORIGEN'); finaltitle = f'ORIGEN DE INGRESOS {nn[1]} {render}'
            _from = f'{workingdirectory}{s}'
            _to = f'{workingdirectory}{folder3}{finaltitle}.pdf'
            shutil.move(_from, _to)
        elif s.__contains__('BURO') or s.__contains__('BURÓ'):
            _from = f'{workingdirectory}{s}'
            _to = f'{workingdirectory}{folder3}{finaltitle}.pdf'
            shutil.move(f'{workingdirectory}{s}', f'{workingdirectory}{folder3}BURÓ {render}.pdf')

def _folder4():
    global folder4
    folder4 = '4. RESULTADOS DE ANÁLISIS'
    os.makedirs(f'{workingdirectory}{folder4}')

def startf(lastfolder, startinfo):
    try: os.startfile(startinfo)
    except:
        lastfolder.config(text='N/A', state='disabled')
        messagebox.showerror('Procesador de documentos', f'El archivo "{render}" no se puede abrir ya que ha sido movido o eliminado.')

def setupfolder(lastfolder):
    dir = workingdirectory.split('/')
    dir.pop()
    parent = dir.pop()
    dir = '/'.join(dir)
    dir += '/'
    os.rename(f'{dir}{parent}', f'{dir}{render}')
    global startinfo
    startinfo = f'{dir}{render}'
    lastfolder.config(state='normal', text=render, command=lambda:startf(lastfolder, startinfo))

def com_pullrequest(stack_inputs, push):
    global workingdirectory
    ifnocache = os.getcwd()
    try:
        cachedirectory = workingdirectory.split('/')
        cachedirectory.pop()
        cachedirectory.pop()
        cachedirectory = '/'.join(cachedirectory)
        workingdirectory = filedialog.askdirectory(title='Procesador de documentos', initialdir=cachedirectory)
    except: workingdirectory = filedialog.askdirectory(title='Procesador de documentos', initialdir=ifnocache)
    if workingdirectory != '':
        workingdirectory += '/'
        for field in range(len(stack_inputs)):
            stack_inputs[field].delete(0, END)
        folderscan()
        for s in scanner:
            if s.__contains__('AFF') or s.__contains__('Aff') or s.__contains__('aff'): source = s
        source = open(f'{workingdirectory}{source}', 'rb')
        _reader = PdfReader(source)
        _page = _reader.pages[1].extract_text().split('\n')

        global r_document, r_id, r_name, r_lname
        r_document = _page[4]
        r_fname = f'{_page[7]} {_page[8]}'.replace('\n','')
        r_fname = r_fname.split(';')
        r_fname = r_fname.pop()
        r_fname = r_fname.split(',')
        r_fname = r_fname[0]
        buildr_id = _page[8].split(',')
        buildr_id = buildr_id[1]
        r_id = ''
        for char in buildr_id:
            if char.isnumeric(): r_id += char
        namesplit(r_fname)
        source.close()
        stack_inputs[0].insert(0, r_document.upper())
        stack_inputs[1].insert(0, r_id.upper())
        stack_inputs[2].insert(0, r_name.upper())
        stack_inputs[3].insert(0, r_lname.upper())
        push.config(state='normal')
    else: messagebox.showinfo('Procesador de documentos', 'Seleccione una carpeta a procesar para empezar.')

def com_pushrequest(stack_inputs, push, lastfolder):
    capsfilenames()
    onlypdfs()
    global render
    push.config(state='disabled')
    render = f'{stack_inputs[1].get()} {stack_inputs[3].get()} {stack_inputs[2].get()} {stack_inputs[0].get()}'
    render = render.upper()
    for s in stack_inputs:
        s.delete(0, END)
    print()
    _folder4()
    _folder3()
    _folder2()
    _folder1()
    _folder0()
    setupfolder(lastfolder)
