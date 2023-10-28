import os, shutil
from tkinter import filedialog as fd
from tkinter import messagebox as mb
from tkinter import END
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from PIL import Image

def folderscan():
    global scanner
    scanner = os.listdir(workingdirectory)

def namesplit(r_fname):
    global r_name, r_lname
    r_name = ''
    r_lname = ''
    r_fname = r_fname.split(' ')
    for useless in r_fname:
        if useless == '' or useless == ' ': r_fname.remove(useless)
    if len(r_fname) == 3:
        r_name = r_fname[0]
        r_lname = f'{r_fname[1]} {r_fname[2]}'
    elif len(r_fname) == 4:
        r_name = f'{r_fname[0]} {r_fname[1]}'
        r_lname = f'{r_fname[2]} {r_fname[3]}'
    elif len(r_fname) > 4:
        r_lname = f'{r_fname[-2]} {r_fname[-1]}'
        r_fname.pop()
        r_fname.pop()
        r_name = ' '.join(r_fname)

def mm_pullrequest(stack_inputs, push):
    global workingdirectory
    ifnocache = os.getcwd()
    try:
        cachedirectory = workingdirectory.split('/')
        cachedirectory.pop()
        cachedirectory.pop()
        cachedirectory = '/'.join(cachedirectory)
        workingdirectory = fd.askdirectory(title='Buscar carpeta', initialdir=cachedirectory)
    except: workingdirectory = fd.askdirectory(title='Buscar carpeta', initialdir=ifnocache)
    if workingdirectory != '':
        workingdirectory += '/'
        for field in range(len(stack_inputs)):
            stack_inputs[field].delete(0, END)
        folderscan()
        for s in scanner:
            ss = s.upper()
            if ss.__contains__('AFFIDAVIT') or ss.__contains__('AFFIDÁVIT'): source = s
        source = open(f'{workingdirectory}{source}', 'rb')
        _reader = PdfReader(source)
        _page = _reader.pages[1].extract_text().split('\n')
        global _docsize
        _docsize = len(_reader.pages)
        global r_document, r_id, r_fname
        requiredlines = (f'{_page[3]} {_page[6]} {_page[7]}')
        requiredlines = requiredlines.split(' ')
        r_document = requiredlines[0]
        requiredlines = requiredlines[2:]
        requiredlines = ' '.join(requiredlines)
        requiredlines = requiredlines.split(',')
        r_fname = requiredlines[0]
        requiredlines = requiredlines[1:]
        requiredlines = ' '.join(requiredlines)
        r_id = ''
        for char in requiredlines:
            if char.isnumeric(): r_id += char
        namesplit(r_fname)
        source.close()
        stack_inputs[0].insert(0, r_document.upper())
        stack_inputs[1].insert(0, r_id.upper())
        stack_inputs[2].insert(0, r_name.upper())
        stack_inputs[3].insert(0, r_lname.upper())
        push.config(state='normal')
    else: mb.showinfo('Procesador de documentos', 'Seleccione una carpeta a procesar para empezar.')

def startf(lastfolder):
    try: os.startfile(startinfo)
    except:
        lastfolder.config(text='N/A', state='disabled')
        mb.showerror('Procesador de documentos', f'El archivo "{summary}" no se puede abrir ya que ha sido movido o eliminado.')

def createpdf(subfolder, titleit, pages):
    folderscan()
    titleit = f'{titleit} {render}'
    global aff
    for s in scanner:
        _s_ = s.upper().replace('PDF', 'pdf').replace('PNG', 'png').replace('JPG', 'jpg').replace('JPEG', 'jpeg')
        if _s_.__contains__('AFF'): aff = _s_
    source = f'{workingdirectory}{aff}'
    pdf = PdfReader(source)
    _writer = PdfWriter()
    for p in pages:
        _writer.add_page(pdf.pages[p])
    with open(f'{workingdirectory}{subfolder}{titleit}.pdf', 'wb') as f:
        _writer.write(f)
        f.close()

def onlypdfs():
    for s in scanner:
        _s_ = s.upper().replace('PNG', 'png').replace('JPG', 'jpg').replace('JPEG', 'jpeg')
        if _s_.__contains__('.png') or _s_.__contains__('.jpg') or _s_.__contains__('.jpeg'):
            img = Image.open(f'{workingdirectory}{s}')
            img = img.convert('RGB')
            onlyfilename = s.replace('.png', '').replace('.jpg', '').replace('.jpeg', '')
            img.save(f'{workingdirectory}{onlyfilename}.pdf')

def _folder0():
    global folder0
    folder0 = '0. OTROS DOCUMENTOS/'
    os.makedirs(f'{workingdirectory}{folder0}')
    folderscan()
    moving = f'{workingdirectory}{folder0}'
    for s in scanner:
        _s_ = s.upper().replace('PDF', 'pdf').replace('PNG', 'png').replace('JPG', 'jpg').replace('JPEG', 'jpeg')
        if _s_.__contains__('AFF'): shutil.move(f'{workingdirectory}{s}', f'{moving}KIT {render}.pdf')
        elif _s_.__contains__('FIRMA'): shutil.move(f'{workingdirectory}{s}', f'{moving}KIT FIRMA REP LEGAL {render}.pdf')
        elif _s_.__contains__('ID1.pdf') or _s_.__contains__('ID2.pdf'):
            os.remove(f'{workingdirectory}{s}')
        elif _s_.__contains__('ID1'):
            ss = s.split('.')
            shutil.move(f'{workingdirectory}{s}', f'{moving}ID CARA 1 {render}.{ss[1]}')
        elif _s_.__contains__('ID2'):
            ss = s.split('.')
            shutil.move(f'{workingdirectory}{s}', f'{moving}ID CARA 2 {render}.{ss[1]}')

def _folder1():
    global folder1
    folder1 = '1. INFORMACIÓN GENERAL/'
    os.makedirs(f'{workingdirectory}{folder1}')
    if _docsize == 26:
        createpdf(folder1, 'AUTORIZACIÓN CIC', [22,24,25])
        createpdf(folder1, 'CICAC', [23,24,25])
        createpdf(folder1, 'CONSENTIMIENTO INFORMADO', [20,24,25])
        createpdf(folder1, 'KYC', [18,19,24,25])
    elif _docsize == 33:
        createpdf(folder1, 'AUTORIZACIÓN CIC', [22,31,32])
        createpdf(folder1, 'CICAC', [23,31,32])
        createpdf(folder1, 'CONSENTIMIENTO INFORMADO', [20,31,32])
        createpdf(folder1, 'CONSENTIMIENTO INFORMADO SMART', [24,25,31,32])
        createpdf(folder1, 'KYC', [18,19,31,32])
        createpdf(folder1, 'KYC SMART', [26,31,32])
    _merge = PdfMerger()
    for s in scanner:
        _s_ = s.upper().replace('.PDF', '.pdf')
        if _s_.__contains__('.pdf'):
            if _s_.__contains__('ID1') or _s_.__contains__('ID2'): _merge.append(f'{workingdirectory}{s}')
    outputidname = f'{workingdirectory}{folder1}ID {render}.pdf'
    with open(outputidname, 'wb') as f:
        _merge.write(f)
        _merge.close()
        f.close()

def _folder2():
    global folder2
    folder2 = '2. APROBACIONES CREDITICIAS/'
    os.makedirs(f'{workingdirectory}{folder2}')
    if _docsize == 26:
        createpdf(folder2, 'CONTRATO', [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,24,25])
        createpdf(folder2, 'PAGARÉ', [17,24,25])
    elif _docsize == 33:
        createpdf(folder2, 'CONTRATO', [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,31,32])
        createpdf(folder2, 'CONTRATO SMART', [27,28,29,30,31,32])
        createpdf(folder2, 'PAGARÉ', [17,31,32])

def _folder3():
    global folder3
    folder3 = '3. INFORMACIÓN PARA ANÁLISIS CAPACIDAD/'
    os.makedirs(f'{workingdirectory}{folder3}')
    if _docsize == 26:
        createpdf(folder3, 'DECLARACIÓN JURADA', [21,24,25])
    elif _docsize == 33:
        createpdf(folder3, 'DECLARACIÓN JURADA', [21,31,32])
    finaltitle = ''
    for s in scanner:
        _s_ = s.upper().replace('.PDF', '').replace('.PNG', '').replace('.JPG', '').replace('.JPEG', '')
        if s.__contains__('pdf') and _s_.__contains__('ORDEN'):
            if s.__contains__('n.pdf') or s.__contains__('N.pdf'): finaltitle = f'ORDEN PATRONAL 1 {render}'
            else: nn = _s_.split('ORDEN'); finaltitle = f'ORDEN PATRONAL {nn[1]} {render}'
            _from = f'{workingdirectory}{s}'
            _to = f'{workingdirectory}{folder3}{finaltitle}.pdf'
            shutil.move(_from, _to)
        elif s.__contains__('pdf') and _s_.__contains__('ORIGEN'):
            if s.__contains__('n.pdf') or s.__contains__('N.pdf'): finaltitle = f'ORIGEN DE INGRESOS 1 {render}'
            else: nn = _s_.split('ORIGEN'); finaltitle = f'ORIGEN DE INGRESOS {nn[1]} {render}'
            _from = f'{workingdirectory}{s}'
            _to = f'{workingdirectory}{folder3}{finaltitle}.pdf'
            shutil.move(_from, _to)
        elif _s_.__contains__('BURO') or _s_.__contains__('BURÓ'):
            _from = f'{workingdirectory}{s}'
            _to = f'{workingdirectory}{folder3}{finaltitle}.pdf'
            shutil.move(f'{workingdirectory}{s}', f'{workingdirectory}{folder3}BURÓ {render}.pdf')

def _folder4():
    global folder4
    folder4 = '4. RESULTADOS DE ANÁLISIS'
    os.makedirs(f'{workingdirectory}{folder4}')

def mm_pushrequest(stack_inputs, push, lastfolder):
    global blank
    blank = False
    for s in stack_inputs:
        if s.get() == '' or s.get() == ' ':
            blank = True
            break
    if blank == False:
        global swallow_inputs, render
        swallow_inputs = []
        for s in stack_inputs:
            swallow_inputs.append(s.get().upper())
            s.delete(0, END)
        render = f'{swallow_inputs[1]} {swallow_inputs[3]} {swallow_inputs[2]} {swallow_inputs[0]}'
        onlypdfs()
        _folder4()
        _folder3()
        _folder2()
        _folder1()
        _folder0()
        folderscan()
        for s in scanner:
            if s.__contains__('.png') or s.__contains__('.jpg') or s.__contains__('.jpeg'): os.remove(f'{workingdirectory}{s}')
        push.config(state='disabled')
        global summary
        summary = f'{swallow_inputs[1]} {swallow_inputs[3]} {swallow_inputs[2]} {swallow_inputs[0]}'
        dir = workingdirectory.split('/')
        dir.pop()
        parent = dir.pop()
        dir = '/'.join(dir)
        dir += '/'
        os.rename(f'{dir}{parent}', f'{dir}{render}')
        global startinfo
        startinfo = f'{dir}{render}'
        lastfolder.config(state='normal', text=summary, command=lambda:startf(lastfolder))
    else:mb.showwarning('Procesador de documentos', 'Hay campos sin rellenar.\n\nAsegúrese de rellenar todos los campos para continuar.')