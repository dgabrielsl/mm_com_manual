import os, shutil
from tkinter import filedialog, messagebox, END
from PyPDF2 import PdfReader, PdfMerger
from PIL import Image
from handler import *

def folderscan():
    global scanner
    scanner = os.listdir(downloadfrom)

def capsfilenames():
    for f in scanner:
        _f = f.upper().replace('.PDF', '.pdf').replace('.PNG', '.png').replace('.JPG', '.jpg').replace('.JPEG', '.jpeg')
        os.rename(f'{downloadfrom}/{f}', f'{downloadfrom}/{_f}')
    folderscan()

def switch(self):
    self[0] = self[0].upper()
    if self[0] == 'ENERO': self[0] = '01'
    if self[0] == 'FEBRERO': self[0] = '02'
    if self[0] == 'MARZO': self[0] = '03'
    if self[0] == 'ABRIL': self[0] = '04'
    if self[0] == 'MAYO': self[0] = '05'
    if self[0] == 'JUNIO': self[0] = '06'
    if self[0] == 'JULIO': self[0] = '07'
    if self[0] == 'AGOSTO': self[0] = '08'
    if self[0] == 'SEPTIEMBRE': self[0] = '09'
    if self[0] == 'OCTUBRE': self[0] = '10'
    if self[0] == 'NOVIEMBRE': self[0] = '11'
    if self[0] == 'DICIEMBRE': self[0] = '12'

def makeid():
    global idstopdf
    idstopdf = []
    idsinpdf = []
    for ids in scanner:
        if ids.__contains__('.png') or ids.__contains__('.jpg') or ids.__contains__('.jpeg'):
            if ids.__contains__('ID'): idstopdf.append(ids)
        if ids.__contains__('ID1.pdf') or ids.__contains__('ID2.pdf'):
            idsinpdf.append(ids)
    if len(idstopdf) == 2:
        for s in idstopdf:
            img = Image.open(f'{downloadfrom}/{s}')
            img = img.convert('RGB')
            onlyfilename = s.replace('.png', '').replace('.jpg', '').replace('.jpeg', '')
            img.save(f'{downloadfrom}/{onlyfilename}.pdf')
    if len(idsinpdf) == 2:
        merger = PdfMerger()
        outputname = f'{downloadfrom}/ID.pdf'
        for id in idsinpdf:
            merger.append(open(f'{downloadfrom}/{id}', 'rb'))
        with open(outputname, 'wb') as finalid:
            merger.write(finalid)
    try:
        for supr in idstopdf:
            os.remove(f'{downloadfrom}/{supr}')
    except: pass
    try:
        for supr in idsinpdf:
            os.remove(f'{downloadfrom}/{supr}')
    except: pass

def onlypdfs():
    makeid()
    folderscan()
    for s in scanner:
        if s.__contains__('.png') or s.__contains__('.jpg') or s.__contains__('.jpeg'):
            if not s.__contains__('ID'):
                img = Image.open(f'{downloadfrom}/{s}')
                img = img.convert('RGB')
                onlyfilename = s.replace('.png', '').replace('.jpg', '').replace('.jpeg', '')
                img.save(f'{downloadfrom}/{onlyfilename}.pdf')
                os.remove(f'{downloadfrom}/{s}')

def startf(lastfolder, startinfo):
    try: os.startfile(startinfo)
    except:
        lastfolder.config(text='N/A', state='disabled')
        messagebox.showerror('Procesador de documentos', f'El archivo "{render}" no se puede abrir ya que ha sido movido o eliminado.')

def setupfolder(lastfolder):
    dir = downloadfrom.split('/')
    dir.pop()
    parent = dir.pop()
    dir = '/'.join(dir)
    dir += '/'
    global startinfo
    startinfo = f'{savingdirectory}/{render}'
    lastfolder.config(state='normal', text=render, command=lambda:startf(lastfolder, startinfo))

def set_savingdirectory(stack_buttons):
    global savingdirectory
    savingdirectory = filedialog.askdirectory(title='Guardar en')
    if savingdirectory != '':
        get_sdir = savingdirectory
        get_sdir = get_sdir.split('/')
        divide = get_sdir.pop().upper()
        stack_buttons[0].config(text=f'GUARDAR EN "{divide}"', fg='#74FF92')
        stack_buttons[1].config(state='normal')
    else:
        stack_buttons[0].config(text=f'GUARDAR EN', fg='#FFFF00')
        stack_buttons[1].config(state='disabled')
        stack_buttons[2].config(state='disabled')
        stack_buttons[3].config(state='disabled')
        messagebox.showerror('Procesador de documentos', 'Debe configurar una ruta de salida para empezar.')

def set_downloadfrom(stack_buttons):
    global downloadfrom
    downloadfrom = filedialog.askdirectory(title='Buscar los documentos a procesar en')
    if downloadfrom != '':
        get_ddir = downloadfrom
        get_ddir = get_ddir.split('/')
        divide = get_ddir.pop().upper()
        stack_buttons[1].config(text=f'BUSCAR EN "{divide}"', fg='#74FF92')
        stack_buttons[2].config(state='normal')
        stack_buttons[3].config(state='normal')
    else:
        stack_buttons[2].config(state='disabled')
        stack_buttons[3].config(state='disabled')
        messagebox.showerror('Procesador de documentos', 'Debe configurar una ruta de entrada para continuar.')

def readabledocs_searching(stack_inputs):
    global readabledoc
    readabledoc = ''
    for f in scanner:
        fF = f.upper()
        if fF.__contains__('KYC2'): readabledoc = f'{downloadfrom}/{f}'
    if readabledoc == '':
        for f in scanner:
            fF = f.upper()
            if fF.__contains__('CIC2'): readabledoc = f'{downloadfrom}/{f}'; break
            elif fF.__contains__('CONSENTIMIENTO2'): readabledoc = f'{downloadfrom}/{f}'; break
            elif fF.__contains__('DECLARACION2') or fF.__contains__('DECLARACIÓN2'): readabledoc = f'{downloadfrom}/{f}'; break
            elif fF.__contains__('PAGARÉ') or fF.__contains__('PAGARE'):
                readabledoc = f'{downloadfrom}/{f}'
                global h_crafted_text
                h_crafted_text = []
                try: makeocr(downloadfrom, scanner, h_crafted_text)
                except: print('Procesador de documentos', 'Se requiere agregar PyTesseract a las variables de entorno.')
                break

def _txtbase():
    global _content
    _content = ''
    try:
        pdf = open(readabledoc, 'rb')
        _reader = PdfReader(pdf)
        _content = _reader.pages[0].extract_text().split('\n')
        pdf.close()
    except: pass

def fillentries(stack_inputs):
    for s in stack_inputs:
        s.delete(0, END)
    stack_inputs[0].insert(0, req_document)
    stack_inputs[1].insert(0, kycdated)
    stack_inputs[2].insert(0, req_id)
    stack_inputs[3].insert(0, req_names)
    stack_inputs[4].insert(0, req_lnames)

def returndata(stack_inputs):
    global select, kycdated, req_document, req_id, req_names, req_lnames
    kycdated = ''; req_document = ''; req_id = ''; req_names = ''; req_lnames = ''
    _txtbase()
    select = readabledoc.upper()
    if select.__contains__('KYC2'):
        kycdated = _content[3]
        kycdated = kycdated.split(' ')
        kycdated = kycdated[8:10]
        try: kycdated[1] = kycdated[1].replace('\xa0', '').replace('Código', '')
        except: pass
        switch(kycdated)
        kycdated = f'{kycdated[0]}-{kycdated[1]}'
        for f in _content:
            if f.__contains__('Producto') and  f.__contains__('Pagar'):
                req_document = f
                req_document = req_document.split(' ')
                req_document = req_document[-1]
            if f.__contains__('Primer') and f.__contains__('nombre') and f.__contains__('apellido'):
                req_full_name = []
                temp_req_full_name = f
                temp_req_full_name = temp_req_full_name.replace('Primer nombre', '').replace('Segundo nombre', '').replace('Primer apellido', '').replace('Segundo apellido', '')
                temp_req_full_name = temp_req_full_name.split(' ')
                for r in temp_req_full_name:
                    if r != '': req_full_name.append(r)
                temp_req_full_name = ''
                if len(req_full_name) == 3:
                    req_names = req_full_name[0]
                    req_lnames = f'{req_full_name[1]} {req_full_name[2]}'
                elif len(req_full_name) == 4:
                    req_names = f'{req_full_name[0]} {req_full_name[1]}'
                    req_lnames = f'{req_full_name[2]} {req_full_name[3]}'
                elif len(req_full_name) > 4:
                    req_lnames = f'{req_full_name[-2]} {req_full_name[-1]}'
                    req_full_name.pop(); req_full_name.pop()
                    req_names = ' '.join(req_full_name)
            if f.__contains__('Número') and f.__contains__('identificación') and f.__contains__('Estado'):
                temp_req_id = f
                for rr in temp_req_id:
                    nn = rr.isnumeric()
                    if nn: req_id += rr

    if select.__contains__('CIC2'):
        for f in _content:
            if f.__contains__('Yo') and  f.__contains__('identificación') and  f.__contains__('autorizo'):
                req_line = f
                req_line = req_line.split(',')
                req_line = f'{req_line[1]} {req_line[2]}'
                req_line = req_line.replace('\xa0', ' ').replace('identificación', '').replace('número', '')
                _req_line = req_line.split(' ')
                req_id = _req_line.pop()
                req_line = []
                for rr in _req_line:
                    if rr != '': req_line.append(rr.upper())
                if len(req_line) == 3:
                    req_names = req_line[0]
                    req_lnames = f'{req_line[1]} {req_line[2]}'
                elif len(req_line) == 4:
                    req_names = f'{req_line[0]} {req_line[1]}'
                    req_lnames = f'{req_line[2]} {req_line[3]}'
                elif len(req_line) > 4:
                    req_lnames = f'{req_line[-2]} {req_line[-1]}'
                    req_line.pop(); req_line.pop()
                    req_names = ' '.join(req_line)

    if select.__contains__('CONSENTIMIENTO2') or select.__contains__('DECLARACION2') or select.__contains__('DECLARACIÓN2'):
        for f in _content:
            if f.__contains__('portador') and  f.__contains__('identidad') and  f.__contains__('manifiesto'):
                req_line = f
                for rr in req_line:
                    nn = rr.isnumeric()
                    if nn: req_id += str(rr)
            if f.__contains__('El suscrit') and  f.__contains__('mayor') and  f.__contains__('vecin'):
                req_line = f
                req_line = req_line.split(', ')
                req_line = req_line[1]
                req_line = req_line.split(' ')
                if len(req_line) == 3:
                    req_names = req_line[0]
                    req_lnames = f'{req_line[1]} {req_line[2]}'
                elif len(req_line) == 4:
                    req_names = f'{req_line[0]} {req_line[1]}'
                    req_lnames = f'{req_line[2]} {req_line[3]}'
                elif len(req_line) > 4:
                    req_lnames = f'{req_line[-2]} {req_line[-1]}'
                    req_line.pop(); req_line.pop()
                    req_names = ' '.join(req_line)

    if select.__contains__('PAGAR'):
        try:
            req_id = h_crafted_text[1]
            req_full_name = h_crafted_text[0]
            req_full_name = req_full_name.split(' ')
            if len(req_full_name) == 3:
                req_names = req_full_name[0]
                req_lnames = f'{req_full_name[1]} {req_full_name[2]}'
            elif len(req_full_name) == 4:
                req_names = f'{req_full_name[0]} {req_full_name[1]}'
                req_lnames = f'{req_full_name[2]} {req_full_name[3]}'
            elif len(req_full_name) > 4:
                req_lnames = f'{req_full_name[-2]} {req_full_name[-1]}'
                req_full_name.pop(); req_full_name.pop()
                req_names = ' '.join(req_full_name)
        except: pass
    fillentries(stack_inputs)

def man_pullrequest(stack_inputs, stack_buttons):
    set_downloadfrom(stack_buttons)

def man_readf(stack_inputs, stack_buttons):
    folderscan()
    readabledocs_searching(stack_inputs)
    returndata(stack_inputs)
    stack_buttons[3].config(state='normal')

def man_pushrequest(stack_inputs, stack_buttons):
    if stack_inputs[0].get() != '' and stack_inputs[2].get() != '' and stack_inputs[3].get() != '' and stack_inputs[4].get() != '':
        capsfilenames()
        onlypdfs()
        onlypdfs()
        global render
        render = f'{stack_inputs[2].get()} {stack_inputs[4].get()} {stack_inputs[3].get()} {stack_inputs[0].get()}'
        kycdated = stack_inputs[1].get().upper()
        render = render.upper()
        pr = f'{savingdirectory}/{render}/'
        folderscan()
        _f0 = '0. OTROS DOCUMENTOS'; _f1 = '1. INFORMACIÓN GENERAL';  _f2 = '2. APROBACIONES CREDITICIAS';  _f3 = '3. INFORMACIÓN PARA ANÁLISIS CAPACIDAD';  _f4 = '4. RESULTADOS DE ANÁLISIS'; 
        try: os.makedirs(f'{pr}'); os.makedirs(f'{pr}{_f0}'); os.makedirs(f'{pr}{_f1}'); os.makedirs(f'{pr}{_f2}'); os.makedirs(f'{pr}{_f3}'); os.makedirs(f'{pr}{_f4}')
        except: pass
        for fm in scanner:
            _from = f'{downloadfrom}/{fm}'
            _to = f'{savingdirectory}/{render}'
            _fm = fm.replace('.pdf', '')
            if fm.__contains__('CHEQUE') or fm.__contains__('ESCRITURA') or fm.__contains__('FOTOS') or fm.__contains__('KUIKI') or fm.__contains__('PASAPORTE') or fm.__contains__('PROFORMA') or fm.__contains__('OTROS') or fm.__contains__('SABIAS') or fm.__contains__('SEGURO') or fm.__contains__('VOU') or fm.__contains__('SABÍAS') or fm.__contains__('DOM'):
                if fm.__contains__('KUIKI'): shutil.move(_from, f'{_to}/{_f0}/{_fm} COMUNICA {render}.pdf')
                elif fm.__contains__('SABIAS') or fm.__contains__('SABÍAS'): shutil.move(_from, f'{_to}/{_f0}/SABÍAS QUE {render}.pdf')
                elif fm.__contains__('VOU'): shutil.move(_from, f'{_to}/{_f0}/{_fm}CHER {render}.pdf')
                elif fm.__contains__('DOM'): shutil.move(_from, f'{_to}/{_f0}/DOMICILIACIÓN {render}.pdf')
                else: shutil.move(_from, f'{_to}/{_f0}/{_fm} {render}.pdf')
            elif fm.__contains__('ID.pdf') or fm.__contains__('CIC') or fm.__contains__('CIC1') or fm.__contains__('CIC2') or fm.__contains__('CICAC') or fm.__contains__('CONSENTIMIENTO') or fm.__contains__('KYC') or fm.__contains__('KYC1') or fm.__contains__('KYC2'):
                if fm.__contains__('CIC.pdf') or fm.__contains__('CIC1.pdf'): shutil.move(_from, f'{_to}/{_f1}/AUTORIZACIÓN CIC 1 {render}.pdf')
                elif fm.__contains__('CIC2'):
                    if kycdated != '': shutil.move(_from, f'{_to}/{_f1}/AUTORIZACIÓN CIC {kycdated} {render}.pdf')
                    else: shutil.move(_from, f'{_to}/{_f1}/AUTORIZACIÓN CIC 2 {render}.pdf')
                elif fm.__contains__('CICAC'): shutil.move(_from, f'{_to}/{_f1}/AUTORIZACIÓN CICAC {render}.pdf')
                elif fm.__contains__('CONSENTIMIENTO.pdf') or fm.__contains__('CONSENTIMIENTO1.pdf'): shutil.move(_from, f'{_to}/{_f1}/CONSENTIMIENTO INFORMADO 1 {render}.pdf')
                elif fm.__contains__('CONSENTIMIENTO2'):
                    if kycdated != '': shutil.move(_from, f'{_to}/{_f1}/CONSENTIMIENTO INFORMADO {kycdated} {render}.pdf')
                    else: shutil.move(_from, f'{_to}/{_f1}/CONSENTIMIENTO INFORMADO 2 {render}.pdf')
                elif fm.__contains__('KYC.pdf') or  fm.__contains__('KYC1.pdf'): shutil.move(_from, f'{_to}/{_f1}/KYC 1 {render}.pdf')
                elif fm.__contains__('KYC2'):
                    if kycdated != '': shutil.move(_from, f'{_to}/{_f1}/KYC {kycdated} {render}.pdf')
                    else: shutil.move(_from, f'{_to}/{_f1}/KYC 2 {render}.pdf')
                else: shutil.move(_from, f'{_to}/{_f1}/{_fm} {render}.pdf')
            elif fm.__contains__('CONTRATO') or fm.__contains__('PAGAR') or fm.__contains__('LETRA'):
                if fm.__contains__('PAGAR') or fm.__contains__('LETRA'): shutil.move(_from, f'{_to}/{_f2}/PAGARÉ {render}.pdf')
                else: shutil.move(_from, f'{_to}/{_f2}/{_fm} {render}.pdf')
            elif fm.__contains__('DECLARAC') or fm.__contains__('ORDEN') or fm.__contains__('ORIGEN'):
                if fm.__contains__('DECLARACION.pdf') or fm.__contains__('DECLARACION1.pdf') or fm.__contains__('DECLARACIÓN.pdf') or fm.__contains__('DECLARACIÓN1.pdf'): shutil.move(_from, f'{_to}/{_f3}/DECLARACIÓN JURADA 1 {render}.pdf')
                elif fm.__contains__('DECLARACION2') or fm.__contains__('DECLARACIÓN2'):
                    if kycdated != '': shutil.move(_from, f'{_to}/{_f3}/DECLARACIÓN JURADA {kycdated} {render}.pdf')
                    else: shutil.move(_from, f'{_to}/{_f3}/DECLARACIÓN JURADA 2 {render}.pdf')
                elif fm.__contains__('ORDEN'):
                    if fm.__contains__('ORDEN.pdf'): shutil.move(_from, f'{_to}/{_f3}/ORDEN PATRONAL 1 {render}.pdf')
                    else:
                        nn = _fm.split('ORDEN')
                        shutil.move(_from, f'{_to}/{_f3}/ORDEN PATRONAL {nn[1]} {render}.pdf')
                elif fm.__contains__('ORIGEN'): 
                    if fm.__contains__('ORIGEN.pdf'): shutil.move(_from, f'{_to}/{_f3}/ORIGEN DE INGRESOS 1 {render}.pdf')
                    else:
                        nn = _fm.split('ORIGEN')
                        shutil.move(_from, f'{_to}/{_f3}/ORIGEN DE INGRESOS {nn[1]} {render}.pdf')
            else:
                myfolder = [_f0, _f1, _f2, _f3, _f4]
                global cod, ended
                cod = []; ended = ''
                cod.append(_fm[0]); cod.append(_fm[1]); cod.append(_fm[2])
                if cod[0].isdigit():
                    if cod[1] == 'F' and cod[2] == ' ':
                        _fm = _fm[3:]
                        ended = f'{_fm} {kycdated} {render}.pdf'
                    else:
                        _fm = _fm[2:]
                        ended = f'{_fm} {render}.pdf'
                    pos = cod[0]
                    pos = int(pos)
                    try: moving = f'{savingdirectory}/{render}/{myfolder[pos]}/{ended}'
                    except: moving = f'{savingdirectory}/{render}/Núm de carpeta ({str(pos)}) incorrecto - {ended}'
                else:
                    if cod[0] == 'F' and cod[1] == ' ':
                        _fm = _fm[2:]
                        ended = f'{_fm} {kycdated} {render}.pdf'
                    else: ended = f'{_fm} {render}.pdf'
                    moving = f'{savingdirectory}/{render}/{ended}'
                os.rename(_from, moving)
        for field in stack_inputs:
            field.delete(0, END)

        stack_buttons[3].config(state='disabled')
        setupfolder(stack_buttons[4])
    else: messagebox.showwarning('Procesador de documentos', 'Hay campos obligatorios sin rellenar\nÚnicamente se admite el campo "Fecha en KYC #2" sin rellenar.')