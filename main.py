import PySimpleGUI as sg
import whisper
from docx import Document
import os, subprocess, platform
from os.path import exists

def find_new_filepath():
    fpath = 'doc-{}.docx'
    new_fpath = ''
    count = 0
    while True:
        new_fpath = fpath.format(count)
        if not exists(new_fpath):
            break
        count = count + 1
    return new_fpath

def launch_path(fpath):
    if platform.system() == 'Darwin':
        subprocess.call(('open', fpath))
    elif platform.system() == 'Windows':
        os.startfile(fpath)
    else:
        subprocess.call(('xdg-open', fpath))

def transcribe(file_path):
    model = whisper.load_model('medium')
    result = model.transcribe(file_path, language='de')
    text = result['text']
    document = Document()
    document.add_paragraph(text)
    fpath = find_new_filepath()
    document.save(fpath)
    launch_path(fpath)

layout = [[sg.Text('Audiodatei auswählen'), sg.FileBrowse(key='-IN-')], [sg.Button("OK")]]
window = sg.Window('Audio Transkribieren', layout, size=(300,200))   

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit': # if user closes window or clicks cancel
        break
    elif event == 'OK':
        path = values['-IN-']
        transcribe(path)

window.close()
