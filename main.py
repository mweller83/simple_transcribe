import PySimpleGUI as sg
import whisper
from docx import Document
import os, subprocess, platform

def transcribe(file_path):
    model = whisper.load_model('medium')
    result = model.transcribe(file_path, language='de')
    text = result['text']
    #print(text)
    document = Document()
    document.add_paragraph(text)
    fpath = 'doc.docx'
    document.save(fpath)
    if platform.system() == 'Darwin':
        subprocess.call(('open', fpath))
    elif platform.system() == 'Windows':
        os.startfile(fpath)
    else:
        subprocess.call(('xdg-open', fpath))

layout = [[sg.Text('Audiodatei auswählen'), sg.FileBrowse(key='-IN-')], [sg.Button("OK")]]
window = sg.Window('Audio Transkribieren', layout, size=(300,200))   

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit': # if user closes window or clicks cancel
        break
    elif event == 'OK':
        path = values['-IN-']
        #print(path)
        transcribe(path)

window.close()
