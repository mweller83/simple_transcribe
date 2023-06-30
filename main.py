import PySimpleGUI as sg
import whisper
from docx import Document
import os, subprocess, platform
from os.path import exists
import threading
import time

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

def update_progress_bar(win, stop):
    while True:
        if stop():
            break
        win.write_event_value('update_bar', 1)
        time.sleep(.5)

def transcribe(file_path, window):
    stop_thread = False
    thread2 = threading.Thread(target=update_progress_bar, args=(window, lambda: stop_thread))
    thread2.start()
    model = whisper.load_model('medium')
    result = model.transcribe(file_path, language='de')
    text = result['text']
    document = Document()
    document.add_paragraph(text)
    fpath = find_new_filepath()
    document.save(fpath)
    launch_path(fpath)
    stop_thread = True
    thread2.join(timeout=0)
    window.write_event_value('-THREAD-', 'Finished')

layout = [[sg.Text('Audiodatei auswählen', size=(30,1)), sg.FileBrowse(key='-IN-')], 
          [sg.Text('', size=(30,1)), sg.Button('Transkribieren')],
          [sg.Text('Fortschritt', key='bar_label'), sg.ProgressBar(100, size=(20, 20), orientation='h', key='bar')],
          ]

window = sg.Window('Audio Transkribieren', layout, size=(400,100), finalize=True)   

thread = None
progress = 0
step = 2

def show_progress_bar(bool):
    window['bar_label'].update(visible=bool)
    window['bar'].update(visible=bool)

show_progress_bar(False)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit': # if user closes window or clicks cancel
        if thread:
            thread.join(timeout=0)
        break
    elif event == 'Transkribieren' and not thread:
        path = values['-IN-']
        if path:
            show_progress_bar(True)
            window['Transkribieren'].update(disabled=True)
            thread = threading.Thread(target=transcribe, args=(path, window))
            thread.start()
    elif event == '-THREAD-':
        thread.join(timeout=0)
        thread = None
        progress = 0
        window['bar'].update_bar(0,0)
        show_progress_bar(False)
        window['Transkribieren'].update(disabled=False)
    elif event == 'update_bar':
        window['bar'].update_bar(progress)
        progress = progress + step
        if progress > 100 or progress < 0:
            step *= -1

window.close()
