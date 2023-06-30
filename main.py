import PySimpleGUI as sg
import whisper
from docx import Document
import os, subprocess, platform
from os.path import exists
from pathlib import Path
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

layout = [[sg.Button('Audiodatei auswählen', key='browse'), sg.Input(expand_x=True, disabled=True, key='file')], 
          [sg.ProgressBar(100, size=(10, 10), expand_x=True, orientation='h', visible=False, key='bar')],
          ]

window = sg.Window('Audio Transkribieren', layout, size=(400,80), finalize=True)   

thread = None
progress = 0
step = 2

while True:
    event, values = window.read()
    print('evt = {}, values = {}'.format(event, values))
    if event == sg.WIN_CLOSED or event == 'Exit':
        if thread:
            thread.join(timeout=0)
        break
    elif event == 'browse' and not thread:
        path = sg.popup_get_file("", no_window=True)
        if not path or not Path(path).is_file():
            continue
        window['file'].update(path)
        window['bar'].update(visible=True)
        window['browse'].update(disabled=True)
        thread = threading.Thread(target=transcribe, args=(path, window))
        thread.start()
    elif event == '-THREAD-':
        thread.join(timeout=0)
        thread = None
        progress = 0
        window['bar'].update_bar(0,0)
        window['bar'].update(visible=False)
        window['browse'].update(disabled=False)
    elif event == 'update_bar':
        window['bar'].update_bar(progress)
        progress = progress + step
        if progress > 100 or progress < 0:
            step *= -1

window.close()
