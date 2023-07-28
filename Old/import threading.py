import threading
import PySimpleGUI as sg
from pynput import keyboard

def f(key):
    window.write_event_value(key, None)

def func():
    with keyboard.GlobalHotKeys({
            '<alt>+<ctrl>+t': lambda key='Hotkey1':f(key),
            '<alt>+<ctrl>+q': lambda key='Hotkey2':f(key)}) as h:
        h.join()

def f_test():
    layout = [[sg.Text("Hello from PySimpleGUI")], [sg.Button("OK")]]
    window = sg.Window("Demo", layout)
    while True:
        event, values = window.read()
        if event == "OK" or event == sg.WIN_CLOSED:
            break
    window.close()

layout = [[sg.Text("")]]
window = sg.Window("Main Script", layout, finalize=True)
window.hide()
threading.Thread(target=func, daemon=True).start()


while True:
    event, values = window.read()
    print(event)
    if event in (sg.WINDOW_CLOSED, 'Hotkey2'):
        break
    elif event == 'Hotkey1':
        f_test()
window.close()
