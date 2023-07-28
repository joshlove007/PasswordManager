# from pynput import keyboard

# def on_press(key):
#     try:
#         print('alphanumeric key {0} pressed'.format(
#             key.char))
#     except AttributeError:
#         print('special key {0} pressed'.format(
#             key))

# def on_release(key):
#     print('{0} released'.format(
#         key))
#     if key == keyboard.Key.esc:
#         # Stop listener
#         return False

# # Collect events until released
# with keyboard.Listener(
#         on_press=on_press,
#         on_release=on_release) as listener:
#     listener.join()

# # ...or, in a non-blocking fashion:
# listener = keyboard.Listener(
#     on_press=on_press,
#     on_release=on_release)
# listener.start()
import pynput
from pynput import keyboard
from pynput.keyboard import Key, Controller



Events = []
while True:
    with keyboard.Events() as events:
        event = events.get(1.0)
        if event == None:
            Events = []
            continue
        elif isinstance(event,keyboard.Events.Release) and hasattr(event.key,'name') and event.key.name in ['ctrl','ctrl_l','ctrl_r','alt','alt_l','alt_r','shift']:
            print('Received event {}'.format(event))
            Events.append(event.key.name)
        elif isinstance(event,keyboard.Events.Release) and hasattr(event.key,'char') and event.key.char == 'v':
            print('Received event {}'.format(event))
            Events.append(event.key.char)
        elif event.key == keyboard.Key.esc:
            break
        if Events:
            print(Events)
        if Events and all(_ in Events for _ in ['ctrl','alt','shift','v']):
            break
 


print(Events)


def on_activate():
    KB = Controller()
    KB.type('worked')
    KB.press('c')
    KB.release('c')
    print('Global hotkey activated!')
    raise pynput._util.AbstractListener.StopException

def for_canonical(f):
    return lambda k: f(l.canonical(k))

hotkey = keyboard.HotKey(keyboard.HotKey.parse('<ctrl>+<alt>+<shift>+v'),on_activate)

l = keyboard.Listener(on_press=None,on_release=for_canonical(hotkey.release))
l.start()
l.wait()
l.join()

GH =  keyboard.GlobalHotKeys({'<ctrl>+<alt>+<shift>+v': on_activate})
GH.start()
GH.wait()
while GH.running:
    GH.wait()