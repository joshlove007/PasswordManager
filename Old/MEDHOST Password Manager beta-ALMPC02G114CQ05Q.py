#region Import Modulesimport base64
import configparser
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import gc
import pynput
import keyboard as kb
from pynput import keyboard
from pynput.keyboard import Key, Controller
from ast import literal_eval
from base64 import b64decode, b64encode
from datetime import datetime, timedelta
from getpass import getuser
from os import getcwd, getenv
from time import sleep

import lastpass
import PySimpleGUI as sg
from Crypto.Cipher import AES
from lastpass.fetcher import make_key
from onlykey.client import MessageField, OnlyKey
from requests import get, post


if not sg.running_mac():
    import win32crypt
#endregion

#region Functions
#region Function GetCyberArkToken
def GetCyberArkToken(ArkPass,ArkUser,ArkURL="https://cyberark.medhost.com/PasswordVault/API",Method = "RADIUS"):
    print("Retrieving CyberArk Token - Sending DUO Push")
   
    Body = {
        "username" : ArkUser,
        "password" : ArkPass
    }

    URL = f'{ArkURL}/auth/{Method}/Logon'

    try:
        Token = (post(url=URL,json=Body)).json()
    except Exception as error:
        return {'ErrorMessage':error}
    return Token 
#endregion

#region Function GetCyberArkActivities
def GetCyberArkActivities(Token,AccountId,ArkURL = "https://cyberark.medhost.com/PasswordVault/api"):
    Header = {"Authorization" : Token}

    URL =  ArkURL + f'/Accounts/{AccountId}/Activities'
        
    Response = (get(url=URL,headers=Header)).json()

    return Response
#endregion

#region Function GetCyberArkSettings
def GetCyberArkSettings(Token,ArkURL = "https://cyberark.medhost.com/PasswordVault/api"):
    Header = {"Authorization" : Token}

    URL =  ArkURL + '/settings/configuration'
        
    Response = (get(url=URL,headers=Header)).json()

    return Response
#endregion

#region Function FindCyberArkAccount
def FindCyberArkAccount(Token,Username=None,SearchType='contains',limit='100',page='1',ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    offset = str(int(limit) * (int(page) - 1))
    if not Username or Username == 'Search for accounts':
        #filter = f"?filter=safeName+eq+{SafeName.split(' '*50)[1]}" if SafeName and SafeName != 'Safe Name' else ''
        
        URL    = ArkURL + "/ExtendedAccounts" +"?limit=" + limit + "&offset=" + offset
        
        Response = (get(url=URL,headers=Header)).json()

        return Response
    else:
        #filter = f"&filter=safeName+eq+{SafeName.split(' '*50)[1]}" if SafeName and SafeName != 'Safe Name' else ''
        
        URL  = ArkURL + "/ExtendedAccounts" + '?search=' + Username + '&searchtype=' + SearchType + f'&limit={limit}' + "&offset=" + offset
        Response = (get(url=URL,headers=Header)).json()
        
        return Response
#endregion

#region Function GetCyberArkSafes
def GetCyberArkSafes(Token,ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    
    URL  = ArkURL + '/Safes'
    
    Response = (get(url=URL,headers=Header)).json()
    
    return Response
#endregion

#region Function GetCyberArkPassword
def GetCyberArkPassword(AccountId,Token,UserName,Reason="MEDHOST Password Manager Update",Case='',ArkURL="https://cyberark.medhost.com/PasswordVault/API"):
    print("Retrieving CyberArk Password for " + UserName)

    Header = {"Authorization" : Token}
    
    Body = {
        "reason"  : Reason,
        "TicketId": Case
    }

    URL = ArkURL + "/Accounts/" + AccountId + "/Password/Retrieve"

    Response = (post(url=URL,json=Body,headers=Header)).json()
    return Response
#endregion

#region Function CheckInCyberArkPassword
def CheckInCyberArkPassword(AccountId,Token,UserName,ArkURL="https://cyberark.medhost.com/PasswordVault/API"):
    print("Checking in CyberArk Password for " + UserName)

    Header = {"Authorization" : Token}
    
    URL = ArkURL + "/Accounts/" + AccountId + "/CheckIn"

    Response = post(url=URL,headers=Header)
    
    return Response
#endregion

#region Function MakeCyberArkTable
def MakeCyberArkTable(ArkAccts):
    CATable = []

    for Acct in ArkAccts['Accounts']:
        AcctP    = Acct['Properties']

        inuse    = '⛔' if 'LockedBy' in AcctP and AcctP['LockedBy'] else '✅' 

        username = AcctP['UserName'    ] if 'UserName'     in AcctP else '' 
        address  = AcctP['Address'     ] if 'Address'      in AcctP else ''
        facility = AcctP['FacilityName'] if 'FacilityName' in AcctP else 'MEDHOST'
        client   = AcctP['ClientNumber'] if 'ClientNumber' in AcctP else '' 
        system   = AcctP['SystemName'  ] if 'SystemName'   in AcctP else ''
        tool     = AcctP['Tool'        ] if 'Tool'         in AcctP else ''
        support  = AcctP['SupportEmail'] if 'SupportEmail' in AcctP else '' 
        safe     = AcctP['Safe'        ] if 'Safe'         in AcctP else '' 
        name     = AcctP['Name'        ] if 'Name'         in AcctP else ''
        lockedby = AcctP['LockedBy'    ] if 'LockedBy'     in AcctP else ''
        UsedBy   = AcctP['LastUsedBy'  ] if 'LastUsedBy'   in AcctP else ''
        UsedDate = AcctP['LastUsedDate'] if 'LastUsedDate' in AcctP else ''
        Id       =  Acct['AccountID'   ] if 'AccountID'    in Acct  else ''

        row = [inuse,username,address,facility,client,system,tool,support,safe,name,lockedby,UsedBy,UsedDate,Id]

        CATable.append(row)
    
    return CATable
#endregion

#region Function get_chrome_datetime
def get_chrome_datetime(chromedate):
    """Return a `datetime.datetime` object from a chrome format datetime
    Since `chromedate` is formatted as the number of microseconds since January, 1601"""
    return datetime(1601, 1, 1) + timedelta(microseconds=chromedate)
#endregion

#region function getkccred
def getkccred(service, account):
    def decode_hex(s):
        s = eval('"' + re.sub(r"(..)", r"\x\1", s) + '"')
        if "" in s: s = s[:s.index("")]
        return s

    cmd = ' '.join([
        "/usr/bin/security",
        " find-generic-password",
        "-g -s '%s' -a '%s'" % (service, account),
        "2>&1 >/dev/null"
    ])
    p = os.popen(cmd)
    s = p.read()
    p.close()
    m = re.match(r"password: (?:0x([0-9A-F]+)\s*)?\"(.*)\"$", s)
    if m:
        hexform, stringform = m.groups()
        if hexform:
            return decode_hex(hexform)
        else:
            return stringform
#endregion

#region Function get_encryption_key
def get_encryption_key(edge=False,macos=False):
    userpath = os.environ["HOME"] if macos else os.environ["USERPROFILE"]
    if macos:
        if edge:
            cred = getkccred('Microsoft Edge Safe Storage','Microsoft Edge')
        else:
            cred = getkccred('Chrome Safe Storage','Chrome')
        safeStorageKey = cred.encode() #base64.b64decode(cred) if isinstance(cred,str) else b''
        
        return hashlib.pbkdf2_hmac('sha1', safeStorageKey, b'saltysalt', 1003)[:16]

    if edge:local_state_path = os.path.join(userpath,"AppData", "Local", "Microsoft", "Edge","User Data", "Local State")
    else:   local_state_path = os.path.join(userpath,"AppData", "Local", "Google", "Chrome","User Data", "Local State")

    with open(local_state_path, "r", encoding="utf-8") as f:
        local_state = f.read()
        local_state = json.loads(local_state)

    # decode the encryption key from Base64
    key = base64.b64decode(local_state["os_crypt"]["encrypted_key"])
    # remove DPAPI str
    key = key[5:]
    # return decrypted key that was originally encrypted
    # using a session key derived from current user's logon credentials
    # doc: http://timgolden.me.uk/pywin32-docs/win32crypt.html
    return win32crypt.CryptUnprotectData(key, None, None, None, 0)[1]
#endregion

#region Function decrypt_password
def decrypt_password(password, key,macos=False):
    if macos:
        iv      = b' ' * 16
        newpassword = password[3:]    
        cipher = AES.new(key, AES.MODE_CBC, iv)
        x = cipher.decrypt(newpassword)
        return x[:-x[-1]].decode('utf8')
    try:
        # get the initialization vector
        iv = password[3:15]
        password = password[15:]
        # generate cipher
        cipher = AES.new(key, AES.MODE_GCM, iv)
        # decrypt password
        return cipher.decrypt(password)[:-16].decode()
    except:
        try:
            return str(win32crypt.CryptUnprotectData(password, None, None, None, 0)[1])
        except:
            # not supported
            return ""
#endregion

#region Function getchromiumvault
def getchromiumvault(edge=False,macos=False):
    # get the AES key
    key = get_encryption_key(edge,macos)
    # local sqlite database path
    userpath = os.environ["HOME"] if macos else os.environ["USERPROFILE"]
    if macos:
        if edge:
            db_path = os.path.join(userpath, "Library", "Application Support", "Microsoft Edge","Default", "Login Data")
        else:
            db_path = os.path.join(userpath, "Library", "Application Support", "Google", "Chrome","Default", "Login Data")
    else:
        if edge:
            db_path = os.path.join(userpath, "AppData", "Local", "Microsoft", "Edge", "User Data", "default", "Login Data")
        else:
            db_path = os.path.join(userpath, "AppData", "Local", "Google", "Chrome", "User Data", "default", "Login Data")
    # copy the file to another location
    # as the database will be locked if chrome is currently running
    filename = "ChromiumData.db"
    shutil.copyfile(db_path, filename)
    # connect to the database
    db = sqlite3.connect(filename)
    cursor = db.cursor()
    cursor.execute("select c.name FROM pragma_table_info('logins') c;")
    columns = {i:cname[0] for i,cname in  enumerate(cursor.fetchall())}
    # `logins` table has the data we need
    cursor.execute("select * from logins")
    # iterate over all rows
    chromiumvault = [{columns[i]:_ for i,_ in enumerate(row)} for row in cursor.fetchall()]

    for i,entry in enumerate(chromiumvault):
        try:
            chromiumvault[i]['password_value']           = decrypt_password(entry['password_value'], key, macos)
            chromiumvault[i]['date_created']             = get_chrome_datetime(entry['date_created'])
            chromiumvault[i]['date_last_used']           = get_chrome_datetime(entry['date_last_used'])
            chromiumvault[i]['date_password_modified']   = get_chrome_datetime(entry['date_password_modified'])
        except:
            continue
                
    cursor.close()
    db.close()
    try:
        # try to remove the copied db file
        os.remove(filename)
    except:
        pass
    return chromiumvault
#endregion

#region Function collapse
def collapse(layout, key:str, visible:bool,pad:int=0,size=(None,None)):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this seciton visible / invisible
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    return sg.pin(sg.Column(layout, key=key,visible=visible,pad=pad,size=size))
#endregion

#region Function collapse
def EventHelper(Input=''):
    return Input
#endregion

#region on_activate
def on_activate():
    print('Global hotkey activated!')
    raise pynput._util.AbstractListener.StopException
#endregion

#region typestring
def typestring(string=''):
    KB = Controller()   
    KB.type(string)
#endregion

#region Function HotKeyListener
def HotKeyListener():
    kb.wait('ctrl+win+v',True,True)
#endregion

#region MoveRow
def MoveRow(window,layout,NewRow,OldRow):
    Nr=NewRow;Or=OldRow;location=window.CurrentLocation()
    layout.insert(Nr,layout.pop(Or))    
    for rI,row in enumerate(layout):
        for eI,elem in enumerate(row):
            layout[rI][eI].Position = (rI,eI)
            layout[rI][eI].ParentContainer = None
    values = window.ReturnValuesDictionary
    listboxes = [{'key':_[1].Key,'values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox']  
    newwindow = sg.Window(Title, layout,finalize=True,location=location)
    newwindow.fill(values)
    for box in listboxes: newwindow[box['key']].update(values=box['values'])
    window.close()
    return newwindow
#endregion

#region Function getdate
def getdate(date):
    return datetime.fromtimestamp(date).strftime('%m-%d-%Y %I:%M %p')
#endregion

#region Function SwapContainers
def SwapContainers(window,C1,C2):
    location  = window.CurrentLocation() if window.Shown else (None,None)
    TabGroups = [_[0] for _ in window.key_dict.items() if _[1].Type == 'tabgroup']
    values    = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0] not in TabGroups}
    listboxes = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox' and _[1].Values]
    tabs      = [{'Key':_[1].Key,'TabID':_[1].TabID} for _ in window.key_dict.items() if _[1].Type == 'tab'and (_[1].TabID or _[1].TabID == 0)]
    tables    = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'table' and _[1].Values]
    metadata  = window.metadata
    
    C1Rows = window[C2].Rows
    C2Rows = window[C1].Rows
    window[C1].Rows = C1Rows
    window[C2].Rows = C2Rows

    for rI,row in enumerate(window.Rows):
        for eI,elem in enumerate(row):
                window.Rows[rI][eI].Position = (rI,eI)
                window.Rows[rI][eI].ParentContainer = None
        
    newwindow = sg.Window(Title, window.Rows,location=location,finalize=True,font=window.Font)
    newwindow.force_focus()
    window.close()
    newwindow.fill(values)
    newwindow.metadata = metadata
    for box in listboxes: newwindow[box['Key']].update(values=box['Values'])
    for tab in tabs:      newwindow[tab['Key']].TabID = tab['TabID']
    for table in tables: 
        ScrollPosition = values[table['Key']][-1] / (len(table['Values']) - 1) if values[table['Key']] and values[table['Key']][-1] != 0 else 0
        newwindow[table['Key']].update(values=table['Values'])
        newwindow[table['Key']].update(select_rows=values[table['Key']])
        newwindow[table['Key']].set_vscroll_position(ScrollPosition)    
    return newwindow
#endregion

#region Function ReloadWindow
def ReloadWindow(window):
    location  = window.CurrentLocation() if window.Shown else (None,None)
    TabGroups = [_[0] for _ in window.key_dict.items() if _[1].Type == 'tabgroup']
    values    = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0] not in TabGroups}
    listboxes = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox' and _[1].Values]
    tabs      = [{'Key':_[1].Key,'TabID':_[1].TabID} for _ in window.key_dict.items() if _[1].Type == 'tab'and (_[1].TabID or _[1].TabID == 0)]
    tables    = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'table' and _[1].Values]
    metadata  = window.metadata

    for rI,row in enumerate(window.Rows):
        for eI,elem in enumerate(row):
                window.Rows[rI][eI].Position = (rI,eI)
                window.Rows[rI][eI].ParentContainer = None
        
    newwindow = sg.Window(Title, window.Rows,location=location,finalize=True,font=window.Font)
    newwindow.force_focus()
    window.close()
    newwindow.fill(values)
    newwindow.metadata = metadata
    try:
        for box in listboxes: newwindow[box['Key']].update(values=box['Values'])
        for tab in tabs:      newwindow[tab['Key']].TabID = tab['TabID']
        for table in tables: 
            ScrollPosition = values[table['Key']][-1] / (len(table['Values']) - 1) if values[table['Key']] and values[table['Key']][-1] != 0 else 0
            newwindow[table['Key']].update(values=table['Values'])
            newwindow[table['Key']].update(select_rows=values[table['Key']])
            newwindow[table['Key']].set_vscroll_position(ScrollPosition)          
    except Exception as error:
        print(error)
    return newwindow
#endregion

#region CopyLayout
def CopyLayout(layout):
    NewLayout = layout   
    for rI,row in enumerate(NewLayout):
        for eI,elem in enumerate(row):
            NewLayout[rI][eI].Position = (rI,eI)
            NewLayout[rI][eI].ParentContainer = None
    return NewLayout
#endregion

#region MoveRows
def MoveRows(window,layout,Rows:list[tuple]):
    '''Rows = list of tuples [(NewRow,OldRow),(NewRow,OldRow)]'''
    location=window.CurrentLocation()
    oldnums = [_[1] for _ in Rows]
    newnums = [_[0] for _ in Rows]
    srows=[row for rI,row in enumerate(layout) if rI not in oldnums]
    newlayout = []
    i = 0
    while i < len(layout):
        if i not in newnums:
            newlayout.append(srows.pop(0))
        else:
            mrow = [_ for _ in Rows if _[0] == i][0]
            newlayout.append(layout[mrow[1]])
        for eI,elem in enumerate(newlayout[i]):
            newlayout[i][eI].Position = (i,eI)
            newlayout[i][eI].ParentContainer = None
        i = i + 1
    listboxes = [{'key':_[1].Key,'values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox']  
    values = window.ReturnValuesDictionary
    newwindow = sg.Window(Title, newlayout,finalize=True,location=location)
    newwindow.fill(values)
    for box in listboxes: newwindow[box['key']].update(values=box['values'])
    window.close()
    return (newwindow,newlayout)
#endregion

#region Function pin_popup
def pin_popup(message, title=None, default_text='', password_char='', offer_reset=False, size=(None, None), button_color=None,
                   background_color=None, text_color=None, icon=None, font=None, no_titlebar=False,
                   grab_anywhere=False, keep_on_top=None, location=(None, None), relative_location=(None, None), image=None, modal=True):
    """
    Display Popup with text entry field. Returns the text entered or None if closed / cancelled

    :param message:          message displayed to user
    :type message:           (str)
    :param title:            Window title
    :type title:             (str)
    :param default_text:     default value to put into input area
    :type default_text:      (str)
    :param password_char:    character to be shown instead of actually typed characters
    :type password_char:     (str)
    :param size:             (width, height) of the InputText Element
    :type size:              (int, int)
    :param button_color:     Color of the button (text, background)
    :type button_color:      (str, str) or str
    :param background_color: background color of the entire window
    :type background_color:  (str)
    :param text_color:       color of the message text
    :type text_color:        (str)
    :param icon:             filename or base64 string to be used for the window's icon
    :type icon:              bytes | str
    :param font:             specifies the  font family, size, etc. Tuple or Single string format 'name size styles'. Styles: italic * roman bold normal underline overstrike
    :type font:              (str or (str, int[, str]) or None)
    :param no_titlebar:      If True no titlebar will be shown
    :type no_titlebar:       (bool)
    :param grab_anywhere:    If True can click and drag anywhere in the window to move the window
    :type grab_anywhere:     (bool)
    :param keep_on_top:      If True the window will remain above all current windows
    :type keep_on_top:       (bool)
    :param location:         (x,y) Location on screen to display the upper left corner of window
    :type location:          (int, int)
    :param relative_location: (x,y) location relative to the default location of the window, in pixels. Normally the window centers.  This location is relative to the location the window would be created. Note they can be negative.
    :type relative_location: (int, int)
    :param image:            Image to include at the top of the popup window
    :type image:             (str) or (bytes)
    :param modal:            If True then makes the popup will behave like a Modal window... all other windows are non-operational until this one is closed. Default = True
    :type modal:             bool
    :return:                 Text entered or None if window was closed or cancel button clicked
    :rtype:                  str | None
    """


    layout = [[sg.Text(message, auto_size_text=True, text_color=text_color, background_color=background_color, font=font)],
               [sg.InputText(default_text=default_text, size=size, key='_INPUT_', password_char=password_char)],
               [sg.Button('Ok', size=(6, 1), bind_return_key=True), sg.Button('Cancel', size=(6, 1))]]
    
    if offer_reset: 
        layout[2].append(sg.Push())
        layout[2].append(sg.Button('Reset Pin'))

    window = sg.Window(title=title or message, layout=layout, icon=icon, auto_size_text=True, button_color=button_color, no_titlebar=no_titlebar,
                    background_color=background_color, grab_anywhere=grab_anywhere, keep_on_top=keep_on_top, location=location, relative_location=relative_location, finalize=True, modal=modal)

    button, values = window.read()
    window.close()
    del window
    if button == 'Cancel':
        return None
    elif button == 'Ok':
        path = values['_INPUT_']
        return path
    elif button == 'Reset Pin':
        return 'reset'
#endregion

#region Function Bool
def Bool(string):
    return bool(literal_eval(str(string)))
#endregion

#region Function array
def array(string):
    return literal_eval(str(string))
#endregion

#region Function SaveToConfig
def SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,Print=True):
        config = configparser.ConfigParser()
        user = getuser()
        Pass = (subprocess.run("hostname",capture_output=True,text=True,shell=True)).stdout.rstrip()
        key = make_key(user,Pass,pin)

        config.add_section('General')
        config['General']['Sync Options Visible']         = str(window["SyncOpt"]._visible)
        config.add_section('CyberArk')
        config['CyberArk']['Network Username']            =     Inputs["ArkUser"]
        config['CyberArk']['Accounts']                    = str(window['CATable'].Values).replace('✅','').replace('⛔','')
        config['CyberArk']['Selected Account Index']      = str(Inputs['CATable'])
        config['CyberArk']['Sync From CyberArk Selected'] = str(Inputs["CAFrom"])
        config['CyberArk']['Login Visible']               = str(window['ArkLoginInfo']._visible)
        config.add_section('LastPass')     
        config['LastPass']['Username']                    =     Inputs["LPUser"]
        config['LastPass']['Sync From LastPass Selected'] = str(Inputs['LPFrom'])
        config['LastPass']['Sync To LastPass Selected']   = str(Inputs["LP"])
        config['LastPass']['LastPass Accounts']           = str(window['LPSelected'].Values)
        config['LastPass']['Login Visible']               = str(window['LPLoginInfo']._visible)
        config.add_section('OnlyKey')     
        config['OnlyKey']['OnlyKey Selected']             = str(Inputs["OK"])
        config['OnlyKey']['Keyword']                      =     Inputs["OK_Keyword"]
        config['OnlyKey']['Keyword Search Selected']      = str(Inputs["KWSearch"])
        config.add_section('App Data')     
        config['App Data']['CyUsPa']                      = EncryptString(Inputs["ArkPass"],key)
        config['App Data']['LaUsPa']                      = EncryptString(Inputs["LPPass"],key)
        config['App Data']['CySeTo']                      = EncryptString(ArkInfo['Token'],key)
        config['App Data']['LaSeTo']                      = EncryptString(LPInfo['Token'],key)
        config['App Data']['LaSeKe']                      = EncryptString(b64encode(LPInfo['Key']).decode(),key)
        config['App Data']['LaSeId']                      = EncryptString(LPInfo['SessionId'],key)
        config['App Data']['LPSeIt']                      = EncryptString(str(LPInfo['Iteration']),key)

        config.add_section('Selected Slots')
        for _ in Slots: config['Selected Slots'][_]  = str(Inputs[f"S_{_}"])
        
        config.add_section('Acct-Slot Mappings')
        for _ in Slots: config['Acct-Slot Mappings'][_]  = str(Inputs[f"C_{_}"])
        
        with open(WorkingDirectory, 'w') as configfile:
            config.write(configfile)
        if Print: print(f"Configuration Saved To: {WorkingDirectory}")
        return 
#endregion

#region DecryptString
def DecryptString(data, encryption_key,b64=False):
    """Decrypts AES-256 CBC bytes."""
    data = data.encode() if isinstance(data,str) else data
    aes = AES.new(encryption_key, AES.MODE_CBC, b64decode(data[1:25]))
    d = aes.decrypt(b64decode(data[26:]))
    # http://passingcuriosity.com/2009/aes-encryption-in-python-with-m2crypto/
    unpad = lambda s: s[0:-ord(d[-1:])]
    if b64:
        return b64decode(unpad(d))
    else:
        return unpad(d).decode()
#endregion

#region Function EncryptString
def EncryptString(string, encryption_key):
    """
    Encrypt AES-256 CBC bytes.
    """
    #https://stackoverflow.com/questions/12524994/encrypt-decrypt-using-pycrypto-aes-256
    pad = lambda s: s+(16-len(s)%16)*chr(16-len(s)%16)
    data = pad(string).encode()
    aes = AES.new(encryption_key, AES.MODE_CBC)
    iv    = b64encode(aes.iv).decode()
    edata = b64encode(aes.encrypt(data)).decode()
    en_string = f"!{iv}|{edata}"
    # http://passingcuriosity.com/2009/aes-encryption-in-python-with-m2crypto/
    return en_string
#endregion

#region Function GetConfigAppData
def GetConfigAppData(WorkingDirectory,pin,Print=True,FB=False):
        config = configparser.ConfigParser()
        config.read(WorkingDirectory)
        Pass = (subprocess.run("hostname",capture_output=True,text=True,shell=True)).stdout.rstrip()
        key = make_key(getuser(),Pass,pin)
        fb = EncryptString('',key)

        AppData = {}

        ArkPass     = _ if bool(_:=config.get('App Data','CyUsPa',fallback=fb)) and not FB else fb            
        LPPass      = _ if bool(_:=config.get('App Data','LaUsPa',fallback=fb)) and not FB else fb     
        ArkToken    = _ if bool(_:=config.get('App Data','CySeTo',fallback=fb)) and not FB else fb     
        LPToken     = _ if bool(_:=config.get('App Data','LaSeTo',fallback=fb)) and not FB else fb     
        LPKey       = _ if bool(_:=config.get('App Data','LaSeKe',fallback=fb)) and not FB else fb     
        LPsId       = _ if bool(_:=config.get('App Data','LaSeId',fallback=fb)) and not FB else fb      
        LPIteration = _ if bool(_:=config.get('App Data','LPSeIt',fallback=fb)) and not FB else fb                

        try:
            AppData['ArkPass']     =               DecryptString((ArkPass    ),key)
            AppData['LPPass']      =               DecryptString((LPPass     ),key)
            AppData['ArkToken']    =               DecryptString((ArkToken   ),key)
            AppData['LPToken']     =               DecryptString((LPToken    ),key)
            AppData['LPKey']       =               DecryptString((LPKey      ),key,b64=True)
            AppData['LPsId']       =               DecryptString((LPsId      ),key)
            AppData['LPIteration'] = int(_) if (_:=DecryptString((LPIteration),key)).isnumeric() else 0
        except Exception as error:
            return 'error'
        if Print: print(f"Configuration Loaded From: {WorkingDirectory}")
        
        if FB:
            config['App Data']['CyUsPa'] = fb
            config['App Data']['LaUsPa'] = fb
            config['App Data']['CySeTo'] = fb
            config['App Data']['LaSeTo'] = fb
            config['App Data']['LaSeKe'] = fb
            config['App Data']['LaSeId'] = fb
            config['App Data']['LPSeIt'] = fb
            with open(WorkingDirectory, 'w') as configfile:
                config.write(configfile)

        return AppData
#endregion

#region Function GetCyberArkLogin
def GetCyberArkLogin(ArkUser,ArkPass):
    for param in (ArkUser,ArkPass):
        if not bool(param):
            print("All fields must be compeleted, fill in all empty fields")
            return
        
    ArkToken  = GetCyberArkToken(ArkPass=ArkPass,ArkUser=ArkUser)
    if 'ErrorMessage' in ArkToken:
        print("Unable Retrieve CyberArk Token")
        print(ArkToken['ErrorMessage'])
        print("Try Re-entering Your CyberArk Credentials and Checking Your Connection")
        return
    else:
        print(f"DUO Push Accepted - {ArkUser} Authenticated")

    return ArkToken
#endregion

#region LastPassSession
def LastPassSession(AppData):
    LastPass = lastpass.Session(id=AppData['LPsId'],key_iteration_count=AppData['LPIteration'],lptoken=AppData['LPToken'],encryption_key=AppData['LPKey'])
    try:
        LastPass.OpenVault(include_shared=True)
    except:
        pass
    if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
        return
    return LastPass
#endregion

#region LastPassLogin
def LastPassLogin(Inputs):    
    try:
        LastPass = lastpass.Session.login(username=Inputs['LPUser'],password=Inputs['LPPass']).OpenVault(include_shared=True)
    except Exception as e:
        return print(e)
    return LastPass
#endregion

#region Function SetLastPassPassword
def SetLastPassPassword(LastPass,Password,Username='',Account=None):
    try:
        if Account:
            LastPass.UpdateAccount(id=Account.id,password=Password,Account=Account)
    
        elif Username:
            UpdateParameters = dict(
                name=Username,
                username=Username,
                password=Password,
                notes="Created By MEDHOST Password Manager",
            )
            LastPass.NewAccount(**UpdateParameters)
        else:
            print('Unable to add or update LastPass entry')
    except Exception as error:
        print(error)
        return LastPass

    return LastPass
#endregion

#region Function UpdateOnlyKey
def UpdateOnlyKey(Password,OK_Keyword,SlotSelections:dict,SlotsTrue:bool):     
    #Check for Connection to OnlyKey
    try:
        onlykey = OnlyKey(connect=True,tries=1)
    except BaseException as Error:
        print(Error)
        return
    
    #Check if OnlyKey Locked
    try:
        onlykey.read_bytes(timeout_ms=1000)
    except BaseException as Error:
        print(Error)
        onlykey.close()
        return

    OKSlots = onlykey.getlabels()

    BGSlots = []
    for Slot in OKSlots:
        if SlotsTrue and Slot.name in SlotSelections or not SlotsTrue and OK_Keyword in Slot.label:
            print(f'Setting Slot {Slot.name} {Slot.label}')
            BGSlots.append(Slot)
            sleep(.2)
            onlykey.setslot(slot_number=Slot.number, message_field=MessageField.PASSWORD, value=Password)
    onlykey.close()
    if not bool(BGSlots):
        print(f"Unable to find any slots containing the keyword \"{OK_Keyword}\"")
    return Password 
#endregion
#endregion

#region Read Config File and get prepopulated data
Slots = ["1a","1b","2a","2b","3a","3b","4a","4b","5a","5b","6a","6b"]

config = configparser.ConfigParser()
try:
    HomeDirectory = getenv("HOME")
except:
    HomeDirectory = None
    
WorkingDirectory = HomeDirectory + '/Config.ini' if sg.running_mac() and HomeDirectory else getcwd() + '/Config.ini'
config.read(WorkingDirectory)

ConfigSyncOpt  =  Bool(config.get('General' ,'Sync Options Visible'        ,fallback='True'))
       
ConfigUsername =       config.get('CyberArk','Network Username'            ,fallback=getuser())
ConfigCATable  = array(config.get('CyberArk','Accounts'                    ,fallback="[]"))
ConfigCAIndex  = array(config.get('CyberArk','Selected Account Index'      ,fallback="[]"))
ConfigCAFrom   =  Bool(config.get('CyberArk','Sync From CyberArk Selected' ,fallback='True'))
ConfigCALogin  =  Bool(config.get('CyberArk','Login Visible'               ,fallback='True'))
       
ConfigLPUser   =       config.get('LastPass','Username'                    ,fallback="")
ConfigLP       =  Bool(config.get('LastPass','Sync To LastPass Selected'   ,fallback="True"))
ConfigLPFrom   =  Bool(config.get('LastPass','Sync From LastPass Selected' ,fallback="False"))
ConfigLPAccts  = array(config.get('LastPass','LastPass Accounts'           ,fallback="[]"))
ConfigLPLogin  =  Bool(config.get('LastPass','Login Visible'               ,fallback='True')) 
       
ConfigOK       =  Bool(config.get('OnlyKey' ,'OnlyKey Selected'            ,fallback="False"))
ConfigOKWord   =       config.get('OnlyKey' ,'Keyword'                     ,fallback="BG*")
ConfigKWTrue   =  Bool(config.get('OnlyKey' ,'Keyword Search Selected'     ,fallback="False"))
ConfigSlots    = {Slot : Bool(config.get('Selected Slots',Slot,fallback="False")) for Slot in Slots}
ConfigMap      = {Slot : config.get('Acct-Slot Mappings',Slot,fallback="") for Slot in Slots}

if ConfigUsername == "": ConfigUsername = getuser()

#AppData Pin '1F512'
test = True
if not test:
    pin = pin_popup('Enter Your Pin',password_char="*",title='MEDHOST Password Manager',offer_reset=True)
    if not pin:
        exit()
    elif pin == 'reset':    
        pin = pin_popup('Enter Your New Pin',password_char="*",title='MEDHOST Password Manager')
        while not pin or not pin.isnumeric() or len(pin) < 6:
            pin = pin_popup('Your pin must be numeric and contain 6 or more characters',password_char="*",title='MEDHOST Password Manager')
            if not pin:
                exit()
                
        pin = int(pin)
        AppData = GetConfigAppData(WorkingDirectory,pin,FB=True)
        if AppData == 'error':
            exit()
    else:
        pin = int(pin)
        AppData = GetConfigAppData(WorkingDirectory,pin)
        while AppData == 'error':
            pin = pin_popup('Try Reentering your pin',password_char="*",title='MEDHOST Password Manager',offer_reset=True)
            if not pin:
                exit()
            elif pin == 'reset':    
                pin = pin_popup('Enter Your New Pin',password_char="*",title='MEDHOST Password Manager')
                while not pin or not pin.isnumeric() or len(pin) < 6:
                    pin = pin_popup('Your pin must be numeric and contain 6 or more characters',password_char="*",title='MEDHOST Password Manager')
                    if not pin:
                        exit()         
                pin = int(pin)
                AppData = GetConfigAppData(WorkingDirectory,pin,FB=True)
                if AppData == 'error':
                    exit()
            else:
                pin = int(pin)
                AppData = GetConfigAppData(WorkingDirectory,pin)
                
else:
    pin = 999999
    AppData = GetConfigAppData(WorkingDirectory,pin)
    if AppData == 'error':
        print('Pin error')
        exit()
        


ArkInfo = {'Token':AppData['ArkToken']}
LPInfo = {
    'Token'    : AppData['LPToken'],
    'Key'      : AppData['LPKey'],
    'SessionId': AppData['LPsId'],
    'Iteration': AppData['LPIteration']
}
if AppData['LPsId'] and AppData['LPIteration'] and AppData['LPToken'] and AppData['LPKey']:
    LastPass = LastPassSession(AppData)
else:
    LastPass = None    
    
if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
    LPList   =  ['Login to see LastPass entries here']
    LVTable  = [['Login to see LastPass entries here']]
    LPTable  = [['Login to see LastPass entries here']]
    LVGroups = ['Group/Folder']
else:
    LPList   = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
    LVTable  = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
    LPTable  = [[_.name,_.username,_.url,_.notes,_.id,_.group] for _ in LastPass.accounts]
    LVGroups = ['Group/Folder']
    for Group in [_[4] for _ in LVTable]:
        try:
            Group = Group.split('\\')[-1]
        except:
            pass
        if Group and Group not in LVGroups:
            LVGroups.append(Group)

ArkToken   = AppData['ArkToken']
if ArkToken:
    try:
        ArkAccts = FindCyberArkAccount(Token=ArkToken)
    except:
        ArkAccts = None
else:
    ArkAccts = None
    
if not ArkAccts or 'ErrorMessage' in ArkAccts:
    CATable = ConfigCATable if ConfigCATable else [[],['','','Login to see'],['','','CyberArk'],['','','Accounts Here']]
else:
    CATable = ConfigCATable if ConfigCATable else MakeCyberArkTable(ArkAccts)
    
GCTable = []
METable = []
macos = True if sg.running_mac() else False
try:
    chromevault = getchromiumvault(macos=macos)
except:
    chromevault = []  
if chromevault:
    for Pass in chromevault:
        if 'android://' not in Pass['signon_realm']:
            url      = Pass['signon_realm'].replace('https://','').replace('http://','').replace('www.','')
            url      = url[:len(url)-1] if url[-1] == '/' else url
            username = Pass['username_value']
            GCTable.append([url,username,Pass['id'],Pass['password_value']])
try:
    edgevault = getchromiumvault(edge=True,macos=macos)
except:
    edgevault = []  
if edgevault:
    for Pass in edgevault:
        if 'android://' not in Pass['signon_realm']:
            url      = Pass['signon_realm'].replace('https://','').replace('http://','').replace('www.','')
            url      = url[:len(url)-1] if url[-1] == '/' else url
            username = Pass['username_value']
            METable.append([url,username,Pass['id'],Pass['password_value']])

sg.DEFAULT_INPUT_ELEMENTS_COLOR = 'white'#sg.DEFAULT_BACKGROUND_COLOR # 
sg.DEFAULT_INPUT_TEXT_COLOR     = 'black'#'white'                     #  
InputColor = dict(background_color=sg.DEFAULT_INPUT_ELEMENTS_COLOR,text_color=sg.DEFAULT_INPUT_TEXT_COLOR)
#endregion

#region Collapsible Section Layouts
#region tooltips
ToggleToolTip   = "Click Here to Collapse or Expand Output Section"
OutputToolTip   = "The Outputs and Error Messages of the Application Appear Here"
ClearToolTip    = "Click Here to Clear Text from the Output Element"
KeywordToolTip  = "OnlyKey Slot Label Keyword: Search Keyword for Labels of Slots You Would Like to Update"
UserNameToolTip = "Your CyberArk Username (sAMAccountName): For Logging into CyberArk"
PasswordToolTip = "Your 16+ Digit Domain Password: For Logging into CyberArk"
BGUserToolTip   = "Select a CyberArk Account to Check-Out and Sync"
SubmitToolTip   = "Click Here to Login to CyberArk and Update OnlyKey Slot Passwords"
SaveToolTip     = "Click Here to Save Field Values and Output Toggle to Config.ini File (Network Password is NOT Saved)"
#endregion

#region Output
ee = dict(enable_events=True)
Fm  = 2 if sg.running_mac() else 0
Sm  = 2 if sg.running_mac() else 0
Im  = 0 if sg.running_mac() else 0
OutPutSection = [
    [sg.Output(size=(63,15), key='-OUTPUT-',visible=True,expand_x=True,echo_stdout_stderr=True,pad=((5,5),(5,3)),tooltip=OutputToolTip,**InputColor)],
    [sg.Button('Clear',key='ClearD',tooltip=ClearToolTip)]
]

OutPutSection2 = [
    [sg.Output(size=(63,15), key='-OUTPUT-2',visible=True,expand_x=True,echo_stdout_stderr=True,pad=((5,5),(5,3)),tooltip=OutputToolTip,**InputColor)],
    [sg.Button('Clear',key='ClearD2',tooltip=ClearToolTip)]
]
#endregion

#region LastPass
LPAccts = ConfigLPAccts
Ls    =  55   if sg.running_mac() else 45
Llip  =  0    if sg.running_mac() else 5
Llip2 =  2    if sg.running_mac() else 0
Llip3 =  4    if sg.running_mac() else 8
Lbp   = (4,3) if sg.running_mac() else (2,2)
LPColor  = dict(disabled_readonly_background_color=sg.DEFAULT_BACKGROUND_COLOR, disabled_readonly_text_color='white')

LastPassOptions = [
    [sg.Text("⩠ Login",justification="Left",size=(5, 1),font=("Helvetica",10+Fm),**ee,visible=True,k='LPLoginToggle'),sg.Text("LastPass",justification="center",size=(35+Sm*2, 1),font=(f"Helvetica {12+Fm} bold")),sg.VPush()],
    [sg.pin(sg.Column([
        [sg.Text('Please enter your LastPass Login Information')],
        [sg.Text('LastPass Username', size=(16+Sm, 1)), sg.Input(default_text=ConfigLPUser     ,key="LPUser",s=(45,1),enable_events=True)],
        [sg.Text('LastPass Password', size=(16+Sm, 1)), sg.Input(default_text=AppData['LPPass'],key="LPPass",s=(45,1),enable_events=True,password_char="*")],
        [sg.Button(button_text='Refresh Account List',key="LPRefresh"),sg.Push(),sg.Button(button_text='Login',key="LPLogin")],
    ],p=(0,(0,Llip)),k="LPLoginInfo",visible=False))],
    [sg.pin(sg.Column([
        [sg.Text('Entry Selection Options'),sg.Radio('Auto','LPEnSel',k='LPAuto',default=True,**ee),sg.Radio('Select Manually','LPEnSel',k='LPExist',**ee),sg.Radio('Create New','LPEnSel',k='LPNew',visible=False,**ee)],
    ],p=(0,(0,Llip2)),k="LPSelMethod",visible=True))],
    [sg.pin(sg.Column([
        [sg.Text('An entry will be automatically selected or created for you')],
        [sg.Text('LastPass Login Status:'),sg.Input('❌  You are not currently logged into LastPass',key="LPStatus",s=(39,1),disabled=True,**LPColor)],
        [sg.Text(' ')],
    ],p=(0,(0,Llip3)),k="LPAutoMethod",visible=False))],
    [sg.Text('Select Account(s) to sync:', size=(28+Sm+Sm, 1),p=((5,0),0)),sg.Text('Selected Accounts:', size=(29+Sm, 1),p=0)],
    [sg.Listbox(values=LPList,k='LPList',s=(30,10),p=((5,0),0),enable_events=True,**InputColor),sg.Listbox(values=ConfigLPAccts,k='LPSelected',s=(30,10),p=((0,4),0),enable_events=True,**InputColor)],
    [sg.Button(button_text='Choose For Me',key="LPAutoChoose",visible=False,p=(5,0)),sg.Push(),sg.Button(button_text='Clear All',key="LPClear",p=(5,Lbp))]
]
#endregion

#region Chrome
Tcs = 3 if sg.running_mac() else 0
Tr  = 6 if sg.running_mac() else 5
Tp  = 5 if sg.running_mac() else 9
GCOptions = [
    [sg.Text("Chrome",justification="center",size=(Ls, 1),p=(5,(4,3)),font=(f"Helvetica {12+Fm} bold"))],
    [sg.Text('Select Google Chrome password(s) to sync:')],
    [sg.Table(GCTable,['URL','Username']    ,auto_size_columns=False,justification='left'  ,visible_column_map=[True,True,False] ,num_rows=Tr, k='GCTable'   ,col_widths=[24+Tcs, 25+Tcs],enable_events=True,**InputColor)],
    [sg.Column([   
        [sg.Table([],['Selected Passwords'],auto_size_columns=False,justification='center',visible_column_map=[True,False,False],num_rows=5, k='GCSelected',col_widths=[37+Tcs*2],enable_events=True,**InputColor)],
    ],p=(0,(0,Tp))),sg.Column([
        [sg.Push(),sg.Button(button_text='Select All',p=(5,(0,3)),key="GCSelect")],
        [sg.Text(' ',s=(10,2))],
        [sg.Button(button_text='Clear Selection',p=(5,(0,3)),key="GCClear")]
    ],p=(0,(0,Tp)))]
]
#endregion

#region Edge
MEOptions = [
    [sg.Text("Edge",justification="center",size=(Ls, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.Text('Select Microsoft Edge password(s) to sync:')],
    [sg.Table(METable,['URL','Username']    ,auto_size_columns=False,justification='left'  ,visible_column_map=[True,True,False] ,num_rows=Tr, k='METable'   ,col_widths=[24+Tcs, 25+Tcs],enable_events=True,**InputColor)],
    [sg.Column([
        [sg.Table([],['Selected Passwords'],auto_size_columns=False,justification='center',visible_column_map=[True,False,False],num_rows=5, k='MESelected',col_widths=[37+Tcs*2],enable_events=True,**InputColor)]
    ],p=(0,(0,Tp+1))),sg.Column([
        [sg.Push(),sg.Button(button_text='Select All',p=(5,(0,3)),key="MESelect")],
        [sg.Text(' ',s=(10,2))],
        [sg.Button(button_text='Clear Selection',p=(5,(0,3)),key="MEClear")]
    ],p=(0,(0,Tp+1)))]
]
#endregion

#region CyberArk
Pd  = 0 if sg.running_mac() else 5
Pcl = 2 if sg.running_mac() else 7
Tm  = 1 if sg.running_mac() else 0
CAeReason = 'Enter business justification for checking out account'
CAc = dict(disabled_readonly_background_color=sg.DEFAULT_BACKGROUND_COLOR, disabled_readonly_text_color='white')
CAsf = dict(font=("Helvetica",9+Fm))
CAOptions = [
    [sg.Text("⩠ Login",justification="Left",size=(5, 1),font=("Helvetica",10+Fm),enable_events=True,visible=True,k='CALoginToggle'),sg.Text("CyberArk",justification="Left",size=(21+Sm*2, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.pin(sg.Column([
        [sg.Text('Please enter your CyberArk login information')],
        [sg.Text('Username:', size=(8, 1),p=((5,0),3),tooltip=UserNameToolTip),sg.Input(default_text=ConfigUsername    ,k="ArkUser",tooltip=UserNameToolTip,s=(17+Sm,1),p=((5,5),3)),
        sg.Text('Password:' , size=(8, 1),p=((5,0),3),tooltip=PasswordToolTip),sg.Input(default_text=AppData['ArkPass'],k="ArkPass",tooltip=PasswordToolTip,s=(17+Sm,1),password_char="*"),
        sg.Push(),sg.Button(button_text='Login',key="CALogin")]
    ],p=(0,(0,Pcl)),k="ArkLoginInfo",visible=False))],
    [sg.Text('Select an Account to check-out and sync:', size=(32,1),p=((5,0),0)),sg.Push(),sg.Text('Page Size:',p=(5,0),**CAsf),sg.Combo(['25','50','100','500','1000'],'100',s=(4, 1),p=(5,0),k='CAPageSize',**CAsf)], #,sg.Button('Refresh',k="CARefresh",font=('Helvetica',10))],
    [sg.Table(CATable,['','Username','Address','FacilityName'],auto_size_columns=False,justification='center',**ee,visible_column_map=[True,True,True,True] ,num_rows=9+Tm, 
        k='CATable',col_widths=[3,14+Sm,12+Sm,20+Sm+Tm],p=(5,(3,0)),**InputColor,tooltip=BGUserToolTip,select_mode=sg.TABLE_SELECT_MODE_BROWSE)
    ],
    [sg.Combo(['Contains','StartsWith'],'Contains',p=((5,5),3),s=(8, 1),k='CAsType',**CAsf),sg.Input('Search for accounts',key="ArkQuery",s=(32+Tm,1)),sg.Button('Search',k="CASearch",bind_return_key=True,target='ArkQuery'),sg.VSep()
    ,sg.Button('◀',k="CA◀",p=((5,0),3)),sg.Input('1',k='CAPage',s=(2,1),p=((3,3),3),disabled=True,**CAc),sg.Button('▶',k="CA▶",p=((0,5),3))],
    [sg.Text('Reason:', size=(6+Sm, 1),p=((5,1),(3,Pd)),tooltip=BGUserToolTip),sg.Input(CAeReason,key="CAReason",p=((6,5),(3,Pd)),s=(45,2)),sg.Input('Case #',key="CACase",s=(9,1),p=(5,(3,Pd)))]
]
#endregion

#region CyberArk Account Details
AcctColor  = dict(background_color=sg.DEFAULT_BACKGROUND_COLOR,text_color='white')
DAcctColor = dict(disabled_readonly_background_color=sg.DEFAULT_BACKGROUND_COLOR, disabled_readonly_text_color='white')
AcctFont   = dict(font=(f"Helvetica {10+Fm} bold"))
Cpm = 7 if sg.running_mac() else 0
CAAcctDetails = [
    [sg.Text('Username'      , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CAUser"    ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Address'       , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CAAddress" ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('SystemName'    , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CASystem"  ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Facility'      , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CAFacility",readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Client'        , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CAClient"  ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Tool'          , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CATool"    ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('SupportEmail'  , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CAEmail"   ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Locked By'     , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CALockBy"  ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Last Used By'  , size=(11, 1)), sg.Input(s=(50+Sm,1),k="CAUsedBy"  ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Last Used Date', size=(11, 1)), sg.Input(s=(50+Sm,1),k="CAUsedDate",readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
]

CAAcctTable = [
    [sg.Table([],['Date/Time','User','Action'],auto_size_columns=False,justification='center'  ,visible_column_map=[True,True,True] ,num_rows=14, 
        k='CAAcctTable',col_widths=[16+Sm,16+Sm,17+Sm],p=((5,6+Cpm),(6,10)),tooltip=BGUserToolTip,select_mode=sg.TABLE_SELECT_MODE_BROWSE)]
]
CATabs = [
    [sg.TabGroup([
        [sg.Tab('Account',CAAcctDetails,k='CAAcct'),sg.Tab('Activities',CAAcctTable,k='CAActivities')]
    ],p=0,enable_events=True,k='CATabs')]
]
#endregion

#region OnlyKey
Cs = ConfigSlots
KWLabelSearchText = [[sg.Text('Label Keyword', size=(15, 1),key="LabelKeyword",tooltip=KeywordToolTip), sg.InputText(default_text=ConfigOKWord,key="OK_Keyword",tooltip=KeywordToolTip,s=(44,1))]]
KWTrue  = ConfigKWTrue
try:
    onlykey = OnlyKey(connect=True,tries=1)
    onlykey.read_bytes(timeout_ms=1000)
    OKSlots = onlykey.getlabels()
    onlykey.close()
except BaseException as Error:
    print(Error)
    OKSlots = None

if OKSlots:
    sl = {Slot : [(Slot+" "+_.label if _.label and _.label != 'ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ' else Slot) for _ in OKSlots if _.name == Slot][0]  for Slot in Slots}
else:
    sl = {Slot : Slot for Slot in Slots}

Ss = 18 if sg.running_mac() else 16
ComboTrue = ConfigLPFrom
Sf = (('Helvetica',8) if ComboTrue else ('Helvetica',10)) if not sg.running_mac() else ('Helvetica',12)

OkCheckLeft  = [[sg.Checkbox(sl['1a'],k='S_1a',font=Sf,default=Cs['1a'],s=(Ss,1),p=((3,1),(0 ,0)),enable_events=True)] , [sg.Checkbox(sl['1b'],k='S_1b',font=Sf,default=Cs['1b'],s=(Ss,1),p=((3,1),0),enable_events=True)],
                [sg.Checkbox(sl['3a'],k='S_3a',font=Sf,default=Cs['3a'],s=(Ss,1),p=((3,1),(22,0)),enable_events=True)] , [sg.Checkbox(sl['3b'],k='S_3b',font=Sf,default=Cs['3b'],s=(Ss,1),p=((3,1),0),enable_events=True)],
                [sg.Checkbox(sl['5a'],k='S_5a',font=Sf,default=Cs['5a'],s=(Ss,1),p=((3,1),(22,0)),enable_events=True)] , [sg.Checkbox(sl['5b'],k='S_5b',font=Sf,default=Cs['5b'],s=(Ss,1),p=((3,1),0),enable_events=True)]]
OkCheckRight = [[sg.Checkbox(sl['2a'],k='S_2a',font=Sf,default=Cs['2a'],s=(Ss,1),p=((3,1),(0 ,0)),enable_events=True)] , [sg.Checkbox(sl['2b'],k='S_2b',font=Sf,default=Cs['2b'],s=(Ss,1),p=((3,1),0),enable_events=True)],
                [sg.Checkbox(sl['4a'],k='S_4a',font=Sf,default=Cs['4a'],s=(Ss,1),p=((3,1),(22,0)),enable_events=True)] , [sg.Checkbox(sl['4b'],k='S_4b',font=Sf,default=Cs['4b'],s=(Ss,1),p=((3,1),0),enable_events=True)],
                [sg.Checkbox(sl['6a'],k='S_6a',font=Sf,default=Cs['6a'],s=(Ss,1),p=((3,1),(22,0)),enable_events=True)] , [sg.Checkbox(sl['6b'],k='S_6b',font=Sf,default=Cs['6b'],s=(Ss,1),p=((3,1),0),enable_events=True)]]

LPOKMap = ConfigMap
Cf = ('Helvetica',9) if not sg.running_mac() else ('Helvetica',12) 
Scs = 10 if sg.running_mac() else 11
OKComboLeft  = [[sg.Combo(LPAccts,k='C_1a',font=Cf,default_value=LPOKMap['1a'],s=(Scs,1),p=((0,2),(0 ,0)),disabled=not Cs['1a'],**InputColor)] , [sg.Combo(LPAccts,k='C_1b',font=Cf,default_value=LPOKMap['1b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['1b'],**InputColor)],
                [sg.Combo(LPAccts,k='C_3a',font=Cf,default_value=LPOKMap['3a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['3a'],**InputColor)] , [sg.Combo(LPAccts,k='C_3b',font=Cf,default_value=LPOKMap['3b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['3b'],**InputColor)],
                [sg.Combo(LPAccts,k='C_5a',font=Cf,default_value=LPOKMap['5a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['5a'],**InputColor)] , [sg.Combo(LPAccts,k='C_5b',font=Cf,default_value=LPOKMap['5b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['5b'],**InputColor)]]
OKComboRight = [[sg.Combo(LPAccts,k='C_2a',font=Cf,default_value=LPOKMap['2a'],s=(Scs,1),p=((0,2),(0 ,0)),disabled=not Cs['2a'],**InputColor)] , [sg.Combo(LPAccts,k='C_2b',font=Cf,default_value=LPOKMap['2b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['2b'],**InputColor)],
                [sg.Combo(LPAccts,k='C_4a',font=Cf,default_value=LPOKMap['4a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['4a'],**InputColor)] , [sg.Combo(LPAccts,k='C_4b',font=Cf,default_value=LPOKMap['4b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['4b'],**InputColor)],
                [sg.Combo(LPAccts,k='C_6a',font=Cf,default_value=LPOKMap['6a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['6a'],**InputColor)] , [sg.Combo(LPAccts,k='C_6b',font=Cf,default_value=LPOKMap['6b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['6b'],**InputColor)]]

OkCheckText = [[sg.Text('Select Slots to Update:', size=(15,2),p=((5,0),5))],[sg.Button(button_text='Refresh Labels',key="OKRefresh")],[sg.Text(" ",size=(1,6))],
               [sg.Button(button_text='Clear Selection',key="OKClear")]]

Osp = 45 if sg.running_mac() else 0 
SlotSelector = [
    [sg.Col(OkCheckText,p=((0,Osp),0),visible=not ComboTrue,k='OkCheckText'),sg.Fr('OnlyKey Slots',
        layout=[
            [sg.Col(OkCheckLeft,p=0),sg.Col(OKComboLeft,p=0,visible=ComboTrue,k='OKComboLeft'),sg.VSep(pad=(0,5)),
            sg.Col(OkCheckRight,p=0),sg.Col(OKComboRight,p=0,visible=ComboTrue,k='OKComboRight')
        ]])]
    ]
SlotSelMethod = [[sg.Text('Slot Selction Method', size=(16+Sm, 1),p=((5,Osp),0)),sg.Push(),sg.Radio('Pick Slots','SMethod',k="PSlotsRadio",enable_events=True,default=not KWTrue),
    sg.Radio('Keyword Search','SMethod',k="KWSearch",default=KWTrue,enable_events=True),sg.Push()]]

Clrp = 70 if sg.running_mac() else 0 
OkSlotSelection = [
    [sg.Text("Onlykey",justification="center",size=(Ls, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.pin(sg.Col(SlotSelMethod,k='SlotSelMethod',p=0,visible=not ComboTrue))],
    [sg.pin(sg.Col([[sg.Text('Map LastPass Accounts Selected Above to Onlykey slots below'),sg.Push(),sg.Button(button_text='Clear All',key="OKClear2",p=((Clrp,5),0))]],k='OK_LPAcctText',p=0,visible=ComboTrue))],
    [collapse(KWLabelSearchText, "KWLabelSearchText",KWTrue)],
    [sg.pin(sg.Column(SlotSelector,k="SlotSelector",p=0,visible=not KWTrue))]
]
#endregion

#region LastPass Account Details
LVmp = 11 if sg.running_mac() else 0
LVAcctDetails = [
    [sg.Text('Name'    , size=(7, 1)), sg.Input(s=(40,1),k="LVName" ,readonly=True,**AcctColor,**DAcctColor,**AcctFont),sg.Button(button_text='Type For Me',p=((5,4),3),k="LVPaste")],
    [sg.Text('Group'   , size=(7, 1)), sg.Input(s=(40,1),k="LVGroup",readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Username', size=(7, 1)), sg.Input(s=(40,1),k="LVUser" ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Password', size=(7, 1)), sg.Input(s=(40,1),k="LVPass" ,readonly=True,**AcctColor,**DAcctColor,**AcctFont,password_char="*"),sg.Button(button_text='Show',p=((5,4),3),k="LVShow"),sg.Button(button_text='Copy',p=((5,4),3),k="LVCopy")],
    [sg.Text('URL'     , size=(7, 1)), sg.Input(s=(55,1),k="LVUrl"  ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Notes',p=(5,(3,0)), size=(5, 1))],
    [sg.Multiline(s=(63,7),p=((3,3),(1,LVmp)),k="LVNote",write_only=True,disabled=True,**AcctColor,**AcctFont)]
]
#endregion

#region Local vault

Vtp = 8  if sg.running_mac() else 5
Vip = 0  if sg.running_mac() else 3
Vcp = 0  if sg.running_mac() else 4
Vr1 = 21 if sg.running_mac() else 20
Vr2 = 12 if sg.running_mac() else 11
VaultLayout = [
    [sg.Text("⩠ Login",justification="Left",size=(5, 1),font=("Helvetica",10+Fm),enable_events=True,visible=True,k='LVLoginToggle'),sg.Text("Local Vault",justification="Center",size=(21+Sm*2, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.pin(sg.Column([
        [sg.Text('Please enter your LastPass Login Information',p=(5,(Vip,2)))],
        [sg.Text('LastPass Username', size=(16+Sm, 1)), sg.InputText(default_text=ConfigLPUser,key="LVlUser",s=(45+Im,1),**InputColor)],
        [sg.Text('LastPass Password', size=(16+Sm, 1)), sg.InputText(password_char="*",key="LVlPass",default_text=AppData['LPPass'],s=(45+Im,1),**InputColor)],
        [sg.Button(button_text='Refresh Account List',key="LVRefresh"),sg.Push(),sg.Button(button_text='Login',key="LVLogin")],
    ],p=(0,(0,Vcp)),k="LVLoginInfo",visible=False))],
    [sg.Table(LVTable,['Name','Username','Url','Id','Group','Notes','Password',],auto_size_columns=False,justification='left'  ,visible_column_map=[True,True,False,False,False,False,False] 
     ,num_rows=Vr1, k='LVTable',col_widths=[24+Tcs, 25+Tcs],enable_events=True,select_mode=sg.TABLE_SELECT_MODE_BROWSE,**InputColor)],
    [sg.Combo(['Type','Password','SecureNote'],'Type',s=(10, 1),p=(5,(0,Vtp)),**ee,k='LVsType'),sg.Input('Search for accounts',key="LVQuery",p=(5,(0,Vtp)),s=(32,1),**ee),
     sg.Combo(LVGroups,'Group/Folder',s=(15, 1),p=(5,(0,Vtp)),k='LVsGroup',**ee)],
    [sg.TabGroup([
        [sg.Tab('Account Details',LVAcctDetails,k='LVAcctDetails'),sg.Tab('Output',OutPutSection2,k='OutPutSection2')]
    ],p=0,enable_events=True,k='Tabs')]
]
#endregion
#endregion

#region Main Application Window Layout
Mfp = 3 if sg.running_mac() else 3
LPFrom = ConfigLPFrom
OKTrue = ConfigOK
LPTrue = ConfigLP
CATrue = ConfigCAFrom
submit = 'Check-Out & Sync' if CATrue else "Sync Passwords"
#Layout
layout = [
    [sg.pin(sg.Column(CAOptions,p=0,visible=True ,k='CAOptions'))],
    [sg.pin(sg.Column(GCOptions,p=0,visible=False,k='GCOptions'))],
    [sg.pin(sg.Column(MEOptions,p=0,visible=False,k='MEOptions'))],
    [sg.pin(sg.Fr('Sync Options', 
        layout=[
        [sg.Text('Sync From:',p=(5,(Mfp,3)),size=(8+Sm, 1)),sg.Radio('CyberArk','PFrom',k='CAFrom',enable_events=True,p=(5,(Mfp,3)),default=CATrue),sg.Radio('LastPass','PFrom',k='LPFrom',enable_events=True,p=(5,(Mfp,3)),default=LPFrom)
        ,sg.Radio('Chrome','PFrom',k='GCFrom',enable_events=True,p=(5,(Mfp,3)),default=False),sg.Radio('Edge','PFrom',k='MEFrom',enable_events=True,p=(5,(Mfp,3)),s=(9+Sm*4,1),default=False),],
        [sg.Text('Sync To:', size=(8+Sm, 1)),sg.pin(sg.Checkbox('LastPass',k='LP',enable_events=True,default=LPTrue,disabled=not OKTrue,visible=not LPFrom)),
        sg.pin(sg.Checkbox('OnlyKey',k='OK',enable_events=True,default=OKTrue,disabled=not LPTrue or LPFrom))],
    ],p=((5,0),(0,1)),k="SyncOpt"))],
    [sg.Button(button_text='Save',key="Save",tooltip=SaveToolTip),sg.Text('⩠ Sync Options',k='SyncOptToggle',**ee),sg.Push(),sg.pin(sg.Button(button_text='Check In',key="CheckIn"))
        ,sg.Submit(submit,tooltip=SubmitToolTip,p=((5,10),3),k='Submit')],
    [sg.HSep()], 
    [sg.TabGroup([
        [sg.Tab('Output',OutPutSection,k='OutPutSection'),sg.Tab('Account Details',CATabs,k='CAAcctDetails'),sg.Tab('LastPass',LastPassOptions,k='LPOptions',visible=LPTrue and not LPFrom), 
        sg.Tab('OnlyKey', OkSlotSelection,k='OKOptions',visible=OKTrue)]
    ],p=0,enable_events=True,k='Tabs')]
]

MainTabs = [
    [sg.TabGroup([
        [sg.Tab('Check-Out & Sync',layout,k='SyncPasswords'),sg.Tab('Local Vault',VaultLayout,k='LocalVault',visible=True)]
    ],p=0,enable_events=True,k='MainTabs')]
]

#alternate up/down arrows ⩢⩠▲▼⩓⩔  locks 🔓🔒❎☒☑⚿⛔🗝️🔑🔐⛓️✔️🚫✅
#Window and Titlebar Options
Title = 'MEDHOST Password Manager beta'

size = (501, 673) if sg.running_mac() else (497, 760)
window = sg.Window(Title, MainTabs,font=("Helvetica", 10+Fm),finalize=True,return_keyboard_events=True,icon=icon) #,size=size)
#endregion

#region Opening Application Tasks
timeout=100
WindowRead = window.read(timeout=0)
window.DisableDebugger()
event      = WindowRead[0]
Inputs     = WindowRead[1]
OKVisible  = OKTrue
CAVisible  = CATrue
ArkAccountPass = None
window['LPFrom'].Value = LPFrom
window['CAFrom'].Value = CATrue
window.metadata = {'LastUpdate':datetime.now()}
window['CASearch'].metadata = {'Searching':False,'LastUpdate':datetime.now(),'SearchStart':datetime.now()}
window['LVShow'].metadata = {'LastUpdate':datetime.now(),'Shown':False}
# window['-OUTPUT-2'].reroute_stdout_to_here()
# window['-OUTPUT-2'].reroute_stderr_to_here()
# window['-OUTPUT-'].reroute_stdout_to_here()
# window['-OUTPUT-'].reroute_stderr_to_here()

if not ConfigSyncOpt:
    window.perform_long_operation(lambda:EventHelper('SyncOptToggle'),'SyncOptToggle')

if ConfigCALogin and not window['ArkLoginInfo']._visible:
    numrows = window['CATable'].NumRows - 4
    window['CATable'].NumRows = numrows
    if sg.running_mac(): window['CATable'].update(num_rows=numrows)
    window['ArkLoginInfo'].update(visible=not window['ArkLoginInfo']._visible)
    window['CALoginToggle'].update(value='⩢ Login')
    if not sg.running_mac(): window['CATable'].update(num_rows=numrows)
    
if ConfigLPLogin:
    window['LPList'].Size = (30,5)
    window['LPSelected'].Size = (30,5)
    if sg.running_mac(): 
        window['LPList'].set_size((30,5))
        window['LPSelected'].set_size((30,5))
    window['LPLoginInfo'].update(visible=not window['LPLoginInfo']._visible)
    window['LPLoginToggle'].update(value='⩢ Login')
    if sg.running_mac(): window['LPSelMethod'].update(visible=False)
    if not sg.running_mac():
        window['LPSelMethod'].update(visible=False) 
        window['LPList'].set_size((30,5))
        window['LPSelected'].set_size((30,5))
        
elif Inputs['LPAuto']:       
    window['LPList'    ].Size = (30,5)
    window['LPSelected'].Size = (30,5)
    window['LPList'    ].set_size((30,5))
    window['LPSelected'].set_size((30,5))
    window['LPAutoMethod'].update(visible=True)
    if LastPass and LastPass.accounts:
        window['LPStatus'].update(value=f'✅ You are logged into LastPass')
  
if Inputs['LPAuto']:       
    window['LPList'    ].update(disabled=True)
    window['LPList'    ].Disabled=True
    window['LPSelected'].update(disabled=True)
    window['LPClear'   ].update(disabled=True)

if LPTrue:
    window['LPOptions'].select()
else:
    window['OKOptions'].select()

CALoginAttempt = 0
if LPFrom:
    window = SwapContainers(window,'LPOptions','CAOptions')
    
if ConfigCAIndex and ConfigCATable and len(ConfigCATable) >= ConfigCAIndex[0]:
    try: 
        window['CATable'].update(select_rows=ConfigCAIndex)
        ScrollPosition = ConfigCAIndex[0] / (len(ConfigCATable) - 1) if ConfigCAIndex[0] != 0 else 0
        window['CATable'].set_vscroll_position(ScrollPosition)
    except:
        pass
if not sg.running_mac():
    window.perform_long_operation(lambda:HotKeyListener(),'HotKeyReturn')  
#endregion

#region Application Behaviour if Statments - Event Loop    
while True:
    try:
        #Read Events(Actions) and Users Inputs From Window 
        WindowRead = window.read(timeout=timeout)
        event      = WindowRead[0]
        Inputs     = WindowRead[1]
        
        #End Script when Window is Closed
        if event == sg.WIN_CLOSED:
            break
        
        if event in ['HotKeyReturn','LVPaste','typesleepreturn','pastesleepreturn']:
            if event in ['HotKeyReturn','typesleepreturn']:
                seconds = 0.3
                if event != 'typesleepreturn':
                    window.perform_long_operation(lambda:sleep(seconds), 'typesleepreturn')
                    continue
                typestring(Inputs['LVPass'])
                if not sg.running_mac(): window.perform_long_operation(lambda:HotKeyListener(),'HotKeyReturn')
  
            if event in ['LVPaste','pastesleepreturn']:
                seconds = 2 
                if event != 'pastesleepreturn':
                    window.perform_long_operation(lambda:sleep(seconds), 'pastesleepreturn')
                    continue
                typestring(Inputs['LVPass'])

        if event:
            if event != '__TIMEOUT__':
                timeout = 100
                window.metadata['LastUpdate'] = datetime.now()
            elif window.metadata['LastUpdate'] + timedelta(minutes=5) < datetime.now():
                timeout = 1000
        
        try:
            if window.find_element_with_focus():
                element = window.find_element_with_focus()
                if element == window['CAReason'] and window['CAReason'].get() == 'Enter business justification for checking out account':
                    window['CAReason'].update(value='')
                if element != window['CAReason'] and window['CAReason'].get() == '':
                    window['CAReason'].update(value='Enter business justification for checking out account')
                    
                if element == window['CACase'] and window['CACase'].get() == 'Case #':
                    window['CACase'].update(value='')
                if element != window['CACase'] and window['CACase'].get() == '':
                    window['CACase'].update(value='Case #')
                    
                if element == window['ArkQuery'] and window['ArkQuery'].get() == 'Search for accounts':
                    window['ArkQuery'].update(value='')
                if element != window['ArkQuery'] and window['ArkQuery'].get() == '':
                    window['ArkQuery'].update(value='Search for accounts')
                    
                if element == window['LVQuery'] and window['LVQuery'].get() == 'Search for accounts':
                    window['LVQuery'].update(value='')
                if element != window['LVQuery'] and window['LVQuery'].get() == '':
                    window['LVQuery'].update(value='Search for accounts')
        except:
            pass
        
        if window['CASearch'].metadata['Searching'] and window['CASearch'].metadata['LastUpdate'] + timedelta(seconds=1) < datetime.now():
            window['CASearch'].metadata['LastUpdate'] = datetime.now()
            CATable = window['CATable'].get()
            if CATable[4][2] != 'Searching..........' and 'Searching' in CATable[4][2]:
                window['CATable'].update(values=[[],[],[],[],['','',CATable[4][2] + '.']])
            else:
                window['CATable'].update(values=[[],[],[],[],['','','Searching']])

        if event == 'SyncOptToggle':
            if window['ArkLoginInfo']._visible:
                numrows = window['CATable'].NumRows + 4
                window['CATable'].NumRows = numrows
                if sg.running_mac(): window['CATable'].update(num_rows=numrows)
                window['CALoginToggle'].update(value='⩠ Login')
                window['ArkLoginInfo'].update(visible=not window['ArkLoginInfo']._visible)
                if not sg.running_mac(): window['CATable'].update(num_rows=numrows)
            numrows = window['CATable'].NumRows
            if window['SyncOpt']._visible:
                numrows = numrows + 5
                window['CATable'].NumRows = numrows
                if not sg.running_mac(): window['CAOptions'].set_size((470, 357)) 
                window['CATable'].update(num_rows=numrows)
                window['SyncOpt'].update(visible=False)
                window['SyncOptToggle'].update(value='⩢ Sync Options')
            elif Inputs and 'SyncOptToggle' not in Inputs or not Inputs['SyncOptToggle']:
                numrows = numrows - 5
                window['CATable'].NumRows = numrows
                if not sg.running_mac(): window['CAOptions'].set_size((470, 277))
                window['CATable'].update(num_rows=numrows)
                window['SyncOpt'].update(visible=True)                    
                window['SyncOptToggle'].update(value='⩠ Sync Options')    

        if window['-OUTPUT-'].get().rstrip() == '' and window['-OUTPUT-'].get().rstrip() != window['-OUTPUT-2'].get().rstrip():
            window['-OUTPUT-2'].update(value='')

        if event == 'CATable' and not window['CAAcctDetails'].Disabled:
            CATable = window['CATable'].get()
            window['CAAcctDetails'].select()
            if CATable and (Inputs['CATable'] or Inputs['CATable'] == 0) and isinstance(Inputs['CATable'][0],int):
                CAAcct  = CATable[Inputs['CATable'][0]]
                Useddate = getdate(int(CAAcct[12])) if CAAcct[11] else ''
                window["CAUser"    ].update(value=CAAcct[1])
                window["CAAddress" ].update(value=CAAcct[2])
                window["CASystem"  ].update(value=CAAcct[5])
                window["CAFacility"].update(value=CAAcct[3])
                window["CAClient"  ].update(value=CAAcct[4])
                window["CATool"    ].update(value=CAAcct[6])
                window["CAEmail"   ].update(value=CAAcct[7])
                window["CALockBy"  ].update(value=CAAcct[10])
                window["CAUsedBy"  ].update(value=CAAcct[11])
                window["CAUsedDate"].update(value=Useddate)
                Activities = GetCyberArkActivities(ArkToken,CAAcct[13])
                if 'Activities' in Activities:
                    CAActTable = [[getdate(_['Date']),_['User'],_['Action']] for _ in Activities['Activities']]
                    window['CAAcctTable'].update(values=CAActTable)

        if event == 'LPAuto':
            window['LPSelected'].update(values=[])  
            window['LPList'    ].Size = (30,5)
            window['LPSelected'].Size = (30,5)
            window['LPList'    ].update(disabled=True)
            window['LPList'    ].Disabled=True
            window['LPSelected'].update(disabled=True)
            window['LPClear'   ].update(disabled=True)
            window['LPList'    ].set_size((30,5))
            window['LPSelected'].set_size((30,5))
            window['LPAutoMethod'].update(visible=True)
            if LastPass and LastPass.accounts:
                window['LPStatus'].update(value=f'✅ You are logged into LastPass')
            else:
                window['LPStatus'].update(value=f'❌  You are not currently logged into LastPass')
                
                
        if event == 'LPExist':
            window['LPList'    ].Size = (30,10)
            window['LPSelected'].Size = (30,10)
            window['LPList'    ].update(disabled=False)
            window['LPList'    ].Disabled=False
            window['LPSelected'].update(disabled=False)
            window['LPClear'   ].update(disabled=False)
            if not sg.running_mac(): window['LPAutoMethod'].update(visible=False)
            window['LPList'    ].set_size((30,10))
            window['LPSelected'].set_size((30,10))
            window['LPAutoMethod'].update(visible=False)
            
        if event == 'LVCopy':
            sg.clipboard_set(window['LVPass'].get())
            
        if event in ['LVQuery','LVsGroup','LVsType']:
            LVQTable = LVTable
            if Inputs['LVQuery'] and Inputs['LVQuery'] != 'Search for accounts':
                LVQTable = [_ for _ in LVQTable if Inputs['LVQuery'] in ''.join([str(x) for x in _]) + ''.join([str(x).title() for x in _]) + ''.join([str(x).lower() for x in _])]
            if Inputs['LVsGroup'] and Inputs['LVsGroup'] != 'Group/Folder':
                LVQTable = [_ for _ in LVQTable if _[4].split('\\')[-1] == Inputs['LVsGroup']]
            if Inputs['LVsType'] and Inputs['LVsType'] != 'Type':
                if Inputs['LVsType'] == 'Password':
                    LVQTable = [_ for _ in LVQTable if _[2] != 'http://sn']
                elif Inputs['LVsType'] == 'SecureNote':
                    LVQTable = [_ for _ in LVQTable if _[2] == 'http://sn']
                    
            window['LVTable'].update(values=LVQTable)
        
        if event in ["LPUser","LPPass"]:
            window['LVlUser'].update(value=Inputs['LPUser'])
            window['LVlPass'].update(value=Inputs['LPPass'])
        
        if event in ['CALoginToggle','CALogin',"CALoginSearch"]:
            numrows = window['CATable'].NumRows
            if window['ArkLoginInfo']._visible:
                numrows = numrows + 4
                window['CATable'].NumRows = numrows
                if sg.running_mac(): window['CATable'].update(num_rows=numrows)
                window['CALoginToggle'].update(value='⩠ Login')
                window['ArkLoginInfo'].update(visible=not window['ArkLoginInfo']._visible)
                if not sg.running_mac(): window['CATable'].update(num_rows=numrows)
            elif not window['ArkLoginInfo']._visible and event not in ['CALogin',"CALoginSearch"]:
                numrows = numrows - 4
                window['CATable'].NumRows = numrows
                if sg.running_mac(): window['CATable'].update(num_rows=numrows)
                window['ArkLoginInfo'].update(visible=not window['ArkLoginInfo']._visible)
                window['CALoginToggle'].update(value='⩢ Login')
                if not sg.running_mac(): window['CATable'].update(num_rows=numrows)
                
        if event in ['LVLogin','LVRefresh']:
            window['LPUser'].update(value=Inputs['LVlUser'])
            window['LPPass'].update(value=Inputs['LVlPass'])
            window['OutPutSection2'].select()
            window['LVTable'].NumRows = Vr1
            if sg.running_mac(): window['LVTable'].update(num_rows=Vr1)
            window['LVLoginToggle'].update(value='⩠ Login')
            window['LVLoginInfo'].update(visible=not window['LVLoginInfo']._visible)
            if not sg.running_mac(): window['LVTable'].update(num_rows=Vr1)
        

        if event == 'LVLoginToggle':
            numrows = window['LVTable'].NumRows
            if window['LVLoginInfo']._visible:
                numrows = numrows + 7
                window['LVTable'].NumRows = numrows
                if sg.running_mac(): window['LVTable'].update(num_rows=numrows)
                window['LVLoginToggle'].update(value='⩠ Login')
                window['LVLoginInfo'].update(visible=not window['LVLoginInfo']._visible)
                if not sg.running_mac(): window['LVTable'].update(num_rows=numrows)
            else:
                numrows = numrows - 7
                window['LVTable'].NumRows = numrows
                if sg.running_mac(): window['LVTable'].update(num_rows=numrows)
                window['LVLoginInfo'].update(visible=not window['LVLoginInfo']._visible)
                window['LVLoginToggle'].update(value='⩢ Login')
                if not sg.running_mac(): window['LVTable'].update(num_rows=numrows)

        if window['LVShow'].metadata['Shown'] and window['LVShow'].metadata['LastUpdate'] + timedelta(seconds=60) < datetime.now():
            event = 'LVShow'

        if event == 'LVShow':
            window['LVShow'].metadata['LastUpdate'] = datetime.now()
            Table    = window['LVTable']
            Table    = Table.Values if isinstance(Table,sg.Table) else []
            if window['LVUrl'].get() == 'http://sn': LVAcct = Table[Inputs['LVTable'][0]]
            if window['LVPass'].PasswordCharacter == '*':
                window['LVShow'].metadata['Shown']      = True
                window['LVPass'].update(password_char='')
                if window['LVUrl'].get() == 'http://sn': window['LVNote'].update(value=LVAcct[5])
            else:
                window['LVShow'].metadata['Shown'] = False
                window['LVPass'].update(password_char='*')
                if window['LVUrl'].get() == 'http://sn':
                    window['LVNote'].update(value=len(LVAcct[5])*'*')                

        if event == 'LVTable':
            window['LVAcctDetails'].select()
            Table    = window['LVTable']
            Table    = Table.Values if isinstance(Table,sg.Table) else []
            window['LVPass'].update(password_char='*')
            if Inputs['LVTable'] and Table:
                LVAcct = Table[Inputs['LVTable'][0]]
                window['LVName' ].update(value=LVAcct[0])
                window['LVGroup'].update(value=LVAcct[4])
                window['LVUser' ].update(value=LVAcct[1])
                window['LVPass' ].update(value=LVAcct[6])
                window['LVUrl'  ].update(value=LVAcct[2])
                if LVAcct[2] == 'http://sn':
                    window['LVNote'].update(value=len(LVAcct[5])*'*')
                else:
                    window['LVNote'].update(value=LVAcct[5])

        if isinstance(event,str) and event.startswith('S_'):
            Slot = event.replace('S_','')
            window[f'C_{Slot}'].Disabled = not Inputs[event]
            window[f'C_{Slot}'].update(disabled=not Inputs[event])
            
        if window['LPFrom'].Value == True and Inputs['LPFrom'] == False:
            window['LPOptions']._visible=Inputs['LP']
            window['LP'       ]._visible=True
            window['LP'       ].Disabled = not OKTrue
            window['OK'       ].Disabled = not Inputs['LP']
            window.ReturnValuesDictionary['OK'] = OKTrue 
            if event != 'CAFrom': 
                window['LPFrom'].Value = False
                window = SwapContainers(window,'LPOptions','CAOptions')
                
        if window['CAFrom'].Value == True and Inputs['CAFrom'] == False:
            window['CAFrom' ].Value = False
            window['Submit' ].ButtonText='Sync Passwords'
            if not sg.running_mac(): 
                for _ in Slots: window[f"S_{_}"].Font=('Helvetica',8)
            window['CAAcctDetails'    ].Disabled=True
            window['CAAcctDetails'    ]._visible=False              
            window['SyncOptToggle'    ]._visible=False                
            window['CheckIn'          ]._visible=False
            window['OkCheckText'      ]._visible=False  
            window['SlotSelMethod'    ]._visible=False
            window['KWLabelSearchText']._visible=False
            window['OK_LPAcctText'    ]._visible=True 
            window['OKComboLeft'      ]._visible=True 
            window['OKComboRight'     ]._visible=True 
            window['SlotSelector'     ]._visible=True 
            if event!= 'LPFrom': window = ReloadWindow(window)
            
        if event == 'GCFrom':
            window['GCOptions'].update(visible=Inputs['GCFrom'])
            window['MEOptions'].update(visible=Inputs['MEFrom'])
            window['CAOptions'].update(visible=Inputs['CAFrom'])
            GCTable    = window['GCSelected']
            GCTable    = GCTable.Values if isinstance(GCTable,sg.Table) else []
            Combolist = [_[0]+' '*50+str(_) for _ in GCTable]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)
            if Inputs['LP']: window['LPOptions'].select()
            else: window['OKOptions'].select()            

        if event == 'MEFrom':
            window['MEOptions'].update(visible=Inputs['MEFrom'])
            window['CAOptions'].update(visible=Inputs['CAFrom'])
            window['GCOptions'].update(visible=Inputs['GCFrom'])
            MESelected = window['MESelected']
            MESelected = MESelected.Values if isinstance(MESelected,sg.Table) else []
            Combolist = [_[0]+' '*50+str(_) for _ in MESelected]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)
            if Inputs['LP']: window['LPOptions'].select()
            else: window['OKOptions'].select()
            
        if event == 'CAFrom' and window['CAFrom'].Value == False:
            window['CAFrom'].Value = True
            if not sg.running_mac():
                for _ in Slots: window[f"S_{_}"].Font=('Helvetica',10)
            window['Submit'           ].ButtonText='Check-Out & Sync'
            window['CAOptions'        ]._visible=Inputs['CAFrom']
            window['MEOptions'        ]._visible=Inputs['MEFrom']
            window['GCOptions'        ]._visible=Inputs['GCFrom']
            window['CheckIn'          ]._visible=True
            window['SlotSelector'     ]._visible=Inputs['PSlotsRadio']
            window['KWLabelSearchText']._visible=Inputs['KWSearch']
            window['CAAcctDetails'    ].Disabled=False
            window['CAAcctDetails'    ]._visible=True
            window['SyncOptToggle'    ]._visible=True 
            window['SlotSelMethod'    ]._visible=True
            window['OkCheckText'      ]._visible=True
            window['OK_LPAcctText'    ]._visible=False
            window['OKComboLeft'      ]._visible=False
            window['OKComboRight'     ]._visible=False
            if window['LPFrom'].Value == True:
                window['LPFrom'].Value = False
                window = SwapContainers(window,'LPOptions','CAOptions')
            else:
                window = ReloadWindow(window)
            if Inputs['LP']: window['LPOptions'].select()
            else: window['OKOptions'].select()
  
        if event == 'LPFrom' and window['LPFrom'].Value== False:
            window['MEOptions']._visible=Inputs['MEFrom']
            window['GCOptions']._visible=Inputs['GCFrom']
            window['CAOptions']._visible=Inputs['LPFrom']
            window['LPFrom'   ].Value   = True
            window['LPOptions']._visible=False
            window['LP'       ]._visible=False
            window['OK'       ].Disabled=True
            OKTrue = Inputs['OK']
            window.ReturnValuesDictionary['OK'] = True
            window = SwapContainers(window,'LPOptions','CAOptions')
            LPSelected = window['LPSelected']
            LPSelected = LPSelected.Values if isinstance(LPSelected,sg.Listbox) else []
            for _ in Slots: 
                window[f"C_{_}"].update(values=LPSelected) 
            window.key_dict['OKOptions'].select()

        if event == 'LP':
            window['OK'].Disabled = not Inputs['LP']
            window['OK'].update(disabled=not Inputs['LP'])
            try:
                window['LPOptions'].update(visible=Inputs['LP'])
            except:
                window['LPOptions']._visible=Inputs['LP']
                window = ReloadWindow(window)
    
        if event == 'OK':
            window['OutPutSection'].select()
            window['LP'].Disabled = not Inputs['OK']
            window['LP'].update(disabled=not Inputs['OK'])
            window['OKOptions'].update(visible=not window['OKOptions'].visible)

        if isinstance(event,str) and event.startswith('OKClear'):
            for _ in Slots: 
                window[f"C_{_}"].update(value='')
                window[f"C_{_}"].update(disabled=True)
                window[f"C_{_}"].Disabled=True
            for _ in Slots: window[f"S_{_}"].update(value=False)    
                     
        if event == 'CARefresh':
            ArkToken   = AppData['ArkToken']
            ArkAccount = FindCyberArkAccount(Token=ArkToken)
            if not ArkAccount or 'ErrorMessage' in ArkAccount:
                window['CATable'].update(values=[['Login to see CyberArk Accounts here']])
                window['OutPutSection'].select()
                window['-OUTPUT-'].update('')
                if ArkAccount and 'ErrorMessage' in ArkAccount:
                    print(ArkAccount['ErrorMessage'])
                if Inputs['ArkUser'] and Inputs['ArkPass']:
                    print('Unable to refresh accounts list, attempting login to get new token')
                    window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "CALoginReturn")
                else:
                    print('Unable to refresh accounts list, login to get new token')
                continue
            else:
                CATable = MakeCyberArkTable(ArkAccts)
                window['CATable'].update(values=CATable)
                
        if event in ['CALogin','CALoginReturn','CALoginCheckIn',"CALoginSearch","CALoginAcctReturn"]:
            if event == 'CALogin':
                window['-OUTPUT-'].update('')
                window['OutPutSection'].select()
                window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "CALoginReturn")
            else:
                if event != "CALoginAcctReturn":
                    ArkToken = Inputs[event]
                    window.perform_long_operation(lambda:FindCyberArkAccount(ArkToken,Inputs['ArkQuery'],Inputs['CAsType'],Inputs['CAPageSize']), "CALoginAcctReturn")
                    window['CATable'].update(values=[[],[],[],[],['','','Searching']])
                    window['CASearch'].metadata = {'Searching':True,'LastUpdate':datetime.now(),'SearchStart':datetime.now()}
                    continue
                ArkAccts = Inputs["CALoginAcctReturn"]
                window['CATable'].update(values=[])
                window['CASearch'].metadata['Searching'] = False
                if not ArkAccts or 'ErrorMessage' in ArkAccts:
                    window['-OUTPUT-'].update('')
                    window['OutPutSection'].select()
                    if ArkAccts and 'ErrorMessage' in ArkAccts:
                        print(ArkAccts['ErrorMessage'])
                        window.perform_long_operation(lambda:EventHelper('CALoginToggle'),'CALoginToggle')
                    print('Login unsuccessful')
                    continue            
                CATable = MakeCyberArkTable(ArkAccts)
                SelectedRows = Inputs['CATable'] if CATable == window['CATable'].Values else []
                ScrollPosition = Inputs['CATable'][0] / (len(window['CATable'].Values) - 1) if Inputs['CATable'] and len(window['CATable'].Values) != 1 and CATable == window['CATable'].Values else 0
                window['CATable'].update(values=CATable)
                window['CATable'].update(select_rows=SelectedRows)
                window['CATable'].set_vscroll_position(ScrollPosition)
                ArkInfo = {'Token': ArkToken}
                SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')

        if event in ["CA◀","CA▶"]:
            if event == "CA▶":
                window['CAPage'].update(value=str(int(Inputs['CAPage']) + 1))
                Inputs['CAPage'] = str(int(Inputs['CAPage']) + 1)
                event = 'CASearchPage'
            if event == "CA◀" and Inputs['CAPage'] != '1':
                window['CAPage'].update(value=str(int(Inputs['CAPage']) - 1))
                Inputs['CAPage'] = str(int(Inputs['CAPage']) - 1)
                event = 'CASearchPage'

        if event in ['CASearch','CALoginSearch','CASearchPage','CASearchReturn']:
            if event != 'CASearchReturn':
                window.perform_long_operation(lambda:EventHelper('SyncOptToggle'),'SyncOptToggle')
                ArkToken = AppData['ArkToken'] if not ArkToken else ArkToken
                if event != 'CASearchPage': 
                    window['CAPage'].update(value=1)
                    Inputs['CAPage'] = 1
                window.perform_long_operation(lambda:FindCyberArkAccount(ArkToken,Inputs['ArkQuery'],Inputs['CAsType'],Inputs['CAPageSize'],Inputs['CAPage']),'CASearchReturn')
                window['CATable'].update(values=[[],[],[],[],['','','Searching']])
                window['CASearch'].metadata = {'Searching':True,'LastUpdate':datetime.now(),'SearchStart':datetime.now()}
                continue
            ArkAccts = Inputs['CASearchReturn']
            window['CATable'].update(values=[])
            window['CASearch'].metadata['Searching'] = False
            if 'ErrorMessage' in ArkAccts:
                window['OutPutSection'].select()
                window['-OUTPUT-'].update('')
                ErrorM = ArkAccts['ErrorMessage']
                print(ErrorM)
                if 'token' in ErrorM or 'logon' in ErrorM or 'terminated' in ErrorM:
                    if Inputs['ArkUser'] and Inputs['ArkPass']:
                        print('Attempting login to get new token')
                        window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "CALoginSearch")                 
                    else:
                        print('Fill in CyberArk username and password fields to login and get new token')
                        window.perform_long_operation(lambda:EventHelper('CALoginToggle'),'CALoginToggle')
                continue
            else:
                CATable = MakeCyberArkTable(ArkAccts) 
                SelectedRows = Inputs['CATable'] if CATable == window['CATable'].Values else []
                ScrollPosition = Inputs['CATable'][0] / (len(window['CATable'].Values) - 1) if Inputs['CATable'] and len(window['CATable'].Values) != 1 and CATable == window['CATable'].Values else 0
                window['CATable'].update(values=CATable)
                window['CATable'].update(select_rows=SelectedRows)
                window['CATable'].set_vscroll_position(ScrollPosition)            
                
        if event in ['CheckIn','CALoginCheckIn']:
            if event == 'CheckIn':window['-OUTPUT-'].update('')
            window['OutPutSection'].select()   
            if not Inputs['CATable']:
                print('You must select an Account in CyberArk to check-out')
                continue
            CATable = window['CATable']
            CATable = CATable.Values if isinstance(CATable,sg.Table) else []
            Acct  = CATable[Inputs['CATable'][0]]
            ArkToken = AppData['ArkToken'] if not ArkToken else ArkToken

            CheckInResponse = CheckInCyberArkPassword(AccountId=Acct[13], Token=ArkToken, UserName=Acct[1])
            
            if CheckInResponse and hasattr(CheckInResponse,'ok') and CheckInResponse.ok:
                print('Account successfully checked in')
                continue
            elif CheckInResponse.content:
                try:
                    Rjson = CheckInResponse.json()
                except:
                    Rjson = {}
                if 'ErrorMessage' in Rjson:
                    ErrorM = CheckInResponse.json()['ErrorMessage']
                    print(ErrorM)
                    if 'token' in ErrorM or 'logon' in ErrorM:
                        if Inputs['ArkUser'] and Inputs['ArkPass']:
                            print('Attempting login to get new token')
                            window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), 'CALoginCheckIn')                 
                        else:
                            print('You will need to login to CyberArk to check in an Account')
                        continue
            print('Unable to check in password')
    
        if event in ['LPRefresh','LVRefresh','LPRefreshReturn']:
            if event != 'LPRefreshReturn':
                window.perform_long_operation(lambda:LastPassSession(AppData), "LPRefreshReturn")
            else:
                LastPass = Inputs['LPRefreshReturn']
            if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                LPList = ['Login to see LastPass entries here']
                window['OutPutSection'].select()
                window['-OUTPUT-'].update('')
                if Inputs['LPUser'] and Inputs['LPPass']:
                    print('Unable to refresh accounts list, attempting login to get new token')
                    window.perform_long_operation(lambda:LastPassLogin(Inputs), "LPLoginReturn")
                else:
                    print('Unable to refresh accounts list, login to get new token')
                continue
            else:
                LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                window['LVTable'].update(values=LVTable)               
                if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                window['LPList' ].update(values=LPList)
                LVGroups = ['Group/Folder']
                for Group in [_[4] for _ in LVTable]:
                    Group = Group.split('\\')[-1]
                    if Group and Group not in LVGroups:
                        LVGroups.append(Group)
                window['LVsGroup'].update(values=LVGroups)
                if window['LPLoginInfo']._visible: event = 'LPLoginToggle'             

                
        if event in ['LPLogin','LPLoginReturn','LVLogin']:
            if event != 'LPLoginReturn':
                window['-OUTPUT-'].update('')
                window['OutPutSection'].select()
                print('Logging into LastPass')
                LPRefreshRedirect=False
                window.perform_long_operation(lambda:LastPassLogin(Inputs), "LPLoginReturn")
            else:
                LastPass = Inputs['LPLoginReturn']
                if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                    window['OutPutSection'].select()
                    print('Login unsuccessful')
                    LastPass = None
                    continue
                else:
                    window['OutPutSection'].select()
                    LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                    LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                    window['LVTable'].update(values=LVTable)
                    if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                    LVGroups = ['Group/Folder']
                    for Group in [_[4] for _ in LVTable]:
                        Group = Group.split('\\')[-1]
                        if Group and Group not in LVGroups:
                            LVGroups.append(Group)           
                    window['LVsGroup'].update(values=LVGroups)

                    LPInfo = {
                        'Token'    : LastPass.token,
                        'Key'      : LastPass.encryption_key,
                        'SessionId': LastPass.id,
                        'Iteration': LastPass.key_iteration_count
                    }
                    SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                    window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')
                    if window['LPOptions'].visible: window['LPOptions'].select()
                    if window['LPLoginInfo']._visible: event = 'LPLoginToggle'


        if event == 'LPLoginToggle':
            if window['LPLoginInfo']._visible:
                window['LPLoginToggle'].update(value='⩠ Login')
                if not Inputs['LPAuto']:
                    window['LPList'].Size = (30,10)
                    window['LPSelected'].Size = (30,10)
                    if sg.running_mac():
                        window['LPSelMethod'].update(visible=True) 
                        window['LPList'].set_size((30,10))
                        window['LPSelected'].set_size((30,10))
                    window['LPLoginToggle'].update(value='⩠ Login')
                    window['LPLoginInfo'].update(visible=not window['LPLoginInfo']._visible)
                    if not sg.running_mac(): 
                        window['LPList'].set_size((30,10))
                        window['LPSelected'].set_size((30,10))
                        window['LPSelMethod'].update(visible=True)
                else:
                    window['LPSelMethod'].update(visible=True)
                    window['LPLoginInfo'].update(visible=False)                    
                    window['LPSelected'].update(values=[])       
                    window['LPList'    ].Size = (30,5)
                    window['LPSelected'].Size = (30,5)
                    window['LPList'    ].update(disabled=True)
                    window['LPList'    ].Disabled=True
                    window['LPSelected'].update(disabled=True)
                    window['LPClear'   ].update(disabled=True)
                    window['LPList'    ].set_size((30,5))
                    window['LPSelected'].set_size((30,5))
                    window['LPAutoMethod'].update(visible=True)
                    if LastPass and LastPass.accounts:
                        window['LPStatus'].update(value=f'✅ You are logged into LastPass')
                    else:
                        window['LPStatus'].update(value=f'❌  You are not currently logged into LastPass')
            else:
                if sg.running_mac():
                    window['LPSelMethod' ].update(visible=False)
                    window['LPAutoMethod'].update(visible=False)
                    window['LPLoginInfo' ].update(visible=True)
                    window['LPList'].Size = (30,5)
                    window['LPSelected'].Size = (30,5) 
                    window['LPList'      ].set_size((30,5))
                    window['LPSelected'  ].set_size((30,5))
                window['LPLoginToggle'].update(value='⩢ Login')
                if not sg.running_mac():
                    window['LPSelMethod'  ].update(visible=False) 
                    window['LPLoginInfo'  ].update(visible=True)
                    window['LPAutoMethod' ].update(visible=False) 
                    window['LPList'     ].set_size((30,5))
                    window['LPSelected' ].set_size((30,5))
                
        if event == 'GCTable':
            GCSelected = window['GCSelected']
            GCSelected = GCSelected.Values if isinstance(GCSelected,sg.Table) else []
            GCTable    = window['GCTable']
            GCTable    = GCTable.Values if isinstance(GCTable,sg.Table) else []
            if Inputs['GCTable'] and GCTable and GCTable[Inputs['GCTable'][0]] not in GCSelected:
                values = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0].startswith('C_')}
                NewList = GCSelected + [GCTable[Inputs['GCTable'][0]]]
                window['GCSelected'].update(values=NewList)
                Combolist = [_[0]+' '*50+str(_) for _ in NewList]
                for _ in Slots: 
                    window[f"C_{_}"].update(values=Combolist)
                window.fill(values)

        if event == 'GCSelected':
            GCSelected = window['GCSelected']
            GCSelected = GCSelected.Values if isinstance(GCSelected,sg.Table) else []
            if Inputs['GCSelected']:
                values = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0].startswith('C_')}
                GCSelected.pop(Inputs['GCSelected'][0])
                window['GCSelected'].update(values=GCSelected)
                Combolist = [_[0]+' '*50+str(_) for _ in GCSelected]
                for _ in Slots: 
                    window[f"C_{_}"].update(values=Combolist)
                window.fill(values)
                
        if event == 'GCSelect':
            GCTable    = window['GCTable']
            GCTable    = GCTable.Values if isinstance(GCTable,sg.Table) else []
            window['GCSelected'].update(values=GCTable)
            Combolist = [_[0]+' '*50+str(_) for _ in GCTable]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)
            
        if event == 'GCClear':
            window['GCSelected'].update(values=[])
            for _ in Slots: 
                window[f"C_{_}"].update(values=[])       

        if event == 'METable':
            MESelected = window['MESelected']
            MESelected = MESelected.Values if isinstance(MESelected,sg.Table) else []
            METable    = window['METable']
            METable    = METable.Values if isinstance(METable,sg.Table) else []
            if Inputs['METable'] and METable and METable[Inputs['METable'][0]] not in MESelected:
                values = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0].startswith('C_')}
                NewList = MESelected + [METable[Inputs['METable'][0]]]
                window['MESelected'].update(values=NewList)
                Combolist = [_[0]+' '*50+str(_) for _ in NewList]
                for _ in Slots: 
                    window[f"C_{_}"].update(values=Combolist)
                window.fill(values)

        if event == 'MESelected':
            MESelected = window['MESelected']
            MESelected = MESelected.Values if isinstance(MESelected,sg.Table) else []
            if Inputs['MESelected']:
                values = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0].startswith('C_')}
                MESelected.pop(Inputs['MESelected'][0])
                window['MESelected'].update(values=MESelected)
                Combolist = [_[0]+' '*50+str(_) for _ in MESelected]
                for _ in Slots: 
                    window[f"C_{_}"].update(values=Combolist)
                window.fill(values)
                
        if event == 'MESelect':
            METable    = window['METable']
            METable    = METable.Values if isinstance(METable,sg.Table) else []
            window['MESelected'].update(values=METable)
            Combolist = [_[0]+' '*50+str(_) for _ in METable]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)          
            
        if event == 'MEClear':
            window['MESelected'].update(values=[])
            for _ in Slots: 
                window[f"C_{_}"].update(values=[])                     
        
        if event == 'LPList':
            LPSelected = window['LPSelected']
            LPSelected = LPSelected.Values if isinstance(LPSelected,sg.Listbox) else []
            if Inputs['LPList'] and Inputs['LPList'][0] not in LPSelected:
                values = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0].startswith('C_')}
                NewList = LPSelected + Inputs['LPList']
                window['LPSelected'].update(values=NewList)
                for _ in Slots: 
                    window[f"C_{_}"].update(values=NewList)
                window.fill(values)
            
        
        if event == 'LPSelected':
            LPSelected = window['LPSelected']
            LPSelected = LPSelected.Values if isinstance(LPSelected,sg.Listbox) else []
            if Inputs['LPSelected']:
                values = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0].startswith('C_') and _[1] != Inputs['LPSelected'][0]}
                LPSelected.remove(Inputs['LPSelected'][0])
                window['LPSelected'].update(values=LPSelected)
                for _ in Slots: 
                    window[f"C_{_}"].update(values=LPSelected)
                window.fill(values)
        
        if event == 'LPClear':
            window['LPSelected'].update(values=[])
            for _ in Slots: 
                window[f"C_{_}"].update(values=[])
                
        if event == 'Tabs':
            if Inputs['Tabs'] == 'OKOptions' and not [True for _ in Slots if len(window[f'S_{_}'].Text) > 2]:
                try:
                    onlykey = OnlyKey(connect=True,tries=1)
                    onlykey.read_bytes(timeout_ms=1000)
                    OKSlots = onlykey.getlabels()
                    onlykey.close()
                except BaseException as Error:
                    if not test: window['OutPutSection'].select()
                    print(Error)
                    OKSlots = None
                if OKSlots:
                    sl = {Slot : [(Slot+" "+_.label if _.label and _.label != 'ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ' else Slot) for _ in OKSlots if _.name == Slot][0]  for Slot in Slots}
                else:
                    if not test: window['OutPutSection'].select()
                    sl = {Slot : Slot for Slot in Slots}
                    
                for Slot in Slots:  window[f'S_{Slot}'].update(text=sl[Slot])

        if event == 'OKRefresh':
            window['-OUTPUT-'].update('')
            try:
                onlykey = OnlyKey(connect=True,tries=1)
                onlykey.read_bytes(timeout_ms=1000)
                OKSlots = onlykey.getlabels()
                onlykey.close()
            except BaseException as Error:
                window['OutPutSection'].select()
                print(Error)
                OKSlots = None
            if OKSlots:
                sl = {Slot : [(Slot+" "+_.label if _.label and _.label != 'ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ' else Slot) for _ in OKSlots if _.name == Slot][0]  for Slot in Slots}
            else:
                window['OutPutSection'].select()
                sl = {Slot : Slot for Slot in Slots}
                
            for Slot in Slots:  window[f'S_{Slot}'].update(text=sl[Slot])
        
        #Toggle Slot Picker Element
        if event in ["PSlotsRadio",'KWSearch']:
            KWTrue = not KWTrue
            window['SlotSelector'].update(visible=not KWTrue)
            window['KWLabelSearchText'].update(visible=KWTrue)


        #Save Current Input Values to Config File
        if event == "Save":
            window['-OUTPUT-'].update('')
            config = configparser.ConfigParser()
            ArkInfo = ArkInfo if ArkInfo else {'Token': ""}
            LPInfo  = LPInfo  if LPInfo  else {'Token':"",'Key':"",'SessionId':'','Iteration':''}
            window.perform_long_operation(lambda:SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window),'SaveReturn')
            window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')

        if event == 'AppDataReturn':
            AppData = Inputs['AppDataReturn']
        
        #Clear Print Statements From Output Element
        if event in ['ClearD2','ClearD']:
            window['-OUTPUT-'].update('')
            continue
        
#endregion
        
#region Password Sync
        if event in ["Submit",'LoginCheckReturn','FunctionReturn','OKCAFromReturn']:
            window['OutPutSection'].select()
            
            if Inputs['MEFrom'] or Inputs['GCFrom']:
                Selected = window['GSSelected'].Values if Inputs['GCFrom'] else window['MESelected'].Values
                
                if Inputs['LP']:
                    if not LastPass and AppData['LPsId'] and AppData['LPIteration'] and AppData['LPToken'] and AppData['LPKey']:
                        LastPass = LastPassSession(AppData)
                        
                    if not Inputs['LPUser'] or not Inputs['LPPass']:
                        if not Inputs["LPUser"]: print('Fill in the LastPass Username field to Login to LastPass')
                        if not Inputs["LPPass"]: print('Fill in the LastPass Password field to Login to LastPass')
                        continue
                        
                    if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                        print ('Attempting LastPass login')
                        LastPass = LastPassLogin(Inputs)
                        if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                            print('Login unsuccessful')
                            LastPass = None
                            continue
                    
                    LPInfo = {
                        'Token'    : LastPass.token,
                        'Key'      : LastPass.encryption_key,
                        'SessionId': LastPass.id,
                        'Iteration': LastPass.key_iteration_count
                    }

                    folder = 'Edge' if Inputs['MEFrom'] else 'Chrome'

                    for Cred in Selected:
                        if Cred[0] in [_.name for _ in LastPass.accounts]:
                            Account = [_ for _ in LastPass.accounts if _.name == Cred[0]][0]
                            LastPass.UpdateAccount(Account.id,password=Cred[3],group=f'MEDHOST Password Manager\\{folder}',Account=Account)
                        else:
                            notes="Created By MEDHOST Password Manager"
                            LastPass.NewAccount(Cred[0],Cred[1],Cred[3],notes=notes,group=f'MEDHOST Password Manager\\{folder}',url=Cred[0])
                    if LastPass.accounts:
                        LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                        window['LVTable'].update(values=LVTable) 
                        LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                        if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)                                                 

            if Inputs['LPFrom']:
                window['-OUTPUT-'].update('')
                if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                    if AppData['LPsId'] and AppData['LPIteration'] and AppData['LPToken'] and AppData['LPKey']:
                        LastPass = LastPassSession(AppData)
                        
                    if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                        if not Inputs['LPUser'] or not Inputs['LPPass']:
                            if not Inputs["LPUser"]: print('Fill in the LastPass Username field to Login to LastPass')
                            if not Inputs["LPPass"]: print('Fill in the LastPass Password field to Login to LastPass')
                            continue  
                        
                        print ('Attempting LastPass login')
                        LastPass = LastPassLogin(Inputs)
                        if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                            print('Login unsuccessful')
                            LastPass = None
                            continue
                    LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                    if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)  
                    LPInfo = {
                        'Token'    : LastPass.token,
                        'Key'      : LastPass.encryption_key,
                        'SessionId': LastPass.id,
                        'Iteration': LastPass.key_iteration_count
                    }
                    
                TrueSlots = {_[0].replace("S_","") : _[1] for _ in Inputs.items() if isinstance(_[0],str) and "S_" in _[0] and _[1] == True}
                
                for slot in TrueSlots:
                    dropdown = Inputs[f'C_{slot}']
                    if dropdown and len(dropdown) > 120 and (' ' * 120) in dropdown: 
                        ID = dropdown.split(' ' * 120)[1]
                        Account = [_ for _ in LastPass.accounts if _.id == int(ID) or _.id == ID]
                        if not Account:
                            name = dropdown.split(' ' * 120)[0]
                            print(f'Unable to find account {name} with Id {ID} in list of accounts')
                            print ('Unknown error')
                            continue
                        else:
                            Account = Account[0]
                            
                        UpdateOnlyKey(Password=Account.password,OK_Keyword='',SlotSelections={slot:True},SlotsTrue=True)
                    else:
                        print(f'No LastPass account is mapped to OnlyKey slot {slot}')
                            

            if Inputs['CAFrom']:
                if event not in  ['FunctionReturn','LoginCheckReturn'] and not AppData['ArkToken']: 
                    window['-OUTPUT-'].update('')
                    if not Inputs["ArkPass"] or not Inputs['CATable'] or not Inputs["ArkUser"] or not Inputs['CAReason'] or Inputs['CAReason'] == CAeReason:   
                        if not Inputs["ArkUser"]: print('Fill in the Username field to Login to CyberArk')
                        if not Inputs["ArkPass"]: print('Fill in the Password field to Login to CyberArk')
                        if not Inputs['CATable']: print('You must select an Account in CyberArk to check-out')
                        if not Inputs['CAReason'] or Inputs['CAReason'] == CAeReason:
                            print('You must enter a business justification to check-out an account')
                        continue
                    CALoginAttempt = 1
                    window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "FunctionReturn")
                    continue
                else:
                    if event not in ['FunctionReturn','LoginCheckReturn','OKCAFromReturn']: window['-OUTPUT-'].update('')
                    if not Inputs['CATable'] or not Inputs['CAReason'] or Inputs['CAReason'] == CAeReason:
                        if not Inputs['CATable']: print('You must select an Account in CyberArk to check-out')
                        if not Inputs['CAReason'] or Inputs['CAReason'] == CAeReason:
                            print('You must enter a business justification to check-out an account')
                        continue
                    if event not in ['LoginCheckReturn','OKCAFromReturn']:
                        FunctionReturn = Inputs['FunctionReturn'] if 'FunctionReturn' in Inputs else None
                        ArkToken = FunctionReturn if FunctionReturn else AppData['ArkToken']
                        ArkInfo  = {'Token': ArkToken}
                        window.perform_long_operation(lambda:GetCyberArkSettings(Token=ArkToken), 'LoginCheckReturn')
                        continue
                    CATable  = window['CATable']
                    CATable  = CATable.Values if isinstance(CATable,sg.Table) else []
                    BGUser   = CATable[Inputs['CATable'][0]]
                    if event == 'LoginCheckReturn':
                        LoginCheck = Inputs['LoginCheckReturn']
                        if LoginCheck and 'ErrorMessage' in LoginCheck and CALoginAttempt == 0 and Inputs["ArkUser"] and Inputs["ArkPass"]:
                            CALoginAttempt = 1
                            window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "FunctionReturn")
                            continue
                        elif LoginCheck and 'ErrorMessage' in LoginCheck and CALoginAttempt != 0:
                            print(LoginCheck['ErrorMessage'])
                            print('CyberArk Login unsuccessful')
                            CALoginAttempt = 0
                            continue
                        elif LoginCheck == None:
                            print("Unable Find CyberArk Account for " + BGUser[1])
                            print("Try Re-entering CyberArk Account Name")
                            continue
                        elif not Inputs["ArkUser"] or not Inputs["ArkPass"]:
                            if not Inputs["ArkUser"]: print('Fill in the Network Username field to Login to CyberArk')
                            if not Inputs["ArkPass"]: print('Fill in the Network Password field to Login to CyberArk')
                            continue
                        
                        CaseNum = Inputs['CACase'] if Inputs['CACase'] != 'Case #' else ''

                        ArkAccountPass = GetCyberArkPassword(AccountId=BGUser[13], Token=ArkToken, UserName=BGUser[1],Reason=Inputs['CAReason'],Case=CaseNum)
                        if ArkAccountPass and isinstance(ArkAccountPass,dict) and 'ErrorMessage' in ArkAccountPass:
                            print(ArkAccountPass['ErrorMessage'])
                            print(f'Unable to retreive password from CyberArk for {BGUser[1]}')
                            continue
                        
                        if not ArkAccountPass:
                            print(f'Unable to retreive password from CyberArk for {BGUser[1]}')
                            continue
                    
                    if Inputs['OK'] and event not in ['OKCAFromReturn']:
                        TrueSlots = {_[0].replace("S_","") : _[1] for _ in Inputs.items() if isinstance(_[0],str) and "S_" in _[0] and _[1] == True}
                        if Inputs["PSlotsRadio"] and not TrueSlots:
                            window['-OUTPUT-'].update('')
                            print('You must select which OnlyKey slots you would like to update')
                            sleep(2)
                            window.key_dict['OKOptions'].select()
                            continue
                        window.perform_long_operation(lambda:UpdateOnlyKey(Password=ArkAccountPass,OK_Keyword=Inputs["OK_Keyword"],SlotSelections=TrueSlots,SlotsTrue=Inputs["PSlotsRadio"]), 'OKCAFromReturn')
                        continue
                    if Inputs['LP']:
                        if not LastPass and AppData['LPsId'] and AppData['LPIteration'] and AppData['LPToken'] and AppData['LPKey']:
                            LastPass = LastPassSession(AppData)
                            
                        if not Inputs['LPUser'] or not Inputs['LPPass']:
                            if not Inputs["LPUser"]: print('Fill in the LastPass Username field to Login to LastPass')
                            if not Inputs["LPPass"]: print('Fill in the LastPass Password field to Login to LastPass')
                            continue
                            
                        if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                            print ('Attempting LastPass login')
                            LastPass = LastPassLogin(Inputs)
                            if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                                print('Login unsuccessful')
                                LastPass = None
                                continue
                        else:
                            LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                            if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                        
                        LPInfo = {
                            'Token'    : LastPass.token,
                            'Key'      : LastPass.encryption_key,
                            'SessionId': LastPass.id,
                            'Iteration': LastPass.key_iteration_count
                        }
                        LPSelected = window['LPSelected'].Values
                        if LPSelected:
                            for Item in LPSelected:
                                Account = [_ for _ in LastPass.accounts if _.id == int(Item.split(' ' * 120)[1])][0]
                                LastPass.UpdateAccount(Account.id,password=ArkAccountPass,group='MEDHOST Password Manager\\CyberArk',Account=Account)
                        elif BGUser[1] in [_.name for _ in LastPass.accounts]:
                            Account = [_ for _ in LastPass.accounts if _.name == BGUser[1]][0]
                            LastPass.UpdateAccount(Account.id,password=ArkAccountPass,group='MEDHOST Password Manager\\CyberArk',Account=Account)
                        else:
                            Account = None    
                            notes="Created By MEDHOST Password Manager"
                            LastPass.NewAccount(BGUser[1],BGUser[1],ArkAccountPass,'MEDHOST Password Manager\\CyberArk',notes=notes)
                        if LastPass.accounts:
                            LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                            window['LVTable'].update(values=LVTable)
                            LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                            if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                            if Account:
                                index = [_[0] for _ in enumerate(LVTable) if _[1][3] == Account.id]
                            else:
                                index = [_[0] for _ in enumerate(LVTable) if _[1][0] == BGUser[0] and _[1][4] == 'MEDHOST Password Manager\\CyberArk']

                            if index:
                                ScrollPosition = index[0] / (len(LVTable) - 1) if LVTable and index and len(LVTable) != 1 else 0
                                window['LVTable'].update(select_rows=index)
                                window['LVTable'].set_vscroll_position(ScrollPosition)
                                window['LocalVault'].select()                           

                    SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                    window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')
                    
                    continue
        # else:
        #     newwindow = window
        #     del window
        #     gc.collect()
        #     window = newwindow
        #     continue
    except Exception as error:
        raise error
        print(error)
        continue    
window.close()
#endregion
