#region Import Modules 
import json
import sqlite3
import win32crypt
import shutil
import base64
import os
from ast import Continue, literal_eval as array
from time import sleep
from tkinter import W
from onlykey.client import OnlyKey,MessageField
from requests import get,post
from getpass import getuser
from os import getcwd,getenv
import configparser
import lastpass
from lastpass.fetcher import make_key
from Crypto.Cipher import AES
import subprocess
from base64 import b64decode,b64encode
import PySimpleGUI as sg
from datetime import timezone, datetime, timedelta
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

#region Function FindCyberArkAccount
def FindCyberArkAccount(Token,Username=None,ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    if not Username:
        URL    = ArkURL + "/Accounts"
        Response = (get(url=URL,headers=Header)).json()
        return Response
    else:
        URL  = ArkURL + "/Accounts" + '?search=' + Username

        Response = (get(url=URL,headers=Header)).json()
        if 'value' in Response:
            ArkAccount = next((User for User in Response['value'] if User['userName'] == Username), None)
        else:
            ArkAccount = Response
        return ArkAccount
#endregion

#region Function GetCyberArkPassword
def GetCyberArkPassword(AccountId,Token,UserName,ArkURL="https://cyberark.medhost.com/PasswordVault/API"):
    print("Retrieving CyberArk Password for " + UserName)

    Header = {"Authorization" : Token}
    
    Body = {
        "reason" : "MEDHOST Password Manager Update"
    }

    URL = ArkURL + "/Accounts/" + AccountId + "/Password/Retrieve"

    Response = (post(url=URL,json=Body,headers=Header)).json()
    return Response
#endregion

#region Function GetCyberArkPassword
def CheckInCyberArkPassword(AccountId,Token,UserName,ArkURL="https://cyberark.medhost.com/PasswordVault/API"):
    print("Checking in CyberArk Password for " + UserName)

    Header = {"Authorization" : Token}
    
    URL = ArkURL + "/Accounts/" + AccountId + "/CheckIn"

    Response = post(url=URL,headers=Header)
    return Response
#endregion

def get_chrome_datetime(chromedate):
    """Return a `datetime.datetime` object from a chrome format datetime
    Since `chromedate` is formatted as the number of microseconds since January, 1601"""
    return datetime(1601, 1, 1) + timedelta(microseconds=chromedate)

def get_encryption_key(edge=False):
    if edge:
        local_state_path = os.path.join(os.environ["USERPROFILE"],
                                        "AppData", "Local", "Microsoft", "Edge",
                                        "User Data", "Local State")
    else:
        local_state_path = os.path.join(os.environ["USERPROFILE"],
                                        "AppData", "Local", "Google", "Chrome",
                                        "User Data", "Local State")
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

def decrypt_password(password, key):
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

def getchromiumvault(edge=False):
    # get the AES key
    key = get_encryption_key(edge)
    # local sqlite database path
    if edge:
        db_path = os.path.join(os.environ["USERPROFILE"], "AppData", "Local",
                            "Microsoft", "Edge", "User Data", "default", "Login Data")
    else:
        db_path = os.path.join(os.environ["USERPROFILE"], "AppData", "Local",
                            "Google", "Chrome", "User Data", "default", "Login Data")
    # copy the file to another location
    # as the database will be locked if chrome is currently running
    filename = "ChromeData.db"
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

    try:
        for i,entry in enumerate(chromiumvault):
            chromiumvault[i]['password_value']           = decrypt_password(entry['password_value'], key)
            chromiumvault[i]['date_created']             = get_chrome_datetime(entry['date_created'])
            chromiumvault[i]['date_last_used']           = get_chrome_datetime(entry['date_last_used'])
            chromiumvault[i]['date_password_modified']   = get_chrome_datetime(entry['date_password_modified'])
    except:
        pass
                
    cursor.close()
    db.close()
    try:
        # try to remove the copied db file
        os.remove(filename)
    except:
        pass
    return chromiumvault


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

#region Function SwapContainers
def SwapContainers(window,C1,C2):
    location  = window.CurrentLocation() if window.Shown else (None,None)
    values    = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and not _[0] == 'Tabs'}
    listboxes = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox' and _[1].Values]
    tabs      = [{'Key':_[1].Key,'TabID':_[1].TabID} for _ in window.key_dict.items() if _[1].Type == 'tab'and (_[1].TabID or _[1].TabID == 0)]

    C1Rows = window[C2].Rows
    C2Rows = window[C1].Rows
    window[C1].Rows = C1Rows
    window[C2].Rows = C2Rows

    for rI,row in enumerate(window.Rows):
        for eI,elem in enumerate(row):
                window.Rows[rI][eI].Position = (rI,eI)
                window.Rows[rI][eI].ParentContainer = None
        
    newwindow = sg.Window(Title, window.Rows,location=location,finalize=True,font=window.Font)
    window.close()
    newwindow.fill(values)
    for box in listboxes: newwindow[box['Key']].update(values=box['Values'])
    for tab in tabs:      newwindow[tab['Key']].TabID = tab['TabID']

    return newwindow
#endregion

#region Function ReloadWindow
def ReloadWindow(window):
    location  = window.CurrentLocation()
    values    = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str)}
    listboxes = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox']
    tabs      = [{'Key':_[1].Key,'TabID':_[1].TabID} for _ in window.key_dict.items() if _[1].Type == 'tab']
    
    for rI,row in enumerate(window.Rows):
        for eI,elem in enumerate(row):
                window.Rows[rI][eI].Position = (rI,eI)
                window.Rows[rI][eI].ParentContainer = None
        
    newwindow = sg.Window(Title, window.Rows,location=location,finalize=True,font=window.Font)
    window.close()
    newwindow.fill(values)
    for box in listboxes: newwindow[box['Key']].update(values=box['Values'])
    for tab in tabs:      newwindow[tab['Key']].TabID = tab['TabID']

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
    from distutils.util import strtobool as stb
    return bool(stb(str(string)))
#endregion

#region Function SaveToConfig
def SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,Print=True):
        config = configparser.ConfigParser()
        user = getuser()
        Pass = (subprocess.run("hostname",capture_output=True,text=True)).stdout.rstrip()
        key = make_key(user,Pass,pin)

        config.add_section('General')
        config.add_section('CyberArk')
        config['CyberArk']['Network Username']            =     Inputs["ArkUser"]
        config['CyberArk']['Account']                     =     Inputs["BGUser"]
        config['CyberArk']['Sync From CyberArk Selected'] = str(Inputs["CAFrom"])
        config.add_section('LastPass')     
        config['LastPass']['Username']                    =     Inputs["LPUser"]
        config['LastPass']['Sync From LastPass Selected'] = str(Inputs['LPFrom'])
        config['LastPass']['Sync To LastPass Selected']   = str(Inputs["LP"])
        config['LastPass']['LastPass Accounts']           = window.key_dict['LPSelected'].Values.__str__()
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
        Pass = (subprocess.run("hostname",capture_output=True,text=True)).stdout.rstrip()
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
                notes="Created By CyberArk Update Tool",
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

#region Read Config File
Slots = ["1a","1b","2a","2b","3a","3b","4a","4b","5a","5b","6a","6b"]

config = configparser.ConfigParser()
try:
    HomeDirectory = getenv("HOME")
except:
    HomeDirectory = None
    
WorkingDirectory = HomeDirectory + '/Config.ini' if sg.running_mac() and HomeDirectory else getcwd() + '/Config.ini'
config.read(WorkingDirectory)
       
ConfigUsername =       config.get('CyberArk','Network Username'            ,fallback=getuser())
ConfigBGUser   =       config.get('CyberArk','Account'                     ,fallback="")
ConfigCAFrom   =  Bool(config.get('CyberArk','Sync From CyberArk Selected' ,fallback='True')) 
       
ConfigLPUser   =       config.get('LastPass','Username'                    ,fallback="")
ConfigLP       =  Bool(config.get('LastPass','Sync To LastPass Selected'   ,fallback="True"))
ConfigLPFrom   =  Bool(config.get('LastPass','Sync From LastPass Selected' ,fallback="False"))
ConfigLPAccts  = array(config.get('LastPass','LastPass Accounts'           ,fallback="[]")) 
       
ConfigOK       =  Bool(config.get('OnlyKey' ,'OnlyKey Selected'            ,fallback="True"))
ConfigOKWord   =       config.get('OnlyKey' ,'Keyword'                     ,fallback="BG*")
ConfigKWTrue   =  Bool(config.get('OnlyKey' ,'Keyword Search Selected'     ,fallback="False"))
ConfigSlots    = {Slot : Bool(config.get('Selected Slots',Slot,fallback="False")) for Slot in Slots}
ConfigMap      = {Slot : config.get('Acct-Slot Mappings',Slot,fallback="") for Slot in Slots}

if ConfigUsername == "": ConfigUsername = getuser()

#AppData Pin
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
        
AppData = AppData if AppData != 'error' else {}

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
    
if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts or test:
    LPList = ['Login to see LastPass entries here']
else:
    LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]

ArkToken   = AppData['ArkToken']
if ArkToken:
    ArkAccount = FindCyberArkAccount(Token=ArkToken)
else:
    ArkAccount = None
    
if not ArkAccount or 'ErrorMessage' in ArkAccount or test:
    CAList = ['Login to see CyberArk Accounts here']
else:
    CAList = [(_['userName'] + (" " * 120) + _['id']) for _ in ArkAccount['value']]
    
try:
    chromevault = getchromiumvault()
except:
    chromevault = []

GCList = []    
if chromevault:
    for Pass in chromevault:
        if 'android://' not in Pass['signon_realm']:
            url      = Pass['signon_realm'].replace('https://','').replace('http://','').replace('www.','')
            url      = url[:len(url)-1] if url[-1] == '/' else url
            url      = (url[:25] if len(url) >= 25 else url + ((25 - len(url)) * " ")) + 5 * " "
            username = Pass['username_value']
            GCList.append(url + username + (100 * " ") + str(Pass['id'])) 
#endregion

#region Collapsible Section Layouts
#tooltips
ToggleToolTip   = "Click Here to Collapse or Expand Output Section"
OutputToolTip   = "The Outputs and Error Messages of the Application Appear Here"
ClearToolTip    = "Click Here to Clear Text from the Output Element"
KeywordToolTip  = "OnlyKey Slot Label Keyword: Search Keyword for Labels of Slots You Would Like to Update"
UserNameToolTip = "Your CyberArk Username (sAMAccountName): For Logging into CyberArk"
PasswordToolTip = "Your 16+ Digit Domain Password: For Logging into CyberArk"
BGUserToolTip   = "Your CyberArk Account Username (sAMAccountName): Account to Search for in CyberArk"
SubmitToolTip   = "Click Here to Login to CyberArk and Update OnlyKey Slot Passwords"
SaveToolTip     = "Click Here to Save Field Values and Output Toggle to Config.ini File (Network Password is NOT Saved)"

#region Output
Fm  = 2 if sg.running_mac() else 0
Sm  = 2 if sg.running_mac() else 0
Im  = 0 if sg.running_mac() else 0
OutPutSection = [
    [sg.Output(size=(62,15), key='-OUTPUT-',visible=True,expand_x=True,echo_stdout_stderr=True,pad=((5,5),(5,3)),tooltip=OutputToolTip)],
    [sg.Button('Clear',key='ClearD',tooltip=ClearToolTip)]
]
#endregion

#region LastPass
LPAccts = ConfigLPAccts
Ls  = 55 if sg.running_mac() else 45
Lbs = 30 if sg.running_mac() else 30
LastPassOptions = [
    [sg.Text("LastPass",justification="center",size=(Ls, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.Text('Please enter your LastPass Information')],
    [sg.Text('LastPass Username', size=(16+Sm, 1)), sg.InputText(default_text=ConfigLPUser,key="LPUser",s=(45+Im,1))],
    [sg.Text('LastPass Password', size=(16+Sm, 1)), sg.InputText(password_char="*",key="LPPass",default_text=AppData['LPPass'],s=(45+Im,1))],
    [sg.Button(button_text='Refresh Account List',key="LPRefresh"),sg.Push(),sg.Button(button_text='Login',key="LPLogin")],
    [sg.Text('Select Account(s) to sync:', size=(28+Sm+Sm, 1),p=((5,0),0)),sg.Text('Selected Accounts:', size=(28+Sm, 1),p=0)],
    [sg.Listbox(values=LPList,k='LPList',s=(30+Im,5),p=((5,0),0),enable_events=True),sg.Listbox(values=ConfigLPAccts,k='LPSelected',s=(Lbs+Im,5),p=((0,4),0),enable_events=True)],
    [sg.Push(),sg.Button(button_text='Clear All',key="LPClear",p=(5,3))]
]
#endregion

#region Chrome
GCOptions = [
    [sg.Text("Chrome",justification="center",size=(Ls, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.Text('Select Google Chrome password(s) to sync:')],
    [sg.Listbox(values=GCList,k='GCList',s=(55,5),font=('Courier New',10),enable_events=True)],
    [sg.Text('Selected Passwords:')],
    [sg.Listbox(values=['example.com           joshlove007@gmail.com'],k='GCSelected',s=(55,5),font=('Courier',10),enable_events=True)],
    [sg.Push(),sg.Button(button_text='Clear All',key="GCClear",p=(5,3))]
]
#endregion

#region CyberArk
Pd = 10 if sg.running_mac() else 9
CAOptions = [
    [sg.Text("CyberArk",justification="center",size=(Ls, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.Text('Please enter your CyberArk Information')],
    [sg.Text('Network Username', size=(16+Sm, 1),tooltip=UserNameToolTip), sg.InputText(default_text=ConfigUsername,key="ArkUser",tooltip=UserNameToolTip)],
    [sg.Text('Network Password', size=(16+Sm, 1),tooltip=PasswordToolTip), sg.InputText(password_char="*",key="ArkPass",tooltip=PasswordToolTip,default_text=AppData['ArkPass'])],
    [sg.Button(button_text='Refresh Account List',key="CARefresh"),sg.Push(),sg.Button(button_text='Login',key="CALogin")],
    [sg.Text('Select an Account to sync:', size=(32,1),p=((5,0),0))],
    [sg.Listbox(values=CAList,key='CAList',size=(63+Im,5),p=((5,0),0),enable_events=True)],
    [sg.Text('Selected Account', size=(16+Sm, 1),p=(5,(3,Pd)),tooltip=BGUserToolTip)  , sg.InputText(default_text=ConfigBGUser,key="BGUser",tooltip=BGUserToolTip)]
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
OKComboLeft  = [[sg.Combo(LPAccts,k='C_1a',font=Cf,default_value=LPOKMap['1a'],s=(Scs,1),p=((0,2),(0 ,0)),disabled=not Cs['1a'])] , [sg.Combo(LPAccts,k='C_1b',font=Cf,default_value=LPOKMap['1b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['1b'])],
                [sg.Combo(LPAccts,k='C_3a',font=Cf,default_value=LPOKMap['3a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['3a'])] , [sg.Combo(LPAccts,k='C_3b',font=Cf,default_value=LPOKMap['3b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['3b'])],
                [sg.Combo(LPAccts,k='C_5a',font=Cf,default_value=LPOKMap['5a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['5a'])] , [sg.Combo(LPAccts,k='C_5b',font=Cf,default_value=LPOKMap['5b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['5b'])]]
OKComboRight = [[sg.Combo(LPAccts,k='C_2a',font=Cf,default_value=LPOKMap['2a'],s=(Scs,1),p=((0,2),(0 ,0)),disabled=not Cs['2a'])] , [sg.Combo(LPAccts,k='C_2b',font=Cf,default_value=LPOKMap['2b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['2b'])],
                [sg.Combo(LPAccts,k='C_4a',font=Cf,default_value=LPOKMap['4a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['4a'])] , [sg.Combo(LPAccts,k='C_4b',font=Cf,default_value=LPOKMap['4b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['4b'])],
                [sg.Combo(LPAccts,k='C_6a',font=Cf,default_value=LPOKMap['6a'],s=(Scs,1),p=((0,2),(25,0)),disabled=not Cs['6a'])] , [sg.Combo(LPAccts,k='C_6b',font=Cf,default_value=LPOKMap['6b'],s=(Scs,1),p=((0,2),0),disabled=not Cs['6b'])]]

OkCheckText = [[sg.Text('Select Slots to Update:', size=(15,2),p=((5,0),5))],[sg.Button(button_text='Refresh Labels',key="OKRefresh")],[sg.Text(" ",size=(1,6))],
               [sg.Button(button_text='Clear Selection',key="OKClear")]]

Osp = 45 if sg.running_mac() else 0 
SlotSelector = [
    [sg.vtop(sg.Col(OkCheckText,p=((0,Osp),0),visible=not ComboTrue,k='OkCheckText')),sg.Fr('OnlyKey Slots',
     layout=[[sg.Col(OkCheckLeft,p=0),sg.Col(OKComboLeft,p=0,visible=ComboTrue,k='OKComboLeft'),sg.VSep(pad=(0,5)),
     sg.Col(OkCheckRight,p=0),sg.Col(OKComboRight,p=0,visible=ComboTrue,k='OKComboRight')]])]
    ]
SlotSelMethod = [[sg.Text('Slot Selction Method', size=(16+Sm, 1),p=((5,Osp),0)),sg.Push(),sg.Radio('Pick Slots','SMethod',k="PSlotsRadio",enable_events=True,default=not KWTrue),
    sg.Radio('Keyword Search','SMethod',k="KWSearch",default=KWTrue,enable_events=True),sg.Push()]]

Clrp = 70 if sg.running_mac() else 0 
OkSlotSelection = [
    [sg.Text("Onlykey",justification="center",size=(Ls, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.pin(sg.Col(SlotSelMethod,k='SlotSelMethod',p=0,visible=not ComboTrue))],
    [sg.pin(sg.Col([[sg.Text('Map LastPass Accounts Selected Above to Onlykey slots below'),sg.Push(),sg.Button(button_text='Clear All',key="OKClear2",p=((Clrp,5),0))]],k='OK_LPAcctText',p=0,visible=ComboTrue))],
    [sg.pin(sg.Column(SlotSelector,k="SlotSelector",p=0,visible=not KWTrue))],
    [collapse(KWLabelSearchText, "KWLabelSearchText",KWTrue)]
]
#endregion
#endregion

#region Main Application Window Layout
LPFrom = ConfigLPFrom
OKTrue = ConfigOK
LPTrue = ConfigLP
CATrue = ConfigCAFrom
#Layout
layout = [
    [collapse(CAOptions,'CAOptions',False)],
    [sg.Column(GCOptions,p=0)],
    [sg.Text('Sync Password(s) From', size=(18+Sm, 1)),sg.Radio('CyberArk','PFrom',k='CAFrom',enable_events=True,default=CATrue),sg.Radio('LastPass','PFrom',k='LPFrom',enable_events=True,default=LPFrom),sg.Push()],
    [sg.Text('Sync Password(s) To', size=(18+Sm, 1)),sg.pin(sg.Checkbox('LastPass',k='LP',enable_events=True,default=LPTrue,disabled=not OKTrue,visible=not LPFrom)),
     sg.pin(sg.Checkbox('OnlyKey',k='OK',enable_events=True,default=OKTrue,disabled=not LPTrue or LPFrom)),sg.Push()],
    [sg.Button(button_text='Save',key="Save",tooltip=SaveToolTip),sg.Push(),sg.Button(button_text='Check In',key="CheckIn")
     ,sg.Submit("Sync Passwords",bind_return_key=True,tooltip=SubmitToolTip)],
    [sg.HSep()],
    [sg.TabGroup([[sg.Tab('Output',OutPutSection,k='OutPutSection'),sg.Tab('LastPass',LastPassOptions,k='LPOptions',visible=LPTrue and not LPFrom), 
        sg.Tab('OnlyKey', OkSlotSelection,k='OKOptions',visible=OKTrue)]],p=0,enable_events=True,k='Tabs')]
]
#alternate up/down arrows ⩢⩠▲▼⩓⩔
#Window and Titlebar Options
Title = 'MEDHOST Password Manager'
window = sg.Window(Title, layout,font=("Helvetica", 10+Fm))
#endregion

#region Application Behaviour if Statments - Event Loop
OKVisible  = OKTrue
CAVisible  = CATrue
CALoginAttempt = 0
if LPFrom:
    window = SwapContainers(window,'LPOptions','CAOptions')

while True:
    try:
        #Read Events(Actions) and Users Inputs From Window 
        WindowRead = window.read()
        event      = WindowRead[0]
        Inputs     = WindowRead[1]
        
        #End Script when Window is Closed
        if event == sg.WIN_CLOSED:
            break
            
        
        if isinstance(event,str) and event.startswith('S_'):
            Slot = event.replace('S_','')
            window[f'C_{Slot}'].Disabled = not Inputs[event]
            window[f'C_{Slot}'].update(disabled=not Inputs[event])
        
        if event == 'LPFrom' and CATrue:
            CATrue = False
            if not sg.running_mac():
                for _ in Slots: window[f"S_{_}"].Font=('Helvetica',8)
            window['SlotSelMethod']._visible=False
            window['OK_LPAcctText']._visible=True
            window['OKComboLeft']._visible=True 
            window['OKComboRight']._visible=True
            window['OkCheckText']._visible=False
            window['SlotSelector']._visible=True
            window['KWLabelSearchText']._visible=False 
            window['CheckIn']._visible=False 
            window['LPOptions']._visible=False
            window['LP'].update(visible=False)
            OKTrue = Inputs['OK']
            window.ReturnValuesDictionary['OK'] = True
            window['OK'].Disabled = True
            window = SwapContainers(window,'LPOptions','CAOptions')
            window.key_dict['OKOptions'].select()

        if event == 'CAFrom' and not CATrue:
            CATrue = True
            if not sg.running_mac():
                for _ in Slots: window[f"S_{_}"].Font=('Helvetica',10)
            window['SlotSelMethod']._visible=True
            window['OK_LPAcctText']._visible=False
            window['OKComboLeft']._visible=False
            window['OKComboRight']._visible=False
            window['OkCheckText']._visible=True
            window['SlotSelector']._visible=Inputs['PSlotsRadio']
            window['KWLabelSearchText']._visible=Inputs['KWSearch']
            window.ReturnValuesDictionary['OK'] = OKTrue
            window['CheckIn']._visible=True 
            window['LPOptions']._visible=Inputs['LP']
            window['LP'].update(visible=True)
            window['OK'].Disabled = not Inputs['LP']
            window = SwapContainers(window,'LPOptions','CAOptions')
        
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
            
        if event == 'CAList':
            if Inputs['CAList']:
                window['BGUser'].update(Inputs['CAList'][0])
                     
        if event == 'CARefresh':
            ArkToken   = AppData['ArkToken']
            ArkAccount = FindCyberArkAccount(Token=ArkToken)
            if not ArkAccount or 'ErrorMessage' in ArkAccount:
                window['CAList'].update(values=['Login to see CyberArk Accounts here'])
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
                CAList = [(_['userName'] + (" " * 120) + _['id']) for _ in ArkAccount['value']]
                window['CAList'].update(values=CAList)
                
        if event in ['CALogin','CALoginReturn','CALoginCheckIn']:
            if event != 'CALoginReturn':
                window['-OUTPUT-'].update('')
                window['OutPutSection'].select()
                window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "CALoginReturn")
            else:
                ArkToken = Inputs['CALoginReturn']
                ArkAccount = FindCyberArkAccount(Token=ArkToken)
                if not ArkAccount or 'ErrorMessage' in ArkAccount:
                    window['-OUTPUT-'].update('')
                    window['OutPutSection'].select()
                    if ArkAccount and 'ErrorMessage' in ArkAccount:
                        print(ArkAccount['ErrorMessage'])
                    print('Login unsuccessful')
                    continue
                CAList = [(_['userName'] + (" " * 120) + _['id']) for _ in ArkAccount['value']]
                window['CAList'].update(values=CAList)
                ArkInfo = {'Token': ArkToken}
                SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                AppData = GetConfigAppData(WorkingDirectory,pin,Print=False)
                AppData = AppData if AppData != 'error' else {}
                
        if event in ['CheckIn','CALoginCheckIn']:
            window['-OUTPUT-'].update('')
            window['OutPutSection'].select()
            if not Inputs["BGUser"]:
                print('You must select or enter the username of an Account in CyberArk')
                continue
            BGUser = Inputs["BGUser"].split(' '*120)[0] if len(Inputs["BGUser"]) > 120 else Inputs["BGUser"]
            ArkToken = AppData['ArkToken']
            if not ArkToken and Inputs["ArkUser"] and Inputs["ArkPass"]:
                window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "CALoginCheckIn")
                continue
            elif not ArkToken:
                print('You will need to login to CyberArk to check in an Account')
                Continue
            if len(Inputs["BGUser"]) < 120:
                ArkAccount = FindCyberArkAccount(Username=BGUser, Token=ArkToken)
                if ArkAccount and 'ErrorMessage' in ArkAccount:
                    print(ArkAccount['ErrorMessage'])
                    print('Unable to check in password')
                    continue
            else:
                ArkAccount = {'id':Inputs["BGUser"].split(' '*120)[1]}
            
            CheckInResponse = CheckInCyberArkPassword(AccountId=ArkAccount['id'], Token=ArkToken, UserName=BGUser)
            
            if CheckInResponse and 'ErrorMessage' in CheckInResponse:
                print(CheckInResponse['ErrorMessage'])
                print('Unable to check in password')
                continue
            if CheckInResponse and hasattr(CheckInResponse,'ok') and CheckInResponse.ok:
                print('Account successfully checked in')
    
        if event == 'LPRefresh':
            LastPass = LastPassSession(AppData)
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
                window['LPList'].update(values=LPList)

                
        if event in ['LPLogin','LPLoginReturn']:
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
                    window['LPList'].update(values=LPList)
                    LPInfo = {
                        'Token'    : LastPass.token,
                        'Key'      : LastPass.encryption_key,
                        'SessionId': LastPass.id,
                        'Iteration': LastPass.key_iteration_count
                    }
                    SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                    AppData = GetConfigAppData(WorkingDirectory,pin,Print=False)
                    AppData = AppData if AppData != 'error' else {}
                    if window['LPOptions'].visible: window['LPOptions'].select()

        if event == 'LPList':
            LPSelected = window['LPSelected']
            LPSelected = LPSelected.Values if isinstance(LPSelected,sg.PySimpleGUI.Listbox) else []
            if Inputs['LPList'] and Inputs['LPList'][0] not in LPSelected:
                values = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0].startswith('C_')}
                NewList = LPSelected + Inputs['LPList']
                window['LPSelected'].update(values=NewList)
                for _ in Slots: 
                    window[f"C_{_}"].update(values=NewList)
                window.fill(values)
            
        
        if event == 'LPSelected':
            LPSelected = window['LPSelected']
            LPSelected = LPSelected.Values if isinstance(LPSelected,sg.PySimpleGUI.Listbox) else []
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
                    window['OutPutSection'].select()
                    print(Error)
                    OKSlots = None
                if OKSlots:
                    sl = {Slot : [(Slot+" "+_.label if _.label and _.label != 'ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ' else Slot) for _ in OKSlots if _.name == Slot][0]  for Slot in Slots}
                else:
                    window['OutPutSection'].select()
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
            window['KWLabelSearchText'].update(visible=KWTrue)
            window['SlotSelector'].update(visible=not KWTrue)

        #Save Current Input Values to Config File
        if event == "Save":
            window['-OUTPUT-'].update('')
            config = configparser.ConfigParser()
            ArkInfo = ArkInfo if ArkInfo else {'Token': ""}
            LPInfo  = LPInfo  if LPInfo  else {'Token':"",'Key':"",'SessionId':'','Iteration':''}
            SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window)
            AppData = GetConfigAppData(WorkingDirectory,pin,Print=False)
            AppData = AppData if AppData != 'error' else {}
        
        #Clear Print Statements From Output Element
        if event in ['ClearU','ClearD']:
            window['-OUTPUT-'].update('')
            continue
        
        if event == "FunctionReturn":
            FunctionReturn = Inputs['FunctionReturn']
        else:
            FunctionReturn = None
        
        #OnlyKey CyberArk Password Update Script
        if event == "Sync Passwords" or FunctionReturn:
            window['OutPutSection'].select()
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
                    window['LPList'].update(values=LPList)  
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
                if not FunctionReturn and not AppData['ArkToken']:
                    window['-OUTPUT-'].update('')
                    if not Inputs["ArkPass"] or not Inputs["BGUser"] or not Inputs["ArkUser"]:   
                        if not Inputs["ArkUser"]: print('Fill in the Network Username field to Login to CyberArk')
                        if not Inputs["ArkPass"]: print('Fill in the Network Password field to Login to CyberArk')
                        if not Inputs["BGUser"]:  print('You must select or enter the username of an Account in CyberArk')
                        continue
                    CALoginAttempt = 1
                    window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "FunctionReturn")
                    continue
                else:
                    if not FunctionReturn: window['-OUTPUT-'].update('')
                    if not Inputs["BGUser"]:
                        print('You must select or enter the username of an Account in CyberArk')
                        continue
                    BGUser = Inputs["BGUser"].split(' '*120)[0] if len(Inputs["BGUser"]) > 120 else Inputs["BGUser"]
                    ArkToken = FunctionReturn if FunctionReturn else AppData['ArkToken']
                    ArkInfo  = {'Token': ArkToken}
                    ArkAccount = FindCyberArkAccount(Username=BGUser, Token=ArkToken)
                    if ArkAccount and 'ErrorMessage' in ArkAccount and CALoginAttempt == 0 and Inputs["ArkUser"] and Inputs["ArkPass"]:
                        CALoginAttempt = 1
                        window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "FunctionReturn")
                        continue
                    elif ArkAccount and 'ErrorMessage' in ArkAccount and CALoginAttempt != 0:
                        print(ArkAccount['ErrorMessage'])
                        print('CyberArk Login unsuccessful')
                        CALoginAttempt = 0
                        continue
                    elif ArkAccount == None:
                        print("Unable Find CyberArk Account for " + BGUser)
                        print("Try Re-entering CyberArk Account Name")
                        continue
                    elif not Inputs["ArkUser"] or not Inputs["ArkPass"]:
                        if not Inputs["ArkUser"]: print('Fill in the Network Username field to Login to CyberArk')
                        if not Inputs["ArkPass"]: print('Fill in the Network Password field to Login to CyberArk')
                        continue
                    
                    ArkAccountPass = GetCyberArkPassword(AccountId=ArkAccount['id'], Token=ArkToken, UserName=BGUser)
                    if ArkAccountPass and isinstance(ArkAccountPass,dict) and 'ErrorMessage' in ArkAccountPass:
                        print(ArkAccountPass['ErrorMessage'])
                        print(f'Unable to retreive password from CyberArk for {BGUser}')
                        continue
                    
                    if not ArkAccountPass:
                        print(f'Unable to retreive password from CyberArk for {BGUser}')
                        continue
                    
                    if Inputs['OK']:
                        TrueSlots = {_[0].replace("S_","") : _[1] for _ in Inputs.items() if isinstance(_[0],str) and "S_" in _[0] and _[1] == True}
                        if Inputs["PSlotsRadio"] and not TrueSlots:
                            window['-OUTPUT-'].update('')
                            print('You must select which OnlyKey slots you would like to update')
                            sleep(2)
                            window.key_dict['OKOptions'].select()
                            continue
                        UpdateOnlyKey(Password=ArkAccountPass,OK_Keyword=Inputs["OK_Keyword"],SlotSelections=TrueSlots,SlotsTrue=Inputs["PSlotsRadio"])
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
                            window['LPList'].update(values=LPList)
                        
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
                                SetLastPassPassword(LastPass,Password=ArkAccountPass,Account=Account)
                        elif BGUser in [_.name for _ in LastPass.accounts]:
                            Account = [_ for _ in LastPass.accounts if _.name == BGUser][0]
                            SetLastPassPassword(LastPass,Password=ArkAccountPass,Account=Account)
                        else:    
                            SetLastPassPassword(LastPass,Password=ArkAccountPass,Username=BGUser)       
                        
                    FunctionReturn = False
                    SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                    AppData = GetConfigAppData(WorkingDirectory,pin,Print=False)
                    AppData = AppData if AppData != 'error' else {}
                    continue
    except Exception as error:
        print(error)
        continue    
window.close()
#endregion