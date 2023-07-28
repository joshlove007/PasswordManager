#region Import Modules 
from ast import literal_eval as array
from optparse import Values
from time import sleep
from tkinter import W
from onlykey.client import OnlyKey,MessageField
from requests import get,post
from getpass import getuser
from os import getcwd
import configparser
import PySimpleGUI as sg
import lastpass
from lastpass.fetcher import make_key
from Crypto.Cipher import AES
import subprocess
from base64 import b64decode,b64encode
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
    print("Retrieving CyberArk Password For " + UserName)

    Header = {"Authorization" : Token}
    
    Body = {
        "reason" : "OnlyKey Update"
    }

    URL = ArkURL + "/Accounts/" + AccountId + "/Password/Retrieve"

    Response = (post(url=URL,json=Body,headers=Header)).json()
    return Response
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
def SaveToConfig(Inputs,Collapsed,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,Print=True):
        config = configparser.ConfigParser()
        user = getuser()
        Pass = (subprocess.run("hostname",capture_output=True,text=True)).stdout.rstrip()
        key = make_key(user,Pass,pin)

        config.add_section('General')
        config['General']['Ouput Collapsed']         = str(Collapsed)
        config.add_section('CyberArk')
        config['CyberArk']['Network Username']       = Inputs["ArkUser"]
        config['CyberArk']['Account']                = Inputs["BGUser"]
        config.add_section('LastPass')
        config['LastPass']['Username']               = Inputs["LPUser"]
        config['LastPass']['LastPass Selected']      = str(Inputs["LP"])
        config['LastPass']['LastPass Accounts']      = window.key_dict['LPSelected'].Values.__str__()
        config.add_section('OnlyKey')
        config['OnlyKey']['OnlyKey Selected']        = str(Inputs["OK"])
        config['OnlyKey']['Keyword']                 = Inputs["OK_Keyword"]
        config['OnlyKey']['Keyword Search Selected'] = str(Inputs["KWSearch"])
        config.add_section('App Data')
        config['App Data']['CyUsPa']                 = EncryptString(Inputs["ArkPass"],key)
        config['App Data']['LaUsPa']                 = EncryptString(Inputs["LPPass"],key)
        config['App Data']['CySeTo']                 = EncryptString(ArkInfo['Token'],key)
        config['App Data']['LaSeTo']                 = EncryptString(LPInfo['Token'],key)
        config['App Data']['LaSeKe']                 = EncryptString(b64encode(LPInfo['Key']).decode(),key)
        config['App Data']['LaSeId']                 = EncryptString(LPInfo['SessionId'],key)
        config['App Data']['LPSeIt']                 = EncryptString(str(LPInfo['Iteration']),key)

        config.add_section('Selected Slots')
        for _ in Slots: config['Selected Slots'][_]  = str(Inputs[f"S_{_}"])
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
        print("Try Re-entering Your CyberArk Credentials or Press EXIT")
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

def LastPassLogin(Inputs):    
    try:
        LastPass = lastpass.Session.login(username=Inputs['LPUser'],password=Inputs['LPPass']).OpenVault(include_shared=True)
    except Exception as e:
        return print(e)
    return LastPass

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
            sleep(.1)
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
WorkingDirectory = getcwd() + '/Config.ini'
config.read(WorkingDirectory)
ConfigOutput   =  Bool(config.get('General' ,'Ouput Collapsed'         ,fallback="False"))
       
ConfigUsername =       config.get('CyberArk','Network Username'        ,fallback=getuser())
ConfigBGUser   =       config.get('CyberArk','Account'                 ,fallback="")
       
ConfigLPUser   =       config.get('LastPass','Username'                ,fallback="")
ConfigLP       =  Bool(config.get('LastPass','LastPass Selected'       ,fallback="False"))
ConfigLPAccts  = array(config.get('LastPass','LastPass Accounts'       ,fallback="[]")) 
   
ConfigOK       =  Bool(config.get('OnlyKey' ,'OnlyKey Selected'        ,fallback="False"))
ConfigOKWord   =       config.get('OnlyKey' ,'Keyword'                 ,fallback="BG*")
ConfigKWTrue   =  Bool(config.get('OnlyKey' ,'Keyword Search Selected' ,fallback="False"))
ConfigSlots    = {Slot : Bool(config.get('Selected Slots',Slot,fallback="False")) for Slot in Slots}
if ConfigUsername == "": ConfigUsername = getuser()

#AppData Pin
test = False
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
        exit()

ArkInfo = {'Token':AppData['ArkToken']}
LPInfo = {
    'Token'    : AppData['LPToken'],
    'Key'      : AppData['LPKey'],
    'SessionId': AppData['LPsId'],
    'Iteration': AppData['LPIteration']
}
LastPass = LastPassSession(AppData)
if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
    LPList = ['Login to see LastPass entries here']
else:
    LPList = [(_.name + (" " * 50) + str(_.id)) for _ in LastPass.accounts]

ArkToken   = AppData['ArkToken']
ArkAccount = FindCyberArkAccount(Token=ArkToken)
if not ArkAccount or 'ErrorMessage' in ArkAccount:
    CAList = ['Login to see CyberArk Accounts here']
else:
    CAList = [(_['userName'] + (" " * 120) + _['id']) for _ in ArkAccount['value']]
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
left_col = [
    [sg.T('⩠ Output:', enable_events=True, k='ToggleSection2',pad=((5,1),(5,5)),tooltip=ToggleToolTip)],
    [sg.Button('Clear',key='ClearU',tooltip=ClearToolTip)]
]
right_col =[[sg.Output(size=(51,3), key='SmallOut',echo_stdout_stderr=True,pad=0,tooltip=OutputToolTip)]]

ColSectionU = [[sg.Column(left_col, element_justification='left',pad=0),sg.Column(right_col, element_justification='left',pad=0)]]

ColSectionD = [
    [sg.T('⩢ Output:', enable_events=True, k='ToggleSection',pad=((5,5),(5,5)),tooltip=ToggleToolTip)],
    [sg.Output(size=(60,8), key='-OUTPUT-',visible=True,echo_stdout_stderr=True,pad=((1,0),(0,3)),tooltip=OutputToolTip)],
    [sg.Button('Clear',key='ClearD',tooltip=ClearToolTip)]
]

Collapsed = ConfigOutput
OutPutSection = [
    [collapse(ColSectionU, "SectionU",visible=Collapsed)],
    [collapse(ColSectionD, "SectionD",visible=(not Collapsed))]
]
#endregion

#region LastPass
LastPassOptions = [
    [sg.Text("LastPass",justification="center",size=(52, 1),font=("Helvetica 12 bold"))],
    [sg.Text('Please enter your LastPass Information')],
    [sg.Text('LastPass Username', size=(16, 1)), sg.InputText(default_text=ConfigLPUser,key="LPUser")],
    [sg.Text('LastPass Password', size=(16, 1)), sg.InputText(password_char="*",key="LPPass",default_text=AppData['LPPass'])],
    [sg.Button(button_text='Refresh Account List',key="LPRefresh"),sg.Push(),sg.Button(button_text='Login',key="LPLogin")],
    [sg.Text('Select Account(s) to sync:', size=(32, 1),p=0),sg.Text('Selected Accounts:', size=(30, 1),p=0)],
    [sg.Listbox(values=LPList,k='LPList',s=(30,5),p=0,enable_events=True),sg.Listbox(values=ConfigLPAccts,k='LPSelected',s=(30,5),p=0,enable_events=True)],
    [sg.Push(),sg.Button(button_text='Clear All',key="LPClear")]
]

#endregion

#region CyberArk
CAOptions = [
    [sg.Text("CyberArk",justification="center",size=(52, 1),font=("Helvetica 12 bold"))],
    [sg.Text('Please enter your CyberArk Information')],
    [sg.Text('Network Username', size=(16, 1),tooltip=UserNameToolTip), sg.InputText(default_text=ConfigUsername,key="ArkUser",tooltip=UserNameToolTip)],
    [sg.Text('Network Password', size=(16, 1),tooltip=PasswordToolTip), sg.InputText(password_char="*",key="ArkPass",tooltip=PasswordToolTip,default_text=AppData['ArkPass'])],
    [sg.Button(button_text='Refresh Account List',key="CARefresh"),sg.Push(),sg.Button(button_text='Login',key="CALogin")],
    [sg.Text('Select an Account to sync:', size=(32, 1),p=0)],
    [sg.Listbox(values=CAList,key='CAList',size=(62,5),p=0,enable_events=True)],
    [sg.Text('Selected Account', size=(16, 1),tooltip=BGUserToolTip)  , sg.InputText(default_text=ConfigBGUser,key="BGUser",tooltip=BGUserToolTip)]
    #[sg.Button(button_text='Select',key="CASelect")]
]
#endregion

#region OnlyKey
Cs = ConfigSlots
KWLabelSearchText = [[sg.Text('Label Keyword', size=(15, 1),key="LabelKeyword",tooltip=KeywordToolTip), sg.InputText(default_text=ConfigOKWord,key="OK_Keyword",tooltip=KeywordToolTip)]]
KWTrue  = ConfigKWTrue


OkCheckLeft = [
    [sg.Checkbox('1a',k='S_1a',default=Cs['1a'],p=((20,5),(5,5))) , sg.Checkbox('1b',k='S_1b',default=Cs['1b'],p=((5,5),(5,5))) ],
    [sg.Checkbox('3a',k='S_3a',default=Cs['3a'],p=((20,5),(5,5))) , sg.Checkbox('3b',k='S_3b',default=Cs['3b'],p=((5,5),(5,5))) ],
    [sg.Checkbox('5a',k='S_5a',default=Cs['5a'],p=((20,5),(5,5))), sg.Checkbox('5b',k='S_5b',default=Cs['5b'],p=((5,5),(5,5)))]
]
OkCheckRight = [
    [sg.Checkbox('2a',k='S_2a',default=Cs['2a'],p=((5,5),(5,5))) , sg.Checkbox('2b',k='S_2b',default=Cs['2b'],p=((5,5),(5,5))) ],
    [sg.Checkbox('4a',k='S_4a',default=Cs['4a'],p=((5,5),(5,5))) , sg.Checkbox('4b',k='S_4b',default=Cs['4b'],p=((5,5),(5,5))) ],
    [sg.Checkbox('6a',k='S_6a',default=Cs['6a'],p=((5,5),(5,5))), sg.Checkbox('6b',k='S_6b',default=Cs['6b'],p=((5,5),(5,5)))]
]
OkCheckText = [[sg.Text('Select Slots to Update', size=(16,3))],[sg.Button(button_text='Clear All',key="OKClear")]]
SlotSelector = [[sg.Column(OkCheckText,p=((0,20),0)),sg.Column(OkCheckLeft,p=0),sg.VerticalSeparator(pad=(25,10)),sg.Column(OkCheckRight,p=0)]]

OkSlotSelection = [
    [sg.Text("Onlykey",justification="center",size=(52, 1),font=("Helvetica 12 bold"))],
    [sg.Text('Slot Selction Method', size=(16, 1)),sg.Push(),sg.Radio('Pick Slots','SMethod',k="PSlotsRadio",enable_events=True,default=not KWTrue),
    sg.Radio('Keyword Search','SMethod',k="KWSearch",default=KWTrue,enable_events=True),sg.Push()],
    [collapse(KWLabelSearchText, "KWLabelSearchText",KWTrue)],
    [collapse(SlotSelector, "SlotSelector",visible=not KWTrue)]
]
#endregion
#endregion

#region Main Application Window Layout

OKTrue = ConfigOK
LPTrue = ConfigLP
CATrue = True
#Layout
layout = [
    [collapse(CAOptions,'CAOptions',True)],
    [sg.Text('Sync Password(s) From', size=(18, 1)),sg.Radio('CyberArk','PFrom',k='CAFrom',enable_events=True,default=True),sg.Radio('LastPass','PFrom',k='LPFrom',enable_events=True,default=False),sg.Push()],
    [sg.Text('Sync Password(s) To', size=(18, 1)),sg.pin(sg.Checkbox('LastPass',k='LP',enable_events=True,default=LPTrue)),sg.pin(sg.Checkbox('OnlyKey',k='OK',enable_events=True,default=OKTrue)),sg.Push()],
    [sg.Submit("Sync Passwords",bind_return_key=True,tooltip=SubmitToolTip),sg.Button(button_text='Save',key="Save",tooltip=SaveToolTip),sg.Push(),
     sg.pin(sg.Button(button_text='LastPass Options',key="LPButton",visible=LPTrue)),sg.pin(sg.Button(button_text='OnlyKey Options',key="OKButton",visible=OKTrue))],
    [sg.HSep()],
    [collapse(OutPutSection, "OutPutSection",visible=True)],
    [collapse(LastPassOptions,'LPOptions',False)],
    [collapse(OkSlotSelection, 'OKOptions',False)]
]
#alternate up/down arrows ⩢⩠▲▼⩓⩔
#Window and Titlebar Options
Title = 'MEDHOST Password Manager'
window = sg.Window(Title, layout)
#endregion

LPVisible  = False
OutVisible = True
OKVisible  = False
CAVisible  = CATrue


#region Application Behaviour if Statments - Event Loop
while True:
    #Read Events(Actions) and Users Inputs From Window 
    WindowRead = window.read()
    event      = WindowRead[0]
    Inputs     = WindowRead[1]
    
    #End Script when Window is Closed
    if event == sg.WIN_CLOSED:
        break
    
    if event == 'LPFrom' and CATrue:
        CATrue = False
        CAVisible = False
        LPVisible = True
        OKVisible = False
        window,layout = MoveRows(window,layout,Rows=[(0,6),(6,0)])
        window['LPOptions'].update(visible=LPVisible)
        window['OKOptions'].update(visible=OKVisible)
        window['CAOptions'].update(visible=CAVisible)
        window['OutPutSection'].update(visible=True)
        window['LPButton'].update(visible=CATrue)
        window['LP'].update(visible=False)

    if event == 'CAFrom' and not CATrue:
        CATrue = True
        LPVisible = False
        CAVisible = True
        OKVisible = False
        window,layout = MoveRows(window,layout,Rows=[(6,0),(0,6)])
        window['CAOptions'].update(visible=CAVisible)
        window['OKOptions'].update(visible=OKVisible)
        window['OutPutSection'].update(visible=True)
        window['LPOptions'].update(visible=LPVisible)
        window['LPButton'].update(visible=CATrue)
        window['LP'].update(visible=True)
    
    if event == 'LP':
        window['LPButton'].update(visible=Inputs["LP"])
        window['LPOptions'].update(visible=False)
  
    if event == 'OK':
        window['OKButton'].update(visible=Inputs["OK"])
        window['OKOptions'].update(visible=False)

    if event == 'OKClear':
        for _ in Slots: window[f"S_{_}"].update(value=False)
        
    if event == 'CAList':
        if Inputs['CAList']:
            window['BGUser'].update(Inputs['CAList'][0])
            
    if event == 'CARefresh':
        ArkToken   = AppData['ArkToken']
        ArkAccount = FindCyberArkAccount(Token=ArkToken)
        if not ArkAccount or 'ErrorMessage' in ArkAccount:
            window['LPList'].update(values=['Login to see CyberArk Accounts here'])
            print('Unable to refresh accounts list, login to get new token')
            continue
        CAList = [(_['userName'] + (" " * 120) + _['id']) for _ in ArkAccount['value']]
        window['CAList'].update(values=CAList)
        
    if event == 'CALogin' or event == 'CALoginReturn':
        if event != 'CALoginReturn':
            LPVisible = False
            OKVisible = False
            window['LPOptions'].update(visible=LPVisible)
            window['OKOptions'].update(visible=OKVisible)
            window['OutPutSection'].update(visible=True)
            window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "CALoginReturn")
        else:
            ArkToken = Inputs['CALoginReturn']
            ArkAccount = FindCyberArkAccount(Token=ArkToken)
            if not ArkAccount:
                print('Login unsuccessful')
                continue
            CAList = [(_['userName'] + (" " * 120) + _['id']) for _ in ArkAccount['value']]
            window['CAList'].update(values=CAList)
            ArkInfo = {'Token': ArkToken}
            SaveToConfig(Inputs,Collapsed,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
            AppData = GetConfigAppData(WorkingDirectory,pin,Print=False)
            AppData = AppData if AppData != 'error' else {}
             
    if event == 'LPLogin' or event == 'LPLoginReturn':
        if event != 'LPLoginReturn':
            LPVisible = False
            OKVisible = False
            window['OKOptions'].update(visible=OKVisible)
            window['OutPutSection'].update(visible=True)
            if Inputs['CAFrom']:
                LPVisible = False
                window['LPOptions'].update(visible=LPVisible)
            window.perform_long_operation(lambda:LastPassLogin(Inputs), "LPLoginReturn")
        else:
            LastPass = Inputs['LPLoginReturn']
            if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                OKVisible = False
                window['OutPutSection'].update(visible=True)
                if Inputs['CAFrom']:
                    LPVisible = False
                    window['LPOptions'].update(visible=LPVisible)
                window['OKOptions'].update(visible=OKVisible)
                print('Login unsuccessful')
                LastPass = None
                continue
            else:
                LPVisible = True
                OKVisible = False
                if Inputs['CAFrom']:
                    window['OutPutSection'].update(visible=not LPVisible)
                window['OKOptions'].update(visible=OKVisible)
                window['LPOptions'].update(visible=LPVisible)
                LPList = [(_.name + (" " * 50) + str(_.id)) for _ in LastPass.accounts]
                window['LPList'].update(values=LPList)
                LPInfo = {
                    'Token'    : LastPass.token,
                    'Key'      : LastPass.encryption_key,
                    'SessionId': LastPass.id,
                    'Iteration': LastPass.key_iteration_count
                }
                SaveToConfig(Inputs,Collapsed,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                AppData = GetConfigAppData(WorkingDirectory,pin,Print=False)
                AppData = AppData if AppData != 'error' else {}

    if event == 'LPList':
        LPSelected = window['LPSelected']
        LPSelected = LPSelected.Values if isinstance(LPSelected,sg.PySimpleGUI.Listbox) else []
        if Inputs['LPList'] and Inputs['LPList'][0] not in LPSelected:
            window['LPSelected'].update(values=LPSelected + Inputs['LPList'])
    
    if event == 'LPSelected':
        LPSelected = window['LPSelected']
        LPSelected = LPSelected.Values if isinstance(LPSelected,sg.PySimpleGUI.Listbox) else []
        if Inputs['LPSelected']:
            LPSelected.remove(Inputs['LPSelected'][0])
            window['LPSelected'].update(values=LPSelected)
    
    if event == 'LPClear':
        window['LPSelected'].update(values=[])       

    #Toggle Slot Picker Element
    if event in ["PSlotsRadio",'KWSearch']:
        KWTrue = not KWTrue
        window['KWLabelSearchText'].update(visible=KWTrue)
        window['SlotSelector'].update(visible=not KWTrue)

    #Toggle Show Slot Selection Element
    if event == 'OKButton':
        OKVisible = not OKVisible
        LPVisible = False
        window['OutPutSection'].update(visible=not OKVisible)
        if Inputs['CAFrom']:
            LPVisible = False
            window['LPOptions'].update(visible=LPVisible)
        if Inputs['LPFrom']:
            CAVisible = False
            window['CAOptions'].update(visible=CAVisible)
        window['OKOptions'].update(visible=OKVisible)
          
    if event == 'LPButton':
        LPVisible = not LPVisible
        OKVisible = False
        window['OutPutSection'].update(visible=not LPVisible)
        window['OKOptions'].update(visible=OKVisible)
        window['LPOptions'].update(visible=LPVisible)
        if LPVisible and not LastPass:
            LastPass = LastPassSession(AppData)
            if not LastPass or not isinstance(LastPass,lastpass.Session) or not LastPass.accounts:
                window['LPList'].update(values=['Login to see LastPass entries here'])
                continue
            LPList = [(_.name + (" " * 50) + str(_.id)) for _ in LastPass.accounts]
            window['LPList'].update(values=LPList)

    
    #Toggle Collapsible Output Element
    if event in ['ToggleSection','ToggleSection2']:
        Collapsed = not Collapsed
        window['SectionD'].update(visible=(not Collapsed))
        window['SectionU'].update(visible=Collapsed)
        
    #Save Current Input Values to Config File
    if event == "Save":
        window['-OUTPUT-'].update('')
        window['SmallOut'].update('')
        config = configparser.ConfigParser()
        ArkInfo = ArkInfo if ArkInfo else {'Token': ""}
        LPInfo  = LPInfo  if LPInfo  else {'Token':"",'Key':"",'SessionId':'','Iteration':''}
        SaveToConfig(Inputs,Collapsed,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window)
    
    #Clear Print Statements From Output Element
    if event in ['ClearU','ClearD']:
        window['-OUTPUT-'].update('')
        window['SmallOut'].update('')
        continue
    
    if event == "FunctionReturn":
        FunctionReturn = Inputs['FunctionReturn']
    else:
        FunctionReturn = None
    
    #OnlyKey CyberArk Password Update Script
    if event == "Sync Passwords" or FunctionReturn:
        window['-OUTPUT-'].update('')
        window['SmallOut'].update('')
        OutVisible = True
        window['OutPutSection'].update(visible=OutVisible)
        window['OKOptions'].update(visible=False)
        window['LPOptions'].update(visible=False)

        if not FunctionReturn and not AppData['ArkToken']:
            window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "FunctionReturn")
            continue
        else:
            ArkToken = FunctionReturn if FunctionReturn else AppData['ArkToken']
            BGUser = Inputs["BGUser"].split(' '*120)[0] if len(Inputs["BGUser"]) > 120 else Inputs["BGUser"]
            ArkAccount = FindCyberArkAccount(Username=BGUser, Token=ArkToken)
            if ArkAccount and 'ErrorMessage' in ArkAccount:
                window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=Inputs["ArkPass"]), "FunctionReturn")
                continue
            
            if ArkAccount == None:
                print("Unable Find CyberArk Account for " + BGUser)
                print("Try Re-entering Your BG Username or Press EXIT")
                continue
            ArkAccountPass = GetCyberArkPassword(AccountId=ArkAccount['id'], Token=ArkToken, UserName=BGUser)
            ArkInfo = {'Token': ArkToken}

            if Inputs['OK']:
                TrueSlots = {_[0].replace("S_","") : _[1] for _ in Inputs.items() if "S_" in _[0] and _[1] == True}
                UpdateOnlyKey(Password=ArkAccountPass,OK_Keyword=Inputs["OK_Keyword"],SlotSelections=TrueSlots,SlotsTrue=Inputs["PSlotsRadio"])
            if Inputs['LP']:
                if not LastPass:
                    LastPass = LastPassLogin(Inputs)
                if not LastPass:
                    continue
                LPInfo = {
                    'Token'    : LastPass.token,
                    'Key'      : LastPass.encryption_key,
                    'SessionId': LastPass.id,
                    'Iteration': LastPass.key_iteration_count
                }
                LPSelected = window['LPSelected'].get_list_values()
                if LPSelected:
                    for Item in LPSelected:
                        Account = [_ for _ in LastPass.accounts if _.id == int(Item.split(' ' * 50)[1])][0]
                        SetLastPassPassword(LastPass,Password=ArkAccountPass,Account=Account)
                else:
                    SetLastPassPassword(LastPass,Password=ArkAccountPass,Username=BGUser)       
                
            FunctionReturn = False
            SaveToConfig(Inputs,Collapsed,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
            AppData = GetConfigAppData(WorkingDirectory,pin,Print=False)
            AppData = AppData if AppData != 'error' else {}
            continue
window.close()
#endregion