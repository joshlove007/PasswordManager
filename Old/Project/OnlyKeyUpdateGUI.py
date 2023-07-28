#region Import Modules
from time import sleep
from onlykey.client import OnlyKey,MessageField
from requests import get,post
from getpass import getuser
from os import getcwd
import configparser
import PySimpleGUI as sg
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
    1
    Token = (post(url=URL,json=Body)).json()
    return Token 
#endregion

#region Function FindCyberArkAccount
def FindCyberArkAccount(Username,Token,ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    
    URL    = ArkURL + "/Accounts" + '?search=' + Username

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
def collapse(layout, key:str, visible:bool,pad:int=0):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this seciton visible / invisible
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    return sg.pin(sg.Column(layout, key=key,visible=visible,pad=pad))
#endregion

#region Function UpdateOnlyKey
def UpdateOnlyKey(ArkUser,ArkPass,BGUsername,OK_Keyword,SlotSelections:dict,SlotsTrue:bool):
    for param in (ArkUser,ArkPass,BGUsername,(SlotSelections if SlotsTrue else OK_Keyword)):
        if not bool(param):
            print("All fields must be compeleted, fill in all empty fields")
            return
        
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
        
    ArkToken  = GetCyberArkToken(ArkPass=ArkPass,ArkUser=ArkUser)
    if 'ErrorMessage' in ArkToken:
        print("Unable Retrieve CyberArk Token")
        print(ArkToken['ErrorMessage'])
        print("Try Re-entering Your CyberArk Credentials or Press EXIT")
        onlykey.close()
        return
    else:
        print(f"DUO Push Accepted - {ArkUser} Authenticated")

    ArkAccount = FindCyberArkAccount(Username=BGUsername, Token=ArkToken)
    if ArkAccount == None:
        print("Unable Find CyberArk Account for " + BGUsername)
        print("Try Re-entering Your BG Username or Press EXIT")
        onlykey.close()
        return

    BGPassword = GetCyberArkPassword(AccountId=ArkAccount['id'], Token=ArkToken, UserName=BGUsername)

    OKSlots = onlykey.getlabels()

    BGSlots = []
    for Slot in OKSlots:
        if SlotsTrue and Slot.name in SlotSelections or not SlotsTrue and OK_Keyword in Slot.label:
            print(f'Setting Slot {Slot.name} {Slot.label}')
            BGSlots.append(Slot)
            sleep(.1)
            onlykey.setslot(slot_number=Slot.number, message_field=MessageField.PASSWORD, value=BGPassword)
    onlykey.close()
    if not bool(BGSlots):
        print(f"Unable to find any slots containing the keyword \"{OK_Keyword}\"")
    return
#endregion
#endregion

#region Read Config File
Slots = ["1a","1b","2a","2b","3a","3b","4a","4b","5a","5b","6a","6b"]

config = configparser.ConfigParser()
WorkingDirectory = getcwd() + '/Config.ini'
config.read(WorkingDirectory)
ConfigUsername = config.get('OnlyKey Update Parameters','Network Username'        ,fallback=getuser())
ConfigBGUser   = config.get('OnlyKey Update Parameters','BG Username'             ,fallback="")
ConfigOKWord   = config.get('OnlyKey Update Parameters','OnlyKey Keyword'         ,fallback="BG*")
ConfigOutput   = config.get('OnlyKey Update Parameters','Ouput Collapsed'         ,fallback="False")
ConfigKWTrue   = config.get('OnlyKey Update Parameters','Keyword Search Selected' ,fallback="False") 
ConfigSlots    = {Slot : (True if config.get('Selected Slots',Slot,fallback="False") == "True" else False) for Slot in Slots}
if ConfigUsername == "": ConfigUsername = getuser()
#endregion

#region Collapsible Section Layouts
#tooltips
ToggleToolTip   = "Click Here to Collapse or Expand Output Section"
OutputToolTip   = "The Outputs and Error Messages of the Application Appear Here"
ClearToolTip    = "Click Here to Clear Text from the Output Element"
KeywordToolTip  = "OnlyKey Slot Label Keyword: Search Keyword for Labels of Slots You Would Like to Update"

#Left Column Uppersection
left_col = [
    [sg.T('⩠ Output:', enable_events=True, k='ToggleSection2',pad=((5,1),(5,5)),tooltip=ToggleToolTip)],
    [sg.Button('Clear',key='ClearU',tooltip=ClearToolTip)]
]
#RightColumn Uppersection
right_col =[[sg.Output(size=(52,3), key='SmallOut',echo_stdout_stderr=True,pad=0,tooltip=OutputToolTip)]]

#Upper collapsible Section - Columns Combined
ColSectionU = [[sg.Column(left_col, element_justification='left',pad=0),sg.Column(right_col, element_justification='left',pad=0)]]

#Lower collapsible section
ColSectionD = [
    [sg.T('⩢ Output:', enable_events=True, k='ToggleSection',pad=((5,5),(5,5)),tooltip=ToggleToolTip)],
    [sg.Output(size=(60,8), key='-OUTPUT-',visible=True,echo_stdout_stderr=True,pad=((1,0),(0,3)),tooltip=OutputToolTip)],
    [sg.Button('Clear',key='ClearD',tooltip=ClearToolTip)]
]

#Outputs
Collapsed = True if ConfigOutput == "True" else False
OutPutSection = [
    [collapse(ColSectionU, "SectionU",visible=Collapsed)],
    [collapse(ColSectionD, "SectionD",visible=(not Collapsed))]
]

#KeyWord TextBox
KWLabelSearchText = [[sg.Text('Label Keyword', size=(16, 1),key="LabelKeyword",tooltip=KeywordToolTip), sg.InputText(default_text=ConfigOKWord,key="OK_Keyword",tooltip=KeywordToolTip)]]

#Slot Section Checkboxes
Cs = ConfigSlots
SlotSelector = [             
    [sg.Push(),sg.Text("Pick Slots To Update Password",justification="center",size=(60, 1),visible=True,pad=((15,5),(5,5))),sg.Push()],      
    [sg.Checkbox('1a',k='S_1a',default=Cs['1a'],p=((20,5),(5,5))) , sg.Checkbox('1b',k='S_1b',default=Cs['1b'],p=((5,5),(5,5))) ,sg.Push(),sg.Checkbox('2a',k='S_2a',default=Cs['2a'],p=((5,5),(5,5))) , sg.Checkbox('2b',k='S_2b',default=Cs['2b'],p=((5,15),(5,5)))],      
    [sg.Checkbox('3a',k='S_3a',default=Cs['3a'],p=((20,5),(5,5))) , sg.Checkbox('3b',k='S_3b',default=Cs['3b'],p=((5,5),(5,5))) ,sg.Push(),sg.Checkbox('4a',k='S_4a',default=Cs['4a'],p=((5,5),(5,5))) , sg.Checkbox('4b',k='S_4b',default=Cs['4b'],p=((5,15),(5,5)))],      
    [sg.Checkbox('5a',k='S_5a',default=Cs['5a'],p=((20,5),(5,20))), sg.Checkbox('5b',k='S_5b',default=Cs['5b'],p=((5,5),(5,20))),sg.Push(),sg.Checkbox('6a',k='S_6a',default=Cs['6a'],p=((5,5),(5,20))), sg.Checkbox('6b',k='S_6b',default=Cs['6b'],p=((5,15),(5,20)))]
]
#endregion

#region Main Application Window Layout
#toolTips
UserNameToolTip = "Your CyberArk Username (sAMAccountName): For Logging into CyberArk"
PasswordToolTip = "Your 16+ Digit Domain Password: For Logging into CyberArk"
BGUserToolTip   = "Your Admin (BG) Username (sAMAccountName): Account to Search for in CyberArk"
SubmitToolTip   = "Click Here to Login to CyberArk and Update OnlyKey Slot Passwords"
SaveToolTip     = "Click Here to Save Field Values and Output Toggle to Config.ini File (Network Password is NOT Saved)"

#Layout
KWTrue  = False if ConfigKWTrue == "False" else True
visible = KWTrue
layout = [
    [sg.Text('Please enter your CyberArk Information')],
    [sg.Text('Network Username', size=(16, 1),tooltip=UserNameToolTip), sg.InputText(default_text=ConfigUsername,key="ArkUser",tooltip=UserNameToolTip)],
    [sg.Text('Network Password', size=(16, 1),tooltip=PasswordToolTip), sg.InputText(password_char="*",key="ArkPass",tooltip=PasswordToolTip)],
    [sg.Text('CyberArk Username', size=(16, 1),tooltip=BGUserToolTip), sg.InputText(default_text=ConfigBGUser,key="BGUser",tooltip=BGUserToolTip)],
    [collapse(KWLabelSearchText, "KWLabelSearchText",visible=KWTrue)],
    [sg.Text('Slot Selction Method', size=(16, 1)),sg.Push(),sg.Radio('Pick Slots','SMethod',key="PSlotsRadio",enable_events=True,default=not KWTrue),sg.Radio('Keyword Search','SMethod',key="KWSearch",default=KWTrue,enable_events=True),sg.Push()],
    [sg.Submit("Submit",bind_return_key=True,tooltip=SubmitToolTip),sg.Button(button_text='Save',key="Save",tooltip=SaveToolTip),sg.Push(),sg.Button(button_text='Show/Hide Slots',key="ShowSlots",visible=not KWTrue)],
    [collapse(OutPutSection, "OutPutSection",visible=visible)],
    [collapse(SlotSelector, "SlotSelector",visible=not visible)]
]
#alternate up/down arrows ⩢⩠▲▼⩓⩔
#Window and Titlebar Options
window = sg.Window('OnlyKey CyberArk Update', layout,font=("Helvetica", 11))
#endregion

#region Application Behaviour if Statments - Event Loop
while True: 
    #Read Events(Actions) and Users Inputs From Window 
    WindowRead = window.read()
    event      = WindowRead[0]
    Inputs     = WindowRead[1]
    
    #End Script when Window is Closed
    if event == sg.WIN_CLOSED:
        break
    
    #Toggle Slot Picker Element
    if event == "PSlotsRadio":
        window['KWLabelSearchText'].update(visible=False)
        window['ShowSlots'].update(visible=True)
    
    #Toggle Keyword Search Element
    if event == "KWSearch":
        visible = True
        window['KWLabelSearchText'].update(visible=visible)
        window['OutPutSection'].update(visible=visible)
        window['SlotSelector'].update(visible=not visible)
        window['ShowSlots'].update(visible=not visible)

    #Toggle Show Slot Selection Element
    if event == 'ShowSlots':
        visible = not visible
        window['OutPutSection'].update(visible=visible)
        window['SlotSelector'].update(visible=not visible)
    
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
        config.add_section('OnlyKey Update Parameters')
        config['OnlyKey Update Parameters']['Network Username']        = Inputs["ArkUser"]
        config['OnlyKey Update Parameters']['BG Username']             = Inputs["BGUser"]
        config['OnlyKey Update Parameters']['OnlyKey Keyword']         = Inputs["OK_Keyword"]
        config['OnlyKey Update Parameters']['Ouput Collapsed']         = str(Collapsed)
        config['OnlyKey Update Parameters']['Keyword Search Selected'] = str(Inputs["KWSearch"])
        config.add_section('Selected Slots')
        for Slot in Slots: config['Selected Slots'][Slot] = str(Inputs[f"S_{Slot}"])
        with open(WorkingDirectory, 'w') as configfile:
            config.write(configfile)
        print(f"Configuration Saved To: {WorkingDirectory}")
    
    #Clear Print Statements From Output Element
    if event in ['ClearU','ClearD']:
        window['-OUTPUT-'].update('')
        window['SmallOut'].update('')
        continue
    
    #OnlyKey CyberArk Password Update Script
    if event == "Submit":
        window['-OUTPUT-'].update('')
        window['SmallOut'].update('')
        visible = True
        window['OutPutSection'].update(visible=visible)
        window['SlotSelector'].update(visible=not visible)
        TrueSlots = {_[0].replace("S_","") : _[1] for _ in Inputs.items() if "S_" in _[0] and _[1] == True}
        ArkUser=Inputs["ArkUser"];ArkPass=Inputs["ArkPass"];BGUser=Inputs["BGUser"];OK_Keyword=Inputs["OK_Keyword"];PSlotsRadio=Inputs["PSlotsRadio"]
        window.perform_long_operation(lambda:UpdateOnlyKey(ArkUser=ArkUser,ArkPass=ArkPass,BGUsername=BGUser,OK_Keyword=OK_Keyword,SlotSelections=TrueSlots,SlotsTrue=PSlotsRadio), "FunctionReturn")
        continue
window.close()
#endregion