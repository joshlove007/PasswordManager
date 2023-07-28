#region Import Modules
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
def collapse(layout, key:str, visible:bool):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this seciton visible / invisible
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    return sg.pin(sg.Column(layout, key=key,visible=visible))
#endregion
#endregion

#region Read Config File
config = configparser.ConfigParser()
WorkingDirectory = getcwd() + '\\Config.ini'
config.read(WorkingDirectory)
ConfigUsername = config.get('OnlyKey Update Parameters','Network Username',fallback=getuser())
ConfigBGUser   = config.get('OnlyKey Update Parameters','BG Username'     ,fallback="")
ConfigOKWord   = config.get('OnlyKey Update Parameters','OnlyKey Keyword' ,fallback="BG*")
ConfigColumn   = config.get('OnlyKey Update Parameters','Ouput Collapsed' ,fallback="False")
if ConfigUsername == "": ConfigUsername = getuser()
#endregion

#region Collapsible Section Layouts
#tooltips
ToggleToolTip = "Click Here to Collapse or Expand Output Section"
OutputToolTip = "The Outputs and Error Messages of the Application Appear Here"
ClearToolTip  = "Click Here to Clear Text from the Output Element"

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
    [sg.Output(size=(61,8), key='-OUTPUT-',visible=True,echo_stdout_stderr=True,pad=((1,0),(0,3)),tooltip=OutputToolTip)],
    [sg.Button('Clear',key='ClearD',tooltip=ClearToolTip)]
]
#endregion

#region Main Application Layout
#toolTips
UserNameToolTip = "Your CyberArk Username (sAMAccountName): For Logging into CyberArk"
PasswordToolTip = "Your 16+ Digit Domain Password: For Logging into CyberArk"
BGUserToolTip   = "Your Admin (BG) Username (sAMAccountName): Account to Search for in CyberArk"
KeywordToolTip  = "OnlyKey Slot Label Keyword: Search Keyword for Labels of Slots You Would Like to Update"
SubmitToolTip   = "Click Here to Login to CyberArk and Update OnlyKey Slot Labels"
SaveToolTip     = "Click Here to Save Field Values and Output Toggle to Config.ini File (Network Password is NOT Saved)"

#Layout
Collapsed = True if ConfigColumn == "True" else False
layout = [
    [sg.Text('Please enter your CyberArk Information')],
    [sg.Text('Network Username', size=(15, 1),tooltip=UserNameToolTip), sg.InputText(default_text=ConfigUsername,key="ArkUser",tooltip=UserNameToolTip)],
    [sg.Text('Network Password', size=(15, 1),tooltip=PasswordToolTip), sg.InputText(password_char="*",key="ArkPass",tooltip=PasswordToolTip)],
    [sg.Text('BG Username', size=(15, 1),tooltip=BGUserToolTip), sg.InputText(default_text=ConfigBGUser,key="BGUser",tooltip=BGUserToolTip)],
    [sg.Text('Label Keyword', size=(15, 1),tooltip=KeywordToolTip), sg.InputText(default_text=ConfigOKWord,key="OK_Keyword",tooltip=KeywordToolTip)],
    [sg.Submit("Submit",bind_return_key=True,tooltip=SubmitToolTip),sg.Push(), sg.Button('Save',key="Save",tooltip=SaveToolTip)],
    [collapse(ColSectionU, "SectionU",visible=Collapsed)],
    [collapse(ColSectionD, "SectionD",visible=(not Collapsed))]
]
#alternate up/down arrows ⩢⩠▲▼⩓⩔
#Window and Titlebar Options
window = sg.Window('OnlyKey CyberArk Update', layout)
#endregion

#region Application Behaviour if Statments - Event Loop
while True: 
    #Read Events(Actions) and Users Inputs From Window 
    WindowRead = window.read()
    event      = WindowRead[0]
    Inputs     = WindowRead[1]
    #event, Inputs = window.read()
    
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
        config['OnlyKey Update Parameters']['Network Username'] = Inputs["ArkUser"]
        config['OnlyKey Update Parameters']['BG Username']      = Inputs["BGUser"]
        config['OnlyKey Update Parameters']['OnlyKey Keyword']  = Inputs["OK_Keyword"]
        config['OnlyKey Update Parameters']['Ouput Collapsed']  = str(Collapsed)
        with open(WorkingDirectory, 'w') as configfile:
            config.write(configfile)
        print(f"Configuration Saved To: {WorkingDirectory}")
        
    #End Script when Window is Closed
    if event == sg.WIN_CLOSED:
        break
    
    #Clear Print Statements From Output Element
    if event in ['ClearU','ClearD']:
        window['-OUTPUT-'].update('')
        window['SmallOut'].update('')
        continue
    
    #region OnlyKey CyberArk Password Update Script
    if event == "Submit":
        window['-OUTPUT-'].update('')
        window['SmallOut'].update('')
        ArkUser    = Inputs["ArkUser"]
        ArkPass    = Inputs["ArkPass"]
        BGUsername = Inputs["BGUser"]
        OK_Keyword = Inputs["OK_Keyword"]
        
        ArkToken  = GetCyberArkToken(ArkPass=ArkPass,ArkUser=ArkUser)
        if 'ErrorMessage' in ArkToken:
            print("Unable Retrieve CyberArk Token")
            print(ArkToken['ErrorMessage'])
            print("Try Re-entering Your CyberArk Credentials or Press EXIT")
            continue
        else:
            print(f"DUO Push Accepted - {ArkUser} Authenticated")

        ArkAccount = FindCyberArkAccount(Username=BGUsername, Token=ArkToken)
        if ArkAccount == None:
            print("Unable Find CyberArk Account for " + BGUsername)
            print("Try Re-entering Your BG Username or Press EXIT")
            continue

        BGPassword = GetCyberArkPassword(AccountId=ArkAccount['id'], Token=ArkToken, UserName=BGUsername)

        OKSlots = OnlyKey().getlabels()

        BGSlots = []
        for Slot in OKSlots:
            if OK_Keyword in Slot.label:
                print(f'Setting Slot {Slot.name} {Slot.label}')
                BGSlots.append(Slot)
                OnlyKey().setslot(slot_number=Slot.number, message_field=MessageField.PASSWORD, value=BGPassword)
        continue  
    #endregion
      
window.close()
#endregion