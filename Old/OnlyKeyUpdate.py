import requests
import subprocess
import onlykey
import getpass
import os
import configparser
import PySimpleGUI as sg

layout = [
    [sg.Text('Please enter your Name, Address, Phone')],
    [sg.Text('Name', size=(15, 1)), sg.InputText()],
    [sg.Text('Address', size=(15, 1)), sg.InputText()],
    [sg.Text('Phone', size=(15, 1)), sg.InputText()],
    [sg.Submit(), sg.Cancel()],
    [sg.Text('Output:')],
    [sg.Output(size=(50,10), key='-OUTPUT-')],
    [sg.Button('Clear'), sg.Button('Exit')]
]

window = sg.Window('OnlyKey CyberArk Update', layout)
event, values = window.read()
window.close()
print(event, values[0], values[1], values[2])

while True:             # Event Loop
    event, values = window.read()
    print(event, values)
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == 'Clear':
        window['-OUTPUT-'].update('')
window.close()


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
    Token = (requests.post(url=URL,json=Body)).json()
    return Token 
#endregion

#region Function FindCyberArkAccount
def FindCyberArkAccount(Username,Token,ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    
    URL    = ArkURL + "/Accounts" + '?search=' + Username

    Response = (requests.get(url=URL,headers=Header)).json()
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

    Response = (requests.post(url=URL,json=Body,headers=Header)).json()
    return Response
#endregion
#endregion

#region Config File
config = configparser.ConfigParser()
WorkingDirectory = os.getcwd()
config.read(WorkingDirectory + '\\Config.ini')
ConfigUsername = config['OnlyKey Update Parameters']['Network Username']
ConfigBGUser   = config['OnlyKey Update Parameters']['BG Username']
#endregion

if ConfigUsername == '':
    PSCommand  = '"$env:username"'
    ArkUser    = (subprocess.run('powershell.exe ' + PSCommand,capture_output=True,text=True)).stdout.rstrip()
else:
    ArkUser = ConfigUsername

print("Network Username: " + ArkUser)
ArkPass = getpass.getpass(prompt='Network Password: ') 

ArkToken  = GetCyberArkToken(ArkPass=ArkPass,ArkUser=ArkUser)
while 'ErrorMessage' in ArkToken:
    print("Unable Retrieve CyberArk Token")
    print(ArkToken['ErrorMessage'])
    print("Try Re-entering Your CyberArk Credentials or Type EXIT to Exit Script")
    ArkUser = input("Network Username: ")
    if ArkUser == "EXIT":
        exit()
    ArkPass = getpass.getpass(prompt='Network Password: ') 
    ArkToken   = GetCyberArkToken(ArkPass=ArkPass,ArkUser=ArkUser)

if ConfigBGUser == '':
    BGUsername = input("BG Username: ")
else:
    BGUsername = ConfigBGUser

ArkAccount = FindCyberArkAccount(Username=BGUsername, Token=ArkToken)
while ArkAccount == None:
    print("Unable Find CyberArk Account for " + BGUsername)
    print("Try Re-entering Your BG Username or Type EXIT to Exit Script")
    BGUsername = input("BG Username: ")
    if BGUsername == "EXIT":
        exit()
    ArkAccount = FindCyberArkAccount(Username=BGUsername, Token=ArkToken)

BGPassword = GetCyberArkPassword(AccountId=ArkAccount['id'], Token=ArkToken, UserName=BGUsername)

OKSlots = onlykey.client.OnlyKey().getlabels()

BGSlots = []
for Slot in OKSlots:
    if "BG*" in Slot.label:
        print(f'Setting Slot {Slot.name} {Slot.label}')
        BGSlots.append(Slot)
        onlykey.client.OnlyKey().setslot(slot_number=Slot.number, message_field=onlykey.client.MessageField.PASSWORD, value=BGPassword)
         
input("Press Enter to Exit Script")
