#region Import Modules
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
import keyring
import base64
from pynput.keyboard import Key, Controller
from ast import literal_eval
from base64 import b64decode, b64encode
from datetime import datetime, timedelta
from getpass import getuser
from os import getcwd, getenv
from time import sleep
from sys import exit

import lastpass
import PySimpleGUI as sg
from Crypto.Cipher import AES
from lastpass.fetcher import make_key
from onlykey.client import MessageField, OnlyKey
from requests import get, post

if not sg.running_mac():
    import keyboard as kb
    import win32crypt
    
icon = b'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SyFXBYwcC6hBg+z6wf/gd7dWcWrSSwolgd4Xx/kYBfp2gVbDcb6PHad1AgSfgSu94681gZlP0hsdLXYEhLeBi+uOpuwBlzvA8JMhm7IrBWkKxSLwfkbflAeGboGBNa+39j5OH4AsdZW+AQ4OgbESZa/7vLu/u7d/z7T7+wFm5HKidV6o5AAAD6BpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+Cjx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IlhNUCBDb3JlIDQuNC4wLUV4aXYyIj4KIDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+CiAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIKICAgIHhtbG5zOmlwdGNFeHQ9Imh0dHA6Ly9pcHRjLm9yZy9zdGQvSXB0YzR4bXBFeHQvMjAwOC0wMi0yOS8iCiAgICB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIKICAgIHhtbG5zOnN0RXZ0PSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VFdmVudCMiCiAgICB4bWxuczpwbHVzPSJodHRwOi8vbnMudXNlcGx1cy5vcmcvbGRmL3htcC8xLjAvIgogICAgeG1sbnM6R0lNUD0iaHR0cDovL3d3dy5naW1wLm9yZy94bXAvIgogICAgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIgogICAgeG1sbnM6dGlmZj0iaHR0cDovL25zLmFkb2JlLmNvbS90aWZmLzEuMC8iCiAgICB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iCiAgIHhtcE1NOkRvY3VtZW50SUQ9ImdpbXA6ZG9jaWQ6Z2ltcDoyOTZjN2MxZi0yMzc0LTQ1YWUtYmNhZC02MzM2OWRlNTE4ZGUiCiAgIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6ZmNmNmI1MjUtZWVhMi00MDhlLWE1OWMtZjNlMDViYzI0OTNjIgogICB4bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ9InhtcC5kaWQ6MmM4M2QzYjYtMTFiNS00OWExLWI3YTUtMDllZmUzOTk1MjYzIgogICBHSU1QOkFQST0iMi4wIgogICBHSU1QOlBsYXRmb3JtPSJNYWMgT1MiCiAgIEdJTVA6VGltZVN0YW1wPSIxNjQ2MDcxMTA0NjMxNTQ1IgogICBHSU1QOlZlcnNpb249IjIuMTAuMjIiCiAgIGRjOkZvcm1hdD0iaW1hZ2UvcG5nIgogICB0aWZmOk9yaWVudGF0aW9uPSIxIgogICB4bXA6Q3JlYXRvclRvb2w9IkdJTVAgMi4xMCI+CiAgIDxpcHRjRXh0OkxvY2F0aW9uQ3JlYXRlZD4KICAgIDxyZGY6QmFnLz4KICAgPC9pcHRjRXh0OkxvY2F0aW9uQ3JlYXRlZD4KICAgPGlwdGNFeHQ6TG9jYXRpb25TaG93bj4KICAgIDxyZGY6QmFnLz4KICAgPC9pcHRjRXh0OkxvY2F0aW9uU2hvd24+CiAgIDxpcHRjRXh0OkFydHdvcmtPck9iamVjdD4KICAgIDxyZGY6QmFnLz4KICAgPC9pcHRjRXh0OkFydHdvcmtPck9iamVjdD4KICAgPGlwdGNFeHQ6UmVnaXN0cnlJZD4KICAgIDxyZGY6QmFnLz4KICAgPC9pcHRjRXh0OlJlZ2lzdHJ5SWQ+CiAgIDx4bXBNTTpIaXN0b3J5PgogICAgPHJkZjpTZXE+CiAgICAgPHJkZjpsaQogICAgICBzdEV2dDphY3Rpb249InNhdmVkIgogICAgICBzdEV2dDpjaGFuZ2VkPSIvIgogICAgICBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOmVhNjEyMDg2LWE0OGMtNDNiNS1hMWRhLTA2YjQ2YTZiZTc5ZSIKICAgICAgc3RFdnQ6c29mdHdhcmVBZ2VudD0iR2ltcCAyLjEwIChNYWMgT1MpIgogICAgICBzdEV2dDp3aGVuPSIyMDIyLTAyLTI4VDExOjU4OjI0LTA2OjAwIi8+CiAgICA8L3JkZjpTZXE+CiAgIDwveG1wTU06SGlzdG9yeT4KICAgPHBsdXM6SW1hZ2VTdXBwbGllcj4KICAgIDxyZGY6U2VxLz4KICAgPC9wbHVzOkltYWdlU3VwcGxpZXI+CiAgIDxwbHVzOkltYWdlQ3JlYXRvcj4KICAgIDxyZGY6U2VxLz4KICAgPC9wbHVzOkltYWdlQ3JlYXRvcj4KICAgPHBsdXM6Q29weXJpZ2h0T3duZXI+CiAgICA8cmRmOlNlcS8+CiAgIDwvcGx1czpDb3B5cmlnaHRPd25lcj4KICAgPHBsdXM6TGljZW5zb3I+CiAgICA8cmRmOlNlcS8+CiAgIDwvcGx1czpMaWNlbnNvcj4KICA8L3JkZjpEZXNjcmlwdGlvbj4KIDwvcmRmOlJERj4KPC94OnhtcG1ldGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAKPD94cGFja2V0IGVuZD0idyI/PhhonKUAAABXUExURQAAAAQqY////wQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqYwQqY6mWtLcAAAABdFJOUwBA5thmAAAAAWJLR0QAiAUdSAAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB+YCHBE6GEfnZhAAAAOCSURBVHja7dzRtqsgDEVRkv//6PN8b7XakAQSVt4R9xzYqoBjUI+lZevo8B4IYSclH7WhQUbsq9rEYFH6QIOF+cVUKwWWhw9CyM8v07VAYKP0AQYF8zsTJOUX50oT2DO+L0HN/J4EwfklrMIF9o7vSVA1f6xAgfixAiXyuxHUzR8lUCa+G4EzgEg1gcr5/QWq5XcXKJffRaB0fl+BkgCyDYCUB6iZ31GgaH4HAQcAaQBQNr+XQDJA0sFyACLONlVgEiDqVHMBMgbAVof+XyA8/5ZHTwTY9/izAAmnF9/FBEDK/3RxgKSZhm0BsmYa1wDkzWGE9hQIoJomsAIgdVVHRQDVGgJWgOT1beUAVHMF+gOEdWgEWLDWF4CYLkMAFIA614ANYEX+h16PBxAAEgHW5I+5BgBwB1AAAKgkAIABYFl+ABQA9669ARQAAAAAAAAA6twIAAAAAIXuhAEAAAAAAAAAAAAAAAAAAABY9zjcqgAAAAAAAAAAAAAAeHUfuvqTIW+20UoUgNSpAACpVs4AUrAcAaRoeQGIdBZ4BJDK5QAgxWsWQKS7wHcAORxApL9AfwCxA4gcIPAFQOQEAQBuAUSOELgFEAAAOELgDkDkEAEArgEEAAAAAOAIAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4M2Rjwawjxv7gLO1DAEww2l+ywgA8wJDze8yAMC8z8K+QWNma4c7gHmrlUa0zAcwbza071Kc29/oDGDfbRvUUHMB7NuNzS21FIDmA2gmgP2DA+aW0x85SAZQ7xhbAdg/ORIp1wVAnbvcCkB7AGgsgLp2CcAyAM0HUAAAAAAA/gUA4E7w1GeBvR6G8p+jeR9w+huhGu8E414m7gXQ/q3wfvMCzAy1mBuM6nKz2WFxb7hkdvjukKEt7V2yQiQEYGapj7nhTmuE/jmyfeFWSsM4gCIFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwLXAMQADAAAAuBI4BWAAcCNwCMA4B0AB+AowzgQY45wh8ATQfQg85u8+BJ4Beg+BF/k/BNoDjAcAPW0AdBZ4l7/xRfAyf1uB1/mbAuh7gJYCv+TvKPBb/n4Cv+bvJvB7/g+CQ37/ewoY87e5DMz5exDMxG9AMBu/OIFH/LrPh+oV/27esFyNyTo7fWWFQb2pPxUzZvruAqieAAAAAElFTkSuQmCC'
#endregion

#region Functions
#region Function GetCyberArkToken
def GetCyberArkToken(ArkPass,ArkUser,ArkURL="https://cyberark.medhost.com/PasswordVault/API",Method = "RADIUS"):
    if ',' not in ArkPass:
        print("Retrieving CyberArk Token - Sending DUO Push")
    else:
        print("Retrieving CyberArk Token - Sending DUO Passcode")
        
    Body = {
        "username"          : ArkUser,
        "password"          : ArkPass,
        "concurrentSessions": "false"
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
        
    try:
        Response = (get(url=URL,headers=Header)).json()
    except Exception as error:
        return {'error':error}
        
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
def FindCyberArkAccount(Token,Username=None,SearchType='contains',limit='100',page='1',timeout=None,ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    offset = str(int(limit) * (int(page) - 1))
    search = '' if not Username or Username == 'Search for accounts' else ('search=' + Username.strip() + '&searchtype=' + SearchType + '&')
    
    URL  = ArkURL + "/ExtendedAccounts" + '?' + search + f'limit={limit}' + "&offset=" + offset
    try:
        Response = (get(url=URL,headers=Header,timeout=timeout)).json()
    except Exception as Error:
        return {'ErrorMessage':Error}
    
    return Response
#endregion

#region Function FindCyberArkPlatform
def FindCyberArkPlatform(Token,PlatformId=None,Search=None,SearchType='contains',limit='1000',page='1',timeout=None,ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    offset = str(int(limit) * (int(page) - 1))
    search = '' if not Search else ('search=' + Search + '&searchtype=' + SearchType + '&')
    platform = f'/{PlatformId}' if PlatformId else ''
    
    URL  = ArkURL + "/Platforms" + platform + '?' + search + f'limit={limit}' + "&offset=" + offset
    try:
        Response = (get(url=URL,headers=Header,timeout=timeout)).json()
    except Exception as Error:
        return {'ErrorMessage':Error}
    
    return Response
#endregion

#region Function FindCACheckedOut
def FindCACheckedOut(Token,Username,timeout=None,ArkURL = "https://cyberark.medhost.com/PasswordVault/API"):
    Header = {"Authorization" : Token}
    URL    = ArkURL + "/ExtendedAccounts" +"?limit=" + str(1000)
    Response = (get(url=URL,headers=Header,timeout=timeout)).json()
    Return = []
    Return += Response['Accounts']
    offset = 1000
    while offset < Response['Total']:
        URL    = ArkURL + "/ExtendedAccounts" +"?limit=" + str(1000) + "&offset=" + str(offset)
        Response = (get(url=URL,headers=Header,timeout=timeout)).json()
        Return += Response['Accounts']
        offset = offset + 1000
    
    Return = [_ for _ in Return if _['Properties']['LockedBy'] in [Username,Username.lower(),Username.lower().title()]]
    return Return
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
        platform = AcctP['PlatformName'] if 'PlatformName' in AcctP else ''
        Id       =  Acct['AccountID'   ] if 'AccountID'    in Acct  else ''

        row = [inuse,username,address,facility,client,system,tool,support,safe,name,lockedby,UsedBy,UsedDate,Id,platform]

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

#region Function EventHelper
def EventHelper(Input='',Sleep=0.0):
    sleep(Sleep)
    return Input
#endregion

#region on_activate
def on_activate():
    print('Global hotkey activated!')
    raise pynput._util.AbstractListener.StopException
#endregion

#region typestring
def typestring(string='',Delay='.05'):
    KB = Controller()
    for key in string:
        if key in '~!@#$%^&*()_}{|:"<>?QWERTYUIOPASDFGHJKLZXCVBNM':
            kb.press("shift")
            kb.press(key.lower())
            kb.release(key.lower())
            kb.release("shift")
        else:
            kb.send(key.lower())
            kb.release(key.lower())
        sleep(float(Delay)) 
#endregion

#region Function HotKeyListener
def HotKeyListener(HotKey):
    try:
        kb.wait(HotKey,True,False)
        return True
    except:
        return False
#endregion

#region Function getdate
def getdate(date):
    if isinstance(date,int): return datetime.fromtimestamp(date).strftime('%m-%d-%Y %I:%M %p')
    
    elif isinstance(datetime.now(),datetime): return date.strftime('%m-%d-%y %I:%M %p')
    
    else:
        return datetime.now()
#endregion

#region failed experiment
# #region Function SwapContainers
# def SwapContainers(window,C1,C2):
#     location  = window.CurrentLocation() if window.Shown else (None,None)
#     TabGroups = [_[0] for _ in window.key_dict.items() if _[1].Type == 'tabgroup']
#     values    = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0] not in TabGroups}
#     listboxes = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox' and _[1].Values]
#     tabs      = [{'Key':_[1].Key,'TabID':_[1].TabID} for _ in window.key_dict.items() if _[1].Type == 'tab'and (_[1].TabID or _[1].TabID == 0)]
#     tables    = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'table' and _[1].Values]
#     metadata  = window.metadata
#     icon      = window.metadata['icon']
    
#     C1Rows = window[C2].Rows
#     C2Rows = window[C1].Rows
#     window[C1].Rows = C1Rows
#     window[C2].Rows = C2Rows

#     for rI,row in enumerate(window.Rows):
#         for eI,elem in enumerate(row):
#                 window.Rows[rI][eI].Position = (rI,eI)
#                 window.Rows[rI][eI].ParentContainer = None
        
#     newwindow = sg.Window(Title, window.Rows,location=location,finalize=True,font=window.Font,icon=icon)
#     newwindow.force_focus()
#     newwindow.fill(values)
#     newwindow.metadata = metadata
#     for box in listboxes: newwindow[box['Key']].update(values=box['Values'])
#     for tab in tabs:      newwindow[tab['Key']].TabID = tab['TabID']
#     for table in tables: 
#         ScrollPosition = values[table['Key']][-1] / (len(table['Values']) - 1) if values[table['Key']] and values[table['Key']][-1] != 0 else 0
#         newwindow[table['Key']].update(values=table['Values'])
#         newwindow[table['Key']].update(select_rows=values[table['Key']])
#         newwindow[table['Key']].set_vscroll_position(ScrollPosition)
#     window.close()
#     newwindow.DisableDebugger()
#     #sleep(1)    
#     return newwindow
# #endregion

# #region Function ReloadWindow
# def ReloadWindow(window):
#     location  = window.CurrentLocation() if window.Shown else (None,None)
#     TabGroups = [_[0] for _ in window.key_dict.items() if _[1].Type == 'tabgroup']
#     values    = {_[0] : _[1] for _ in window.ReturnValuesDictionary.items() if isinstance(_[0], str) and _[0] not in TabGroups}
#     listboxes = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox' and _[1].Values]
#     tabs      = [{'Key':_[1].Key,'TabID':_[1].TabID} for _ in window.key_dict.items() if _[1].Type == 'tab'and (_[1].TabID or _[1].TabID == 0)]
#     tables    = [{'Key':_[1].Key,'Values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'table' and _[1].Values]
#     metadata  = window.metadata
#     icon      = window.metadata['icon']

#     for rI,row in enumerate(window.Rows):
#         for eI,elem in enumerate(row):
#                 window.Rows[rI][eI].Position = (rI,eI)
#                 window.Rows[rI][eI].ParentContainer = None
        
#     newwindow = sg.Window(Title, window.Rows,location=location,finalize=True,font=window.Font,icon=icon)
#     newwindow.force_focus()
#     window.close()
#     newwindow.fill(values)
#     newwindow.metadata = metadata
#     try:
#         for box in listboxes: newwindow[box['Key']].update(values=box['Values'])
#         for tab in tabs:      newwindow[tab['Key']].TabID = tab['TabID']
#         for table in tables: 
#             ScrollPosition = values[table['Key']][-1] / (len(table['Values']) - 1) if values[table['Key']] and values[table['Key']][-1] != 0 else 0
#             newwindow[table['Key']].update(values=table['Values'])
#             newwindow[table['Key']].update(select_rows=values[table['Key']])
#             newwindow[table['Key']].set_vscroll_position(ScrollPosition)          
#     except Exception as error:
#         print(error)
#     newwindow.DisableDebugger()
#     return newwindow
# # #endregion

# #region CopyLayout
# def CopyLayout(layout):
#     NewLayout = layout   
#     for rI,row in enumerate(NewLayout):
#         for eI,elem in enumerate(row):
#             NewLayout[rI][eI].Position = (rI,eI)
#             NewLayout[rI][eI].ParentContainer = None
#     return NewLayout
# #endregion

# #region MoveRow
# def MoveRow(window,layout,NewRow,OldRow):
#     Nr=NewRow;Or=OldRow;location=window.CurrentLocation()
#     layout.insert(Nr,layout.pop(Or))    
#     for rI,row in enumerate(layout):
#         for eI,elem in enumerate(row):
#             layout[rI][eI].Position = (rI,eI)
#             layout[rI][eI].ParentContainer = None
#     values = window.ReturnValuesDictionary
#     listboxes = [{'key':_[1].Key,'values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox']  
#     newwindow = sg.Window(Title, layout,finalize=True,location=location)
#     newwindow.fill(values)
#     for box in listboxes: newwindow[box['key']].update(values=box['values'])
#     window.close()
#     return newwindow
# #endregion

# #region MoveRows
# def MoveRows(window,layout,Rows:list[tuple]):
#     '''Rows = list of tuples [(NewRow,OldRow),(NewRow,OldRow)]'''
#     location=window.CurrentLocation()
#     oldnums = [_[1] for _ in Rows]
#     newnums = [_[0] for _ in Rows]
#     srows=[row for rI,row in enumerate(layout) if rI not in oldnums]
#     newlayout = []
#     i = 0
#     while i < len(layout):
#         if i not in newnums:
#             newlayout.append(srows.pop(0))
#         else:
#             mrow = [_ for _ in Rows if _[0] == i][0]
#             newlayout.append(layout[mrow[1]])
#         for eI,elem in enumerate(newlayout[i]):
#             newlayout[i][eI].Position = (i,eI)
#             newlayout[i][eI].ParentContainer = None
#         i = i + 1
#     listboxes = [{'key':_[1].Key,'values':_[1].Values} for _ in window.key_dict.items() if _[1].Type == 'listbox']  
#     values = window.ReturnValuesDictionary
#     newwindow = sg.Window(Title, newlayout,finalize=True,location=location)
#     newwindow.fill(values)
#     for box in listboxes: newwindow[box['key']].update(values=box['values'])
#     window.close()
#     return (newwindow,newlayout)
# #endregion
#endregion

#region Function pin_popup
def pin_popup(message, title=None, default_text='', password_char='', offer_reset=False, size=(None, None), button_color=None,
                   background_color=None, text_color=None, icon=None, font=None, no_titlebar=False,pass_through=None,
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
               [sg.InputText(default_text=default_text, size=size, key='PIN_INPUT_', password_char=password_char)],
               [sg.Button('Ok', size=(6, 1), bind_return_key=True), sg.Button('Cancel', size=(6, 1))]]
    
    if offer_reset: 
        layout[2].append(sg.Push())
        layout[2].append(sg.Button('Reset Pin'))

    window = sg.Window(title=title or message, layout=layout, icon=icon, auto_size_text=True, button_color=button_color, no_titlebar=no_titlebar,
                    background_color=background_color, grab_anywhere=grab_anywhere, keep_on_top=keep_on_top, location=location, relative_location=relative_location, finalize=True, modal=modal)
    if pass_through:
        return window
    else:
        button, values = window.read()
        window.close()
        del window
        if button == 'Cancel':
            return None
        elif button == 'Ok':
            path = values['PIN_INPUT_']
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
    try:
        Return = literal_eval(str(string))
    except:
        Return = []
    return Return
#endregion

#region Function SaveToConfig
def SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,Print=True):
    config = configparser.ConfigParser()
    user = getuser()
    salt        = (subprocess.run("hostname",capture_output=True,text=True,shell=True)).stdout.rstrip()
    iterations  = int(pin) if len(str(pin)) < 8 else int(str(pin)[:7])
    keyringcred = keyring.get_password('MEDHOST Password Manager', 'Key')
    Pass        = keyringcred + getuser() + str(pin)
    key         = make_key(Pass,salt,iterations)

    config.add_section('General')
    config['General']['Sync Options Visible']         = str(window["SyncOpt"]._visible)
    config['General']['New Install']                  = 'False'
    config['General']['Type For Me Delay']            =     Inputs['LVtfmDelay']
    config['General']['HotKey Delay']                 =     Inputs['LVhkDelay']
    config['General']['KeyPress Delay']               =     Inputs['LVtDelay']
    config['General']['HotKey']                       =     Inputs['LVHotKey']
    config['General']['Auto Start Listener']          = str(Inputs['LVAutoStart'])
    config.add_section('CyberArk')
    config['CyberArk']['Network Username']            =     Inputs["ArkUser"]
    config['CyberArk']['Accounts']                    = str(window['CATable'].Values).replace('✅','').replace('⛔','')
    config['CyberArk']['Selected Account Index']      = str(Inputs['CATable'])
    config['CyberArk']['MFA Push Selected']           = str(Inputs["CAPush"])
    config['CyberArk']['Sync From CyberArk Selected'] = str(Inputs["CAFrom"])
    config['CyberArk']['Login Visible']               = str(window['ArkLoginInfo']._visible)
    config['CyberArk']['Page Size']                   =     Inputs['CAPageSize']
    config['CyberArk']['CheckOut AutoUpdate']         = str(Inputs['CACOAutoUpdate'])
    config['CyberArk']['CheckOut Update TimeOut']     =     Inputs['CACOAutoTimeOut']
    config['CyberArk']['CheckOut Update Interval']    =     Inputs['CACOAutoInterval']
    config['CyberArk']['Checked Out Accounts']        = str(window['CACOTable'].Values).replace('✅','').replace('⛔','') if window['CACOTable'].Values and window['CACOTable'].Values[0] else '[]'
    config['CyberArk']['Auto Check-In Enabled']       = str(Inputs['CACOAutoCheckIn'])
    config['CyberArk']['Auto Check-In Per Account']   = str(Inputs['CACOPerAccount'])
    config['CyberArk']['Check-In Global Delay']       =     Inputs['CACODelay']
    config['CyberArk']['Per Account Delay Values']    = str(window['CADelay'].Values)
    config['CyberArk']['AutoDelete on Check-In']      = str(Inputs['CACOAutoDelete'])
    config.add_section('LastPass')     
    config['LastPass']['Username']                    =     Inputs["LPUser"]
    config['LastPass']['Sync From LastPass Selected'] = str(Inputs['LPFrom'])
    config['LastPass']['Sync To LastPass Selected']   = str(Inputs["LP"])
    config['LastPass']['LastPass Accounts']           = str(window['LPSelected'].Values)
    config['LastPass']['Login Visible']               = str(window['LPLoginInfo']._visible)
    config['LastPass']['MFA Push Selected']           = str(Inputs['LPPush'])
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

#region Function Random Password
def RandomPassword(length):
    import random
    import array
    
    # maximum length of password needed
    # this can be changed to suit your password length
    MAX_LEN = length
    
    # declare arrays of the character that we need in out password
    # Represented as chars to enable easy string concatenation
    DIGITS = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] 
    LOCASE_CHARACTERS = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h',
                        'i', 'j', 'k', 'm', 'n', 'o', 'p', 'q',
                        'r', 's', 't', 'u', 'v', 'w', 'x', 'y',
                        'z']
    
    UPCASE_CHARACTERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H',
                        'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q',
                        'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',
                        'Z']
    
    SYMBOLS = ['@', '#', '$', '%', '=', ':', '?', '.', '/', '|', '~', '>',
            '*', '(', ')', '<']
    
    # combines all the character arrays above to form one array
    COMBINED_LIST = DIGITS + UPCASE_CHARACTERS + LOCASE_CHARACTERS + SYMBOLS
    
    # randomly select at least one character from each character set above
    rand_digit = random.choice(DIGITS)
    rand_upper = random.choice(UPCASE_CHARACTERS)
    rand_lower = random.choice(LOCASE_CHARACTERS)
    rand_symbol = random.choice(SYMBOLS)
    
    # combine the character randomly selected above
    # at this stage, the password contains only 4 characters but
    # we want a 12-character password
    temp_pass = rand_digit + rand_upper + rand_lower + rand_symbol
    
    
    # now that we are sure we have at least one character from each
    # set of characters, we fill the rest of
    # the password length by selecting randomly from the combined
    # list of character above.
    for x in range(MAX_LEN - 4):
        temp_pass = temp_pass + random.choice(COMBINED_LIST)
    
        # convert temporary password into array and shuffle to
        # prevent it from having a consistent pattern
        # where the beginning of the password is predictable
        temp_pass_list = array.array('u', temp_pass)
        random.shuffle(temp_pass_list)
    
    # traverse the temporary password array and append the chars
    # to form the password
    password = ""
    for x in temp_pass_list:
            password = password + x
    return password
#endregion

#region Function GetConfigAppData
def GetConfigAppData(WorkingDirectory,pin,Print=True,FB=False):
        try:
            config = configparser.ConfigParser()
            config.read(WorkingDirectory)
            salt        = (subprocess.run("hostname",capture_output=True,text=True,shell=True)).stdout.rstrip()
            iterations  = int(pin) if len(str(pin)) < 8 else int(str(pin)[:7])
            keyringcred = keyring.get_password('MEDHOST Password Manager', 'Key') if not FB else RandomPassword(56)
            Pass        =  keyringcred + getuser() + str(pin)
            key         = make_key(Pass,salt,iterations)
            fb          = EncryptString('',key)

            AppData = {}

            ArkPass     = _ if bool(_:=config.get('App Data','CyUsPa',fallback=fb)) and not FB else fb            
            LPPass      = _ if bool(_:=config.get('App Data','LaUsPa',fallback=fb)) and not FB else fb     
            ArkToken    = _ if bool(_:=config.get('App Data','CySeTo',fallback=fb)) and not FB else fb     
            LPToken     = _ if bool(_:=config.get('App Data','LaSeTo',fallback=fb)) and not FB else fb     
            LPKey       = _ if bool(_:=config.get('App Data','LaSeKe',fallback=fb)) and not FB else fb     
            LPsId       = _ if bool(_:=config.get('App Data','LaSeId',fallback=fb)) and not FB else fb      
            LPIteration = _ if bool(_:=config.get('App Data','LPSeIt',fallback=fb)) and not FB else fb                

        
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
            config = configparser.ConfigParser()
            keyring.set_password('MEDHOST Password Manager', 'Key', keyringcred)
            config.add_section('App Data')  
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
        if ',' not in ArkPass:
            print(f"DUO Push Accepted - {ArkUser} Authenticated")
        else:
            print(f"DUO Passcode Accepted - {ArkUser} Authenticated")

    return ArkToken
#endregion

#region LastPassLogin
def LastPassLogin(Inputs,silent=False) -> lastpass.Session:
    if ',' in Inputs['LPPass']:
        Pass    = Inputs['LPPass'].split(',')[0]
        MFAPass = Inputs['LPPass'].split(',')[1]
    else:
        Pass    = Inputs['LPPass']
        MFAPass = None  
    try:
        LastPass = lastpass.Session.login(username=Inputs['LPUser'],password=Pass,multifactor_password=MFAPass).OpenVault(include_shared=True)
    except Exception as e:
        LastPass = lastpass.Session('','','','')
        if not silent: print(e)
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

    
WorkingDirectory = getenv("HOME") + '/Config.ini' if sg.running_mac() else getenv('LOCALAPPDATA') + '\\MEDHOST\\MEDHOST Password Manager\\Config.ini'
PathCheck = os.path.exists(WorkingDirectory) if not sg.running_mac() else True
if not PathCheck:
    try: 
        os.mkdir(getenv('LOCALAPPDATA') + '\\MEDHOST')
    except:
        pass
    try:
        os.mkdir(getenv('LOCALAPPDATA') + '\\MEDHOST\\MEDHOST Password Manager')
    except:
        pass

config.read(WorkingDirectory)

ConfigSyncOpt  =  Bool(config.get('General' ,'Sync Options Visible'        ,fallback='True'))
ConfigNew      =  Bool(config.get('General' ,'New Install'                 ,fallback='True'))
ConfigTFMDelay =       config.get('General' ,'Type For Me Delay'           ,fallback='2')
ConfigHKDelay  =       config.get('General' ,'HotKey Delay'                ,fallback='0.1')
ConfigTDelay   =       config.get('General' ,'KeyPress Delay'              ,fallback='0.05')
ConfigHotKey   =       config.get('General' ,'HotKey'                      ,fallback='insert+v')
ConfigHKAuto   =  Bool(config.get('General' ,'Auto Start Listener'         ,fallback='True'))
       
ConfigUsername =       config.get('CyberArk','Network Username'            ,fallback=getuser())
ConfigCATable  = array(config.get('CyberArk','Accounts'                    ,fallback="[]"))
ConfigCAIndex  = array(config.get('CyberArk','Selected Account Index'      ,fallback="[]"))
ConfigCAPush   =  Bool(config.get('CyberArk','MFA Push Selected'           ,fallback='True'))
ConfigCAFrom   =  Bool(config.get('CyberArk','Sync From CyberArk Selected' ,fallback='True'))
ConfigCALogin  =  Bool(config.get('CyberArk','Login Visible'               ,fallback='True'))
ConfigCAPSize  =       config.get('CyberArk','Page Size'                   ,fallback='100')
ConfigCACOAuto =  Bool(config.get('CyberArk','CheckOut AutoUpdate'         ,fallback='True'))
ConfigCACOTime =       config.get('CyberArk','CheckOut Update TimeOut'     ,fallback='8')
ConfigCACOInt  =       config.get('CyberArk','CheckOut Update Interval'    ,fallback='15')
ConfigCACOActs = array(config.get('CyberArk','Checked Out Accounts'        ,fallback="[]"))
ConfigCACOaChI =  Bool(config.get('CyberArk','Auto Check-In Enabled'       ,fallback='False'))
ConfigCAPerAct =  Bool(config.get('CyberArk','Auto Check-In Per Account'   ,fallback='True'))
ConfigCAGDelay =       config.get('CyberArk','Check-In Global Delay'       ,fallback='2:00 Hrs')
ConfigCAPDelay = array(config.get('CyberArk','Per Account Delay Values'    ,fallback="['Disabled','0:30 Hrs','1:00 Hrs','1:30 Hrs','2:00 Hrs']"))
ConfigCADelete =  Bool(config.get('CyberArk','AutoDelete on Check-In'      ,fallback='True'))
       
ConfigLPUser   =       config.get('LastPass','Username'                    ,fallback="")
ConfigLP       =  Bool(config.get('LastPass','Sync To LastPass Selected'   ,fallback="True"))
ConfigLPFrom   =  Bool(config.get('LastPass','Sync From LastPass Selected' ,fallback="False"))
ConfigLPAccts  = array(config.get('LastPass','LastPass Accounts'           ,fallback="[]"))
ConfigLPLogin  =  Bool(config.get('LastPass','Login Visible'               ,fallback='True'))
ConfigLPPush   =  Bool(config.get('CyberArk','MFA Push Selected'           ,fallback='True')) 
       
ConfigOK       =  Bool(config.get('OnlyKey' ,'OnlyKey Selected'            ,fallback="False"))
ConfigOKWord   =       config.get('OnlyKey' ,'Keyword'                     ,fallback="BG*")
ConfigKWTrue   =  Bool(config.get('OnlyKey' ,'Keyword Search Selected'     ,fallback="False"))
ConfigSlots    = {Slot : Bool(config.get('Selected Slots',Slot,fallback="False")) for Slot in Slots}
ConfigMap      = {Slot : config.get('Acct-Slot Mappings',Slot,fallback="") for Slot in Slots}

if ConfigUsername == "": ConfigUsername = getuser()

if ConfigCACOActs and ConfigCACOActs[0]:
    ConfigCACOActs = [_ for _ in ConfigCACOActs if len(_) >= 13 and str(_[12]).isnumeric() and int(_[12]) > int((datetime.now() - timedelta(hours=4)).timestamp())]

#AppData Pin 
test = False
if not test:
    if not ConfigNew:
        pin = pin_popup('Enter Your Pin',password_char="*",title='MEDHOST Password Manager',offer_reset=True,icon=icon)
        if not pin:
            exit()
    else:
        pin = 'reset'

    if pin == 'reset':    
        pin = pin_popup('Enter Your New Pin',password_char="*",title='MEDHOST Password Manager',icon=icon)
        while not pin or not pin.isnumeric() or len(pin) < 6:
            pin = pin_popup('Your pin must be numeric and contain 6 or more characters',password_char="*",title='MEDHOST Password Manager',icon=icon)
            if not pin:
                exit()
                
        AppData = GetConfigAppData(WorkingDirectory,pin,FB=True)
        if AppData == 'error':
            exit()
    else:
        AppData = GetConfigAppData(WorkingDirectory,pin)
        while AppData == 'error':
            pin = pin_popup('Try Reentering your pin',password_char="*",title='MEDHOST Password Manager',offer_reset=True,icon=icon)
            if not pin:
                exit()
            elif pin == 'reset':    
                pin = pin_popup('Enter Your New Pin',password_char="*",title='MEDHOST Password Manager')
                while not pin or not pin.isnumeric() or len(pin) < 6:
                    pin = pin_popup('Your pin must be numeric and contain 6 or more characters',password_char="*",title='MEDHOST Password Manager',icon=icon)
                    if not pin:
                        exit()         
                AppData = GetConfigAppData(WorkingDirectory,pin,FB=True)
                if AppData == 'error':
                    exit()
            else:
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
    LastPass = lastpass.Session(AppData['LPsId'],AppData['LPIteration'],AppData['LPToken'],AppData['LPKey'])
else:
    LastPass = lastpass.Session('','','','')

if LastPass.login_check():
    try:
        LastPass.OpenVault(True)
    except:
        pass    
    
if LastPass.loggedin and LastPass.accounts:
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
else:
    LPList   =  ['Login to see LastPass entries here']
    LVTable  = [['Login to see LastPass entries here']]
    LPTable  = [['Login to see LastPass entries here']]
    LVGroups = ['Group/Folder']

ArkToken   = AppData['ArkToken']
if ArkToken:
    try:
        ArkAccts = FindCyberArkAccount(Token=ArkToken,timeout=20)
    except:
        ArkAccts = None
else:
    ArkAccts = None
    
if not ArkAccts or 'ErrorMessage' in ArkAccts:
    CATable   = ConfigCATable if ConfigCATable else [[],['','','Login to see'],['','','CyberArk'],['','','Accounts Here']]
    CACOTable = ConfigCACOActs if ConfigCACOActs else [[],['','','Login to see'],['','','CyberArk'],['','','Accounts Here']]
else:
    CATable   = ConfigCATable if ConfigCATable else MakeCyberArkTable(ArkAccts)
    CACOTable = ConfigCACOActs if ConfigCACOActs else []
    
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
CAsf    = dict(font=("Helvetica",9+Fm))
Ls    =  55   if sg.running_mac() else 45
Llip  =  0    if sg.running_mac() else 5
Llip2 =  2    if sg.running_mac() else 0
Llip3 =  4    if sg.running_mac() else 8
Lbp   = (4,3) if sg.running_mac() else (0,0)
lpfb  = ("Helvetica",12) if sg.running_mac() else ("Helvetica",9) 
LPColor  = dict(disabled_readonly_background_color=sg.DEFAULT_BACKGROUND_COLOR, disabled_readonly_text_color='white')

LastPassOptions = [
    [sg.Text("⩠ Login",justification="Left",size=(5, 1),font=("Helvetica",10+Fm),**ee,visible=True,k='LPLoginToggle'),sg.Text("LastPass",justification="center",size=(35+Sm*2, 1),font=(f"Helvetica {12+Fm} bold")),sg.VPush()],
    [sg.pin(sg.Column([
        [sg.Text('Please enter your LastPass Login Information')],
        [sg.Text('LastPass Username', size=(16+Sm, 1)), sg.Input(default_text=ConfigLPUser     ,key="LPUser",s=(45,1),enable_events=True)],
        [sg.Text('LastPass Password', size=(16+Sm, 1)), sg.Input(default_text=AppData['LPPass'],key="LPPass",s=(45,1),enable_events=True,password_char="*")],
        [sg.Button(button_text='Refresh Account List',key="LPRefresh"),sg.Push(),sg.Text('MFA:',p=(0,3)),sg.Radio('Push','LPMFAType',k='LPPush',**ee,default=ConfigLPPush),sg.Radio('Passcode','LPMFAType',k='LPPasscode',**ee,default=not ConfigLPPush),sg.Push(),sg.Button(button_text='Login',key="LPLogin")],
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
    [sg.Push(),sg.Button(button_text='Clear All',key="LPClear",p=(5,Lbp),font=lpfb)]
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
Pcl = 0 if sg.running_mac() else 3
Tm  = 1 if sg.running_mac() else 0
CAeReason = 'Enter business justification for checking out account'
CAc = dict(disabled_readonly_background_color=sg.DEFAULT_BACKGROUND_COLOR, disabled_readonly_text_color='white')
CAOptions = [
    [sg.Text("⩠ Login",justification="Left",size=(5, 1),font=("Helvetica",10+Fm),enable_events=True,visible=True,k='CALoginToggle'),sg.Text("CyberArk",justification="Left",size=(21+Sm*2, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.pin(sg.Column([
        [sg.Text('Enter your CyberArk login information:'),sg.Push(),sg.Text('MFA:',p=(0,3)),sg.Radio('Push','CAMFAType',k='CAPush',**ee,default=ConfigCAPush),sg.Radio('Passcode','CAMFAType',k='CAPasscode',**ee,default=not ConfigCAPush)],
        [sg.Text('Username:', size=(8, 1),p=((5,0),3),tooltip=UserNameToolTip),sg.Input(default_text=ConfigUsername    ,k="ArkUser",tooltip=UserNameToolTip,s=(17+Sm,1),p=((5,5),3)),
        sg.Text('Password:' , size=(8, 1),p=((5,0),3),tooltip=PasswordToolTip),sg.Input(default_text=AppData['ArkPass'],k="ArkPass",tooltip=PasswordToolTip,s=(17+Sm,1),password_char="*"),
        sg.Push(),sg.Button(button_text='Login',key="CALogin")]
    ],p=(0,(0,Pcl)),k="ArkLoginInfo",visible=False))],
    [sg.Text('Select an Account to check-out and sync:', size=(32,1),p=((5,0),0)),sg.Push(),sg.Text('Page Size:',p=(5,0),**CAsf),sg.Combo(['25','50','100','500','1000'],ConfigCAPSize,s=(4, 1),p=(5,0),k='CAPageSize',**CAsf)], #,sg.Button('Refresh',k="CARefresh",font=('Helvetica',10))],
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

#region CyberArk Checked Out
COtp     = 8 if sg.running_mac() else 8
PCACOSet = 13 if sg.running_mac() else 17
CACheckedOut = [
    [sg.Text("⩠ Settings",justification="Left",size=(8, 1),font=("Helvetica",10+Fm),**ee,visible=True,k='CASettingsToggle'),sg.Text("Checked Out Accounts",justification="Left",size=(24+Sm*2, 1),p=(5,(3,COtp)),font=(f"Helvetica {12+Fm} bold"))],
    [sg.pin(sg.Column([
        [sg.Toggle('Auto Check-In:',ConfigCACOaChI,k='CACOAutoCheckIn',**ee),sg.Radio('Per Account','CheckInR',k='CACOPerAccount',enable_events=True,default=ConfigCAPerAct),
            sg.Radio('Global:','CheckInR',k='CACOGlobal',p=((5,0),3),enable_events=True,default=not ConfigCAPerAct),sg.Combo(['0:30 Hrs','1:00 Hrs','1:30 Hrs','2:00 Hrs'],ConfigCAGDelay,p=((0,5),3),s=(8, 1),k='CACODelay',**ee,**CAsf)],
        [sg.Toggle('Delete Vault Entry On Check-In',ConfigCADelete,k='CACOAutoDelete',**ee)], #[sg.Checkbox('Delete Vault Entry On Check-In',k='CACOAutoDelete',p=(5,1),default=ConfigCADelete,enable_events=True)],
        [sg.Toggle('Auto Refresh:',ConfigCACOAuto,k='CACOAutoUpdate',**ee),sg.Text('Refresh Interval (Min):'),sg.Input(s=(3,1),p=((0,5),3),default_text=ConfigCACOInt,k="CACOAutoInterval",justification='center',**AcctColor),
            sg.Text('Refresh Timeout (Hrs):'),sg.Input(s=(3,1),p=(0,3),default_text=ConfigCACOTime,k="CACOAutoTimeOut",justification='center',**AcctColor),sg.Push()]
    ],p=(0,(0,PCACOSet)),k="CACOSettings",visible=False))],
    [sg.Table(CACOTable,['','Username','Address','FacilityName'],auto_size_columns=False,justification='center',**ee,visible_column_map=[True,True,True,True] ,num_rows=12+Tm, 
        k='CACOTable',col_widths=[3,14+Sm,12+Sm,20+Sm+Tm],p=(5,(0,0)),**InputColor,tooltip=BGUserToolTip,select_mode=sg.TABLE_SELECT_MODE_BROWSE)
    ],
    [sg.Button('Refresh',k="CACORefresh"),sg.Push(),sg.Input(s=(13+Sm,1),k="CACOStatus",readonly=True,justification='center',**DAcctColor),sg.Input(s=(16,1),k="CACOLastUpdated",readonly=True,justification='center',**DAcctColor),sg.Push(),
        sg.Button(button_text='View in Vault',disabled=True,key="CACOView"),sg.pin(sg.Button(button_text='Check In',key="CheckIn",disabled=True))],
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
    [sg.pin(sg.Col(OkCheckText,p=((0,Osp),0),visible=not ComboTrue,k='OkCheckText')),sg.Fr('OnlyKey Slots',k='OkFrame',
        layout=[
            [sg.pin(sg.Col(OkCheckLeft,p=0,k='OkCheckLeft')),sg.pin(sg.Col(OKComboLeft,p=0,visible=ComboTrue,k='OKComboLeft')),sg.pin(sg.VSep(pad=(0,5))),
            sg.pin(sg.Col(OkCheckRight,p=0,k='OkCheckRight')),sg.pin(sg.Col(OKComboRight,p=0,visible=ComboTrue,k='OKComboRight'))
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
LVmp = 11 if sg.running_mac() else 11
LVAcctDetails = [
    [sg.Text('Name'    , size=(7, 1)), sg.Input(s=(40,1),k="LVName" ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Group'   , size=(7, 1)), sg.Input(s=(40,1),k="LVGroup",readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Username', size=(7, 1)), sg.Input(s=(40,1),k="LVUser" ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Password', size=(7, 1)), sg.Input(s=(40,1),k="LVPass" ,readonly=True,**AcctColor,**DAcctColor,**AcctFont,password_char="*"),sg.Button(button_text='Show',p=((5,4),3),k="LVShow"),sg.Button(button_text='Copy',p=((5,4),3),k="LVCopy")],
    [sg.Text('URL'     , size=(7, 1)), sg.Input(s=(55,1),k="LVUrl"  ,readonly=True,**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Text('Notes',p=(5,(3,0)), size=(5, 1))],
    [sg.Multiline(s=(63,7),p=((3,3),(1,LVmp)),k="LVNote",write_only=True,disabled=True,**AcctColor,**AcctFont)]
]
#endregion

#region Local Vault Settings
LVSettings = [
    [sg.Text('Pressing the specified HotKey or HotKey Combo simulates keypresses')],
    [sg.Text('*HotKey support currently only works on Windows')],
    [sg.Text('HotKey Delay (Seconds):', size=(18, 1)), sg.Input(ConfigHKDelay ,s=(3,1),k="LVhkDelay" ,**AcctColor,**AcctFont),sg.Text('  Delay Between Keypresses:', size=(22, 1)), sg.Input(ConfigTDelay ,s=(4,1),k="LVtDelay" ,**AcctColor,**AcctFont)],
    [sg.Text('HotKey' , size=(7, 1)) , sg.Input(ConfigHotKey ,s=(40,1),k="LVHotKey"  ,**AcctColor,**AcctFont)],
    [sg.Button(button_text='Restart HotKey Listener',k="LVRestartL"),sg.Button(button_text='Stop HotKey Listener',k="LVStop"),sg.Checkbox('Auto start listener',key='LVAutoStart',default=ConfigHKAuto)],
    [sg.Text('Test Typing by Pressing Type for me and selecting the text box below')],
    [sg.Text('Type for Me Delay (Seconds)', size=(22, 1)), sg.Input(ConfigTFMDelay,s=(3,1),k="LVtfmDelay",**AcctColor,**DAcctColor,**AcctFont)],
    [sg.Button(button_text='Type for Me',p=((5,4),3),k="LVPaste")],
    [sg.Multiline(s=(63,3),p=((3,3),(1,LVmp)),k="LVNote")]
]
#endregion

#region Local vault
Vtp = 7  if sg.running_mac() else 5
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
        [sg.Tab('Account Details',LVAcctDetails,k='LVAcctDetails'),sg.Tab('HotKey Settings',LVSettings,k='LVSettings'),sg.Tab('Output',OutPutSection2,k='OutPutSection2')]
    ],p=0,enable_events=True,k='Tabs')]
]
#endregion

#region LVOptions
NumRows = 12 if sg.running_mac() else 11  #17 if sg.running_mac() else 16
MultiSelKey = 'Command'  if sg.running_mac() else 'Ctrl'
BPad = 1 if sg.running_mac() else 3

LVOptions = [
    [sg.Text("⩠ Login",justification="Left",size=(5, 1),font=("Helvetica",10+Fm),enable_events=True,visible=True,k='LVLoginToggle2'),sg.Text("LastPass",justification="Center",size=(21+Sm*2, 1),font=(f"Helvetica {12+Fm} bold"))],
    [sg.pin(sg.Column([
        [sg.Text('Please enter your LastPass Login Information',p=(5,(Vip,2)))],
        [sg.Text('LastPass Username', size=(16+Sm, 1)), sg.InputText(default_text=ConfigLPUser,key="LVlUser2",s=(45+Im,1))],
        [sg.Text('LastPass Password', size=(16+Sm, 1)), sg.InputText(password_char="*",key="LVlPass2",default_text=AppData['LPPass'],s=(45+Im,1))],
        [sg.Button(button_text='Refresh Account List',key="LVRefresh2"),sg.Push(),sg.Button(button_text='Login',key="LVLogin2")],
    ],p=(0,(0,Vcp)),k="LVLoginInfo2",visible=False))],
    [sg.Text(f'Select Vault Entry(s) to sync to your OnlyKey (Hold {MultiSelKey}/Shift to MultiSelect):',p=(5,1))],
    [sg.Table(LVTable,['Name','Username','Url','Id','Group','Notes','Password',],auto_size_columns=False,justification='left'  ,visible_column_map=[True,True,False,False,False,False,False] 
        ,num_rows=NumRows, k='LVTable2',col_widths=[24+Tcs, 25+Tcs],enable_events=True,select_mode=sg.TABLE_SELECT_MODE_EXTENDED,**InputColor)],
    [sg.Combo(['Type','Password','SecureNote'],'Type',s=(10, 1),p=(5,(0,BPad)),**ee,k='LVsType2'),sg.Input('Search for accounts',key="LVQuery2",p=(5,(0,BPad)),s=(32,1),click_events=True,**ee),
     sg.Combo(LVGroups,'Group/Folder',s=(15, 1),p=(5,(0,BPad)),k='LVsGroup2',**ee)]
]
#endregion
#endregion

#region Main Application Window Layout
Mfp = 3 if sg.running_mac() else 3
LPFrom = False #ConfigLPFrom
OKTrue = ConfigOK
LPTrue = ConfigLP
CATrue = True #ConfigCAFrom
submit = 'Check-Out & Sync' if CATrue else "Sync Passwords"
#Layout
layout = [
    [sg.pin(sg.Column(CAOptions,p=0,visible=True ,k='CAOptions'))],
    [sg.pin(sg.Column(LVOptions,p=0,visible=False,k='LVOptions'))],
    [sg.pin(sg.Column(GCOptions,p=0,visible=False,k='GCOptions'))],
    [sg.pin(sg.Column(MEOptions,p=0,visible=False,k='MEOptions'))],
    [sg.pin(sg.Fr('Sync Options', 
        layout=[
        [sg.Text('Sync From:',p=(5,(Mfp,3)),size=(8+Sm, 1)),sg.Radio('CyberArk','PFrom',k='CAFrom',enable_events=True,p=(5,(Mfp,3)),default=CATrue),sg.Radio('LastPass','PFrom',k='LPFrom',enable_events=True,p=(5,(Mfp,3)),default=LPFrom)
        ,sg.Radio('Chrome','PFrom',k='GCFrom',enable_events=True,p=(5,(Mfp,3)),default=False),sg.Radio('Edge','PFrom',k='MEFrom',enable_events=True,p=(5,(Mfp,3)),s=(9+Sm*4,1),default=False),],
        [sg.Text('Sync To:', size=(8+Sm, 1)),sg.pin(sg.Checkbox('LastPass',k='LP',enable_events=True,default=LPTrue,disabled_color='dark grey',disabled=not OKTrue,visible=not LPFrom)),
        sg.pin(sg.Checkbox('OnlyKey',k='OK',enable_events=True,default=OKTrue,disabled_color='dark grey',disabled=not LPTrue or LPFrom))],
    ],p=((5,0),(0,1)),k="SyncOpt"))],
    [sg.Button(button_text='Save',key="Save",tooltip=SaveToolTip),sg.pin(sg.Button(button_text='Sync Options',key="SyncOptToggle",mouseover_colors=('white',sg.DEFAULT_BACKGROUND_COLOR),button_color=('white',sg.DEFAULT_BACKGROUND_COLOR))),sg.Push(),sg.pin(sg.Text('Auto Check-In:',p=(2,3),k='CAAutoCheckIn')),
        sg.pin(sg.Combo(ConfigCAPDelay,'Disabled',p=(0,3),s=(8, 1),k='CADelay',**CAsf)),sg.pin(sg.Submit(submit,tooltip=SubmitToolTip,disabled=True,p=((5,10),3),k='Submit'))],
    [sg.HSep()], 
    [sg.TabGroup([
        [sg.Tab('Output',OutPutSection,k='OutPutSection'),sg.Tab('Account Details',CATabs,k='CAAcctDetails'),sg.Tab('Checked Out Accounts',CACheckedOut,k='CACheckedOut'),sg.Tab('LastPass',LastPassOptions,k='LPOptions',visible=LPTrue and not LPFrom), 
        sg.Tab('OnlyKey', OkSlotSelection,k='OKOptions',visible=OKTrue)]
    ],p=0,enable_events=True,k='Tabs')]
]
#sg.Text('⩠ Sync Options',k='SyncOptToggle',**ee)

MainTabs = [
    [sg.TabGroup([
        [sg.Tab('Check-Out & Sync',layout,k='SyncPasswords'),sg.Tab('Local Vault',VaultLayout,k='LocalVault',visible=True)]
    ],p=0,enable_events=True,k='MainTabs')]
]

#alternate up/down arrows ⩢⩠▲▼⩓⩔  locks 🔓🔒❎☒☑⚿⛔🗝️🔑🔐⛓️✔️🚫✅
#Window and Titlebar Options
Title = 'MEDHOST Password Manager beta'

closeattempted = False if sg.running_mac() else True

scaling=1.5

size = (501, 694) if sg.running_mac() else (497, 750)
window = sg.Window(Title, MainTabs,font=("Helvetica", 10+Fm),finalize=True,return_keyboard_events=True,alpha_channel=0,enable_close_attempted_event=closeattempted,icon=icon,size=size)
#endregion

#region Opening Application Tasks
window.DisableDebugger()
timeout        = 100
WindowRead     = window.read(timeout=0)
event          = WindowRead[0]
Inputs         = WindowRead[1]
OKVisible      = OKTrue
CAVisible      = CATrue
ArkAccountPass = None
pinwindow      = None

window['LPFrom'].Value = LPFrom
window['CAFrom'].Value = CATrue

window.metadata = {'WindowOpened':datetime.now(),'LastUpdate':datetime.now(),'icon':icon,'ElementWithFocus':None,'CALoggedIn':False,'CACILastCheck':datetime.now()-timedelta(minutes=30),'PinPopup':False}

window.perform_long_operation(lambda:EventHelper('WindowReappear',0.5),'WindowReappear')

window['OKRefresh'].metadata = {'LastCheck':datetime.now()}
window['CASearch'].metadata = {'Searching':False,'LastUpdate':datetime.now(),'SearchStart':datetime.now()}
window['LVShow'].metadata = {'LastUpdate':datetime.now(),'Shown':False}
window['CACOTable'].metadata = {'Searching':False,'LastUpdate':datetime.now(),'SearchStart':datetime.now(),'Accounts':[],'Page':1,'Total':'0','Recieved':'0','Updating':False}
window['CACOAutoUpdate'].metadata = {'TimedOut':False}

if CATrue and not ConfigSyncOpt:
    window.perform_long_operation(lambda:EventHelper('SyncOptToggle'),'SyncOptToggle')
else:
    if sg.running_mac():
        window['SyncOptToggle'].TKStyle.configure(window['SyncOptToggle'].ttk_style_name ,relief='sunken')
    else: 
        window['SyncOptToggle'].TKButton.configure(relief='sunken')

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
    if LastPass.loggedin:
        window['LPStatus'].update(value=f'✅ You are logged into LastPass')
  
if Inputs['LPAuto']:       
    window['LPList'    ].update(disabled=True)
    window['LPList'    ].Disabled=True
    window['LPSelected'].update(disabled=True)
    window['LPClear'   ].update(disabled=True)

CALoginAttempt = 0
# if LPFrom:
#     window['LPFrom'].Value = False
#     window.perform_long_operation(lambda:EventHelper('LPFrom'),'LPFrom')
    
if ConfigCAIndex and ConfigCATable and len(ConfigCATable) >= ConfigCAIndex[0]:
    try: 
        window['CATable'].update(select_rows=ConfigCAIndex)
        ScrollPosition = ConfigCAIndex[0] / (len(ConfigCATable) - 1) if ConfigCAIndex[0] != 0 else 0
        window['CATable'].set_vscroll_position(ScrollPosition)
    except:
        pass
if not sg.running_mac() and Inputs['LVAutoStart']:
    HotKey = Inputs['LVHotKey']
    HotKeyThread = window.perform_long_operation(lambda:HotKeyListener(HotKey),'HotKeyReturn')
    
if LPTrue:
    window['LPOptions'].select()
else:
    window['OKOptions'].select()

window.perform_long_operation(lambda:EventHelper('CACOSilentRefresh'),'CACOSilentRefresh')


if (not Inputs['CACOAutoUpdate'] or not Inputs['CACOAutoCheckIn']) and not window['CACOPerAccount'].Disabled:
    window.perform_long_operation(lambda:EventHelper(Inputs['CACOAutoUpdate']),'CACOAutoUpdate')
#endregion

#region Application Behaviour if Statments - Event Loop
while True:
    try:
        #Read Events(Actions) and Users Inputs From Window 
        WindowRead = window.read(timeout=timeout)
        event      = WindowRead[0]
        Inputs     = WindowRead[1]
        
        #End Script when Window is Closed
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT:
            try:
                SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
            except:
                pass
            break
        
        if event == sg.WINDOW_CLOSED:
            break

        if event == 'WindowReappear':
            window.reappear()

        if event in ['HotKeyReturn','LVPaste','typesleepreturn','pastesleepreturn'] and not sg.running_mac():
            if event == 'HotKeyReturn' and Inputs['HotKeyReturn'] or event == 'typesleepreturn':
                seconds = float(Inputs['LVhkDelay'])
                if event != 'typesleepreturn':
                    window.perform_long_operation(lambda:sleep(seconds), 'typesleepreturn')
                    continue
                typestring(Inputs['LVPass'],Inputs['LVtDelay'])
                HotKey = Inputs['LVHotKey']
                window.perform_long_operation(lambda:HotKeyListener(HotKey),'HotKeyReturn')
  
            if event in ['LVPaste','pastesleepreturn']:
                seconds = float(Inputs['LVtfmDelay']) 
                if event != 'pastesleepreturn':
                    window.perform_long_operation(lambda:sleep(seconds), 'pastesleepreturn')
                    continue
                typestring(Inputs['LVPass'],Inputs['LVtDelay'])
        
        if event in ['LVRestartL','LVRestartReturn'] and not sg.running_mac():
            kb.unhook_all()
            HotKey = Inputs['LVHotKey']
            window.perform_long_operation(lambda:HotKeyListener(HotKey),'HotKeyReturn')

        if event == 'LVStop' and not sg.running_mac():
            kb.unhook_all()
            
        if window['LVTable2'].get() != window['LVTable'].get():
            window['LVTable2'].update(values=window['LVTable'].get())

        if event == 'CASettingsToggle':
            numrows = window['CACOTable'].NumRows
            if not window["CACOSettings"].visible:
                #if sg.running_mac(): window['CACOTable'].update(num_rows=numrows) ⩢
                window["CACOSettings"].update(visible=True)
                window['CACOTable'].NumRows = numrows - 6
                window['CACOTable'].update(num_rows=numrows - 6)
                window['CASettingsToggle'].update(value='⩢ Settings')
            elif window["CACOSettings"].visible:
                window["CACOSettings"].update(visible=False)
                window['CACOTable'].NumRows = numrows + 6
                window['CACOTable'].update(num_rows=numrows + 6)
                window['CASettingsToggle'].update(value='⩠ Settings')

            
        if event in ["CACOView",'CACOTable']:
            CACOTable = window['CACOTable'].get()
            LVTable   = window['LVTable'].get()
            if Inputs['CACOTable'] and CACOTable and CACOTable[0]:
                window["CheckIn"].update(disabled=False)
                if LVTable and LVTable[0]:
                    CACOSelected = CACOTable[Inputs['CACOTable'][0]]
                    Account = None
                    if Account:
                        index = [_[0] for _ in enumerate(LVTable) if _[1][3] == Account.id]
                    else:
                        index = [_[0] for _ in enumerate(LVTable) if _[1][0] == CACOSelected[9] and _[1][4] == 'MEDHOST Password Manager\\CyberArk']

                    if index:
                        window["CACOView"].update(disabled=False)
                        if event == "CACOView":
                            ScrollPosition = index[0] / (len(LVTable) - 1) if LVTable and index and len(LVTable) != 1 else 0
                            window['LVTable'].update(select_rows=index)
                            window['LVTable'].set_vscroll_position(ScrollPosition)
                            window['LocalVault'].select()
                    else:
                        window["CACOView"].update(disabled=True)    

        if (window['CACOTable'].metadata['Searching'] or window['CACOTable'].metadata['Updating']) and window['CACOTable'].metadata['LastUpdate'] + timedelta(seconds=1) < datetime.now():
            window['CACOTable'].metadata['LastUpdate'] = datetime.now()
            CACOTable = window['CACOTable'].get()
            CACOPage = window['CACOTable'].metadata['Page']
            offset = 1000 * CACOPage
            AcctString = 'Accounts' if window['CACOTable'].metadata['Total'] else ''
            offsetstring = '' if not window['CACOTable'].metadata['Total'] else str(window['CACOTable'].metadata['Recieved']) + ' of ' + str(window['CACOTable'].metadata['Total'])
            if window['CACOTable'].metadata['Searching']:
                if CACOTable[4][2] != 'Searching..........' and 'Searching' in CACOTable[4][2]:
                    window['CACOTable'].update(values=[[],[],['','',AcctString],['','',offsetstring],['','',CACOTable[4][2] + '.']])
                else:
                    window['CACOTable'].update(values=[[],[],['','',AcctString],['','',offsetstring],['','','Searching']])
                
            if window['CACOTable'].metadata['Updating']:
                if Inputs['CACOStatus'] != 'Updating..........' and 'Updating' in Inputs['CACOStatus']:
                    window['CACOStatus'].update(value=Inputs['CACOStatus'] + '.')
                else:
                    window['CACOStatus'].update(value='Updating')
                if window['CACOTable'].metadata['Total']:
                    window['CACOLastUpdated'].update(value=offsetstring)
                    
            if window['CACOTable'].metadata['SearchStart'] + timedelta(minutes=5) < datetime.now():
                window['CACOTable'].metadata['Searching'] = False
                window['CACOTable'].metadata['Updating']  = False
                window['CACOTable'].metadata['Accounts'] = []
                window['CACOTable'].update(values=[])
                window['CACOStatus'].update(value='Timed Out')
                window['CACOLastUpdated'].update(value='')

        if (event in ['CACORefresh',"CACOLogin",'CACOSilentRefresh'] and not window['CACOTable'].metadata['Searching'] and not window['CACOTable'].metadata['Updating']) or (isinstance(event,str) and event.startswith("CACOReturn")):
            if event in ['CACORefresh',"CACOLogin",'CACOSilentRefresh']:
                window['CACOAutoUpdate'].metadata['TimedOut'] = False
                if event == "CACOLogin":
                    if 'ErrorMessage' not in Inputs["CACOLogin"]:
                        window['CACheckedOut'].select()
                        ArkToken = Inputs["CACOLogin"]
                        ArkInfo  = {'Token': ArkToken}
                        SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                        window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')
                    else:
                        continue
                else:
                    ArkToken = AppData['ArkToken'] if not ArkToken else ArkToken

                CACOTable  = window['CACOTable']
                CACOTable = CACOTable.Values if isinstance(CACOTable,sg.Table) else []
                if event != 'CACOSilentRefresh' and ((not CACOTable or not CACOTable[0]) or not window.metadata['CALoggedIn']):  
                    window.perform_long_operation(lambda:FindCyberArkAccount(ArkToken,limit='1'), "CACOReturn")
                    window['CACOTable'].metadata = {'Searching':True,'LastUpdate':datetime.now(),'SearchStart':datetime.now(),'Accounts':[],'Page':0,'Total':'','Recieved':'0','Updating':True}                    
                    window['CACOTable'].update(values=[[],[],['','',''],['','',''],['','','Searching']])
                else:
                    window.perform_long_operation(lambda:FindCyberArkAccount(ArkToken,limit='1'), "CACOReturnSilent")
                    window['CACOTable'].metadata = {'Searching':False,'LastUpdate':datetime.now(),'SearchStart':datetime.now(),'Accounts':[],'Page':0,'Total':'','Recieved':'0','Updating':True}

                window['CACOStatus'].update(value='Updating')
                continue
            else:
                Response = Inputs[event]
                if 'ErrorMessage' in Response:
                    window.metadata['CALoggedIn'] = False
                    window['CACOTable'].metadata['Searching'] = False
                    window['CACOTable'].metadata['Updating']  = False
                    window['CACOStatus'].update(value='Login Issue')
                    window['CACOLastUpdated'].update(value='Press Refresh')
                    window['CACOAutoUpdate'].metadata['TimedOut'] = True
                    if event != "CACOReturnSilent":
                        window['CACOTable'].update(values=[])
                        window['OutPutSection'].select()
                        window['-OUTPUT-'].update(value='')
                        ErrorM = Response['ErrorMessage']
                        print(ErrorM)
                        if 'token' in ErrorM or 'logon' in ErrorM or 'terminated' in ErrorM:
                            if Inputs['ArkUser'] and Inputs['ArkPass']:
                                print('Attempting login to get new token')
                                if Inputs['CAPasscode']:
                                    location = window.current_location()
                                    if sg.running_mac():
                                        pinwinlocation = (location[0]+99, location[1]+307)
                                    else:
                                        pinwinlocation = (location[0]+74, location[1]+328)
                                    Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                                    if not Passcode:
                                        print('Login Canceled')
                                        continue
                                    ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                                else:
                                    ArkPass  = Inputs["ArkPass"]
                                window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), "CACOLogin")                 
                            else:
                                print('Fill in CyberArk username and password fields to login and get new token')
                                window.perform_long_operation(lambda:EventHelper('CALoginToggle'),'CALoginToggle')
                        continue
                else:
                    if event in ["CACOReturn","CACOReturnSilent",'CACOReturnSleep']:
                        if event in ["CACOReturn","CACOReturnSilent"]: window.metadata['CALoggedIn'] = True
                        if window.metadata['CALoggedIn']:
                            window['CACOTable'].metadata['Total']    = Response['Total']
                            window['CACOTable'].metadata['Recieved'] = len(window['CACOTable'].metadata['Accounts'])                     
                            window['CACOTable'].metadata['Page'] = window['CACOTable'].metadata['Page'] + 1 
                            window.perform_long_operation(lambda:FindCyberArkAccount(ArkToken,limit='1000',page=str(window['CACOTable'].metadata['Page'])),"CACOReturn2_" + str(window['CACOTable'].metadata['Page']))
                            window['CACOTable'].metadata['LastUpdate'] = datetime.now()
                            offset = window['CACOTable'].metadata['Page'] * 1000
                            if offset < Response['Total']: window.perform_long_operation(lambda:EventHelper({'Total':Response['Total']},1),'CACOReturnSleep')
                    elif event.startswith("CACOReturn2"):
                        try:
                            window['CACOTable'].metadata['Accounts'] += Response['Accounts']
                            window['CACOTable'].metadata['Total']    = Response['Total']
                            window['CACOTable'].metadata['Recieved'] = len(window['CACOTable'].metadata['Accounts'])              
                            if len(window['CACOTable'].metadata['Accounts']) == Response['Total']:
                                OldTable = window['CACOTable'].get()
                                window['CACOTable'].metadata['Searching'] = False
                                window['CACOTable'].metadata['Updating']  = False
                                window['CACOStatus'].update(value='Done')
                                window['CACOTable'].metadata['LastUpdate'] = datetime.now()
                                window['CACOLastUpdated'].update(value=getdate(datetime.now()))
                                Accounts = window['CACOTable'].metadata['Accounts']
                                CACOAccts = {'Accounts':[_ for _ in Accounts if _['Properties']['LockedBy'] in [Inputs['ArkUser'],Inputs['ArkUser'].lower(),Inputs['ArkUser'].lower().title()]]}
                                CACOTable = MakeCyberArkTable(CACOAccts)
                                if len(OldTable) >= 1 and OldTable[0]:
                                    OldToKeep = [_ for _ in OldTable if len(_) >= 13 and _[9] not in [_[9] for _ in CACOTable] and str(_[12]).isnumeric() and int(_[12]) > int(datetime.timestamp(window['CACOTable'].metadata['SearchStart']))]
                                    CACOTable = CACOTable + OldToKeep
                                window['CACOTable'].update(values=CACOTable)
                                window['CACOTable'].metadata['Accounts'] = []
                                if CACOTable: event = 'CheckInCycleStart'
                        except Exception as error:
                            print(error)
                            print(Response)

            if window.metadata['CALoggedIn'] and event == 'CheckInCycleStart' and Inputs['CACOAutoUpdate'] and Inputs['CACOAutoCheckIn'] and LastPass and LastPass.accounts:
                window.metadata['CACILastCheck'] = datetime.now()
                if Inputs['CACOPerAccount']:
                    for entry in LVTable:
                        entryhash = literal_eval(entry[5].split('\nNotes:')[1].rstrip() if len(entry) >= 6 and entry[5] and "'AutoCheckIn'" in entry[5] and "\nNotes:{" in entry[5] else '{}')
                        if entryhash and entryhash['AutoCheckIn']:
                            if entryhash['CheckInDelay'] and ' Hrs' in entryhash['CheckInDelay']:
                                try:
                                    Interval = entryhash['CheckInDelay'].split(' Hrs')[0].split(':')
                                    Hours    = int(Interval[0])
                                    Minutes  = int(Interval[1])
                                    delay    = timedelta(hours=Hours,minutes=Minutes) 
                                except:
                                    delay = None
                            else:
                                delay = None

                            if delay:
                                CAAcct = None
                                CACOTable = window['CACOTable'].get() if window['CACOTable'].get() and window['CACOTable'].get()[0] else []
                                CAAcct = [_ for _ in CACOTable if len(_) >= 14 and _[13] == entryhash['CyberArkId']]
                                if isinstance(CAAcct,list) and len(CAAcct[0]) >= 13 and CAAcct[0][0] and datetime.fromtimestamp(int(CAAcct[0][12])) + delay <= datetime.now():
                                    window.perform_long_operation(lambda:EventHelper(CAAcct[0]),'CASilentCheckIn')
                elif Inputs['CACOGlobal'] and Inputs['CACODelay'] != 'Disabled':
                    CACOTable = window['CACOTable'].get() if window['CACOTable'].get() and window['CACOTable'].get()[0] else []
                    if CACOTable:
                        try:
                            Interval = Inputs['CACODelay'].split(' Hrs')[0].split(':')
                            Hours    = int(Interval[0])
                            Minutes  = int(Interval[1])
                            delay    = timedelta(hours=Hours,minutes=Minutes) 
                        except:
                            delay = None

                        if delay:
                            for CAAcct in CACOTable:    
                                if isinstance(CAAcct,list) and len(CAAcct) >= 13 and CAAcct[0] and datetime.fromtimestamp(int(CAAcct[12])) + delay <= datetime.now():
                                    window.perform_long_operation(lambda:EventHelper(CAAcct),'CASilentCheckIn')

        if event in ['CACOAutoUpdate','CACOAutoCheckIn']:
            if Inputs['CACOAutoUpdate']:
                window['CACOAutoCheckIn'].update(disabled=False)
                if Inputs['CACOAutoCheckIn']:
                    window['CACOPerAccount'].update(disabled=False)
                    window['CACOGlobal'].update(disabled=False)
            if not Inputs['CACOAutoUpdate'] or not Inputs['CACOAutoCheckIn']:
                window['CACOPerAccount'].update(disabled=True)
                window['CACOGlobal'    ].update(disabled=True)
                if not Inputs['CACOAutoUpdate']:
                    window['CACOAutoCheckIn'].update(disabled=True)

        if event == 'CACOPerAccount' or not window['CACODelay'].Disabled and (Inputs['CACOPerAccount'] or not Inputs['CACOAutoUpdate'] or not Inputs['CACOAutoCheckIn']):
            window['CACODelay'].update(disabled = True)

        if event == 'CACOGlobal' or Inputs['CACOGlobal'] and window['CACODelay'].Disabled and Inputs['CACOAutoUpdate'] and Inputs['CACOAutoCheckIn']:
            window['CACODelay'].update(disabled = False)

        if event == 'CACODelay':
            window['CACOTable'].metadata['LastUpdate'] = datetime.now() - timedelta(minutes=float(Inputs['CACOAutoInterval'])-1)

        if event == 'GetPin' or window.metadata['PinPopup'] and pinwindow:
            location = window.current_location()
            if sg.running_mac():
                pinwinlocation = (location[0]+99, location[1]+307)
            else:
                pinwinlocation = (location[0]+74, location[1]+328)
                
            if event == 'GetPin':
                window.metadata['PinPopup'] = True
                window.hide()
                pinwindow = pin_popup('Enter Your Pin',password_char="*",title='MEDHOST Password Manager',location=pinwinlocation,icon=icon,pass_through=True)
                continue
            else:
                Pinbutton, Pinvalues = pinwindow.read(timeout=0)
                if Pinbutton in ['Cancel',sg.WINDOW_CLOSED]:
                    pinwindow.close()
                    break
                elif Pinbutton == 'Ok':
                    pin = Pinvalues['PIN_INPUT_']
                    AppData = GetConfigAppData(WorkingDirectory,pin,False)
                    if AppData == 'error':
                        pinwindow.close()
                        pinwindow = pin_popup('Try Reentering your pin',password_char="*",title='MEDHOST Password Manager',location=pinwinlocation,icon=icon,pass_through=True)
                    else:
                        pinwindow.close()
                        window.un_hide()
                        pinwindow                     = None
                        window.metadata['PinPopup']   = False
                        window.metadata['LastUpdate'] = datetime.now()


        #region Time Dependent actions
        if event:
            try:
                element = window.find_element_with_focus()
            except:
                element = None       

            try:
                CACOAutoTimeOut  = float(Inputs['CACOAutoTimeOut'])
                CACOAutoInterval = float(Inputs['CACOAutoInterval'])
            except:
                CACOAutoTimeOut  = 8
                CACOAutoInterval = 15

            if element != window.metadata['ElementWithFocus']:
                timeout = 100
                window.metadata['LastUpdate'] = datetime.now()
            elif window.metadata['LastUpdate'] and not isinstance(window.metadata['LastUpdate'],bytes) and window.metadata['LastUpdate'] + timedelta(minutes=1) < datetime.now() and timeout and timeout < 1000:
                timeout = 1000
                if test and '1 Minute TimeOut' not in window['-OUTPUT-'].get():
                    print('1 Minute TimeOut')
            # elif window.metadata['LastUpdate'] and window.metadata['LastUpdate'] + timedelta(minutes=5) < datetime.now() and timeout and timeout < 2000:
            #     timeout = 2000
            #     if test and '5 Minute TimeOut' not in window['-OUTPUT-'].get():
            #         print('5 Minute TimeOut')
            elif isinstance(window.metadata['LastUpdate'],datetime) and window.metadata['LastUpdate'] + timedelta(hours=CACOAutoTimeOut) < datetime.now() and not window['CACOAutoUpdate'].metadata['TimedOut']:
                window['CACOAutoUpdate'].metadata['TimedOut'] = True
                window['CACOStatus'].update(value='Timed Out')
                window['CACOLastUpdated'].update(value='Press Refresh')
                if test and (Inputs['CACOAutoTimeOut'] + ' Hour TimeOut') not in window['-OUTPUT-'].get():
                    print(Inputs['CACOAutoTimeOut'] + ' Hour TimeOut')
            
            if element == window.metadata['ElementWithFocus'] and window.metadata['LastUpdate'] and window.metadata['LastUpdate'] + timedelta(hours=4) < datetime.now() and not window.metadata['PinPopup']:
                window.perform_long_operation(lambda:EventHelper('GetPin'),'GetPin')
                
            if Inputs['CACOAutoUpdate'] and window['CACOTable'].metadata['LastUpdate'] + timedelta(minutes=CACOAutoInterval) < datetime.now() and not window['CACOAutoUpdate'].metadata['TimedOut']:
                window.perform_long_operation(lambda:EventHelper('CACOSilentRefresh'),'CACOSilentRefresh')
            
            if not window['Submit'].Disabled:
                if not Inputs['CATable'] and Inputs['CAFrom']: 
                    window['Submit'].update(disabled=True)
            else:
                if Inputs['CATable'] or not Inputs['CAFrom']:
                    window['Submit'].update(disabled=False)
            
            if window['CADelay'].Disabled and Inputs['CATable'] and Inputs['CACOPerAccount']:
                window['CADelay'].update(disabled=False)
                window['CAAutoCheckIn'].update(text_color='white')
            elif not window['CADelay'].Disabled and not Inputs['CATable'] and Inputs['CACOPerAccount']:      
                window['CADelay'].update(disabled=True)
                window['CAAutoCheckIn'].update(text_color='dark grey')
            
            if not window['CADelay'].visible and Inputs['CACOAutoUpdate'] and Inputs['CACOAutoCheckIn'] and Inputs['LP'] and Inputs['CAFrom'] and not Inputs['CACOGlobal']:
                window['CADelay'].update(visible=True)
                window['CAAutoCheckIn'].update(visible=True)
            elif window['CADelay'].visible and (Inputs['CACOGlobal'] or not Inputs['CACOAutoUpdate'] or not Inputs['CACOAutoCheckIn'] or not Inputs['LP'] or not Inputs['CAFrom']):
                window['CADelay'].update(visible=False)
                window['CAAutoCheckIn'].update(visible=False)

            if not Inputs['CACOTable']:
                if not window["CACOView"].Disabled: window["CACOView"].update(disabled=True)
                if not window["CheckIn"].Disabled: window["CheckIn"].update(disabled=True)

            if window['CACOTable'].metadata['Updating']:
                if not window["CACORefresh"].Disabled: window["CACORefresh"].update(disabled=True)
            else:
                if window["CACORefresh"].Disabled: window["CACORefresh"].update(disabled=False)
                    
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

            if element != window['LVQuery2'] and Inputs['LVQuery2'] == '':
                window['LVQuery2'].update(value='Search for accounts') 

            window.metadata['ElementWithFocus'] = element
        #endregion


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
                if not sg.running_mac(): 
                    window['SyncOptToggle'].TKButton.configure(relief='raised')
                    window['CAOptions'].set_size((470, 357))
                else:
                    window['SyncOptToggle'].TKStyle.configure(window['SyncOptToggle'].ttk_style_name ,relief='raised') 
                window['CATable'].update(num_rows=numrows)
                window['SyncOpt'].update(visible=False)
                #window['SyncOptToggle'].update(text='⩢ Sync Options')
            elif Inputs and 'SyncOptToggle' not in Inputs or not Inputs['SyncOptToggle']:
                numrows = numrows - 5
                window['CATable'].NumRows = numrows
                if not sg.running_mac():
                    window['SyncOptToggle'].TKButton.configure(relief='sunken') 
                    window['CAOptions'].set_size((470, 277))
                else:
                    window['SyncOptToggle'].TKStyle.configure(window['SyncOptToggle'].ttk_style_name ,relief='sunken') 
                window['CATable'].update(num_rows=numrows)
                window['SyncOpt'].update(visible=True)                    
                #window['SyncOptToggle'].update(text='⩠ Sync Options')    

        if window['-OUTPUT-'].get().rstrip() == '' and window['-OUTPUT-'].get().rstrip() != window['-OUTPUT-2'].get().rstrip():
            window['-OUTPUT-2'].update(value='')

        if event == 'CATable' and not window['CAAcctDetails'].Disabled:
            CATable = window['CATable'].get()
            if CATable and (Inputs['CATable'] or Inputs['CATable'] == 0) and isinstance(Inputs['CATable'][0],int):
                window['CAAcctDetails'].select()
                window['Submit'].update(disabled=False)
                CAAcct  = CATable[Inputs['CATable'][0]]
                if isinstance(CAAcct,list) and len(CAAcct) > 14:
                    Useddate = getdate(int(CAAcct[12])) if CAAcct[12] else ''
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
                    window.perform_long_operation(lambda:GetCyberArkActivities(ArkToken,CAAcct[13]), 'CAActivitiesReturn')
                    try:
                        plat = CAAcct[14]
                        window.perform_long_operation(lambda:FindCyberArkPlatform(ArkToken,plat), 'CAPlatformReturn')
                    except:
                        pass

        if event == 'CAPlatformReturn' and window.metadata['CALoggedIn']:
            Response = Inputs['CAPlatformReturn']
            if 'ErrorMessage' in Response:
                window.metadata['CALoggedIn'] = False
                window['OutPutSection'].select()
                window['-OUTPUT-'].update(value='')
                ErrorM = Response['ErrorMessage']
                print(ErrorM)
                if 'token' in ErrorM or 'logon' in ErrorM or 'terminated' in ErrorM:
                    if Inputs['ArkUser'] and Inputs['ArkPass']:
                        print('Attempting login to get new token')
                        if Inputs['CAPasscode']:
                            location = window.current_location()
                            if sg.running_mac():
                                pinwinlocation = (location[0]+99, location[1]+307)
                            else:
                                pinwinlocation = (location[0]+74, location[1]+328)                            
                            Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                            if not Passcode:
                                print('Login Canceled')
                                continue                           
                            ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                        else:
                            ArkPass  = Inputs["ArkPass"]
                        window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), "CACOLogin")                 
                    else:
                        print('Fill in CyberArk username and password fields to login and get new token')
                        window.perform_long_operation(lambda:EventHelper('CALoginToggle'),'CALoginToggle')
                continue
            Platform = Response
            Delay    = int(int(Platform['Details']['MinValidityPeriod'])/60)
            hours    = int(Delay/2) if Delay <= 16 else 8
            ComboArray = ['Disabled']
            for hour in range(1,hours+1):
                ComboArray += [f'{hour-1}:30 Hrs',f'{hour}:00 Hrs']
            window['CADelay'].update(values=ComboArray)
            window['CADelay'].update(value='Disabled')



        if event == 'CAActivitiesReturn':
            Activities = Inputs['CAActivitiesReturn']     
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

        if event == 'click_LVQuery2':
            if Inputs['LVQuery2'] == 'Search for accounts':
                window['LVQuery2'].update(value='')
        
        if event in ['LVQuery2','LVsGroup2','LVsType2']:
            LVQTable = LVTable
            if Inputs['LVQuery2'] and Inputs['LVQuery2'] != 'Search for accounts':
                LVQTable = [_ for _ in LVQTable if Inputs['LVQuery2'] in ''.join([str(x) for x in _]) + ''.join([str(x).title() for x in _]) + ''.join([str(x).lower() for x in _])]
            if Inputs['LVsGroup2'] and Inputs['LVsGroup2'] != 'Group/Folder':
                LVQTable = [_ for _ in LVQTable if _[4].split('\\')[-1] == Inputs['LVsGroup2']]
            if Inputs['LVsType2'] and Inputs['LVsType2'] != 'Type':
                if Inputs['LVsType2'] == 'Password':
                    LVQTable = [_ for _ in LVQTable if _[2] != 'http://sn']
                elif Inputs['LVsType2'] == 'SecureNote':
                    LVQTable = [_ for _ in LVQTable if _[2] == 'http://sn']
                    
            window['LVTable2'].update(values=LVQTable)

        if event in ["LPUser","LPPass"]:
            window['LVlUser'].update(value=Inputs['LPUser'])
            window['LVlPass'].update(value=Inputs['LPPass'])
            window['LVlUser2'].update(value=Inputs['LPUser'])
            window['LVlPass2'].update(value=Inputs['LPPass'])
        
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
            window['LVlUser2'].update(value=Inputs['LVlUser'])
            window['LVlPass2'].update(value=Inputs['LVlPass'])
            Inputs['LPUser']=Inputs['LVlUser']
            Inputs['LPPass']=Inputs['LVlPass']
            window['OutPutSection2'].select()
            window['LVTable'].NumRows = Vr1
            if sg.running_mac(): window['LVTable'].update(num_rows=Vr1)
            window['LVLoginToggle'].update(value='⩠ Login')
            window['LVLoginInfo'].update(visible=not window['LVLoginInfo']._visible)
            if not sg.running_mac(): window['LVTable'].update(num_rows=Vr1)

        if event in ['LVLogin2','LVRefresh2']:
            window['LPUser'].update(value=Inputs['LVlUser2'])
            window['LPPass'].update(value=Inputs['LVlPass2'])
            window['LVlUser'].update(value=Inputs['LVlUser2'])
            window['LVlPass'].update(value=Inputs['LVlPass2'])
            Inputs['LPUser']=Inputs['LVlUser2']
            Inputs['LPPass']=Inputs['LVlPass2']
            window['OutPutSection'].select()
            numrows = window['LVTable2'].NumRows + 7
            window['LVTable2'].NumRows = numrows 
            if sg.running_mac(): window['LVTable2'].update(num_rows=numrows)
            window['LVLoginToggle2'].update(value='⩠ Login')
            window['LVLoginInfo2'].update(visible=not window['LVLoginInfo2']._visible)
            if not sg.running_mac(): window['LVTable2'].update(num_rows=numrows)


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

        if event == 'LVLoginToggle2':
            numrows = window['LVTable2'].NumRows
            if window['LVLoginInfo2']._visible:
                numrows = numrows + 7
                window['LVTable2'].NumRows = numrows
                if sg.running_mac(): window['LVTable2'].update(num_rows=numrows)
                window['LVLoginToggle2'].update(value='⩠ Login')
                window['LVLoginInfo2'].update(visible=not window['LVLoginInfo2']._visible)
                if not sg.running_mac(): window['LVTable2'].update(num_rows=numrows)
            else:
                numrows = numrows - 7
                window['LVTable2'].NumRows = numrows
                if sg.running_mac(): window['LVTable2'].update(num_rows=numrows)
                window['LVLoginInfo2'].update(visible=not window['LVLoginInfo2']._visible)
                window['LVLoginToggle2'].update(value='⩢ Login')
                if not sg.running_mac(): window['LVTable2'].update(num_rows=numrows)

        if window['LVShow'].metadata['Shown'] and window['LVShow'].metadata['LastUpdate'] + timedelta(seconds=60) < datetime.now():
            event = 'LVShow'

        if event == 'LVShow':
            window['LVShow'].metadata['LastUpdate'] = datetime.now()
            Table    = window['LVTable'].Values
            Table    = Table if isinstance(Table,list) else []
            if window['LVUrl'].get() == 'http://sn': LVAcct = Table[Inputs['LVTable'][0]]
            if window['LVPass'].PasswordCharacter == '*':
                window['LVShow'].metadata['Shown']      = True
                window['LVPass'].update(password_char='')
                if window['LVUrl'].get() == 'http://sn' or 'NoteType:' in LVAcct[5]: window['LVNote'].update(value=LVAcct[5])
            else:
                window['LVShow'].metadata['Shown'] = False
                window['LVPass'].update(password_char='*')
                if window['LVUrl'].get() == 'http://sn' or 'NoteType:' in LVAcct[5]:
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
                if LVAcct[2] == 'http://sn' or 'NoteType:' in LVAcct[5]:
                    window['LVNote'].update(value=len(LVAcct[5])*'*')
                else:
                    window['LVNote'].update(value=LVAcct[5])
                    
        if event == 'LVTable2':
            LVTable    = window['LVTable2'].get()
            LVSelected = [LVTable[_] for _ in Inputs['LVTable2']]
            Combolist = [_[0]+' '*120+str(_) for _ in LVSelected]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)
            

        if isinstance(event,str) and event.startswith('S_'):
            Slot = event.replace('S_','')
            window[f'C_{Slot}'].Disabled = not Inputs[event]
            window[f'C_{Slot}'].update(disabled=not Inputs[event])
            
        if event == 'TabSelect':
            window[Inputs['TabSelect']].select()
            
        if window['LPFrom'].Value == True and Inputs['LPFrom'] == False:
            if sg.running_mac(): window['OutPutSection'].select()
            window['LPFrom'].Value = False
            window.ReturnValuesDictionary['OK'] = OKTrue 
            Inputs['OK'] = OKTrue 
            window['LPOptions'].update(visible=Inputs['LP'])
            window['OKOptions'].update(visible=Inputs['OK'])
            window['LP'].update(visible=True)
            window['OK'].update(value=OKTrue)
            window['LP'].update(disabled = True if not OKTrue and Inputs['LP'] else False)
            window['OK'].update(disabled = True if not Inputs['LP'] and OKTrue else False)
                
        if window['CAFrom'].Value == True and Inputs['CAFrom'] == False:
            window['CAFrom' ].Value = False
            window['Submit' ].update(text='Sync Passwords')
            if not sg.running_mac(): 
                for _ in Slots: window[f"S_{_}"].update(font=('Helvetica',8))
            else:
                window['OutPutSection'].select()
            window['CAAcctDetails'    ].update(disabled=True)
            window['CAAcctDetails'    ].update(visible=False)
            window['CACheckedOut'     ].update(disabled=True)
            window['CACheckedOut'     ].update(visible=False)
            window['CADelay'          ].update(visible=False)
            window['CAAutoCheckIn'    ].update(visible=False)                
            window['SyncOptToggle'    ].update(visible=False)               
            #window['CheckIn'         ].update(visible=False)
            window['OkCheckText'      ].update(visible=False)  
            window['SlotSelMethod'    ].update(visible=False)
            window['KWLabelSearchText'].update(visible=False)
            window['OK_LPAcctText'    ].update(visible=True )
            window['OKComboLeft'      ].update(visible=True )
            window['OKComboRight'     ].update(visible=True )
            window['SlotSelector'     ].update(visible=True )
            #if event!= 'LPFrom': window = ReloadWindow(window)
            
        if event == 'GCFrom':
            window['GCOptions'].update(visible=Inputs['GCFrom'])
            window['MEOptions'].update(visible=Inputs['MEFrom'])
            window['CAOptions'].update(visible=Inputs['CAFrom'])
            window['LVOptions'].update(visible=Inputs['LPFrom'])
            window['LPAuto'  ].update(disabled=False)
            window['LPExist' ].update(disabled=True)
            window['LPAuto'  ].update(value=True)
            window['LPExist' ].update(value=False)
            GCTable    = window['GCSelected']
            GCTable    = GCTable.Values if isinstance(GCTable,sg.Table) else []
            Combolist = [_[0]+' '*120+str(_) for _ in GCTable]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)     
            if Inputs['LP']: window['LPOptions'].select()
            else: window['OKOptions'].select()
            if window['LPLoginInfo'].visible: window.perform_long_operation(lambda:EventHelper('LPLoginToggle'),'LPLoginToggle')
            window.perform_long_operation(lambda:EventHelper('LPAuto'),'LPAuto')        

        if event == 'MEFrom':
            window['MEOptions'].update(visible=Inputs['MEFrom'])
            window['CAOptions'].update(visible=Inputs['CAFrom'])
            window['GCOptions'].update(visible=Inputs['GCFrom'])
            window['LVOptions'].update(visible=Inputs['LPFrom'])
            window['LPAuto'  ].update(disabled=False)
            window['LPExist' ].update(disabled=True)
            window['LPAuto'  ].update(value=True)
            window['LPExist' ].update(value=False)
            MESelected = window['MESelected']
            MESelected = MESelected.Values if isinstance(MESelected,sg.Table) else []
            Combolist = [_[0]+' '*120+str(_) for _ in MESelected]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)
            if Inputs['LP']: window['LPOptions'].select()
            else: window['OKOptions'].select()
            if window['LPLoginInfo'].visible: window.perform_long_operation(lambda:EventHelper('LPLoginToggle'),'LPLoginToggle')
            window.perform_long_operation(lambda:EventHelper('LPAuto'),'LPAuto')
            
        if event == 'CAFrom' and window['CAFrom'].Value == False:
            window['CAFrom'].Value = True
            if not sg.running_mac():
                for _ in Slots: window[f"S_{_}"].update(font=('Helvetica',10))
            window['Submit'           ].update(text='Check-Out & Sync')
            window['CAOptions'        ].update(visible=Inputs['CAFrom'])
            window['MEOptions'        ].update(visible=Inputs['MEFrom'])
            window['GCOptions'        ].update(visible=Inputs['GCFrom'])
            window['LVOptions'        ].update(visible=Inputs['LPFrom'])
            window['CheckIn'          ].update(visible=True)
            window['SlotSelector'     ].update(visible=Inputs['PSlotsRadio'])
            window['KWLabelSearchText'].update(visible=Inputs['KWSearch'])
            window['CAAcctDetails'    ].update(disabled=False)
            window['CACheckedOut'     ].update(disabled=False)
            window['CACheckedOut'     ].update(visible=True) 
            window['CAAcctDetails'    ].update(visible=True)
            window['CADelay'          ].update(visible=True)
            window['CAAutoCheckIn'    ].update(visible=True)  
            window['SyncOptToggle'    ].update(visible=True)
            window['SlotSelMethod'    ].update(visible=True)
            window['OkCheckText'      ].update(visible=True)
            window['OK_LPAcctText'    ].update(visible=False)
            window['OKComboLeft'      ].update(visible=False)
            window['OKComboRight'     ].update(visible=False)
            window['LPAuto'           ].update(disabled=False)
            window['LPExist'          ].update(disabled=False)
            window['LPAuto'           ].update(value=True)
            window['LPExist'          ].update(value=False)
            if Inputs['LP']: window['LPOptions'].select()
            else: window['OKOptions'].select()
            if window['LPLoginInfo'].visible: window.perform_long_operation(lambda:EventHelper('LPLoginToggle'),'LPLoginToggle')
            window.perform_long_operation(lambda:EventHelper('LPAuto'),'LPAuto')
  
        if event == 'LPFrom' and window['LPFrom'].Value == False:
            window['OutPutSection'].select()
            window['LPFrom'].Value  = True
            window['MEOptions'].update(visible=Inputs['MEFrom'])
            window['GCOptions'].update(visible=Inputs['GCFrom'])
            window['CAOptions'].update(visible=Inputs['CAFrom'])
            window['LVOptions'].update(visible=Inputs['LPFrom'])
            window['OKOptions'].update(visible=True)
            window['OKOptions'].select()
            window['LPOptions'].update(visible=False)
            window['LP'       ].update(visible=False)
            window['OK'       ].update(disabled=True)
            OKTrue = Inputs['OK']
            window['OK'].update(value=True)
            LVTable2 = window['LVTable2']
            Combolist = [_[0]+' '*120+str(_) for _ in [LVTable[_] for _ in Inputs['LVTable2']]]
            for _ in Slots: 
                window[f"C_{_}"].update(values=Combolist)

        if event == 'LP':
            if sg.running_mac(): window['OutPutSection'].select()
            window['OK'].Disabled = not Inputs['LP']
            window['OK'].update(disabled=not Inputs['LP'])
            try:
                window['LPOptions'].update(visible=Inputs['LP'])
            except:
                window['LPOptions']._visible=Inputs['LP']
            if Inputs['LP']: window.perform_long_operation(lambda:EventHelper('LPOptions'),'TabSelect')
    
        if event == 'OK':
            if sg.running_mac(): window['OutPutSection'].select()
            window['LP'].Disabled = not Inputs['OK']
            window['LP'].update(disabled=not Inputs['OK'])
            window['OKOptions'].update(visible=not window['OKOptions'].visible)
            #if Inputs['OK'] and not Inputs['LPFrom']: window.perform_long_operation(lambda:EventHelper('OKOptions'),'TabSelect')
            

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
                window['-OUTPUT-'].update(value='')
                if ArkAccount and 'ErrorMessage' in ArkAccount:
                    print(ArkAccount['ErrorMessage'])
                if Inputs['ArkUser'] and Inputs['ArkPass']:
                    print('Unable to refresh accounts list, attempting login to get new token')
                    if Inputs['CAPasscode']:
                        location = window.current_location()
                        if sg.running_mac():
                            pinwinlocation = (location[0]+99, location[1]+307)
                        else:
                            pinwinlocation = (location[0]+74, location[1]+328)
                        Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                        if not Passcode:
                            print('Login Canceled')
                            continue
                        ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                    else:
                        ArkPass  = Inputs["ArkPass"]
                    window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), "CALoginReturn")
                else:
                    print('Unable to refresh accounts list, login to get new token')
                continue
            else:
                window.perform_long_operation(lambda:EventHelper('CACOSilentRefresh'),'CACOSilentRefresh')
                CATable = MakeCyberArkTable(ArkAccts)
                window['CATable'].update(values=CATable)
                
        if event in ['CALogin','CALoginReturn',"CALoginSearch","CALoginAcctReturn"]:
            if event == 'CALogin':
                window['-OUTPUT-'].update('')
                window['OutPutSection'].select()
                if Inputs['CAPasscode']:
                    location = window.current_location()
                    if sg.running_mac():
                        pinwinlocation = (location[0]+99, location[1]+307)
                    else:
                        pinwinlocation = (location[0]+74, location[1]+328)
                    Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                    if not Passcode:
                        print('Login Canceled')
                        continue
                    ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                else:
                    ArkPass  = Inputs["ArkPass"]
                window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), "CALoginReturn")
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
                    window['-OUTPUT-'].update(value='')
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
                if SelectedRows:
                    window['CATable'].update(select_rows=SelectedRows)
                    window['CATable'].set_vscroll_position(ScrollPosition)
                ArkInfo = {'Token': ArkToken}
                SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')
                window.perform_long_operation(lambda:EventHelper('CACOSilentRefresh'),'CACOSilentRefresh')

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
                        if Inputs['CAPasscode']:
                            location = window.current_location()
                            if sg.running_mac():
                                pinwinlocation = (location[0]+99, location[1]+307)
                            else:
                                pinwinlocation = (location[0]+74, location[1]+328)
                            Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                            if not Passcode:
                                print('Login Canceled')
                                continue
                            ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                        else:
                            ArkPass  = Inputs["ArkPass"]
                        window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), "CALoginSearch")                 
                    else:
                        print('Fill in CyberArk username and password fields to login and get new token')
                        window.perform_long_operation(lambda:EventHelper('CALoginToggle'),'CALoginToggle')
                continue
            else:
                CATable = MakeCyberArkTable(ArkAccts) 
                SelectedRows = Inputs['CATable'] if CATable == window['CATable'].Values else []
                ScrollPosition = Inputs['CATable'][0] / (len(window['CATable'].Values) - 1) if Inputs['CATable'] and len(window['CATable'].Values) != 1 and CATable == window['CATable'].Values else 0
                window['CATable'].update(values=CATable)
                if SelectedRows:
                    window['CATable'].update(select_rows=SelectedRows)
                    window['CATable'].set_vscroll_position(ScrollPosition)
                CACOTable  = window['CACOTable'].get()
                for CAAcct in CATable:
                    if isinstance(CAAcct,list) and CAAcct[10] in [Inputs['ArkUser'],Inputs['ArkUser'].lower(),Inputs['ArkUser'].lower().title()] and (not CACOTable or not CACOTable[0] or CAAcct[9] not in [_[9] for _ in CACOTable]):
                        if not CACOTable or not CACOTable[0]:
                            window['CACOTable'].metadata['Searching'] = False
                            window['CACOTable'].update(values=[CAAcct])
                        else:
                            NewList = CACOTable + [CAAcct]
                            window['CACOTable'].update(values=NewList)

                SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                window.perform_long_operation(lambda:EventHelper('CACOSilentRefresh'),'CACOSilentRefresh')            
                
        if event in ['CheckIn','CALoginCheckIn','CASilentCheckIn']:
            if event == 'CheckIn':window['-OUTPUT-'].update(value='')
            if event != 'CASilentCheckIn':
                window['OutPutSection'].select()   
                if not Inputs['CACOTable']:
                    print('You must select an Account in CyberArk to check-out')
                    continue
                CACOTable = window['CACOTable']
                CACOTable = CACOTable.Values if isinstance(CACOTable,sg.Table) else []
                Acct  = CACOTable[Inputs['CACOTable'][0]]
            else:
                Acct = Inputs['CASilentCheckIn']
                
            if event == 'CALoginCheckIn':                        
                if 'ErrorMessage' not in Inputs["CALoginCheckIn"]:
                    ArkToken = Inputs["CALoginCheckIn"]
                    ArkInfo  = {'Token': ArkToken}
                    SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                    window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')
                else:
                    print(Inputs["CALoginCheckIn"]['ErrorMessage'])
                    continue 
            else:         
                ArkToken = AppData['ArkToken'] if not ArkToken else ArkToken

            CheckInResponse = CheckInCyberArkPassword(AccountId=Acct[13], Token=ArkToken, UserName=Acct[1])
            
            if CheckInResponse and hasattr(CheckInResponse,'ok') and CheckInResponse.ok:
                CACOTable = window['CACOTable'].get()
                if CACOTable and CACOTable[0] and Acct[9] in [_[9] for _ in CACOTable]:
                    for i,CAAcct in enumerate(CACOTable):
                        if isinstance(CAAcct,list) and  CAAcct[9] == Acct[9]:
                            CACOTable[i][12] = str(int(datetime.now().timestamp()))
                            window['CACOTable'].update(values=CACOTable)
                
                if Inputs['CACOAutoDelete']:
                    index = [_[0] for _ in enumerate(LVTable) if _[1][0] == Acct[9] and f"'CyberArkId': '{Acct[13]}'" in _[1][5] and _[1][4] == 'MEDHOST Password Manager\\CyberArk']
                    if index and LastPass and LastPass.accounts:
                        LPEntry = LVTable[index[0]]
                        Account = [_ for _ in LastPass.accounts if _.name == Acct[9] and f"'CyberArkId': '{Acct[13]}'" in _.notes and _.group == 'MEDHOST Password Manager\\CyberArk'][0]
                        LastPass.DeleteAccount(LPEntry[3],Account)
                        LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                        window['LVTable'].update(values=LVTable)
                        LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                        if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                        else: window['LPList' ].update(values=LPList)

                print(f'CyberArk Account: {Acct[9]} successfully checked in')
                window.perform_long_operation(lambda:EventHelper('CACOSilentRefresh',240),'CACOSilentRefresh')
                continue
            elif CheckInResponse.content:
                try:
                    Rjson = CheckInResponse.json()
                except:
                    Rjson = {}
                if 'ErrorMessage' in Rjson:
                    ErrorM = CheckInResponse.json()['ErrorMessage']
                    print(ErrorM)
                    window.metadata['CALoggedIn'] = False
                    if ('token' in ErrorM or 'logon' in ErrorM) and event != 'CASilentCheckIn':
                        if Inputs['ArkUser'] and Inputs['ArkPass']:
                            print('Attempting login to get new token')
                            if Inputs['CAPasscode']:
                                location = window.current_location()
                                if sg.running_mac():
                                    pinwinlocation = (location[0]+99, location[1]+307)
                                else:
                                    pinwinlocation = (location[0]+74, location[1]+328)
                                Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                                if not Passcode:
                                    print('Login Canceled')
                                    continue
                                ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                            else:
                                ArkPass  = Inputs["ArkPass"]
                            window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), 'CALoginCheckIn')                 
                        else:
                            print('You will need to login to CyberArk to check in an Account')
                        continue
            print('Unable to check in password')
    
        if event in ['LPRefresh','LVRefresh','LPRefreshReturn','LVRefresh2']:
            if event != 'LPRefreshReturn':
                window.perform_long_operation(lambda:LastPass.OpenVault(True), "LPRefreshReturn")
                continue
            else:
                LastPass = Inputs['LPRefreshReturn']
                LastPass.login_check()
                if not LastPass.loggedin:
                    window['OutPutSection'].select()
                    window['-OUTPUT-'].update(value='')
                    if Inputs['LPUser'] and Inputs['LPPass']:
                        print('Unable to refresh accounts list, attempting login to get new token')
                        if Inputs['LPPasscode']:
                            location = window.current_location()
                            if sg.running_mac():
                                pinwinlocation = (location[0]+99, location[1]+307)
                            else:
                                pinwinlocation = (location[0]+74, location[1]+328)
                            Passcode          = pin_popup('Enter MFA Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                            if not Passcode:
                                print('Login Canceled')
                                continue
                            Inputs["LPPass"]  = Inputs["LPPass"] + ',' + Passcode       
                        window.perform_long_operation(lambda:LastPassLogin(Inputs), "LPLoginReturn")
                    else:
                        print('Unable to refresh accounts list, login to get new token')
                    continue
                else:
                    LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                    LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                    window['LVTable'].update(values=LVTable)
                    window['LVTable2'].update(values=LVTable)               
                    if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                    window['LPList' ].update(values=LPList)
                    LVGroups = ['Group/Folder']
                    for Group in [_[4] for _ in LVTable]:
                        Group = Group.split('\\')[-1]
                        if Group and Group not in LVGroups:
                            LVGroups.append(Group)
                    window['LVsGroup'].update(values=LVGroups)
                    if window['LPLoginInfo']._visible: event = 'LPLoginToggle'
                    continue            

                
        if event in ['LPLogin','LPLoginReturn','LVLogin','LVLogin2']:
            if event != 'LPLoginReturn':
                window['-OUTPUT-'].update(value='')
                window['OutPutSection'].select()
                print('Logging into LastPass')
                LPRefreshRedirect=False
                if Inputs['LPPasscode']:
                    location = window.current_location()
                    if sg.running_mac():
                        pinwinlocation = (location[0]+99, location[1]+307)
                    else:
                        pinwinlocation = (location[0]+74, location[1]+328)
                    Passcode          = pin_popup('Enter MFA Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                    if not Passcode:
                        print('Login Canceled')
                        continue
                    Inputs["LPPass"]  = Inputs["LPPass"] + ',' + Passcode             
                window.perform_long_operation(lambda:LastPassLogin(Inputs), "LPLoginReturn")
                continue
            else:
                LastPass = Inputs['LPLoginReturn']
                if not LastPass.loggedin:
                    window['OutPutSection'].select()
                    print('Login unsuccessful')
                    continue
                else:
                    print('LastPass Successfully Logged In')
                    window['OutPutSection'].select()
                    LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                    window['LVTable'].update(values=LVTable)
                    window['LVTable2'].update(values=LVTable)
                    LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                    if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                    else: window['LPList' ].update(values=LPList)
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
                    if LastPass and LastPass.loggedin:
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
                Combolist = [_[0]+' '*120+str(_) for _ in NewList]
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
                Combolist = [_[0]+' '*120+str(_) for _ in GCSelected]
                for _ in Slots: 
                    window[f"C_{_}"].update(values=Combolist)
                window.fill(values)
                
        if event == 'GCSelect':
            GCTable    = window['GCTable']
            GCTable    = GCTable.Values if isinstance(GCTable,sg.Table) else []
            window['GCSelected'].update(values=GCTable)
            Combolist = [_[0]+' '*120+str(_) for _ in GCTable]
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
                Combolist = [_[0]+' '*120+str(_) for _ in NewList]
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
                Combolist = [_[0]+' '*120+str(_) for _ in MESelected]
                for _ in Slots: 
                    window[f"C_{_}"].update(values=Combolist)
                window.fill(values)
                
        if event == 'MESelect':
            METable    = window['METable']
            METable    = METable.Values if isinstance(METable,sg.Table) else []
            window['MESelected'].update(values=METable)
            Combolist = [_[0]+' '*120+str(_) for _ in METable]
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
                if sg.running_mac():
                    time = 0.1
                else:
                    time = 0
    
                window.perform_long_operation(lambda:EventHelper('OKRefresh',time),'OKRefresh')
        
            if Inputs['Tabs'] == 'LPOptions':
                    if LastPass.loggedin and LastPass.details.lastlogincheck + timedelta(minutes=1) < datetime.now():
                        LastPass.login_check()

                    if LastPass.loggedin:
                        window['LPStatus'].update(value=f'✅ You are logged into LastPass')
                    else:
                        window['LPStatus'].update(value=f'❌  You are not currently logged into LastPass')

        if event == 'OKRefresh':
            try:
                onlykey = OnlyKey(connect=True,tries=1)
                onlykey.read_bytes(timeout_ms=1000)
                OKSlots = onlykey.getlabels()
                onlykey.close()
            except Exception as Error:
                window['-OUTPUT-'].update(value='')
                window.perform_long_operation(lambda:EventHelper('OutPutSection',3),'TabSelect')
                print(Error)
                OKSlots = None
                
            if OKSlots:
                sl = {Slot : [(Slot+" "+_.label if _.label and _.label != 'ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ' else Slot) for _ in OKSlots if _.name == Slot][0]  for Slot in Slots}
            else:
                sl = {Slot : Slot for Slot in Slots}
                
            for Slot in Slots:  window[f'S_{Slot}'].update(text=sl[Slot])
        
        #Toggle Slot Picker Element
        if event in ["PSlotsRadio",'KWSearch']:
            KWTrue = not KWTrue
            window['SlotSelector'].update(visible=not KWTrue)
            window['KWLabelSearchText'].update(visible=KWTrue)


        #Save Current Input Values to Config File
        if event == "Save":
            window['-OUTPUT-'].update(value='')
            window['OutPutSection'].select()
            print('Saving Configuration...')
            config = configparser.ConfigParser()
            ArkInfo = ArkInfo if ArkInfo else {'Token': ""}
            LPInfo  = LPInfo  if LPInfo  else {'Token':"",'Key':"",'SessionId':'','Iteration':''}
            window.perform_long_operation(lambda:SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window),'SaveReturn')
            window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')

        if event == 'AppDataReturn':
            AppData = Inputs['AppDataReturn']
        
        #Clear Print Statements From Output Element
        if event in ['ClearD2','ClearD']:
            window['-OUTPUT-'].update(value='')
            continue
        
#endregion
        
#region Password Sync
        if event in ["Submit",'LoginCheckReturn','FunctionReturn','OKCAFromReturn']:
            window['OutPutSection'].select()
            
            if Inputs['MEFrom'] or Inputs['GCFrom']:
                Selected = window['GSSelected'].Values if Inputs['GCFrom'] else window['MESelected'].Values
                
                if Inputs['LP']:
                        
                    if not Inputs['LPUser'] or not Inputs['LPPass']:
                        if not Inputs["LPUser"]: print('Fill in the LastPass Username field to Login to LastPass')
                        if not Inputs["LPPass"]: print('Fill in the LastPass Password field to Login to LastPass')
                        continue

                    LastPass.login_check()
                        
                    if not LastPass.loggedin:
                        print ('Attempting LastPass login')
                        if Inputs['LPPasscode']:
                            location = window.current_location()
                            if sg.running_mac():
                                pinwinlocation = (location[0]+99, location[1]+307)
                            else:
                                pinwinlocation = (location[0]+74, location[1]+328)
                            Passcode          = pin_popup('Enter MFA Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                            if not Passcode:
                                print('Login Canceled')
                                continue
                            Inputs["LPPass"]  = Inputs["LPPass"] + ',' + Passcode       
                        LastPass = LastPassLogin(Inputs)
                        if not LastPass.loggedin:
                            print('Login unsuccessful')
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
                            else: window['LPList' ].update(values=LPList)
                                 
                if Inputs['OK']:
                    Selected = window['GSSelected'].Values if Inputs['GCFrom'] else window['MESelected'].Values
                    TrueSlots = {_[0].replace("S_","") : _[1] for _ in Inputs.items() if isinstance(_[0],str) and "S_" in _[0] and _[1] == True}
                    for slot in TrueSlots:
                        dropdown = Inputs[f'C_{slot}']
                        if dropdown and len(dropdown) > 120 and (' ' * 120) in dropdown: 
                            Account = literal_eval(dropdown.split(' ' * 120)[1])
                            if not Account:
                                name = dropdown.split(' ' * 120)[0]
                                print(f'Unable to find account {name} with Id {ID} in list of accounts')
                                print ('Unknown error')
                                continue
                                
                            UpdateOnlyKey(Password=Account[3],OK_Keyword='',SlotSelections={slot:True},SlotsTrue=True)
                        else:
                            print(f'No LastPass account is mapped to OnlyKey slot {slot}')                                       

            if Inputs['LPFrom']:
                window['-OUTPUT-'].update(value='')
                LastPass.login_check()  
                if not LastPass.loggedin:
                    if not Inputs['LPUser'] or not Inputs['LPPass']:
                        if not Inputs["LPUser"]: print('Fill in the LastPass Username field to Login to LastPass')
                        if not Inputs["LPPass"]: print('Fill in the LastPass Password field to Login to LastPass')
                        continue  
                    
                    print ('Attempting LastPass login')
                    if Inputs['LPPasscode']:
                        location = window.current_location()
                        if sg.running_mac():
                            pinwinlocation = (location[0]+99, location[1]+307)
                        else:
                            pinwinlocation = (location[0]+74, location[1]+328)
                        Passcode          = pin_popup('Enter MFA Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                        if not Passcode:
                            print('Login Canceled')
                            continue
                        Inputs["LPPass"]  = Inputs["LPPass"] + ',' + Passcode       
                    LastPass = LastPassLogin(Inputs)
                    if not LastPass.loggedin:
                        print('Login unsuccessful')
                        continue
                    LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                    if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                    else: window['LPList' ].update(values=LPList)
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
                        Account = literal_eval(dropdown.split(' ' * 120)[1])
                        if not Account:
                            name = dropdown.split(' ' * 120)[0]
                            print(f'Unable to find account {name} with in list of accounts')
                            print ('Unknown error')
                            continue
                            
                        UpdateOnlyKey(Password=Account.password,OK_Keyword='',SlotSelections={slot:True},SlotsTrue=True)
                    else:
                        print(f'No LastPass account is mapped to OnlyKey slot {slot}')
                            

            if Inputs['CAFrom']:
                if event not in  ['FunctionReturn','LoginCheckReturn'] and not AppData['ArkToken']: 
                    window['-OUTPUT-'].update(value='')
                    if not Inputs["ArkPass"] or not Inputs['CATable'] or not Inputs["ArkUser"] or not Inputs['CAReason'] or Inputs['CAReason'] == CAeReason:   
                        if not Inputs["ArkUser"]: print('Fill in the Username field to Login to CyberArk')
                        if not Inputs["ArkPass"]: print('Fill in the Password field to Login to CyberArk')
                        if not Inputs['CATable']: print('You must select an Account in CyberArk to check-out')
                        if not Inputs['CAReason'] or Inputs['CAReason'] == CAeReason:
                            print('You must enter a business justification to check-out an account')
                        continue
                    CALoginAttempt = 1
                    if Inputs['CAPasscode']:
                        location = window.current_location()
                        if sg.running_mac():
                            pinwinlocation = (location[0]+99, location[1]+307)
                        else:
                            pinwinlocation = (location[0]+74, location[1]+328)
                        Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                        if not Passcode:
                            print('Login Canceled')
                            continue
                        ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                    else:
                        ArkPass  = Inputs["ArkPass"]
                    window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), "FunctionReturn")
                    continue
                else:
                    if event not in ['FunctionReturn','LoginCheckReturn','OKCAFromReturn']: 
                        window['-OUTPUT-'].update(value='')
                        print('Attempting CyberArk CheckOut')
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
                    CATable  = window['CATable'].Values
                    BGUser   = CATable[Inputs['CATable'][0]]
                    if event == 'LoginCheckReturn':
                        LoginCheck = Inputs['LoginCheckReturn']
                        if LoginCheck and 'ErrorMessage' in LoginCheck and CALoginAttempt == 0 and Inputs["ArkUser"] and Inputs["ArkPass"]:
                            CALoginAttempt = 1
                            if Inputs['CAPasscode']:
                                location = window.current_location()
                                if sg.running_mac():
                                    pinwinlocation = (location[0]+99, location[1]+307)
                                else:
                                    pinwinlocation = (location[0]+74, location[1]+328)
                                Passcode = pin_popup('Enter Duo Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                                if not Passcode:
                                    print('Login Canceled')
                                    continue
                                ArkPass  = Inputs["ArkPass"] + ',' + Passcode
                            else:
                                ArkPass  = Inputs["ArkPass"]
                            window.perform_long_operation(lambda:GetCyberArkLogin(ArkUser=Inputs["ArkUser"],ArkPass=ArkPass), "FunctionReturn")
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
                        Platform       = FindCyberArkPlatform(ArkToken,BGUser[14])
                        if ArkAccountPass and isinstance(ArkAccountPass,dict) and 'ErrorMessage' in ArkAccountPass:
                            print(ArkAccountPass['ErrorMessage'])
                            print(f'Unable to retreive password from CyberArk for {BGUser[1]}')
                            continue
                        
                        if not ArkAccountPass:
                            print(f'Unable to retreive password from CyberArk for {BGUser[1]}')
                            continue

                        CACOTable  = window['CACOTable']
                        CACOTable = CACOTable.Values if isinstance(CACOTable,sg.Table) else []
                        CAIndex = Inputs['CATable'][0]
                        CATable[CAIndex][0]  = '⛔'
                        CATable[CAIndex][12] = str(int(datetime.now().timestamp()))
                        if Inputs['CATable'] and CATable and CATable[CAIndex] not in CACOTable:
                            if not CACOTable or not CACOTable[0]:
                                window['CACOTable'].metadata['Searching'] = False
                                CACOTable = [CATable[CAIndex]]
                                window['CACOTable'].update(values=CACOTable)
                            elif CATable[CAIndex][9] not in [_[9] for _ in CACOTable]:
                                CACOTable = CACOTable + [CATable[CAIndex]]
                                window['CACOTable'].update(values=CACOTable)
                        
                        if CACOTable and CACOTable[0] and CATable[CAIndex][9] in [_[9] for _ in CACOTable]:
                            for i,CAAcct in enumerate(CACOTable):
                                if CAAcct[9] == CATable[CAIndex][9]:
                                    CACOTable[i][12] = str(int(datetime.now().timestamp()))
                                    window['CACOTable'].update(values=CACOTable)


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
                        LastPass.login_check()
                        if not LastPass.loggedin:
                            if not Inputs['LPUser'] or not Inputs['LPPass']:
                                if not Inputs["LPUser"]: print('Fill in the LastPass Username field to Login to LastPass')
                                if not Inputs["LPPass"]: print('Fill in the LastPass Password field to Login to LastPass')
                                window.perform_long_operation(lambda:EventHelper('LPOptions',3),'TabSelect') 
                                if not window['LPLoginInfo']._visible: event = window.perform_long_operation(lambda:EventHelper('LPLoginToggle',3.1),'LPLoginToggle') 
                                continue
                            else:
                                print ('Attempting LastPass login')
                                if Inputs['LPPasscode']:
                                    location = window.current_location()
                                    if sg.running_mac():
                                        pinwinlocation = (location[0]+99, location[1]+307)
                                    else:
                                        pinwinlocation = (location[0]+74, location[1]+328)
                                    Passcode          = pin_popup('Enter MFA Passcode',title='MEDHOST Password Manager',icon=icon,location=pinwinlocation)
                                    if not Passcode:
                                        print('Login Canceled')
                                        continue
                                    Inputs["LPPass"]  = Inputs["LPPass"] + ',' + Passcode       
                                LastPass = LastPassLogin(Inputs)
                                if not LastPass.loggedin:
                                    print('Login unsuccessful')
                                    continue
                                else:
                                    LPInfo = {
                                        'Token'    : LastPass.token,
                                        'Key'      : LastPass.encryption_key,
                                        'SessionId': LastPass.id,
                                        'Iteration': LastPass.key_iteration_count
                                    }

                        LPSelected = window['LPSelected'].Values

                        AutoCheckIn = False if Inputs['CADelay'] == 'Disabled' else True
                        
                        CheckinTime = Platform['Details']['MinValidityPeriod'] if Platform and 'Details' in Platform else ''

                        Hash = {
                            'FacilityName' : BGUser[3],
                            'Address'      : BGUser[2],
                            'Client'       : BGUser[4],
                            'Tool'         : BGUser[6],
                            'CyberArkId'   : BGUser[13],
                            'Reason'       : Inputs['CAReason'],
                            'TicketNumber' : CaseNum,
                            'LastCheckOut' : getdate(datetime.now()),
                            'AutoCheckIn'  : AutoCheckIn,
                            'CheckInDelay' : Inputs['CADelay'],
                            'MinValidity'  : CheckinTime,
                            'CreatedBy'    : 'MEDHOST Password Manager',
                            'Notes'        : ''
                        }
                        
                        notes=f"NoteType:Server\nLanguage:en-US\nHostname:{BGUser[5]}\nUsername:{BGUser[1]}\nPassword:{ArkAccountPass}\nNotes:{Hash}\n"
                        if Inputs['LPExist'] and LPSelected:
                            for Item in LPSelected:
                                Account = [_ for _ in LastPass.accounts if _.id == int(Item.split(' ' * 120)[1])][0]
                                Item = Account
                                if Account.url != "http://sn" or 'NoteType:Server' in Account.notes:
                                    try:
                                        NotesHash = literal_eval(Account.notes.split('\nNotes:')[1].rstrip() if Account.notes and "'CreatedBy'" in Account.notes and "NoteType:Server" in Account.notes else '{}')
                                    except:
                                        NotesHash = {}

                                    if NotesHash:
                                        CreatedBy = NotesHash['CreatedBy']
                                        Notes     = NotesHash['Notes'] if 'Notes' in NotesHash else ''
                                    else:
                                        CreatedBy = 'Other'
                                        Notes     = Account.notes if "\nNotes:{" not in Account.notes else ''
                                        

                                    Hash['CreatedBy'] = CreatedBy
                                    Hash['Notes']     = Notes
                                    notes=f"NoteType:Server\nLanguage:en-US\nHostname:{BGUser[5]}\nUsername:{BGUser[1]}\nPassword:{ArkAccountPass}\nNotes:{Hash}\n"
                                    
                                    LastPass.UpdateAccount(Account.id,password=ArkAccountPass,Account=Account,notes=notes)
                                else:
                                    print('Unable to Update Secure Notes - ' + Item.name)
                        elif BGUser[9] in [_.name for _ in LastPass.accounts]:
                            Account       = [_ for _ in LastPass.accounts if _.name == BGUser[9]][0]
                            try:
                                NotesHash = literal_eval(Account.notes.split('\nNotes:')[1].rstrip() if Account.notes and "'Notes'" in Account.notes and "NoteType:Server" in Account.notes else '{}')
                            except:
                                NotesHash = {}
                                
                            if NotesHash:
                                Notes     = NotesHash['Notes'] if 'Notes'         in NotesHash     else ''
                            else:
                                Notes     = Account.notes      if "\nNotes:{" not in Account.notes else ''
                                
                            Hash['Notes'] = Notes
                            notes=f"NoteType:Server\nLanguage:en-US\nHostname:{BGUser[5]}\nUsername:{BGUser[1]}\nPassword:{ArkAccountPass}\nNotes:{Hash}\n"
                            LastPass.UpdateAccount(Account.id,password=ArkAccountPass,url="http://sn",group='MEDHOST Password Manager\\CyberArk',Account=Account,notes=notes)
                        else:
                            Account = None
                            LastPass.NewAccount(BGUser[9],BGUser[1],ArkAccountPass,'MEDHOST Password Manager\\CyberArk',url="http://sn",notes=notes)
                        if LastPass.accounts:
                            LVTable = [[_.name,_.username,_.url,_.id,_.group,_.notes,_.password] for _ in LastPass.accounts]
                            window['LVTable'].update(values=LVTable)
                            LPList = [(_.name + (" " * 120) + str(_.id)) for _ in LastPass.accounts]
                            if window['LPList'].Disabled: window['LPList' ].update(disabled=False);window['LPList' ].update(values=LPList);window['LPList' ].update(disabled=True)
                            else: window['LPList' ].update(values=LPList)
                            
                            if Account:
                                index = [_[0] for _ in enumerate(LVTable) if _[1][3] == Account.id]
                            else:
                                index = [_[0] for _ in enumerate(LVTable) if _[1][0] == BGUser[9] and _[1][4] == 'MEDHOST Password Manager\\CyberArk']

                            if index:
                                ScrollPosition = index[0] / (len(LVTable) - 1) if LVTable and index and len(LVTable) != 1 else 0
                                window['LVTable'].update(select_rows=index)
                                window['LVTable'].set_vscroll_position(ScrollPosition)
                                window['LocalVault'].select()                           
                   
                    ScrollPosition = CAIndex / (len(CATable) - 1) if CATable and len(CATable) != 1 else 0
                    window['CATable'].update(values=CATable)
                    window['CATable'].update(select_rows=[CAIndex])
                    window['CATable'].set_vscroll_position(ScrollPosition)
                    window['CACase'].update(value='')
                    window['CAReason'].update(value='')
                    SaveToConfig(Inputs,Slots,WorkingDirectory,pin,ArkInfo,LPInfo,window,False)
                    window.perform_long_operation(lambda:GetConfigAppData(WorkingDirectory,pin,Print=False),'AppDataReturn')
                    Sleep = 180 if window['CACOTable'].metadata['Updating'] else 0
                    window.perform_long_operation(lambda:EventHelper('CACOSilentRefresh',Sleep),'CACOSilentRefresh')
                    
                    continue
        # else:
        #     gc.collect()
        #     continue
    except Exception as error:
        print(error)
        continue    
window.close()
#endregion
