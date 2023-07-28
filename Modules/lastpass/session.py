from datetime import datetime
from . import fetcher
from .parser import ConvertToLPString,codecs
from .account import Account
import requests
from . import vault
from xml.etree import ElementTree as etree

# coding: utf-8
class Session(object):
    
    @classmethod
    def login(cls, username, password, multifactor_password=None, client_id=None, outofband=None,retryid=None,MFAProvider=None):
        key_iteration_count = fetcher.request_iteration_count(username)
        return fetcher.request_login(username, password, key_iteration_count, multifactor_password, client_id, outofband,None,retryid,MFAProvider)

    def login_check(self,passthrough=False):
        check     = fetcher.login_check(self)
        
        loggedin  = True if check['ok_response'] else False
        checkinfo = check['ok_response'] if check['ok_response'] else check['error_response']

        if loggedin: 
            for attr in check['session'].__dict__.items():
                if attr[1]: self.__setattr__(attr[0],attr[1])
        else:
            SessionDetails(lastlogincheck=datetime.now(),**checkinfo)
            self.loggedin = False

        return SessionDetails(loggedin=loggedin,checkinfo=checkinfo) if passthrough else loggedin
    
    def __init__(self, id, key_iteration_count,lptoken,encryption_key,**kwargs):
        self.id                  = id
        self.key_iteration_count = key_iteration_count
        self.token               = lptoken
        self.encryption_key      = encryption_key
        self.loggedin            = True if kwargs else None
        self.accounts            = []
        self.details             = SessionDetails(lastlogincheck=datetime.now(),**kwargs) if kwargs else SessionDetails(lastlogincheck=None)

    def __eq__(self, other):
        return (
            self.id                  == other.id                  and 
            self.key_iteration_count == other.key_iteration_count and 
            self.token               == other.token               and
            self.encryption_key      == other.encryption_key)
        
    def __SetAccount__(self,id,name,group,notes,username,password,url,Account,delete):
        encryption_key = self.encryption_key
        
        if Account:
            id       = Account.id       if id       == "" else id         
            name     = Account.name     if name     == "" else name           
            group    = Account.group    if group    == "" else group                      
            notes    = Account.notes    if notes    == "" else notes            
            username = Account.username if username == "" else username               
            password = Account.password if password == "" else password
            url      = Account.url      if url      == "" else url              
    
        body = dict(
            aid		 = id,
            name	 = ConvertToLPString(name    ,encryption_key),
            grouping = ConvertToLPString(group   ,encryption_key),
            extra	 = ConvertToLPString(notes   ,encryption_key),
            username = ConvertToLPString(username,encryption_key),
            password = ConvertToLPString(password,encryption_key),
            url		 = codecs.encode(url.encode(),'hex_codec').decode(),
            delete   = delete,
            extjs	 = 1,
            token	 = self.token,
            method	 = 'cli'
        )

        response = requests.post(url='https://lastpass.com/show_website.php',data=body,cookies={'PHPSESSID': self.id})
        
        try: 
            self.OpenVault()
        except:
            pass
        
        try:
            LPMsg =_.attrib.get('msg') if (_:=etree.XML(response.text).find('result')) != None else None
        
            if LPMsg == 'accountupdated':
                print(f'LastPass Entry {name} Updated')
            elif LPMsg == 'accountadded':
                print(f'LastPass Entry {name} Added')
            elif LPMsg == 'accountdeleted':
                if name:
                    print(f'LastPass Entry {name} Deleted')
                else:
                    print('LastPass Entry Deleted')
            elif LPMsg:
                print('LastPass: ' + LPMsg)
            else:
                print('Error while parsing response from LastPass')  
            
            return response
        except:
            return print('Error while parsing response from LastPass')
    
    def OpenVault(self,include_shared=False,passthrough=False):
        """Fetches accounts from the server and returns a list of Account objects"""
        self.accounts = fetcher.fetch_accounts(self)
        
        if include_shared: self.GetSharedAccounts()
        
        if passthrough: 
            return self.accounts
        else:
            return self

    def UpdateAccount(self,id,name="",group="",notes="",username="",password="",url="",Account:Account=None):
        #User friendly helper function to set the correct attribs for the __SetAccount__ function
        return self.__SetAccount__(id=id,name=name,username=username,password=password,group=group,notes=notes,url=url,Account=Account,delete=0)

    def NewAccount(self,name,username="",password="",group="",notes="",url="http://"):
        #User friendly helper function to set the correct attribs for the __SetAccount__ function
        return self.__SetAccount__(id=0,name=name,username=username,password=password,group=group,notes=notes,url=url,Account=None,delete=0)
    
    def DeleteAccount(self,id,Account:Account=None):
        #User friendly helper function to set the correct attribs for the __SetAccount__ function
        return self.__SetAccount__(id=id,name="",group="",notes="",username="",password="",url="",Account=Account,delete=1)

    def GetSharedAccounts(self):
        blob  = fetcher.fetch(self)
        Vault = vault.Vault(blob,self.encryption_key,output_strings=True,shar_only=True)
        self.accounts = self.accounts + Vault.accounts
        return Vault.accounts
    
    
class SessionDetails(object):
    def __init__(self,**kwargs):
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])