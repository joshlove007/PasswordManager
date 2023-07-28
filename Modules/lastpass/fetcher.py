# coding: utf-8
from . import parser
from time import sleep
import hashlib
from base64 import b64decode
from binascii import hexlify
import requests
from xml.etree import ElementTree as etree
from . import blob
from .version import __version__
from .exceptions import *
from .session import Session
from typing import Literal

headers = {'user-agent': 'lastpass-python/{}'.format(__version__)}

def login(username, password, multifactor_password=None, client_id=None, outofband=0):
    key_iteration_count = request_iteration_count(username)
    return request_login(username, password, key_iteration_count, multifactor_password, client_id, outofband)

def logout(session):
    response = requests.get(
        'https://lastpass.com/logout.php',
        cookies={'PHPSESSID': session.id}
    )

    if response.status_code != requests.codes.ok:
        raise NetworkError()

def fetch(session):
    body = {
        "requestsrc" : 'cli',
        "mobile"     : '1',
        "hasplugin"  : '1.3.3',
        'b64'        : '1'
    }
     
    response = requests.get('https://lastpass.com/getaccts.php',cookies={'PHPSESSID': session.id},params=body)

    if response.status_code != requests.codes.ok:
        tries = 1
        while response.status_code != requests.codes.ok and tries < 4:
            sleep(5)
            response = requests.get('https://lastpass.com/getaccts.php',cookies={'PHPSESSID': session.id},params=body)
            tries = tries + 1
        if response.status_code != requests.codes.ok:
            raise NetworkError()

    return blob.Blob(decode_blob(response.content), session.key_iteration_count)

def fetch_accounts(session):
    response = requests.get('https://lastpass.com/getaccts.php',cookies={'PHPSESSID': session.id})
    
    if response.status_code != requests.codes.ok:
        tries = 1
        while response.status_code != requests.codes.ok and tries < 4:
            sleep(2)
            response = requests.get('https://lastpass.com/getaccts.php',cookies={'PHPSESSID': session.id})
            tries = tries + 1
        if response.status_code != requests.codes.ok:
            raise NetworkError()
    
    parsed_response = etree.XML(response.text)

    return parser.decode_accounts(parsed_response,session)

def request_iteration_count(username):
    response = requests.post('https://lastpass.com/iterations.php',
                               data={'email': username},
                               headers=headers)
    if response.status_code != requests.codes.ok:
        raise NetworkError()

    try:
        count = int(response.content)
    except:
        raise InvalidResponseError('Key iteration count is invalid')

    if count > 0:
        return count
    raise InvalidResponseError('Key iteration count is not positive')

def request_login(username, password, key_iteration_count, multifactor_password=None, client_id=None, outofband:Literal[0,1]=None,retry=None,retryid=None,MFAProvider=None):
    body = {
        'method'               : 'cli',
        'xml'                  : 2,
        'username'             : username,
        'password'             : password,
        'hash'                 : make_hash(username, password, key_iteration_count),
        'iterations'           : key_iteration_count,
        "outofbandsupported"   : 1,
        "includeprivatekeyenc" : 1
    }

    if outofband is not None:
        body["outofbandrequest"] = outofband

    if multifactor_password:
        body['otp'] = multifactor_password

    if retry and 'retryid' in retry:
        body['outofbandretry']   = 1
        body['outofbandretryid'] = retry['retryid'] if not retryid else retryid

    if MFAProvider:
        body["provider"] = MFAProvider

    if client_id:
        body['imei'] = client_id
        
    response = requests.post('https://lastpass.com/login.php',data=body,headers=headers)
    
    if response.status_code != requests.codes.ok:
        reason = response.reason if hasattr(response,'reason') else 'Error Connecting to LastPass'
        raise NetworkError(reason)
   
    try:
        parsed_response = etree.fromstring(response.content)
    except etree.ParseError:
        raise InvalidResponseError()
    
    encryption_key = make_key(username, password, key_iteration_count)

    session = create_session(parsed_response, encryption_key)
    if not session:
        return login_error(parsed_response,body,retry)
    elif outofband == 1:
        print("Multifactor authentication completed")

    return session

def login_check(session):
    """checks if session id is still valid"""
    response = requests.get('https://lastpass.com/login_check.php',cookies={'PHPSESSID': session.id})

    if response.status_code != requests.codes.ok:
        pass
        #raise NetworkError()

    try:
        parsed_response = etree.XML(response.content)
    except etree.ParseError:
        parsed_response = etree.XML('<error message="Error while parsing response from LastPass" />')
        #raise InvalidResponseError()
        
    error = _.attrib if (_:=parsed_response.find('error')) != None else {}
    ok    = _.attrib if (_:=parsed_response.find('ok'))    != None else {}
   
    session = create_session(parsed_response, session.encryption_key)

    return {"session":session,'ok_response':ok,'error_response':error}

def create_session(parsed_response, encryption_key):
    if parsed_response.tag != 'ok':
        try:
            lp_info = parsed_response.find("ok").attrib
        except:
            lp_info = None
    else:
        lp_info = parsed_response

    if lp_info and isinstance(lp_info['sessionid'], str):
        return Session(lp_info['sessionid'],int(lp_info['iterations']),lp_info['token'],encryption_key,**lp_info)

def login_error(parsed_response,body,retry):
    error = None if parsed_response.tag != 'response' else parsed_response.find('error')
    
    if error is None or len(error.attrib) == 0:
        raise UnknownResponseSchemaError()

    exceptions = {
        "unknownemail"              : LastPassUnknownUsernameError,
        "unknownpassword"           : LastPassInvalidPasswordError,
        "googleauthrequired"        : LastPassIncorrectGoogleAuthenticatorCodeError,
        "googleauthfailed"          : LastPassIncorrectGoogleAuthenticatorCodeError,
        "yubikeyrestricted"         : LastPassIncorrectYubikeyPasswordError,
        "microsoftauthrequired"     : LastPassIncorrectMicrosoftAuthenticatorCodeError,
        "microsoftauthfailed"       : LastPassIncorrectMicrosoftAuthenticatorCodeError,
    }

    cause        = error.attrib.get('cause')
    message      = error.attrib.get('message')
    capabilities = error.attrib.get('capabilities')
    retryid      = error.attrib.get('retryid')
    oobrequest   = body.get('outofbandrequest')

    if cause == 'outofbandrequired':
        mfa = error.attrib.get('outofbandname')
        if oobrequest == 0:
            raise LastPassMFAError('Multifactor authentication required!',MFAInfo=error.attrib)
        elif retryid or capabilities and 'push' in capabilities:
            if mfa:
                print(f"Complete multifactor authentication through {mfa}")
                retry={'firsttry':error.attrib}
            else:
                print("MFA request timed out, retrying...")
                retry={**error.attrib,'firsttry':retry['firsttry']}
                
            return request_login(body['username'],body['password'],body['iterations'],outofband=1,retry=retry)
        elif capabilities and 'passcode' in capabilities:
            raise LastPassMissingMFACodeError('Multifactor authentication required!',MFAInfo=error.attrib)
    elif cause == 'multifactorresponsefailed':
        if retry and 'push' in retry['firsttry']['capabilities']:
            raise LastPassMFAPushError(message,MFAInfo=retry)
        elif 'otp' in body:
            raise LastPassIncorrectMFACodeError(message,MFAInfo=error.attrib)
    elif cause:
        raise exceptions.get(cause, LastPassUnknownError)(message or cause,MFAInfo=error.attrib)
    
    raise InvalidResponseError(message)


def decode_blob(blob):
    return b64decode(blob)

def make_key(username, password, key_iteration_count):
    # type: (str, str, int) -> bytes
    if key_iteration_count == 1:
        return hashlib.sha256(username.encode('utf-8') + password.encode('utf-8')).digest()
    else:
        return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), username.encode('utf-8'), key_iteration_count, 32)

def make_hash(username, password, key_iteration_count):
    # type: (str, str, int) -> bytes
    if key_iteration_count == 1:
        return bytearray(hashlib.sha256(hexlify(make_key(username, password, 1)) + password.encode('utf-8')).hexdigest(), 'ascii')
    else:
        return hexlify(hashlib.pbkdf2_hmac(
            'sha256',
            make_key(username, password, key_iteration_count),
            password.encode('utf-8'),
            1,
            32
        ))
