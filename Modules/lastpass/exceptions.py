# coding: utf-8


class Error(Exception):
    """Base class for all errors, should not be raised"""
    pass


#
# Generic errors
#

class NetworkError(Error):
    """Something went wrong with the network"""
    pass


class InvalidResponseError(Error):
    """Server responded with something we don't understand"""
    pass


class UnknownResponseSchemaError(Error):
    """Server responded with XML we don't understand"""
    pass

class MFADetails(object):
    def __init__(self,**kwargs):
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

#
# LastPass returned errors
#

class LastPassUnknownUsernameError(Error):
    """LastPass error: unknown username"""
    def __init__(self,ErrorMessage,**kwargs):
        self.ErrorMessage = ErrorMessage
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])


class LastPassInvalidPasswordError(Error):
    """LastPass error: invalid password"""
    def __init__(self,ErrorMessage,**kwargs):
        self.ErrorMessage = ErrorMessage
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])


class LastPassIncorrectGoogleAuthenticatorCodeError(Error):
    """LastPass error: missing or incorrect Google Authenticator code"""
    def __init__(self,ErrorMessage,MFAInfo=None,**kwargs):
        self.ErrorMessage = ErrorMessage
        self.MFAInfo      = MFADetails(**MFAInfo)
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

class LastPassIncorrectMicrosoftAuthenticatorCodeError(Error):
    """LastPass error: missing or incorrect Microsoft Authenticator code"""
    def __init__(self,ErrorMessage,MFAInfo=None,**kwargs):
        self.ErrorMessage = ErrorMessage
        self.MFAInfo      = MFAInfo
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

class LastPassMFAError(Error):
    """LastPass error: Multifactor Authentication is required but outofband was set to 0 and no MFA Code was provided"""
    def __init__(self,ErrorMessage,MFAInfo=None,**kwargs):
        self.ErrorMessage = ErrorMessage
        self.MFAInfo      = MFADetails(**MFAInfo)
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

class LastPassIncorrectMFACodeError(Error):
    """LastPass error: incorrect Multifactor Authentication code"""
    def __init__(self,ErrorMessage,MFAInfo=None,**kwargs):
        self.ErrorMessage = ErrorMessage
        self.MFAInfo      = MFADetails(**MFAInfo)
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

class LastPassMissingMFACodeError(Error):
    """LastPass error: missing Multifactor Authentication code"""
    def __init__(self,ErrorMessage,MFAInfo=None,**kwargs):
        self.ErrorMessage = ErrorMessage
        self.MFAInfo      = MFADetails(**MFAInfo)
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

class LastPassMFAPushError(Error):
    """LastPass error: Multifactor Authentication Push Timed Out or Was Rejected"""
    def __init__(self,ErrorMessage,MFAInfo=None,**kwargs):
        self.ErrorMessage = ErrorMessage
        self.MFAInfo      = MFADetails(**MFAInfo)
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

class LastPassIncorrectYubikeyPasswordError(Error):
    """LastPass error: missing or incorrect Yubikey password"""
    def __init__(self,ErrorMessage,MFAInfo=None,**kwargs):
        self.ErrorMessage = ErrorMessage
        self.MFAInfo      = MFADetails(**MFAInfo)
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])

class LastPassUnknownError(Error):
    """LastPass error we don't know about"""
    def __init__(self,ErrorMessage,**kwargs):
        self.ErrorMessage = ErrorMessage
        for arg in kwargs.items():
            setattr(self,arg[0],arg[1])
