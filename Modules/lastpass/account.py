# coding: utf-8
class Account(object):
    def __init__(self, **kwargs):
        self.id       = kwargs['id']
        self.name     = kwargs['name']
        self.username = kwargs['username']
        self.password = kwargs['password']
        self.url      = kwargs['url']
        self.group    = kwargs['group']
        self.notes    = kwargs['notes']
        
        for arg in kwargs.items():
            if arg[0] not in ['id','name','username','password','url','group','notes']:
                self.__setattr__(arg[0],arg[1])