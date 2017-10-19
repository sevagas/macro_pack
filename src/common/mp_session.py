#!/usr/bin/env python
# encoding: utf-8


class MpSession():
    """ Represents the state of the current macro_pack run """
    def __init__(self, workingPath, version, mpType):
        self.workingPath = workingPath
        self.version = version
        self.mpType = mpType
        
        self.vbomEncode = False
        self.avBypass = False
        self.obfuscateForm =  False  
        self.obfuscateNames =  False 
        self.obfuscateStrings =  False 
        self.persist = False
        self.keepAlive = False
        self.trojan = False
        self.stealth = False
        self.vbaInput = None
        self.startFunction = None
        self.fileOutput = False
        self.excelFilePath = None   
        self.excel97FilePath = None   
        self.wordFilePath = None 
        self.word97FilePath = None
        self.pptFilePath = None
        self.vbaFilePath = None
        self.stdinContent = None
        self.template = None
        self.ddeMode = False # attack using Dynamic Data Exchange (DDE) protocol (see https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/)
        self.dcom = False
        self.dcomTarget = None