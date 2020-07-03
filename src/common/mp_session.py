#!/usr/bin/env python
# encoding: utf-8

from common.utils import MSTypes

class MpSession:
    """ Represents the state of the current macro_pack run """
    def __init__(self, workingPath, version, mpType):
        self.workingPath = workingPath
        self.version = version
        self.mpType = mpType

        # Attrs depending on getter/setter
        self._outputFilePath = ""
        self._outputFileType = MSTypes.UNKNOWN    

        # regular Attrs
        self.uacBypass  = False
        self.obfuscateForm =  False
        self.obfuscateNames =  False
        self.obfuscateStrings =  False
        self.obfOnlyMain = False
        self.doNotObfConst = False
        self.ObfReplaceConstants = True
        
        self.fileInput = None
        self.startFunction = None
        self.stdinContent = None
        self.template = None
        self.ddeMode = False # attack using Dynamic Data Exchange (DDE) protocol (see https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/)
        self.dosCommand = None
        self.icon = "%windir%\system32\imageres.dll,67" # by default JPG immage icon

       
        self.runTarget = None
        self.runVisible = False
        self.forceYes = False
        self.printFile = False
        self.unicodeRtlo = None

        self.listen = False
        self.listenPort = 80
        self.listenRoot = "."
        self.embeddedFilePath = None
        
        self.isTrojanMode = False
        self.htaMacro = False

        self.Wlisten = False
        self.WRoot = "."
        
        self.vbModulesList = []

    @property
    def outputFileType(self):
        return self._outputFileType

    @property
    def outputFilePath(self):
        return self._outputFilePath

    @outputFilePath.setter
    def outputFilePath(self, outputFilePath):
        self._outputFilePath = outputFilePath
        self._outputFileType = MSTypes.guessApplicationType(self._outputFilePath)
