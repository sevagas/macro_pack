#!/usr/bin/env python
# encoding: utf-8
import string

from common import utils
from common.definitions import LOGLEVEL
from common.utils import MSTypes

class MpSession:
    """ Represents the state of the current macro_pack run """
    def __init__(self, workingPath, version, mpType):
        self.workingPath = workingPath
        self.version = version
        self.mpType = mpType
        self.logLevel = LOGLEVEL

        # Attrs depending on getter/setter
        self._outputFilePath = ""
        self._outputFileType = MSTypes.UNKNOWN    

        # regular Attrs
        self.uacBypass = False
        self.obfuscateForm =  False
        self.obfuscateNames =  False
        self.obfuscateStrings =  False
        self.obfuscateDeclares = False

        self.obfOnlyMain = False
        self.doNotObfConst = False
        self.noStringsObfuscation = False
        self.noNamesObfuscation = False
        self.noSpaceStrip = False
        self.ObfReplaceConstants = True

        self.obfuscatedNamesMinLen = 6
        self.obfuscatedNamesMaxLen = 18
        self._obfuscatedNamesCharset = string.ascii_lowercase
        self.nameObfuscationCallback = utils.randomStringBasedOnCharset
        
        self.fileInput = None
        self.startFunction = None
        self.stdinContent = None
        self.template = None
        self.ddeMode = False # attack using Dynamic Data Exchange (DDE) protocol (see https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/)
        self.dosCommand = None
        self.icon = "%windir%\system32\imageres.dll,67" # by default JPG image icon

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


    """
    https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules
    Use the following rules when you name procedures, constants, variables, and arguments in a Visual Basic module:
     - You must use a letter as the first character.
     - You can't use a space, period (.), exclamation mark (!), or the characters @, &, $, # in the name.
     - Name can't exceed 255 characters in length
    """

    @property
    def obfuscatedNamesCharset(self):
        return self._obfuscatedNamesCharset

    @obfuscatedNamesCharset.setter
    def obfuscatedNamesCharset(self, charset):
        if charset == "alpha":
            self._obfuscatedNamesCharset = string.ascii_lowercase
        elif charset == "alphanum":
            self._obfuscatedNamesCharset = string.ascii_lowercase + string.digits
        elif charset == "complete":
            self._obfuscatedNamesCharset = string.ascii_lowercase + string.digits + r"_"
        else:
            self._obfuscatedNamesCharset = charset



