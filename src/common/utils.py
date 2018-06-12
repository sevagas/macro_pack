#!/usr/bin/env python
# encoding: utf-8


from random import choice
import string
import logging
from termcolor import colored
import os, sys
from collections import OrderedDict



class ColorLogFiler(logging.StreamHandler):
    """ Override logging class to enable terminal colors """
    def emit(self, record):
        try:
            msg = self.format(record)
            msg = msg.replace("[+]",colored("[+]", "green"))
            msg = msg.replace("[-]",colored("[-]", "green"))
            msg = msg.replace("[!]",colored("[!]", "red"))
            stream = self.stream
            stream.write(msg)
            stream.write(self.terminator)
            self.flush()
        except Exception:
            self.handleError(record)

def randomAlpha(length):
    """ Returns a random alphabetic string of length 'length' """
    key = ''
    for i in range(length): # @UnusedVariable
        key += choice(string.ascii_lowercase)
    return key


def getRunningApp():
    if getattr(sys, 'frozen', False):
        return sys.executable
    else:
        import __main__ as main # @UnresolvedImport To get the real origin of the script not the location of current file 
        return os.path.abspath(main.__file__)

class MSTypes():
    XL="Excel"
    XL97="Excel97"
    WD="Word"
    WD97="Word97"
    PPT="PowerPoint"
    PPT97="PowerPoint97"
    MPP = "MSProject"
    PUB="Publisher"
    VSD="Visio"
    VSD97="Visio97"
    VBA="VBA"
    VBS="Visual Basic Script"
    HTA="HTML Application"
    SCT="Windows Script Component"
    WSF="Windows Script File"
    LNK="Shell Link"
    GLK = "Groove Shortcut"
    SCF="Explorer Command File"
    XSL="XSLT Stylesheet"
    URL="URL Shortcut"
    SETTINGS_MS="Settings Shortcut"
    UNKNOWN = "Unknown"
    
    MS_OFFICE_FORMATS = [ XL, XL97, WD, WD97, PPT, MPP, VSD, VSD97] # Formats supported by macro_pack
    VB_FORMATS = [VBA, VBS, HTA, SCT, WSF, XSL ]
    VB_FORMATS.extend(MS_OFFICE_FORMATS)
    Shortcut_FORMATS = [LNK, GLK, SCF, URL, SETTINGS_MS]
    
    # OrderedDict([("target_url",None),("download_path",None)]) 
    EXTENSION_DICT = OrderedDict([ (LNK,".lnk"),( GLK,".glk"),( SCF,".scf"),( URL,".url"), (SETTINGS_MS,".SettingContent-ms"),( XL,".xlsm"),( XL97,".xls"),( WD,".docm"),
                                  (WD97,".doc"),( PPT,".pptm"),( PPT97,".ppt"),( MPP,".mpp"),( PUB,".pub"),( VSD,".vsdm"),( VSD97,".vsd"),
                                  (VBA,".vba"),( VBS,".vbs"),( HTA,".hta"),( SCT,".wsc"),( WSF,".wsf"),( XSL,".xsl") ]) 
    
    #EXTENSION_DICT = { LNK:".lnk", GLK:".glk", SCF:".scf", URL:".url", XL:".xlsm", XL97:".xls", WD:".docm",
    #                   WD97:".doc", PPT:".pptm", PPT97:".ppt", MPP:".mpp", PUB:".pub", VSD:".vsdm", VSD97:".vsd", 
    #                   VBA:".vba", VBS:".vbs", HTA:".hta", SCT:".wsc", WSF:".wsf", XSL:".xsl" }
    
    @classmethod
    def guessApplicationType(self, documentPath):
        """ Guess MS application type based on extension """
        result = ""
        extension = os.path.splitext(documentPath)[1]
        if ".xls" == extension.lower():
            result = self.XL97
        elif ".xlsx" == extension.lower() or extension.lower() == ".xlsm":
            result = self.XL
        elif ".doc" ==  extension.lower():
            result = self.WD97
        elif ".docx" ==  extension.lower() or extension.lower() == ".docm":
            result = self.WD
        elif ".hta" ==  extension.lower():
            result = self.HTA
        elif ".mpp" ==  extension.lower():
            result = self.MPP
        elif ".ppt" ==  extension.lower():
            result = self.PPT97
        elif ".pptm" ==  extension.lower() or extension.lower() == ".pptx":
            result = self.PPT
        elif ".vsd" ==  extension.lower():
            result = self.VSD97
        elif ".vsdm" ==  extension.lower() or extension.lower() == ".vsdx":
            result = self.VSD
        elif ".pub" ==  extension.lower():
            result = self.PUB
        elif ".vba" ==  extension.lower():
            result = self.VBA
        elif ".vbs" ==  extension.lower():
            result = self.VBS
        elif ".sct" ==  extension.lower() or extension.lower() == ".wsc":
            result = self.SCT
        elif ".wsf" == extension.lower():
            result = self.WSF
        elif ".url" ==  extension.lower():
            result = self.URL
        elif ".glk" ==  extension.lower():
            result = self.GLK    
        elif ".lnk" ==  extension.lower():
            result = self.LNK
        elif ".settingcontent-ms" == extension.lower():
            result = self.SETTINGS_MS
        elif ".scf" ==  extension.lower():
            result = self.SCF
        elif ".xsl" ==  extension.lower():
            result = self.XSL
        else:
            result = self.UNKNOWN
        return result


    
        





