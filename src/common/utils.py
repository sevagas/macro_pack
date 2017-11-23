#!/usr/bin/env python
# encoding: utf-8


from random import choice
import string
import logging
from termcolor import colored
import os, sys



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
        return os.path.abspath(__file__)

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
    HTA="HTA"
    UNKNOWN = "Unknown"
        
    @classmethod
    def guessApplicationType(self, documentPath):
        """ Guess MS office application type based on extension """
        result = ""
        extension = os.path.splitext(documentPath)[1]
        if ".xls" == extension:
            result = self.XL97
        elif ".xlsx" == extension or extension == ".xlsm":
            result = self.XL
        elif ".doc" ==  extension:
            result = self.WD97
        elif ".docx" ==  extension or extension == ".docm":
            result = self.WD
        elif ".hta" ==  extension:
            result = self.HTA
        elif ".mpp" ==  extension:
            result = self.MPP
        elif ".ppt" ==  extension:
            result = self.PPT97
        elif ".pptm" ==  extension or extension == ".pptx":
            result = self.PPT
        elif ".vsd" ==  extension:
            result = self.VSD97
        elif ".vsdm" ==  extension or extension == ".vsdx":
            result = self.VSD
        elif ".pub" ==  extension:
            result = self.PUB
        elif ".vba" ==  extension:
            result = self.VBA
        else:
            result = self.UNKNOWN
        return result






