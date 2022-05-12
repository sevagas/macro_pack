#!/usr/bin/env python
# encoding: utf-8


import contextlib
from random import choice
import string
import logging
from termcolor import colored
import os, sys
import socket
from collections import OrderedDict
import importlib.util
import psutil
from datetime import datetime


VBAMAXLINELEN = 400 # max char for a vba line
VBAMAXNBLINE = 100 # Max nm line in a vba method

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
    return ''.join(choice(string.ascii_lowercase) for _ in range(length))


def randomStringBasedOnCharset(length, charset):
    """ Returns a random alphabetic string of length 'length' """
    key = choice('aaaabbcddeeeeeffgghhiiiijkllmmnnnoooppqrrrrsstttuvwy')  # Name has to start with a letter
    for _ in range(length):
        key += choice(charset)
    return key


def extractStringsFromText(text):      
    import re
    result = ""
    if '"' in text:
        matches=re.findall(r'\"(.+?)\"',text)
        # matches is now ['String 1', 'String 2', 'String3']
        result = ",".join(matches)  
    return result


def extractWordInString(strToParse, index):
    """ Extract word (space separated ) at current index"""
    i = index
    while i!=0 and strToParse[i-1] not in " \t\n&|":
        i = i-1
    leftPart = strToParse[i:index]
    i = index
    while i!=len(strToParse) and strToParse[i] not in " \t\n&|":
        i = i+1
    rightPart = strToParse[index:i]
    extractedWord = leftPart+rightPart
    #logging.debug("     [-] extracted Word: %s" % extractedWord)
    return extractedWord


def extractPreviousWordInString(strToParse, index):
    """ Extract the word (space separated ) preceding the one at current index"""
    # Look for beginning or word
    i = index
    if strToParse[i] not in " \t\n":
        while i!=0 and strToParse[i-1] not in " \t\n&|":
            i = i-1
    if i > 2:
        while i!=0 and strToParse[i-1] in " \t\n\",;": # Skip spaces nd special char before previous word
            i = i-1
    previousWord = extractWordInString(strToParse, i) if i > 2 else ""
    logging.debug("     [-] extracted previous Word: %s" % previousWord)
    return previousWord


def extractNextWordInString(strToParse, index):
    """ Extract the word (space separated) following the one at current index"""
    # Look for beginning or word
    i = index
    while i!=len(strToParse) and strToParse[i] not in " \t\n&|":
        i = i+1
    if len(strToParse)-i > 2:
        while i!=0 and strToParse[i] in " \t\n\",;": # Skip spaces nd special char befor previous word
            i = i+1
    if len(strToParse)-i > 2:
        nextWord = extractWordInString(strToParse, i)
    else:
        nextWord = ""
    logging.debug("     [-] Extracted next Word: %s" % nextWord)
    return nextWord


def getHostIp():
    """ return current facing IP address """
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # doesn't have to be reachable
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP


def getRunningApp():
    if getattr(sys, 'frozen', False):
        return sys.executable
    import __main__ as main # @UnresolvedImport To get the real origin of the script not the location of current file
    return os.path.abspath(main.__file__)

def randomAlphaWithSeed(length, seed):
    """ Returns a random alphabetic string of length 'length' """
    key = ''
    cpt = 0
    for i in range(length): # @UnusedVariable
        if i in [0, 2, 4]:
            key += seed[cpt]
            cpt +=1
        else:
            key += choice(string.ascii_lowercase)
    return key

def checkIfProcessRunning(processName):
    """
    Check if there is any running process that contains the given name processName.
    """
    #Iterate over the all the running process
    for proc in psutil.process_iter():
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Check if process name contains the given name string.
            if processName.lower() in proc.name().lower():
                return True
    return False




def yesOrNo(question):
    answer = input(question + "(y/n): ").lower().strip()
    print("")
    while answer not in ["y", "yes", "n", "no"]:
        print("Input yes or no")
        answer = input(f"{question}(y/n):").lower().strip()
        print("")
    return answer[0] == "y"

   
def forceProcessKill(processName):
    """
    Force kill a process (only work on windows)
    """
    os.system("taskkill /f /im  %s >nul 2>&1" % processName)

  
def checkModuleExist(name):
    r"""Returns if a top-level module with :attr:`name` exists *without**
    importing it. This is generally safer than try-catch block around a
    `import X`. It avoids third party libraries breaking assumptions of some of
    our tests, e.g., setting multiprocessing start method when imported
    (see librosa/#747, torchvision/#544).
    """
    spec = importlib.util.find_spec(name)
    return spec is not None 


def validateDate(date_text):
    try:
        if date_text != datetime.strptime(date_text, "%Y-%m-%d").strftime('%Y-%m-%d'):
            raise ValueError
        return True
    except ValueError:
        return False


class MPParam():
    def __init__(self,name,optional=False):
        self.name = name
        self.value = ""
        self.optional = optional


def getParamValue(paramArray, paramName):
    result = ""
    i = 0
    while i < len(paramArray):
        if paramArray[i].name == paramName:
            result = paramArray[i].value
            break
        i += 1
    return result


def progressBar(iterable, prefix='', suffix='', decimals=1, length=100, fill='█', printEnd="\r", disableProgressBar=False):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    if not disableProgressBar:
        total = len(iterable)

        # Progress Bar Printing Function
        def printProgressBar(iteration):
            percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
            filledLength = int(length * iteration // total)
            bar = fill * filledLength + '-' * (length - filledLength)
            print(f'\r{prefix} |{bar}| {percent}% {suffix}', end=printEnd)
        # Initial Call
        printProgressBar(0)
        # Update Progress Bar
        for i, item in enumerate(iterable):
            yield item
            printProgressBar(i + 1)
        # Print New Line on Complete
        print()
    else:
        for i, item in enumerate(iterable):
            yield item
            

textchars = bytearray({7,8,9,10,12,13,27} | set(range(0x20, 0x100)) - {0x7f}) # https://stackoverflow.com/questions/898669/how-can-i-detect-if-a-file-is-binary-non-text-in-python
isBinaryString = lambda bytes: bool(bytes.translate(None, textchars))


class MSTypes:
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
    ACC="Access"
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
    IQY="Excel Web Query"
    SETTINGS_MS="Settings Shortcut"
    SYLK="SYmbolic LinK"
    CHM="Compressed HTML Help"
    LIBRARY_MS="MS Library"
    INF="Setup Information"
    CSPROJ="Visual Studio Project"
    CMD="Command line"
    EXE="Portable Executable"
    DLL="Portable Executable (DLL)"
    MSI="Installer"
    UNKNOWN = "Unknown"

    WORD_AND_EXCEL_FORMATS = [XL, XL97, WD, WD97]
    MS_OFFICE_BASIC_FORMATS =  WORD_AND_EXCEL_FORMATS + [PPT] 
    MS_OFFICE_FORMATS = MS_OFFICE_BASIC_FORMATS + [MPP, VSD, VSD97, ACC] # Formats supported by macro_pack
    VBSCRIPTS_BASIC_FORMATS = [VBS, HTA, SCT, WSF]
    VBSCRIPTS_FORMATS = VBSCRIPTS_BASIC_FORMATS + [XSL]
    VB_FORMATS = VBSCRIPTS_FORMATS + MS_OFFICE_FORMATS
    VB_FORMATS_EXT = VB_FORMATS + [VBA] # VBA format is non executable
    
    Shortcut_FORMATS = [LNK, GLK, SCF, URL, SETTINGS_MS, LIBRARY_MS, INF, IQY, SYLK, CHM, CMD, CSPROJ]
    
    ProMode_FORMATS =  [SYLK, CHM]
    HtaMacro_FORMATS = [LNK, CHM, INF, SYLK, CSPROJ]
    Trojan_FORMATS = MS_OFFICE_BASIC_FORMATS + [MPP, VSD, VSD97,CHM, CSPROJ, LNK, HTA]
    PE_FORMATS = [EXE, DLL]

    # OrderedDict([("target_url",None),("download_path",None)])
    EXTENSION_DICT = OrderedDict([(LNK,".lnk"),(GLK,".glk"),(SCF,".scf"),(URL,".url"), (SETTINGS_MS,".SettingContent-ms"),(LIBRARY_MS,".library-ms"),(INF,".inf"),(IQY, ".iqy"),
                                  (SYLK,".slk"),(CHM,".chm"),(CMD,".cmd"),(CSPROJ,".csproj"),
                                  (XL,".xlsm"),(XL97,".xls"),(WD,".docm"),
                                  (WD97,".doc"),(PPT,".pptm"),(PPT97,".ppt"),(MPP,".mpp"),( PUB,".pub"),( VSD,".vsdm"),(VSD97,".vsd"),
                                  (VBA,".vba"),(VBS,".vbs"),(HTA,".hta"),(SCT,".sct"),(WSF,".wsf"),(XSL,".xsl"),(ACC,".accdb"), (ACC,".mdb"),
                                   (EXE,".exe"),(DLL,".dll"),(MSI,".msi")])



    @classmethod
    def guessApplicationType(cls, documentPath):
        """ Guess MS application type based on extension """
        result = ""
        extension = os.path.splitext(documentPath)[1]
        if extension.lower() == ".xls":
            return cls.XL97
        elif extension.lower() in (".xlsx", ".xlsm", ".xltm"):
            return cls.XL
        elif extension.lower() == ".doc":
            return cls.WD97
        elif extension.lower() in (".docx", ".docm", ".dotm"):
            return cls.WD
        elif extension.lower() == ".hta":
            return cls.HTA
        elif extension.lower() == ".mpp":
            return cls.MPP
        elif extension.lower() == ".ppt":
            return cls.PPT97
        elif extension.lower() in (".pptx", ".pptm", ".potm"):
            return cls.PPT
        elif extension.lower() == ".vsd":
            return cls.VSD97
        elif extension.lower() in [".vsdm", ".vsdx"]:
            return cls.VSD
        elif extension.lower() in (".accdb", ".accde", ".mdb"):
            return cls.ACC
        elif extension.lower() == ".pub":
            return cls.PUB
        elif extension.lower() == ".vba":
            return cls.VBA
        elif extension.lower() == ".vbs":
            return cls.VBS
        elif extension.lower() in [".sct", ".wsc"]:
            return cls.SCT
        elif extension.lower() == ".wsf":
            return cls.WSF
        elif extension.lower() == ".url":
            return cls.URL
        elif extension.lower() == ".glk":
            return cls.GLK
        elif extension.lower() == ".lnk":
            return cls.LNK
        elif extension.lower() == ".settingcontent-ms":
            return cls.SETTINGS_MS
        elif extension.lower() == ".library-ms":
            return cls.LIBRARY_MS
        elif extension.lower() == ".inf":
            return cls.INF
        elif extension.lower() == ".scf":
            return cls.SCF
        elif extension.lower() == ".xsl":
            return cls.XSL
        elif extension.lower() == ".iqy":
            return cls.IQY
        elif extension.lower() == ".slk":
            return cls.SYLK
        elif extension.lower() == ".chm":
            return cls.CHM
        elif extension.lower() == ".csproj":
            return cls.CSPROJ
        elif extension.lower() in [".cmd", ".bat"]:
            return cls.CMD
        elif extension.lower() in (".dll", ".ocx"):
            return cls.DLL
        elif extension.lower() in (".exe"):
            return cls.EXE
        elif extension.lower() in (".msi"):
            return cls.MSI
        else:
            return cls.UNKNOWN
    

