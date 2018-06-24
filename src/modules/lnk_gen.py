#!/usr/bin/env python
# encoding: utf-8
import sys
import logging
from modules.mp_generator import Generator
from collections import OrderedDict
if sys.platform == "win32":
    from win32com.client import Dispatch  # @UnresolvedImport


class LNKGenerator(Generator):
    """ Module used to generate malicious Explorer Command File"""
    
    def check(self):
        if sys.platform != "win32":
            logging.error("  [!] You have to run on Windows OS to build this file format.")
            return False
        else:    
            return True
        
    def buildLnkWithWscript(self, target, targetArgs=None, iconPath=None, workingDirectory = ""):
        """ Build an lnk shortcut using WScript wrapper """
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(self.outputFilePath)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = workingDirectory
        if targetArgs:
            shortcut.Arguments = targetArgs
        if iconPath:
            shortcut.IconLocation = iconPath
        shortcut.save()
        
    
    def generate(self):
        """ Generate LNK file """
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        paramDict = OrderedDict([("Shortcut_Target",None), ("Shortcut_Icon",None) ]) # ("Work_Directory",None)      
        self.fillInputParams(paramDict)
        
        # Get needed parameters
        iconPath = paramDict["Shortcut_Icon"]
        #workingDirectory = paramDict["Work_Directory"]
        # Extract shortcut arguments
        CmdLine = paramDict["Shortcut_Target"].split(' ', 1)
        target = CmdLine[0]
        if len(CmdLine) == 2:
            targetArgs = CmdLine[1]
        else:
            targetArgs = None
        # Create lnk file
        self.buildLnkWithWscript(target, targetArgs, iconPath) # ("Work_Directory",None)
        
        logging.info("   [-] Generated %s file: %s" % (self.outputFileType, self.outputFilePath))
        logging.info("   [-] Test with: \nBrowse %s dir to trigger icon resolution. Click on file to trigger shortcut.\n" % self.outputFilePath)
        

        
        
        