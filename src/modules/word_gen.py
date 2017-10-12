#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport
    import winreg # @UnresolvedImport

import logging
from modules.mp_module import MpModule


class WordGenerator(MpModule):
    """ Module used to generate MS Word file from working dir content"""
    
    def enableVbom(self):
        # Enable writing in macro (VBOM)
        # First fetch the application version
        objWord = win32com.client.Dispatch("Word.Application")
        objWord.Visible = False # do the operation in background 
        self.version = objWord.Application.Version
        # IT is necessary to exit office or value wont be saved
        objWord.Application.Quit()
        del objWord
        # Next change/set AccessVBOM registry value to 1
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\Word\\Security"
        logging.info("   [-] Set %s to 1..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,1) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    
    def disableVbom(self):
        # Disable writing in VBA project
        #  Change/set AccessVBOM registry value to 0
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\Word\\Security"
        logging.info("   [-] Set %s to 0..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,0) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    
    def run(self):
        logging.info(" [+] Generating MS Word document...")
        
        self.enableVbom()

        logging.info("   [-] Open document...")
        # open up an instance of Word with the win32com driver
        word = win32com.client.Dispatch("Word.Application")
        # do the operation in background without actually opening Excel
        word.Visible = False
        document = word.Documents.Add()

        logging.info("   [-] Inject VBA...")
        # Read generated files
        for vbaFile in self.getVBAFiles():
            if vbaFile == self.getMainVBAFile():       
                with open (vbaFile, "r") as f:
                    macro=f.read()
                    # Add the main macro- into ThisDocument part of Word document
                    wordModule = document.VBProject.VBComponents("ThisDocument")
                    wordModule.CodeModule.AddFromString(macro)
            else: # inject other vba files as modules
                with open (vbaFile, "r") as f:
                    macro=f.read()
                    wordModule = document.VBProject.VBComponents.Add(1)
                    wordModule.Name = os.path.splitext(os.path.basename(vbaFile))[0]
                    wordModule.CodeModule.AddFromString(macro)
        
        word.DisplayAlerts=False
        # Remove Informations
        logging.info("   [-] Remove hidden data and personal info...")
        wdRDIAll=99
        document.RemoveDocumentInformation(wdRDIAll)
        
        logging.info("   [-] Save Document...")
        wdFormatXMLDocumentMacroEnabled = 13
        wdFormatDocument = 0
        if self.word97FilePath is not None:
            document.SaveAs(self.word97FilePath, FileFormat=wdFormatDocument)
        if self.wordFilePath is not None:
            document.SaveAs(self.wordFilePath, FileFormat=wdFormatXMLDocumentMacroEnabled)
        
        # save the document and close
        document.Save()
        document.Close()
        word.Application.Quit()
        # garbage collection
        del word
        
        self.disableVbom()
        
        if self.word97FilePath is not None:
            logging.info("   [-] Generated Word file path: %s" % self.word97FilePath)
        
        if self.wordFilePath is not None:
            logging.info("   [-] Generated Word file path: %s" % self.wordFilePath)
         
        
        