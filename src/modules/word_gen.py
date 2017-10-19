#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
from common.utils import MSTypes
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

        logging.info("   [-] Save document format...")
        wdFormatDocument = 0
        wdFormatXMLDocument = 12
        wdFormatXMLDocumentMacroEnabled = 13
        
        if MSTypes.WD97 == self.outputFileType:
            document.SaveAs(self.outputFilePath, FileFormat=wdFormatDocument)
        elif MSTypes.WD == self.outputFileType and ".docx" in self.outputFilePath:
            document.SaveAs(self.outputFilePath, FileFormat=wdFormatXMLDocument)
        elif MSTypes.WD == self.outputFileType and ".docm" in self.outputFilePath:
            document.SaveAs(self.outputFilePath, FileFormat=wdFormatXMLDocumentMacroEnabled)
                    

        logging.info("   [-] Inject VBA...")
        # Read generated files
        for vbaFile in self.getVBAFiles():
            if vbaFile == self.getMainVBAFile():       
                with open (vbaFile, "r") as f:
                    # Add the main macro- into ThisDocument part of Word document
                    wordModule = document.VBProject.VBComponents("ThisDocument")
                    macro=f.read()
                    wordModule.CodeModule.AddFromString(macro)
                    """
                    # Handle big macros by inserting them chink by chunk
                    if len(macro) > 1000000:
                        macrolines = macro.split("\n")
                        for i in range(0,len(macrolines),500):
                            logging.info("i:%s" % str(i))
                            wordModule.CodeModule.insertLines(i+1, "\n".join(macrolines[i:i+500]))
                            #word.DisplayAlerts=False
                            document.Application.Options.Pagination = False
                            document.UndoClear()
                            document.Repaginate()
                            document.Application.ScreenUpdating = True
                            document.Application.ScreenRefresh()
                            document.Save()
  
                    else:
                        wordModule.CodeModule.AddFromString(macro)
                    """
            else: # inject other vba files as modules
                with open (vbaFile, "r") as f:
                    macro=f.read()
                    wordModule = document.VBProject.VBComponents.Add(1)
                    wordModule.Name = os.path.splitext(os.path.basename(vbaFile))[0]
                    wordModule.CodeModule.AddFromString(macro)
        
        #word.DisplayAlerts=False
        # Remove Informations
        logging.info("   [-] Remove hidden data and personal info...")
        wdRDIAll=99
        document.RemoveDocumentInformation(wdRDIAll)
        
        # save the document and close
        document.Save()
        document.Close()
        word.Application.Quit()
        # garbage collection
        del word
        self.disableVbom()

        logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
        
