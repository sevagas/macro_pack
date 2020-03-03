#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
from common import utils
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport
    import winreg # @UnresolvedImport

import logging
from modules.vba_gen import VBAGenerator
import vbLib


class AccessGenerator(VBAGenerator):
    """ Module used to generate MS Access file from working dir content"""
    
    temp_macro_file = 'access_autoexec_macro.txt'

    def getAutoOpenVbaFunction(self):
        return "AutoExec"
    
    def getAutoOpenVbaSignature(self):
        return "Sub AutoExec()"

    def changeSubToFunction(self, macro):
        ind = macro.find(self.getAutoOpenVbaSignature())
        macro
        return macro[:ind] + macro[ind:].replace('Sub', 'Function', 2)
        
    def enableVbom(self):
        # Enable writing in macro (VBOM)
        # First fetch the application version
        objAccess = win32com.client.Dispatch("Access.Application")
        objAccess.Visible = False # do the operation in background 
        self.version = objAccess.Application.Version
        # IT is necessary to exit office or value wont be saved
        objAccess.Application.Quit()
        del objAccess
        # Next change/set AccessVBOM registry value to 1
        keyval = "Software\\Microsoft\\Office\\"  + self.version + "\\Access\\Security"
        logging.info("   [-] Set %s to 1..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,1) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    
    def disableVbom(self):
        # Disable writing in VBA project
        #  Change/set AccessVBOM registry value to 0
        keyval = "Software\\Microsoft\\Office\\"  + self.version + "\\Access\\Security"
        logging.info("   [-] Set %s to 0..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,0) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
    
    
    def check(self):
        logging.info("   [-] Check feasibility...")
        try:
            objAccess = win32com.client.Dispatch("Access.Application")
            objAccess.Application.Quit()
            del objAccess
        except:
            logging.error("   [!] Cannot access Access.Application object. Is software installed on machine? Abort.")
            return False  
        return True
     
    
    def generate(self):
        
        logging.info(" [+] Generating MS Access document...")
        try:
            self.enableVbom()
            
            # open up an instance of Access with the win32com driver\        \\
            access = win32com.client.Dispatch("Access.Application")
            # do the operation in background without actually opening Access
            access.Visible = False
            # open the Access database from the specified file or create if file does not exist
            logging.info("   [-] Open database...")
            access.NewCurrentDatabase(self.outputFilePath)
            
            self.resetVBAEntryPoint()
            logging.info("   [-] Inject VBA...")

            # Read generated files
            for cnt, vbaFile in enumerate(self.getVBAFiles()):
                module_name = "Module%i" % (cnt + 1)
                if vbaFile == self.getMainVBAFile():
                    with open (vbaFile, "r") as f:
                        macro = f.read()
                        # Add the main module into this part of Access file
                        access.DoCmd.RunCommand(139)  # acCmdNewObjectModule = 139
                        access.DoCmd.Save(5, module_name) # AcObjectType.AcModule = 5
                        macro = self.changeSubToFunction(macro)
                        access.Modules.Item(cnt).AddFromString(macro)
                else: # inject other vba files as modules
                    with open (vbaFile, "r") as f:
                        macro = f.read()
                        access.DoCmd.RunCommand(139)  # acCmdNewObjectModule = 139
                        access.DoCmd.Save(5, module_name) # AcObjectType.AcModule = 5
                        access.Modules.Item(cnt).AddFromString(macro)

            content = vbLib.templates.ACCESS_MACRO_TEMPLATE
            macro_file = os.path.join(self.workingPath, self.temp_macro_file)
            with open(macro_file, 'w') as tmp:
                tmp.write(content)

            access.LoadFromText(
                4,  # AcObjectType.AcMacro = 4
                self.getAutoOpenVbaFunction(),
                macro_file
            )
            access.DoCmd.Close(
                4,  # AcObjectType.AcMacro = 4
                self.getAutoOpenVbaFunction(),
                1  # AcCloseSave.AcSaveYes = 1
            )

            # save the database and close
            access.CloseCurrentDatabase()
            access.Quit()
            # garbage collection
            del access

            self.disableVbom()

            logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
            logging.info("   [-] To create a compiled file, open %s and save as .accde!" % self.outputFilePath)
            logging.info("   [-] Test with : \n%s --run %s\n" % (utils.getRunningApp(),self.outputFilePath))
            
        except Exception:
            logging.exception(" [!] Exception caught!")
            logging.error(" [!] Hints: Check if MS office is really closed and Antivirus did not catch the files")
            logging.error(" [!] Attempt to force close MS Access applications...")
            objAccess = win32com.client.Dispatch("Access.Application")
            objAccess.Application.Quit()
            del objAccess


        