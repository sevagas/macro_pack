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
from modules.mp_generator import Generator


class ExcelGenerator(Generator):
    """ Module used to generate MS excel file from working dir content"""
    
    def getAutoOpenVbaFunction(self):
        return "Workbook_Open"
    
    def getAutoOpenVbaSignature(self):
        return "Sub Workbook_Open()"
    
        
    def enableVbom(self):
        # Enable writing in macro (VBOM)
        # First fetch the application version
        objExcel = win32com.client.Dispatch("Excel.Application")
        objExcel.Visible = False # do the operation in background 
        self.version = objExcel.Application.Version
        # IT is necessary to exit office or value wont be saved
        objExcel.Application.Quit()
        del objExcel
        # Next change/set AccessVBOM registry value to 1
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\Excel\\Security"
        logging.info("   [-] Set %s to 1..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,1) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    
    def disableVbom(self):
        # Disable writing in VBA project
        #  Change/set AccessVBOM registry value to 0
        keyval = "Software\\Microsoft\Office\\"  + self.version + "\\Excel\\Security"
        logging.info("   [-] Set %s to 0..." % keyval)
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,0) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
    
    
    def check(self):
        logging.info("   [-] Check feasibility...")
        try:
            objExcel = win32com.client.Dispatch("Excel.Application")
            objExcel.Application.Quit()
            del objExcel
        except:
            logging.error("   [!] Cannot access Excel.Application object. Is software installed on machine? Abort.")
            return False  
        return True
    
    
    def generate(self):
        
        logging.info(" [+] Generating MS Excel document...")
        self.enableVbom()
        
        # open up an instance of Excel with the win32com driver\        \\
        excel = win32com.client.Dispatch("Excel.Application")
        # do the operation in background without actually opening Excel
        excel.Visible = False
        # open the excel workbook from the specified file or create if file does not exist
        logging.info("   [-] Open workbook...")
        workbook = excel.Workbooks.Add()
        
        self.resetVBAEntryPoint()
        logging.info("   [-] Inject VBA...")
        # Read generated files
        for vbaFile in self.getVBAFiles():
            if vbaFile == self.getMainVBAFile():       
                with open (vbaFile, "r") as f:
                    macro=f.read()
                    # Add the main macro- into ThisWorkbook part of excel file
                    excelModule = workbook.VBProject.VBComponents("ThisWorkbook")
                    excelModule.CodeModule.AddFromString(macro)
            else: # inject other vba files as modules
                with open (vbaFile, "r") as f:
                    macro=f.read()
                    excelModule = workbook.VBProject.VBComponents.Add(1)
                    excelModule.Name = os.path.splitext(os.path.basename(vbaFile))[0]
                    excelModule.CodeModule.AddFromString(macro)
        
        excel.DisplayAlerts=False
        # Remove Informations
        logging.info("   [-] Remove hidden data and personal info...")
        xlRDIAll=99
        workbook.RemoveDocumentInformation(xlRDIAll)
        
        logging.info("   [-] Save workbook...")
        xlOpenXMLWorkbookMacroEnabled = 52
        xlExcel8 = 56
        if self.outputFileType == MSTypes.XL97:
            workbook.SaveAs(self.outputFilePath, FileFormat=xlExcel8)
        elif self.outputFileType == MSTypes.XL:
            workbook.SaveAs(self.outputFilePath, FileFormat=xlOpenXMLWorkbookMacroEnabled)
        # save the workbook and close
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()
        # garbage collection
        del excel
        
        self.disableVbom()

        logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
        logging.info("   [-] Test with : \nmacro_pack.exe --run %s\n" % self.outputFilePath)

        