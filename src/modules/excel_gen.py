#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport
    import winreg # @UnresolvedImport

import logging
from modules.mp_module import MpModule


class ExcelGenerator(MpModule):
    """ Module used to generate MS excel file from working dir content"""
    
    def __init__(self,workingPath, startFunction,excelFilePath=None,excel97FilePath=None):
        self.excelFilePath = excelFilePath
        self.excel97FilePath = excel97FilePath
        super().__init__(workingPath, startFunction)
        
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
    
    
    
    def run(self):
        logging.info(" [+] Generating MS Excel document...")
        
        self.enableVbom()
        
        # open up an instance of Excel with the win32com driver
        excel = win32com.client.Dispatch("Excel.Application")
        # do the operation in background without actually opening Excel
        excel.Visible = False
        # open the excel workbook from the specified file or create if file does not exist
        logging.info("   [-] Open workbook...")
        workbook = excel.Workbooks.Add()
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
                    excelModule.CodeModule.AddFromString(macro)
        
        logging.info("   [-] Save workbook...")
        xlOpenXMLWorkbookMacroEnabled = 52
        xlExcel8 = 56
        if self.excel97FilePath is not None:
            workbook.SaveAs(self.excel97FilePath, FileFormat=xlExcel8)
        
        if self.excelFilePath is not None:
            workbook.SaveAs(self.excelFilePath, FileFormat=xlOpenXMLWorkbookMacroEnabled)
        # save the workbook and close
        excel.DisplayAlerts=False
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()
        # garbage collection
        del excel
        
        self.disableVbom()
        
        if self.excel97FilePath is not None:
            logging.info("   [-] Generated Excel file path: %s" % self.excel97FilePath)
        if self.excelFilePath is not None:
            logging.info("   [-] Generated Excel file path: %s" % self.excelFilePath)
        