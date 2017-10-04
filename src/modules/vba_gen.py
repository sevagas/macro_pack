#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_module import MpModule
import shutil
import os


class VBAGenerator(MpModule):
    """ Module used to generate MS excel file from working dir content"""
    
    def __init__(self,workingPath, startFunction,vbaFilePath=None, fileOutput=False):
        self.vbaFilePath = vbaFilePath
        self.fileOutput = fileOutput
        super().__init__(workingPath, startFunction)
    
    def run(self):
        logging.info(" [+] Analyzing generated VBA files...")
        if self.vbaFilePath is not None:
            if len(self.getVBAFiles())==1:
                shutil.copy2(self.getMainVBAFile(), self.vbaFilePath)
                logging.info("   [-] Generated VBA file: %s" % self.vbaFilePath) 
            else:
                logging.info("   [!] More then one VBA file generated, files will be copied in same dir as %s" % self.vbaFilePath)
                for vbaFile in self.getVBAFiles():
                    shutil.copy2(vbaFile, os.path.join(os.path.dirname(self.vbaFilePath),os.path.basename(vbaFile)))
                    logging.info("   [-] Generated VBA file: %s" % os.path.join(os.path.dirname(self.vbaFilePath),os.path.basename(vbaFile)))   
                 
        
        if self.fileOutput == False:
            logging.info(" [+] Generated VBA code:\n")
            if len(self.getVBAFiles())==1: 
                vbaFile = self.getMainVBAFile() 
                with open(vbaFile,'r') as f:
                    print(f.read())
            else:
                logging.info("   [!] More then one VBA file generated")
                for vbaFile in self.getVBAFiles():
                    with open(vbaFile,'r') as f:
                        print(" =======================  %s  ======================== " % vbaFile)
                        print(f.read())