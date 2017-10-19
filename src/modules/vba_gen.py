#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_module import MpModule
import shutil
import os
from common.utils import MSTypes


class VBAGenerator(MpModule):
    """ Module used to generate MS excel file from working dir content"""
    
    def run(self):
        if len(self.getVBAFiles())>0:
            logging.info(" [+] Analyzing generated VBA files...")
            if self.outputFilePath is not None and self.outputFileType == MSTypes.VBA:
                if len(self.getVBAFiles())==1:
                    shutil.copy2(self.getMainVBAFile(), self.outputFilePath)
                    logging.info("   [-] Generated VBA file: %s" % self.outputFilePath) 
                else:
                    logging.info("   [!] More then one VBA file generated, files will be copied in same dir as %s" % self.outputFilePath)
                    for vbaFile in self.getVBAFiles():
                        shutil.copy2(vbaFile, os.path.join(os.path.dirname(self.outputFilePath),os.path.basename(vbaFile)))
                        logging.info("   [-] Generated VBA file: %s" % os.path.join(os.path.dirname(self.outputFilePath),os.path.basename(vbaFile)))   
                     
            
            if self.outputFilePath == None:
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