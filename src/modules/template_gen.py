#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import shlex
import os
import logging
from modules.mp_module import MpModule
from common import templates, utils


class TemplateToVba(MpModule):
    """ Generate a VBA document from a given template """
        
    def _fillTemplate(self, content, values):
        for value in values:
            content = content.replace("<<<TEMPLATE>>>", value, 1)
        
        # generate random file name
        vbaFile = os.path.abspath(os.path.join(self.workingPath,utils.randomAlpha(9)+".vba"))
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile)) 
        # Write in new file 
        f = open(vbaFile, 'w')
        f.write(content)
        f.close()

    
    def run(self):
        logging.info(" [+] Generating VBA document from template...")
        if self.template is None:
            logging.info("   [!] No template defined")
            return
        
        if self.template == "HELLO":
            content = templates.HELLO
        elif self.template == "DROPPER":
            content = templates.DROPPER
        elif self.template == "DROPPER2":
            content = templates.DROPPER2
        elif self.template == "DROPPER_PS":
            content = templates.DROPPER_PS
        elif self.template == "METERPRETER":
            content = templates.METERPRETER
        else: # if not one of default template suppose its a custom template
            if os.path.isfile(self.template):
                f = open(self.template, 'r')
                content = f.read()
                f.close()
            else:
                logging.info("   [!] Template is not recognized as file or default template.")
                return
         
        # open file containing template values       
        mainFile = self.getMainVBAFile()
        if mainFile != "":
            f = open(mainFile, 'r')
            valuesFileContent = f.read()
            f.close()
            self._fillTemplate(content, shlex.split(valuesFileContent)) # split on space but preserve what is between quotes
            # remove file containing template values
            os.remove(mainFile)
            logging.info("   [-] OK!") 
        else:
            logging.info("   [!] Could not find main file!")