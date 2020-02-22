#!/usr/bin/env python
# encoding: utf-8

import re
import logging
from modules.mp_module import MpModule


class ObfuscateForm(MpModule):

    def _removeComments(self, macroLines):
        # Identify function, subs and variables names
        keyWords = []
        for line in macroLines:
            matchObj = re.match( r".*('.+)$", line, re.M|re.I) 
            if matchObj:
                keyWords.append(matchObj.groups()[0])
    
        # Replace functions by random string
        for keyWord in keyWords:
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keyWord, "")
        return macroLines
    
    
    def _removeSpaces(self, macroLines):
        """Remove tabs space and function separations """
        
        for n,line in enumerate(macroLines):
            macroLines[n] = line.lstrip()
        return macroLines
    
    def _removeTabs(self, macroLines):
        """  Replace all tabs by space """
        for n,line in enumerate(macroLines):
            macroLines[n] = line.replace('\t', ' ')
        return macroLines
    
    
    def run(self):
        logging.info(" [+] VBA form obfuscation ...") 
        logging.info("   [-] Remove spaces...")
        logging.info("   [-] Remove comments...")
        for vbaFile in self.getVBAFiles():
            f = open(vbaFile)
            content = f.readlines()
            f.close()
            
            # Remove comments
            content = self._removeComments(content) # must remove comments before space to avoir empty lines
            # Replace all tabs by space
            content =  self._removeTabs(content)
            # Remove spaces
            content =  self._removeSpaces(content)
        
            # Write in new file 
            f = open(vbaFile, 'w')
            f.writelines(content)
            f.close()
        logging.info("   [-] OK!") 
            
            

            