#!/usr/bin/env python
# encoding: utf-8

import logging
from common.utils import randomAlpha
from modules.vbs_gen import VBSGenerator

WSF_TEMPLATE = \
r"""<?XML version="1.0"?>
<job id="<<<random>>>">
    <script language="VBScript">
        <![CDATA[
            <<<VBS>>>
            <<<MAIN>>>  
        ]]>
    </script>
</job>
"""

class WSFGenerator(VBSGenerator):
    """ Module used to generate WSF file from working dir content
    """
        
        
    def generate(self):
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        self.vbScriptConvert()
        f = open(self.getMainVBAFile()+".vbs")
        vbsContent = f.read()
        f.close()
        
        #vbsContent = vbsContent.replace("WScript.Echo ", "MsgBox ")
        
        # Write VBS in template
        wsfContent = WSF_TEMPLATE
        wsfContent = wsfContent.replace("<<<random>>>", randomAlpha(8))
        wsfContent = wsfContent.replace("<<<VBS>>>", vbsContent)
        wsfContent = wsfContent.replace("<<<MAIN>>>", self.startFunction)
        # Write in new HTA file
        f = open(self.outputFilePath, 'w')
        f.writelines(wsfContent)
        f.close()
        logging.info("   [-] Generated Windows Script File: %s" % self.outputFilePath)
        logging.info("   [-] Test with : \nwscript %s\n" % self.outputFilePath)
        
        
        
        
        