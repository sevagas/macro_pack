#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.vbs_gen import VBSGenerator

HTA_TEMPLATE = \
r"""
<!DOCTYPE html>
<html>
<head>
<HTA:APPLICATION />
<script type="text/vbscript">
<<<VBS>>>
<<<MAIN>>>
Close
</script>
</head>
<body>
</body>
</html>

"""

class HTAGenerator(VBSGenerator):
    """ Module used to generate HTA file from working dir content"""
        
        
    def generate(self):
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        self.vbScriptConvert()
        f = open(self.getMainVBAFile()+".vbs")
        vbsContent = f.read()
        f.close()
        
        vbsContent = vbsContent.replace("WScript.Echo ", "MsgBox ")
        
        # Write VBS in template
        htaContent = HTA_TEMPLATE
        htaContent = htaContent.replace("<<<VBS>>>", vbsContent)
        htaContent = htaContent.replace("<<<MAIN>>>", self.startFunction)
        # Write in new HTA file
        f = open(self.outputFilePath, 'w')
        f.writelines(htaContent)
        f.close()
        logging.info("   [-] Generated HTA file: %s" % self.outputFilePath)
        logging.info("   [-] Test with : \nmshta %s\n" % self.outputFilePath)
        

        
        
        