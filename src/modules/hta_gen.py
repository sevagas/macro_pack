#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.vbs_gen import VBSGenerator

HTA_TEMPLATE = \
r"""
<!DOCTYPE html>
<html>
<head>
<HTA:APPLICATION icon="#" WINDOWSTATE="minimize" SHOWINTASKBAR="no" SYSMENU="no"  CAPTION="no" />
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
        
    def vbScriptConvert(self):
        super().vbScriptConvert()
        f = open(self.getMainVBAFile()+".vbs")
        vbsContent = f.read()
        f.close()
        logging.info("   [-] Convert VBScript to HTA...")
        vbsContent = vbsContent.replace("WScript.Echo ", "MsgBox ")
        vbsContent = vbsContent.replace('WScript.Sleep(1000)','CreateObject("WScript.Shell").Run "cmd /c ping localhost -n 1",0,True')
        vbsContent = vbsContent.replace('Wscript.Quit 0', 'Self.Close')
        vbsContent = vbsContent.replace('Wscript.ScriptFullName', 'self.location.pathname')
        
        # Write in new VBS file
        f = open(self.getMainVBAFile()+".vbs", 'w')
        f.writelines(vbsContent)
        f.close()
        
        
    def generate(self):
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        self.vbScriptConvert()
        f = open(self.getMainVBAFile()+".vbs")
        vbsContent = f.read()
        f.close()
        
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
        

        
        
        