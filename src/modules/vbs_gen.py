#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_module import MpModule
import re

VBS_TEMPLATE = \
r"""
<<<VBS>>>
<<<MAIN>>>
"""

class VBSGenerator(MpModule):
    """ Module used to generate VBS file from working dir content"""
    
    def vbScriptCheck(self):
        logging.info("   [-] Check if VBA->VBScript is possible...")
        # Check nb of source file
        if len(self.getVBAFiles())>1:
            logging.info("   [-] This module cannot handle more than one source file. Abort!")
            return False
        
        f = open(self.getMainVBAFile())
        content = f.readlines()
        f.close()
        # Check there are no call to Application object
        for line in content:
            if line.find("Application.") != -1:
                logging.info("   [-] You cannot access Application object in VBScript. Abort!")
                return False  
        
        # Check there are no DLL import
        for line in content:
            matchObj = re.match( r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"\s*.*', line, re.M|re.I) 
            if matchObj:
                logging.info("   [-] VBScript does not support DLL import. Abort!")
                return False 
        return True
    
    
    def vbScriptConvert(self):
        logging.info("   [-] Convert VBA to VBScript...")
        translators = [("Val(","CInt("),(" Chr$"," Chr"),(" Mid$"," Mid"),("On Error GoTo","'//On Error GoTo"),("byebye:",""), ("Next ", "Next '//")]
        translators.extend([(" As String"," "),(" As Object"," "),(" As Long"," "),(" As Integer"," "),(" As Variant"," ")])
        translators.extend([ ("MsgBox ", "WScript.Echo ")])
        f = open(self.getMainVBAFile())
        content = f.readlines()
        f.close()
        isUsingEnviron = False
        for n,line in enumerate(content):
            # Do easy translations
            for translator in translators:
                line = line.replace(translator[0],translator[1])
            
            # Check if ENVIRON is used
            if line.find("Environ(")!= -1:
                isUsingEnviron = True
                line = re.sub('Environ\("([A-Z_]+)"\)',r'wshShell.ExpandEnvironmentStrings( "%\1%" )' , line, flags=re.I)
            content[n] = line
            # ENVIRON("COMPUTERNAME") ->   
            #Set wshShell = CreateObject( "WScript.Shell" )
            #strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
        
        # Write in new VBS file
        f = open(self.getMainVBAFile()+".vbs", 'a')
        if isUsingEnviron:
            f.write('Set wshShell = CreateObject( "WScript.Shell" )\n')
        f.writelines(content)
        f.close()
        
        
    def genVBS(self):
        logging.info("   [-] Generating VBS file...")
        f = open(self.getMainVBAFile()+".vbs")
        vbsTmpContent = f.read()
        f.close()
        # Write VBS in template
        vbsContent = VBS_TEMPLATE
        vbsContent = vbsContent.replace("<<<VBS>>>", vbsTmpContent)
        vbsContent = vbsContent.replace("<<<MAIN>>>", self.startFunction)
        # Write in new HTA file
        f = open(self.outputFilePath, 'w')
        f.writelines(vbsContent)
        f.close()
        logging.info("   [-] Generated VBS file: %s" % self.outputFilePath)
        
    
    def run(self):
        logging.info(" [+] Generating VBS file from VBA...")
        if not self.vbScriptCheck():
            return 
        self.vbScriptConvert()
        self.genVBS()
        
        
        