#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.vba_gen import VBAGenerator
import re, os


VBS_TEMPLATE = \
r"""
<<<VBS>>>
<<<MAIN>>>
"""

class VBSGenerator(VBAGenerator):
    """ Module used to generate VBS file from working dir content"""
    
    def check(self):
        logging.info("   [-] Check if VBA->VBScript is possible...")
        # Check nb of source file
        vbaFiles = self.getVBAFiles()
        if len(vbaFiles)>1:
            logging.warning("   [!] This module has more than one source file. They will be concatenated into a single VBS file.")
            
    
        for vbaFile in vbaFiles:    
            f = open(vbaFile)
            content = f.readlines()
            f.close()
            # Check there are no call to Application object
            for line in content:
                if line.find("Application.Run") != -1:
                    logging.error("   [-] You cannot access Application object in VBScript. Abort!")
                    logging.error("   [-] Line: %s" % line)
                    return False  
            
            # Check there are no DLL import
            for line in content:
                matchObj = re.match(r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"\s*.*', line, re.M|re.I)
                if matchObj:
                    logging.error("   [-] VBScript does not support DLL import. Abort!")
                    logging.error("   [-] Line: %s" % line)
                    return False 
        return True
    
    
    
    def printFile(self):
        """ Display generated code on stdout """
        if os.path.isfile(self.outputFilePath):
            logging.info(" [+] Generated file content:\n") 
            with open(self.outputFilePath,'r') as f:
                print(f.read())
    
    
    def vbScriptConvert(self):
        logging.info("   [-] Convert VBA to VBScript...")
        translators = [("Val(","CInt("),(" Chr$"," Chr"),(" Mid$"," Mid"),("On Error GoTo","'//On Error GoTo"),("byebye:",""), ("Next ", "Next '//")]
        translators.extend([("() As String"," "),("CVar","")])
        translators.extend([(" As String"," "),(" As Object"," "),(" As Long"," "),(" As Integer"," "),(" As Variant"," "), (" As Boolean", " "), (" As Byte", " "), (" As Excel.Application", " "), (" As Word.Application", " ")])
        translators.extend([("MsgBox ", "WScript.Echo "), ('Application.Wait Now + TimeValue("0:00:01")', 'WScript.Sleep(1000)')])
        translators.extend([('ChDir ', 'createobject("WScript.Shell").currentdirectory =  ')])
        content = []
        for vbaFile in self.getVBAFiles():  
            f = open(vbaFile)
            content.extend(f.readlines())
            f.close()
        isUsingEnviron = False
        for n,line in enumerate(content):
            # Do easy translations
            for translator in translators:
                line = line.replace(translator[0],translator[1])
            
            if line.strip() == "End":
                line = "Wscript.Quit 0 \n"
            
            # Check if ENVIRON is used
            if line.find("Environ(")!= -1:
                isUsingEnviron = True
                line = re.sub('Environ\("([A-Z_]+)"\)',r'wshShell.ExpandEnvironmentStrings( "%\1%" )', line, flags=re.I)
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
        
    
    def generate(self):
                
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        self.vbScriptConvert()
        
        f = open(self.getMainVBAFile()+".vbs")
        vbsTmpContent = f.read()
        f.close()
        
        # Write VBS in template
        vbsContent = VBS_TEMPLATE
        vbsContent = vbsContent.replace("<<<VBS>>>", vbsTmpContent)
        vbsContent = vbsContent.replace("<<<MAIN>>>", self.startFunction)
             
        # Write in new VBS file
        f = open(self.outputFilePath, 'w')
        f.writelines(vbsContent)
        f.close()
        
        logging.info("   [-] Generated VBS file: %s" % self.outputFilePath)
        logging.info("   [-] Test with : \nwscript %s\n" % self.outputFilePath)
        

        
        
        