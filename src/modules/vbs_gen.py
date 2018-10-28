#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.vba_gen import VBAGenerator
import re, os
from vbLib import Base64ToBin, CreateBinFile
import base64

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
                matchObj = re.match( r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"\s*.*', line, re.M|re.I) 
                if matchObj:
                    logging.error("   [-] VBScript does not support DLL import. Abort!")
                    logging.error("   [-] Line: %s" % line)
                    return False 
        return True
    
    
    def vbScriptConvert(self):
        logging.info("   [-] Convert VBA to VBScript...")
        translators = [("Val(","CInt("),(" Chr$"," Chr"),(" Mid$"," Mid"),("On Error GoTo","'//On Error GoTo"),("byebye:",""), ("Next ", "Next '//")]
        translators.extend([(" As String"," "),(" As Object"," "),(" As Long"," "),(" As Integer"," "),(" As Variant"," "), (" As Boolean", " "), (" As Byte", " ")])
        translators.extend([ ("MsgBox ", "WScript.Echo "), ('Application.Wait Now + TimeValue("0:00:01")', 'WScript.Sleep(1000)')])
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
        
    
    
    def embedFile(self):
        """
        Embed the content of  self.embeddedFilePath inside the generated target file
        """
        logging.info("   [-] Embedding file %s..." % self.embeddedFilePath)
        if not os.path.isfile(self.embeddedFilePath):
            logging.warning("   [!] Could not find %s! " % self.embeddedFilePath)
            return
        
        f = open(self.embeddedFilePath, 'rb')
        content = f.read()
        f.close()
        encodedBytes = base64.b64encode(content)
        base64Str= encodedBytes.decode("utf-8")  
       
        # Shorten size if needed
        VBAMAXLINELEN = 100 # VBA will fail if line is too long
        cpt = 0
        newPackedMacro = ""
        nbIter = int(len(base64Str) / VBAMAXLINELEN)
        # Create a VBA string builder containing all encoded macro
        while cpt < nbIter:
            newPackedMacro += base64Str[cpt * VBAMAXLINELEN:(cpt+1) * VBAMAXLINELEN] + "\" \n str = str & \"" 
            cpt += 1
        newPackedMacro += base64Str[cpt * VBAMAXLINELEN:] 
        packedMacro= "\"" + newPackedMacro + "\"" 
    
        newContent = Base64ToBin.VBA + "\n"
        newContent += CreateBinFile.VBA + "\n"
        newContent += "Sub DumpFile(strFilename)"
        newContent += "\n Dim str \n str = %s \n readEmbed = Base64ToBin(str) \n CreateBinFile strFilename, readEmbed \n" % (packedMacro) 
        newContent += "End Sub \n \n"       
        
        
        self.addVBAModule(newContent)
        return  
    
    
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
        

        
        
        