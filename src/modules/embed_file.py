#!/usr/bin/env python
# encoding: utf-8


import base64, os
from modules.mp_module import MpModule
import logging
from common.utils import MSTypes
from vbLib import WriteBytes
from vbLib import Base64ToBin, CreateBinFile


class Embedder(MpModule):

    def embedFileVBA(self):
        """
        Embed the content of  self.mpSession.embeddedFilePath inside the generated target file
        """
        logging.info("   [-] Embedding file %s..." % self.mpSession.embeddedFilePath)
        if not os.path.isfile(self.mpSession.embeddedFilePath):
            logging.error("   [!] Could not find %s " % self.mpSession.embeddedFilePath)
            raise Exception("Invalid file path")
            return
        
        infile = open(self.mpSession.embeddedFilePath, 'rb')
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
    
    
    def embedFileVBS(self):
        """
        Embed the content of  self.mpSession.embeddedFilePath inside the generated target file
        """
        logging.info("   [-] Embedding file %s..." % self.mpSession.embeddedFilePath)
        if not os.path.isfile(self.mpSession.embeddedFilePath):
            logging.warning("   [!] Could not find %s! " % self.mpSession.embeddedFilePath)
            return
        
        f = open(self.mpSession.embeddedFilePath, 'rb')
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
    
    
    def run(self):
        if self.outputFileType in MSTypes.MS_OFFICE_FORMATS:
            self.embedFileVBA()
        else:
            self.embedFileVBS() 
        
