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
    
    
    def genExtractEmbeddedFileVBA(self, identificationTag):
        """
        Creates a new source file containing function to extract and return embedded content as a string
        identificationTag is used to identify the embedded content. 
        Returns name of generated method
        """
        raise NotImplementedError
    
    
    def embedFile(self, identificationTag):
        """
        Embed the content of  self.embeddedFilePath inside the generated target file
        identificationTag is used to identify the embedded content.
        """
        raise NotImplementedError
    
    
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
        # Obfuscate VBA files
        self.runObfuscators()
        # generate
        self.generate()
        
        