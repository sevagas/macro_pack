#!/usr/bin/env python
# encoding: utf-8

import os, mmap, logging,re
from common import utils
from common.utils import MSTypes
import shlex

class MpModule:
    def __init__(self,mpSession):
        self.mpSession = mpSession
        self.workingPath = mpSession.workingPath
        self._startFunction = mpSession.startFunction
        self.outputFilePath = mpSession.outputFilePath
        self.outputFileType = mpSession.outputFileType
        self.template = mpSession.template
        
        self.reservedFunctions = []
        if self._startFunction is not None:
            self.reservedFunctions.append(self._startFunction)
        self.reservedFunctions.append("AutoOpen")
        self.reservedFunctions.append("AutoExec")
        self.reservedFunctions.append("Workbook_Open")
        self.reservedFunctions.append("Document_Open")
        self.reservedFunctions.append("Auto_Open")    
        self.reservedFunctions.append("Document_DocumentOpened")
        self.potentialStartFunctions = []
        self.potentialStartFunctions.append("AutoOpen")
        self.potentialStartFunctions.append("AutoExec")
        self.potentialStartFunctions.append("Workbook_Open")
        self.potentialStartFunctions.append("Document_Open")  
        self.potentialStartFunctions.append("Auto_Open") 
        self.potentialStartFunctions.append("Document_DocumentOpened")    
        
    @property
    def startFunction(self):
        """ Return start function, attempt to find it in vba files if _startFunction is not set """
        result = None
        if self._startFunction is not None:
            result =  self._startFunction
        else:
             
            vbaFiles = self.getVBAFiles()
            for vbaFile in vbaFiles:
                if os.stat(vbaFile).st_size != 0:
                    with open(vbaFile, 'rb', 0) as file, mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
                        for potentialStartFunction in self.potentialStartFunctions:
                            if s.find(potentialStartFunction.encode()) != -1:
                                self._startFunction = potentialStartFunction
                                if self._startFunction not in self.reservedFunctions:
                                    self.reservedFunctions.append(self._startFunction)
                                result = potentialStartFunction
                                break                
        return result
    
    
    def getVBAFiles(self):
        """ Returns path of all vba files in working dir """
        vbaFiles = []
        vbaFiles += [os.path.join(self.workingPath,each) for each in os.listdir(self.workingPath) if each.endswith('.vba')]
        return vbaFiles
    

    
    def getCMDFile(self):
        """ Return command line file (for DDE mode)"""
        if os.path.isfile(os.path.join(self.workingPath,"command.cmd")):
            return os.path.join(self.workingPath,"command.cmd")
        else:
            return ""
        
    
    def fillInputParams(self, paramDict):
        """ 
        Fill parameters dictionary using given input. If input is missing, ask for input to user.
        """
        # Fill parameters based on input file
        cmdFile = self.getCMDFile()
        if cmdFile is not None and cmdFile != "":
            f = open(cmdFile, 'r')
            valuesFileContent = f.read()
            logging.debug("    -> CMD file content: \n%s" % valuesFileContent)
            f.close()
            os.remove(cmdFile)
            if self.mpSession.fileInput is None or len(paramDict) > 1:# if values where passed by input pipe or in a file but there are multiple params
                inputValues = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
            else: 
                inputValues = [valuesFileContent] # value where passed using -f option
            if len(inputValues) >= len(paramDict): 
                i = 0  
                # Fill entry parameters
                for key, value in paramDict.items():
                    paramDict[key] = inputValues[i]
                    i += 1
            else:
                logging.error("   [!] Incorrect number of provided input parameters (%d provided where this features needs %d).\n    Required parameters: %s\n" % (len(inputValues),len(paramDict), list(paramDict.keys())))
                return
        else:
            # if input was not provided
            logging.warning("   [!] Could not find input parameters. Please provide the next values:")
            for key, value in paramDict.items():
                if value is None or value == "" or value.isspace():
                    newValue = None
                    while newValue is None or newValue == "" or newValue.isspace():
                        newValue = input("    %s:" % key)
                    paramDict[key] = newValue
    
    
    def fillInputParams2(self, paramArray):
        """ 
        Fill parameters dictionary using given input. If input is missing, ask for input to user.
        """
        # Fill parameters based on input file
        allMandatoryParamFilled = False
        i = 0
        mandatoryParamLen= 0
        while i < len(paramArray):
            if not paramArray[i].optional:
                mandatoryParamLen += 1 
            i += 1
        if mandatoryParamLen == 0:
            allMandatoryParamFilled = True
        
        cmdFile = self.getCMDFile()
        if cmdFile is not None and cmdFile != "":
            f = open(cmdFile, 'r')
            valuesFileContent = f.read()
            logging.debug("    -> CMD file content: \n%s" % valuesFileContent)
            f.close()
            os.remove(cmdFile)
            if self.mpSession.fileInput is None or len(paramArray) > 1:# if values where passed by input pipe or in a file but there are multiple params
                inputValues = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
            else: 
                inputValues = [valuesFileContent] # value where passed using -f option
            
           
            if len(inputValues) >= mandatoryParamLen: 
                i = 0  
                # Fill entry parameters
                while i < len(inputValues):
                    paramArray[i].value = inputValues[i]
                    i += 1
                    allMandatoryParamFilled = True
                        
        if not allMandatoryParamFilled:
            # if input was not provided
            logging.warning("   [!] Could not find some mandatory input parameters. Please provide the next values:")
            for param in paramArray:
                
                if param.value is None or param.value == "" or param.value.isspace():
                    newValue = None
                    while newValue is None or newValue == "" or newValue.isspace():
                        newValue = input("    %s:" % param.name)
                    param.value = newValue
     

     
    
    def getMainVBAFile(self):
        """ return main vba file (the one containing macro entry point) """
        result = ""
        vbaFiles = self.getVBAFiles()
        if len(vbaFiles)==1:
            result = vbaFiles[0]
        else:
            if self.startFunction is not None:
                for vbaFile in vbaFiles:
                    if os.stat(vbaFile).st_size != 0:
                        with open(vbaFile, 'rb', 0) as file, mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
                            if s.find(self.startFunction.encode()) != -1:
                                result = vbaFile
                                break
        logging.debug("    [*] Start function:%s" % self.startFunction)
        logging.debug("    [*] Main VBA file:%s" % result)
        return result
    
    
    def getFileContainingString(self, strToFind):
        """ Returns fist VB file containing the string to find"""
        result = ""
        vbaFiles = self.getVBAFiles()

        for vbaFile in vbaFiles:
            if os.stat(vbaFile).st_size != 0:
                with open(vbaFile, 'rb', 0) as file, mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
                    if s.find(strToFind.encode()) != -1:
                        result = vbaFile
                        break
                            
        return result
    
    
    def addVBAModule(self, moduleContent, moduleName=None):
        """ 
        Add a new VBA module file containing moduleContent and with random name
        Returns name of new VBA file
        """
        if moduleName is None:
            moduleName = utils.randomAlpha(9)
            modulePath = os.path.join(self.workingPath,moduleName+".vba")
        else:
            modulePath = os.path.join(self.workingPath,utils.randomAlpha(9)+".vba")
        if moduleName in self.mpSession.vbModulesList:
            logging.debug("    [,] %s module already loaded" % moduleName)
        else:
            logging.debug("    [,] Adding module: %s" % moduleName)
            self.mpSession.vbModulesList.append(moduleName)
            f = open(modulePath, 'w')
            f.write(moduleContent)
            f.close()
        return modulePath
    
    
    def addVBLib(self, vbaLib):
        """ 
        Add a new VBA Library module depending on the current context 
        """
        if self.outputFileType in MSTypes.MS_OFFICE_FORMATS:
            if MSTypes.WD in self.outputFileType and hasattr(vbaLib, 'VBA_WD'):
                moduleContent = vbaLib.VBA_WD
            elif MSTypes.XL in self.outputFileType and hasattr(vbaLib, 'VBA_XL'): 
                moduleContent = vbaLib.VBA_XL
            elif MSTypes.PPT in self.outputFileType and hasattr(vbaLib, 'VBA_PPT'): 
                moduleContent = vbaLib.VBA_PPT
            else:
                moduleContent = vbaLib.VBA
        else:
            if self.outputFileType in [MSTypes.HTA, MSTypes.SCT] and hasattr(vbaLib, 'VBS_HTA'):
                moduleContent = vbaLib.VBS_HTA
            else:
                if hasattr(vbaLib, 'VBS'):
                    moduleContent = vbaLib.VBS
                else:
                    moduleContent = vbaLib.VBA
        newModuleName = self.addVBAModule(moduleContent, vbaLib.__name__)
        return newModuleName
    
    
    
    def insertVbaCode(self, targetModule, targetFunction,targetLine, vbaCode):
        """
        Insert some code at targetLine (number) at targetFunction in targetModule
        """
        logging.debug("     [,] Opening "+ targetFunction + " in " + targetModule + " to inject code...")
        f = open(targetModule)
        content = f.readlines()
        f.close()
        
        for n,line in enumerate(content):
            matchObj = re.match(r'.*(Sub|Function)\s+%s\s*\(.*\).*' % targetFunction, line, re.M|re.I)
            if matchObj:  
                logging.debug("     [,] Found " + targetFunction + " ") 
                content[n+targetLine] = content[n+targetLine]+ vbaCode+"\n"
                break
        
        
        logging.debug("     [,] New content: \n " + "".join(content) + "\n\n ") 
        f = open(targetModule, 'w')
        f.writelines(content)
        f.close()
    
    
    def getAutoOpenFunction(self):
        """ Return the VBA Function/Sub name which triggers auto open for the current outputFileType """
        if MSTypes.WD in self.outputFileType:
            result = "AutoOpen"
        elif MSTypes.XL in self.outputFileType:
            result = "Workbook_Open"
        elif MSTypes.PPT in self.outputFileType:
            result = "AutoOpen"
        elif MSTypes.MPP in self.outputFileType:
            result = "Auto_Open"
        elif MSTypes.VSD in self.outputFileType:
            result = "Document_DocumentOpened"
        elif MSTypes.ACC in self.outputFileType:
            result = "AutoExec"
        elif MSTypes.PUB in self.outputFileType:
            result = "Document_Open"
        else:
            result = "AutoOpen"
        return result
            
    
        
    def run(self):
        """ Run the module """
        raise NotImplementedError
    


    

    