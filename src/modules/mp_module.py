#!/usr/bin/env python
# encoding: utf-8

import os, mmap, logging
from common import utils
from common.utils import MSTypes

class MpModule():
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
        self.reservedFunctions.append("Workbook_Open")
        self.reservedFunctions.append("Document_Open")
        self.reservedFunctions.append("Auto_Open")    
        self.reservedFunctions.append("Document_DocumentOpened")
        self.potentialStartFunctions = []
        self.potentialStartFunctions.append("AutoOpen")
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
                if  os.stat(vbaFile).st_size != 0:  
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
    
    def getAutoOpenVbaFunction(self):
        raise NotImplementedError
    
    def getAutoOpenVbaSignature(self):
        raise NotImplementedError 
    
    def getCMDFile(self):
        """ Return command line file (for DDE mode)"""
        if os.path.isfile(os.path.join(self.workingPath,"command.cmd")):
            return os.path.join(self.workingPath,"command.cmd")
        else:
            return ""
        
    
    def getMainVBAFile(self):
        """ return main vba file (the one containing macro entry point) """
        result = ""
        vbaFiles = self.getVBAFiles()
        if len(vbaFiles)==1:
            result = vbaFiles[0]
        else:
            if self.startFunction is not None:
                for vbaFile in vbaFiles:
                    if  os.stat(vbaFile).st_size != 0:  
                        with open(vbaFile, 'rb', 0) as file, mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
                            if s.find(self.startFunction.encode()) != -1:
                                result  = vbaFile
                                break
                            
        return result
    
    def addVBAModule(self, moduleContent):
        """ 
        Add a new VBA module file containing moduleContent and with random name
        Returns name of new VBA file
        """
        newModuleName = os.path.join(self.workingPath,utils.randomAlpha(9)+".vba")
        f = open(newModuleName, 'w')
        f.write(moduleContent)
        f.close()
        return newModuleName
    
    
    @classmethod
    def getAutoOpenFunction(self):
        """ Return the VBA Function/Sub name which triggers autoopen for the current outputFileType """
        result = ""
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
        elif MSTypes.PUB in self.outputFileType:
            result = "Document_Open"
        return result
            
    
    
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
                    if line.find(self.startFunction) != -1:    
                        content[n] = self.getAutoOpenVbaSignature() + "\n"
                f = open(mainFile, 'w')
                f.writelines(content)
                f.close()   
                # 2 Change  cure module start function
                self._startFunction = self.getAutoOpenVbaFunction()
        
    def run(self):
        """ Run the module """
        raise NotImplementedError
    


    

    