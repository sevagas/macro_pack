#!/usr/bin/env python
# encoding: utf-8

import os, mmap
from common import utils


class MpModule():
    def __init__(self,workingPath, startFunction, ):
        self.workingPath = workingPath
        self._startFunction = startFunction
        self.reservedFunctions = []
        if startFunction is not None:
            self.reservedFunctions.append(self._startFunction)
        self.reservedFunctions.append("AutoOpen")
        self.reservedFunctions.append("Workbook_Open")
        self.reservedFunctions.append("Document_Open")
        self.reservedFunctions.append("Auto_Open")    
        self.potentialStartFunctions = []
        self.potentialStartFunctions.append("AutoOpen")
        self.potentialStartFunctions.append("Workbook_Open")
        self.potentialStartFunctions.append("Document_Open")  
        self.potentialStartFunctions.append("Auto_Open")    
        
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
        """ Add a new VBA module file containing moduleContent and with random name """
        newModuleName = os.path.join(self.workingPath,utils.randomAlpha(9)+".vba")
        f = open(newModuleName, 'w')
        f.write(moduleContent)
        f.close()
    
    def run(self):
        """ Run the module """
        raise NotImplementedError