#!/usr/bin/env python
# encoding: utf-8

import re
from modules.mp_module import MpModule
from common.utils import randomAlpha
import logging
from random import randint


class ObfuscateNames(MpModule):

    vbaFunctions = []

    def _findAllFunctions(self):
        for vbaFile in self.getVBAFiles():
            f = open(vbaFile)
            content = f.readlines()
            f.close()
            
            for line in content:
                matchObj = re.match( r'.*(Sub|Function)\s+([a-zA-Z0-9_]+)\s*\(.*\).*', line, re.M|re.I) 
                if matchObj:
                    keyword = matchObj.groups()[1]
                    if keyword not in self.reservedFunctions:
                        self.vbaFunctions.append(keyword)
                        self.reservedFunctions.append(keyword)
                else:
                    matchObj = re.match( r'.*(Sub|Function)\s+([a-zA-Z0-9_]+)\s*As\s*.*', line, re.M|re.I) 
                    if matchObj:
                        keyword = matchObj.groups()[1]
                        if keyword not in self.reservedFunctions:
                            self.vbaFunctions.append(keyword)
                            self.reservedFunctions.append(keyword)
        
            # Write in new file 
            f = open(vbaFile, 'w')
            f.writelines(content)
            f.close()
            

    def _replaceFunctions(self):
        # Identify function, subs and variables names
        logging.info("   [-] Rename functions...")
        self._findAllFunctions()
        
        # Different situation surrounding variables
        varDelimitors=[(" "," "),("\t"," "),("\t","("),(" ","("),("(","("),(" ","\n"),(" ",","),(" "," ="),("."," "),(".","\"")]
        
        # Replace functions and function calls by random string
        for keyWord in self.vbaFunctions:
            keyTmp = randomAlpha(randint(8, 20)) # Generate random names with random size
            #logging.info("|%s|->|%s|" %(keyWord,keyTmp))
            self.reservedFunctions.append(keyTmp)
        
            for vbaFile in self.getVBAFiles():
                f = open(vbaFile)
                content = f.readlines()
                f.close()
    
                for varDelimitor in varDelimitors:
                    newKeyWord = varDelimitor[0] + keyTmp +varDelimitor[1]
                    keywordTmp = varDelimitor[0] + keyWord +varDelimitor[1]
                    for n,line in enumerate(content):
                        matchObj = re.match( r'.*".*%s.*".*' %keyWord, line, re.M|re.I) # check if word is inside a string
                        if matchObj:
                            if "Application.Run" in line or "Application.OnTime" in line: # dynamic function call detected
                                content[n] = line.replace(keywordTmp, newKeyWord)
                        else:
                            
                            content[n] = line.replace(keywordTmp, newKeyWord)
                                
                # Write in new file 
                f = open(vbaFile, 'w')
                f.writelines(content)
                f.close()


    def _replaceLibImports(self, macroLines):
        logging.info("   [-] Rename API imports...")
        # Identify function, subs and variables names
        keyWords = []
        for line in macroLines:
            matchObj = re.match( r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"\s*.*', line, re.M|re.I) 
            if matchObj:
                keyword = matchObj.groups()[1]
                keyWords.append(keyword)
        
        # Remove duplicates
        keyWords = list(set(keyWords))
    
        # Replace functions and function calls by random string
        for keyWord in keyWords:
            keyTmp = randomAlpha(randint(8, 20)) # Generate random names with random size
            for n,line in enumerate(macroLines):
                if "Lib " in line and keyWord + " " in line: # take care of declaration
                    if "Alias " in line: # if fct already has an alias we can change the original keyword
                        macroLines[n] = line.replace(keyWord, keyTmp)
                    else:
                        # We have to create a new alias
                        matchObj = re.match( r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"(\s*).*', line, re.M|re.I) 
                        line =  line.replace(keyWord, keyTmp)
                        macroLines[n] = line.replace(matchObj.groups()[2],matchObj.groups()[2] + "\" Alias \"%s" % keyWord) 
                else:
                    matchObj = re.match( r'.*".*%s.*".*'%keyWord, line, re.M|re.I) # check if word is inside a string
                    if matchObj:
                        if "Application.Run" in line: # dynamic function call detected
                            macroLines[n] = line.replace(keyWord, keyTmp)
                        # else word is part of normal string so we do not touch
                    else:
                        macroLines[n] = line.replace(keyWord, keyTmp)
        return macroLines


    def _replaceVariables(self,macroLines):
        logging.info("   [-] Rename variables...")
        #  variables names
        keyWords = []
        # format something As ...
        for line in macroLines:
            findList = re.findall( r'([a-zA-Z0-9_]+)(\(\))?\s+As\s+(String|Integer|Long|Object|Byte|Variant|Boolean|Any|Word.Application|Excel.Application)', line, re.I) 
            if findList:
                for keyWord in findList:
                    if keyWord[0] not in self.reservedFunctions: # prevent erase of previous variables and function names
                        keyWords.append(keyWord[0])
                        self.reservedFunctions.append(keyWord[0])
        # format Set <something> =  ...
        for line in macroLines:
            findList = re.findall( r'Set\s+([a-zA-Z0-9_]+)\s+=', line, re.I) 
            if findList:
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)
        # format Const <something> =  ...
        for line in macroLines:
            findList = re.findall( r'Const\s+([a-zA-Z0-9_]+)\s+=', line, re.I) 
            if findList:
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)
        # format Type <something>
        for line in macroLines:
            findList = re.findall( r'Type\s+([a-zA-Z0-9_]+)$', line, re.I) 
            if findList:
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)

        #logging.info(str(keyWords))
        
        # Different situation surrounding variables
        varDelimitors=[(" "," "),(" ","."),(" ","("),(" ","\n"),(" ",","),(" ",")"),(" "," =")]
        varDelimitors.extend([("\t"," "),("\t","."),("\t","("),("\t","\n"),("\t",","),("\t",")"),("\t"," =")])
        varDelimitors.extend([("(",")"),("(",","),("("," +"),("("," As"),("(",".")])
        varDelimitors.extend([("="," "),("=",","),("=","\n"),("Set "," =")])
        
        # replace all keywords by random name
        for keyWord in keyWords:
            keyTmp = randomAlpha(randint(8, 20)) # Generate random names with random size
            #logging.info("|%s|->|%s|" %(keyWord,keyTmp))
            for varDelimitor in varDelimitors:
                newKeyWord = varDelimitor[0] + keyTmp +varDelimitor[1]
                keywordTmp = varDelimitor[0] + keyWord +varDelimitor[1]
                for n,line in enumerate(macroLines):
                    macroLines[n] = line.replace(keywordTmp, newKeyWord)
                
        return macroLines


    def _replaceConsts(self,macroLines):
        logging.info("   [-] Rename numeric constants...")
        # Identify and replace constants
        constList = ["0","1", "2"]
        
        for constant in constList:
            # Create random string to replace constant 
            keyTmp = randomAlpha(10)
            constDeclaration = "Const " + keyTmp +" = " + constant + "\n"
            macroLines.insert(0, constDeclaration)
            newKeyWord = " " + keyTmp + ", "
            keywordTmp =  " " + constant + ", " 
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord) 
            newKeyWord = "," + keyTmp + ","
            keywordTmp =  "," + constant + "," 
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord)    
            newKeyWord = ", " + keyTmp + ")"
            keywordTmp =  ", " + constant + ")" 
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord) 
            newKeyWord = "(" + keyTmp + ","
            keywordTmp =  "(" + constant + "," 
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord)            
        return macroLines


    def run(self):
        logging.info(" [+] VBA names obfuscation ...") 
        
        # Obfuscate function name
        content = self._replaceFunctions()
        
        for vbaFile in self.getVBAFiles():
            f = open(vbaFile)
            content = f.readlines()
            f.close()
            
            # Obfuscate variables name
            content = self._replaceVariables(content)
            # replace numerical consts
            content = self._replaceConsts (content)
            #replace lib imports
            content = self._replaceLibImports (content)
        
            # Write in new file 
            f = open(vbaFile, 'w')
            f.writelines(content)
            f.close()
        logging.info("   [-] OK!") 
        
