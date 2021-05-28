#!/usr/bin/env python
# encoding: utf-8

import re
from modules.mp_module import MpModule
from common.utils import extractStringsFromText, randomStringBasedOnCharset
import logging
from random import randint
from random import choice
import string

class ObfuscateNames(MpModule):

    vbaFunctions = []
    win32Functions = []

    def _findAllFunctions(self):
        
        for vbaFile in self.getVBAFiles():
            if self.mpSession.obfOnlyMain:
                if vbaFile != self.getMainVBAFile():
                    continue
            f = open(vbaFile)
            content = f.readlines()
            f.close()
            
            for line in content:
                matchObj = re.match(r'.*(Sub|Function)\s+([a-zA-Z0-9_]+)\s*\(.*\).*', line, re.M|re.I)
                if matchObj:
                    keyword = matchObj.groups()[1]
                    if keyword not in self.reservedFunctions:
                        self.vbaFunctions.append(keyword)
                        self.reservedFunctions.append(keyword)
                else:
                    matchObj = re.match(r'.*(Sub|Function)\s+([a-zA-Z0-9_]+)\s*As\s*.*', line, re.M|re.I)
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
        self._findAllFunctions()
        
        # Different situation surrounding variables
        varDelimitors=[(" "," "),("\t"," "),("\t","("),("\t"," ="),(" ","("),(", ",")"),("AddressOf ",")"),("(","("),(" ","\n"),(" ",","),(" "," ="),("."," "),(".","\"")]
        
        # Replace functions and function calls by random string
        for keyWord in self.vbaFunctions:
            keyTmp = self._generateRandomVbaName() # Generate random names with random size
            self.reservedFunctions.append(keyTmp)
        
            for vbaFile in self.getVBAFiles():
                if self.mpSession.obfOnlyMain:
                    if vbaFile != self.getMainVBAFile():
                        continue
                f = open(vbaFile)
                content = f.readlines()
                f.close()
    
                for varDelimitor in varDelimitors:
                    newKeyWord = varDelimitor[0] + keyTmp +varDelimitor[1]
                    keywordTmp = varDelimitor[0] + keyWord +varDelimitor[1]

                    for n,line in enumerate(content):
                        extractedStrings=extractStringsFromText(line)
                        if keyWord in extractedStrings:
                            if "Application.Run" in line or "Application.OnTime" in line: # dynamic function call detected
                                content[n] = line.replace(keywordTmp, newKeyWord)
                        else:
                            content[n] = line.replace(keywordTmp, newKeyWord)
                                
                # Write in new file 
                f = open(vbaFile, 'w')
                f.writelines(content)
                f.close()


    def _generateRandomVbaName(self):
        """
        Generate random names with random size
        :return:
        """
        vbaName = choice(string.ascii_lowercase) # Name has to start with a letter
        return vbaName + randomStringBasedOnCharset(randint(self.mpSession.obfuscatedNamesMinLen - 1, self.mpSession.obfuscatedNamesMaxLen - 1), self.mpSession.obfuscatedNamesCharset)


    def _findAllWin32Api(self):

        for vbaFile in self.getVBAFiles():
            if self.mpSession.obfOnlyMain:
                if vbaFile != self.getMainVBAFile():
                    continue
            f = open(vbaFile)
            content = f.readlines()
            f.close()

            for line in content:
                matchObj = re.match(r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"\s*.*', line, re.M | re.I)
                if matchObj:
                    keyword = matchObj.groups()[1]
                    if keyword not in self.reservedFunctions:
                        self.win32Functions.append(keyword)
                        self.reservedFunctions.append(keyword)



    def _replaceLibImports(self):

        self._findAllWin32Api()

        # Replace functions and function calls by random string
        for keyWord in self.win32Functions:
            
            keyTmp = self._generateRandomVbaName()
            self.reservedFunctions.append(keyTmp)
            for vbaFile in self.getVBAFiles():
                if self.mpSession.obfOnlyMain:
                    if vbaFile != self.getMainVBAFile():
                        continue
                f = open(vbaFile)
                macroLines = f.readlines()
                f.close()

                #logging.info("Keyword:%s,keyTmp:%s "%(keyWord,keyTmp))
                for n,line in enumerate(macroLines):
                    if "Lib " in line and keyWord + " " in line: # take care of declaration
                        if "Alias " in line: # if fct already has an alias we can change the original keyword
                            #logging.debug(line)
                            macroLines[n] = line.replace(" %s " % keyWord," %s " %  keyTmp, 1)
                            #logging.debug(macroLines[n])
                        else:
                            # We have to create a new alias
                            matchObj = re.match(r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"(\s*).*', line, re.M|re.I)
                            #logging.debug(line)
                            line =  line.replace(" %s " % keyWord, " %s " % keyTmp)
                            #logging.debug(line+"\n")
                            macroLines[n] = line.replace(matchObj.groups()[2],matchObj.groups()[2] + "\" Alias \"%s" % keyWord)
                    else:
                        matchObj = re.match(r'.*".*%s.*".*' % keyWord, line, re.M|re.I) # check if word is inside a string
                        if matchObj:
                            if "Application.Run" in line: # dynamic function call detected
                                macroLines[n] = line.replace(keyWord, keyTmp)

                            # else word is part of normal string, so we do not touch
                        else:
                            if keyWord + " " in line or keyWord + "(" in line:
                                #logging.info(line)
                                macroLines[n] = line.replace(keyWord, keyTmp)

                # Write in new file
                f = open(vbaFile, 'w')
                f.writelines(macroLines)
                f.close()



    def _replaceVariables(self,macroLines):
        
        #  will contain variables names
        keyWords = []
        # format something As ...
        for line in macroLines:
            findList = re.findall(r'([a-zA-Z0-9_]+)(\(\))?\s+As\s+(String|Integer|Long|Object|Byte|Variant|Boolean|Single|Any|Collection|Word.Application|Excel.Application|VbVarType)', line, re.I)
            if findList:
                for keyWord in findList:
                    if keyWord[0] not in self.reservedFunctions: # prevent erase of previous variables and function names
                        keyWords.append(keyWord[0])
                        self.reservedFunctions.append(keyWord[0])
        # format Set <something> =  ...
        for line in macroLines:
            findList = re.findall(r'Set\s+([a-zA-Z0-9_]+)\s+=', line, re.I)
            if findList:
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)
        # format Const <something> =  ...
        for line in macroLines:
            findList = re.findall(r'Const\s+([a-zA-Z0-9_]+)\s+=', line, re.I)
            if findList:
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)
        # format Type <something>
        for line in macroLines:
            findList = re.findall(r'Type\s+([a-zA-Z0-9_]+)$', line, re.I)
            if findList:
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)

        #logging.info(str(keyWords))
        
        # Different situation surrounding variables
        varDelimitors=[(" "," "),(" ","."),(" ","("),(" ","\n"),(" ",","),(" ",")"),(" "," =")]
        varDelimitors.extend([("."," ="),("."," A"),("."," O"),(".",")"),(".",","),("."," ")])
        varDelimitors.extend([("#"," "),("#",",")])
        varDelimitors.extend([("\t"," "),("\t","."),("\t","("),("\t","\n"),("\t",","),("\t",")"),("\t"," =")])
        varDelimitors.extend([("(",")"),("(","("),("(",","),("("," +"),("("," *"),("("," &"),("("," -"),("("," ="),("("," As"),("("," And"),("("," To"),("("," Or"),("(",".")])
        varDelimitors.extend([("="," "),("=",","),("=","\n"),("Set "," =")])
        
        # replace all keywords by random name
        for keyWord in keyWords:
            keyTmp = self._generateRandomVbaName()
            #logging.info("|%s|->|%s|" %(keyWord,keyTmp))
            for varDelimitor in varDelimitors:
                newKeyWord = varDelimitor[0] + keyTmp + varDelimitor[1]
                keywordTmp = varDelimitor[0] + keyWord + varDelimitor[1]
                for n,line in enumerate(macroLines):
                    macroLines[n] = line.replace(keywordTmp, newKeyWord)
                
        return macroLines


    def _replaceConsts(self, macroLines):

        # Identify and replace constants
        constList = ["0","1", "2"]
        
        for constant in constList:
            # Create random string to replace constant 
            keyTmp = self._generateRandomVbaName()
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
        logging.info("   [-] Rename functions...")
        self._replaceFunctions()
        
        logging.info("   [-] Rename variables...")
        if self.mpSession.ObfReplaceConstants:
            logging.info("   [-] Rename some numeric const...")

        # go through each file
        for vbaFile in self.getVBAFiles():
            if self.mpSession.obfOnlyMain:
                if vbaFile != self.getMainVBAFile():
                    continue
            f = open(vbaFile)
            content = f.readlines()
            f.close()
            # Obfuscate variables name
            content = self._replaceVariables(content)
            if self.mpSession.ObfReplaceConstants:
                # replace numerical consts
                if ",0," in content or " 0," in "".join(content):
                    content = self._replaceConsts(content)

            # Write in new file
            f = open(vbaFile, 'w')
            f.writelines(content)
            f.close()

        # replace lib imports
        if self.mpSession.obfuscateDeclares:
            logging.info("   [-] Rename API imports...")
            self._replaceLibImports()
        logging.info("   [-] OK!") 
        
