#!/usr/bin/env python
# encoding: utf-8

import re
import codecs
from modules.mp_module import MpModule
from random import randint
import logging


class ObfuscateStrings(MpModule):
    
    hexToStringRoutine = \
'''Function HexToStr(ByVal hexString As String) As String
Dim counter As Long
For counter = 1 To Len(hexString) Step 2
HexToStr = HexToStr & Chr$(Val("&H" & Mid$(hexString, counter, 2)))
Next counter
End Function
'''


    def _splitStrings(self, macroLines):
        
        # Find strings and randomly split them in half 
        for n,line in enumerate(macroLines):
            #Check if string is not preprocessor instruction, const or contain escape quotes
            if len(line) > 6 and "\"\"" not in line and "PtrSafe Function" not in line and "Declare Function" not in line and "Declare Sub" not in line and "PtrSafe Sub" not in line and "Environ" not in line:
                # Find strings in line
                findList = re.findall(r'"(.+?)"', line, re.I)
                if findList:
                    for detectedString in findList:
                        if len(detectedString) > 4:
                            # Compute value to cut string randomly
                            randomValue = randint(2, len(detectedString)-2)
                            #if len(detectedString[:randomValue])<2:
                            #    logging.error("!!!! 1 byte string split detected for left string: %s \n" % detectedString)
                            #if len(detectedString[randomValue:]) < 2:
                            #    logging.error("!!!! 1 byte string split detected for right string: %s (%s) \n" % (detectedString,detectedString[randomValue:]))
                            newStr = detectedString[:randomValue] + "\" & \"" + detectedString[randomValue:] 
                            line = line.replace(detectedString, newStr)
                    macroLines[n] = line
        return macroLines
    
    
    
    def _maskStrings(self,macroLines, newFunctionName):
        """ Mask string in VBA by encoding them """
        # Find strings and replace them by hex encoded version
        for n,line in enumerate(macroLines):
            #Check if string is not preprocessor instruction, const or contain escape quoting
            if line.lstrip() != "" and line.lstrip()[0] != '#' and "Const" not in line and "\"\"" not in line and "PtrSafe Function" not in line and "Declare Function" not in line and "PtrSafe Sub" not in line and "Declare Sub" not in line and "Environ" not in line:
                # Find strings in line
                findList = re.findall(r'"(.+?)"', line, re.I)
                if findList:
                    for detectedString in findList: 
                        # Hex encode string
                        encodedBytes = codecs.encode(bytes(detectedString, "utf-8"), 'hex_codec')
                        newStr = newFunctionName + "(\"" + encodedBytes.decode("utf-8") + "\")"
                        wordToReplace =  "\"" + detectedString + "\""
                        line = line.replace(wordToReplace, newStr)
                # Replace line if result is not too big
                if len(line) < 1024:
                    macroLines[n] = line
        
        return macroLines
    
    
    
    def run(self):
        if not self.mpSession.noStringsObfuscation:
            logging.info(" [+] VBA strings obfuscation ...")
            logging.info("   [-] Split strings...")
            logging.info("   [-] Encode strings...")
            # Compute new random function and variable names for HexToStr
            if self.mpSession.obfuscateNames:
                newFunctionName = self.mpSession.nameObfuscationCallback(14, self.mpSession.obfuscatedNamesCharset)
                newVarName1 = self.mpSession.nameObfuscationCallback(9, self.mpSession.obfuscatedNamesCharset)
                newVarName2 = self.mpSession.nameObfuscationCallback(8, self.mpSession.obfuscatedNamesCharset)
            else:
                newFunctionName = "HexToStr"
                newVarName1 = "counter"
                newVarName2 = "hexString"
            for vbaFile in self.getVBAFiles():
                # Check if there are strings in file
                with open(vbaFile) as fileToCheck:
                    data = fileToCheck.read()
                if '"' not in data:
                    continue

                f = open(vbaFile)
                content = f.readlines()
                f.close()

                # Split string
                content = self._splitStrings(content)
                # mask string
                content = self._maskStrings(content, newFunctionName)

                # Write in new file
                f = open(vbaFile, 'w')
                f.writelines(content)
                f.close()

            # Add decode routine
            if self.mpSession.mpType == "Pro":
                from pro_vbLib.vbautils import HexToString
                hexDecodeBlock = self.getVBLibContent(HexToString)
            else:
                hexDecodeBlock = self.hexToStringRoutine
            if self.mpSession.obfuscateNames:
                hexDecodeBlock = hexDecodeBlock.replace("HexToStr", newFunctionName).replace("counter", newVarName1).replace("hexString", newVarName2)
                hexDecodeBlock = hexDecodeBlock.replace("wordInter", self.mpSession.nameObfuscationCallback(12, self.mpSession.obfuscatedNamesCharset)).replace("prefix", self.mpSession.nameObfuscationCallback(12, self.mpSession.obfuscatedNamesCharset))
            #logging.info(hexDecodeBlock)
            self.addVBAModule(hexDecodeBlock)

            logging.info("   [-] OK!")
            