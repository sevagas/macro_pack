#!/usr/bin/env python
# encoding: utf-8
import sys, os
import logging
import psutil, time
import __main__ as main # @UnresolvedImport To get the real origin of the script not the location of current file 
from termcolor import colored
import colorama
from collections import OrderedDict
colorama.init() # for nice colored output on windows
import tabulate # easy_install tabulate

MACRO_PACK_PATH = os.path.abspath(os.path.join(os.path.abspath(main.__file__), '..', '..'))

BUILD_SCRIPT = os.path.join(MACRO_PACK_PATH, "build.bat")
BIN_PATH=os.path.join(MACRO_PACK_PATH, "bin")

# Import files from pacro_pack src
SRC_PATH=os.path.join(MACRO_PACK_PATH, "src")
TEST_PATH=os.path.join(MACRO_PACK_PATH, "test")
MP_MAIN=os.path.join(SRC_PATH, "macro_pack.py")

sys.path.append(SRC_PATH)
from common import utils
from common.utils import MSTypes


fileToGenerate = os.path.abspath("resultFile.gug")
fileToGenerateContent = u"This file is the result of a test"

VBA = \
"""

Sub AutoOpen()
    CreateTxtFile "%s", "%s"
End Sub

 'Create A  Text and fill it
 ' Will overwrite existing file
Private Sub CreateTxtFile(FilePath As String, FileContent As String)
   
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(FilePath, True, True)
    Fileout.Write FileContent
    Fileout.Close

End Sub
""" % (fileToGenerate,fileToGenerateContent )


testSummary = OrderedDict()


def testBuild():
    """ Test we can build macro_pack """
    logging.info(" [+] Testing macro_path build...")
    try:
        os.system("cd %s && %s > nul" % (MACRO_PACK_PATH, BUILD_SCRIPT))
        assert(os.path.isfile(os.path.join(BIN_PATH,"macro_pack.exe")))
        logging.info("   [-] Success!\n")
        testSummary["Macro Pack Build"] = "[OK]"
    except:
        logging.exception("   [!] Error!\n")
        testSummary["Macro Pack Build"] = "[KO]"
        return False
     

def testLnkGenerators():
    result = True
    for testFormat in MSTypes.Shortcut_FORMATS:
        testFile = utils.randomAlpha(8)+MSTypes.EXTENSION_DICT[testFormat]
        try:
            logging.info(" [+] Testing generation of %s file..." % testFormat)
            os.system("echo shortcut_dest icon_file  | %s %s -G %s -q" % (sys.executable,MP_MAIN, testFile))
            assert(os.path.isfile(testFile))
            logging.info("   [-] Success!\n")
            testSummary[testFormat] = "[OK]"
        except:
            result = False
            testSummary[testFormat] = "[KO]"
            logging.exception("   [!] Error!\n")
            
        if os.path.isfile(testFile):
            os.remove(testFile)
        
        
    return result



def _executeFile(testFile,testFormat):
    """ Executes an MSType format file based on its format """
    absPathTestFile = os.path.abspath(testFile)
    if testFormat in MSTypes.MS_OFFICE_FORMATS:
        os.system("%s %s  --run=%s -q" % (sys.executable,MP_MAIN, testFile))
    elif testFormat == MSTypes.VBS or testFormat == MSTypes.WSF:
        os.system("wscript %s" % (testFile))
    elif testFormat == MSTypes.HTA:
        os.system("mshta %s" % (absPathTestFile))
    elif testFormat == MSTypes.SCT:
        os.system("regsvr32 /u /n /s /i:%s scrobj.dll" % (testFile))
    elif testFormat == MSTypes.XSL:
        os.system("wmic os get /FORMAT:\"%s\"" % (testFile))
    else:
        os.system("cmd.exe /c %s" % (testFile))


def _clearTextVBGenerationTest(testFile, testFormat):
    vbaTestFile = "testmacro.vba"
    logging.info(" [+] Testing generation of %s file..." % testFormat)
    os.system("%s %s -f %s -G %s -q" % (sys.executable,MP_MAIN,vbaTestFile, testFile))
    assert(os.path.isfile(testFile))
    logging.info("   [-] Success!\n")
    
    if testFormat not in [MSTypes.VBA]:
        logging.info(" [+] Testing run of %s file..." % testFormat)
        _executeFile(testFile,testFormat)
        # Check result
        assert(os.path.isfile(fileToGenerate))
        with open(fileToGenerate, 'rb') as infile:
            content = infile.read().decode('utf-16')
        #logging.info("Content:|%s|  - fileToGenerateContent:|%s|" % (content, fileToGenerateContent))
        assert(content == fileToGenerateContent)
        testSummary[testFormat+ " in Clear Text"] = "[OK]"
        logging.info("   [-] Success!\n")
        os.remove(fileToGenerate)


def _obfuscatedVBGenerationTest(testFile, testFormat):
    vbaTestFile = "testmacro.vba"
    logging.info(" [+] Testing generation of %s obfuscated file..." % testFormat)
    os.system("%s %s -f %s -G %s -o -q" % (sys.executable,MP_MAIN,vbaTestFile, testFile))
    assert(os.path.isfile(testFile))
    
    logging.info("   [-] Success!\n")

    if testFormat not in [MSTypes.VBA]:
        logging.info(" [+] Testing run of %s obfuscated file..." % testFormat)
        _executeFile(testFile,testFormat)
        # Check result
        assert(os.path.isfile(fileToGenerate))
        with open(fileToGenerate, 'rb') as infile:
            content = infile.read().decode('utf-16')
        assert(content == fileToGenerateContent)
        logging.info("   [-] Success!\n")
        testSummary[testFormat+ " obfuscated"] = "[OK]"
        os.remove(fileToGenerate)


def _obfuscatedVBCmdTemplateGenerationTest(testFile, testFormat):
    logging.info(" [+] Testing generation of %s CMD template file..." % testFormat)
    os.system("echo calc.exe | %s %s -G %s -o -q --template=CMD" % (sys.executable,MP_MAIN, testFile))
    assert(os.path.isfile(testFile))
    
    logging.info("   [-] Success!\n")

    if testFormat not in [MSTypes.VBA]:
        logging.info(" [+] Testing run of %s CMD template file..." % testFormat)
        logging.info("   [-] Kill existing calc process")
        for proc in psutil.process_iter():
            processObj = psutil.Process(proc.pid)
            pname = processObj.name()
            if pname in [ "calc", "Calculator", "calc.exe", "Calculator.exe"]:
                proc.kill()
        
        _executeFile(testFile,testFormat)
        # Check result (calc.exe should have popped)
        time.sleep(0.5)
        calcProcessFound = False
        for proc in psutil.process_iter():
            processObj = psutil.Process(proc.pid)
            pname = processObj.name()
            if pname in [ "calc", "Calculator", "calc.exe", "Calculator.exe"]:
                calcProcessFound = True
                proc.kill()
        assert(calcProcessFound == True)
        logging.info("   [-] Success!\n")
        testSummary[testFormat+ " CMD template"] = "[OK]"



def testVBGenerators():
    """ 
    will run test of MS Office and VBS based formats 
    The tests consist into creating the documents, then running them triggering a file creation macro. Then checking the file is well created
    The tests are run in both cleartext and obfuscated mode.
    A third test will check the correct generation of CMD template (pop calc.exe and check it)
    """
    result = True
    vbaTestFile = "testmacro.vba"
    logging.info(" [+] Build macro test file...")
    with open(vbaTestFile, 'w') as outfile:
        outfile.write(VBA)
    for testFormat in MSTypes.VB_FORMATS:
        testFile = utils.randomAlpha(8)+MSTypes.EXTENSION_DICT[testFormat]
        try:
            _clearTextVBGenerationTest(testFile, testFormat)  
        except:
            result = False
            testSummary[testFormat+ " in Clear Text"] = "[KO]"
            logging.exception("   [!] Error!\n")
            
        if os.path.isfile(testFile):
            os.remove(testFile)
        
        testFile = utils.randomAlpha(8)+MSTypes.EXTENSION_DICT[testFormat]
        try:
            _obfuscatedVBGenerationTest(testFile, testFormat)     
        except:
            result = False
            testSummary[testFormat+ " obfuscated"] = "[KO]"
            logging.exception("   [!] Error!\n")
            
        if os.path.isfile(testFile):
            os.remove(testFile)
            
        testFile = utils.randomAlpha(8)+MSTypes.EXTENSION_DICT[testFormat]
        try:
            _obfuscatedVBCmdTemplateGenerationTest(testFile, testFormat)     
        except:
            result = False
            testSummary[testFormat+ " CMD template"] = "[KO]"
            logging.exception("   [!] Error!\n")
        
        try:  
            if os.path.isfile(testFile):
                os.remove(testFile)
        except:
            logging.exception("   [!] Error while attempting to remove %s!\n" % testFile)
        
    os.remove(vbaTestFile)
    return result

    
    
def main():
    logLevel = "INFO"
    logging.basicConfig(level=getattr(logging, logLevel),format="%(message)s", handlers=[utils.ColorLogFiler()])
    finalResult = True
    
    logging.info(" [+] Current interpreter: %s" % sys.executable)
    
    if not testLnkGenerators(): finalResult = False
    if not testVBGenerators(): finalResult = False
    if not testBuild(): finalResult = False
    
    logging.info(" [+] Tests Summary:")
    tableData = []
    index = ["Test","Result"]
    for key, value in testSummary.items():
        value = value.replace("[KO]", colored("[KO]", "magenta", attrs = ["bold"])).replace("[OK]", colored("[OK]", "green", attrs = ["bold"]))
        newLine = [key, value]
        tableData.append(newLine)

   
    logging.info(tabulate.tabulate(tableData,index,tablefmt="grid") + "\n")
    finalResult = " [+] Complete test success: %s \n\n" % str(finalResult)
    
    logging.info(finalResult.replace("False", colored("[KO]", "red", attrs = ["bold"])).replace("True", colored("[OK]", "green", attrs = ["bold"])))
    
    

    
if __name__ == '__main__':
    main()
    
    