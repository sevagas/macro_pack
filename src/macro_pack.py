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

from common import utils, mp_session, help
from common.utils import MSTypes
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
    from pro_modules.dcom_run import DcomGenerator
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
    

def main(argv):   
    
    logLevel = "INFO"
    # initialize macro_pack session object
    mpSession = mp_session.MpSession(WORKING_DIR, VERSION, MP_TYPE)
         
    try:
        longOptions = ["quiet", "input-file=","vba-output=", "mask-strings", "encode","obfuscate","obfuscate-form", "obfuscate-names", "obfuscate-strings", "file=","template=", "start-function=", "dde"] 
        shortOptions= "s:f:t:v:x:X:w:W:P:G:hqmo"
        # only for Pro release
        if MP_TYPE == "Pro":
            longOptions.extend(["vbom-encode", "persist","keep-alive", "av-bypass", "trojan=", "stealth", "dcom="])
            shortOptions += "T:"
        # Only enabled on windows
        if sys.platform == "win32":
            longOptions.extend([ "generate=", "excel-output=", "word-output=", "excel97-output=", "word97-output=", "ppt-output="])
            
        opts, args = getopt.getopt(argv, "s:f:t:v:x:X:w:W:P:G:hqmo", longOptions) # @UnusedVariable
    except getopt.GetoptError:          
        help.printUsage(BANNER, sys.argv[0], mpSession)                     
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
            if arg is None or arg.startswith("-") or  arg == "help" or arg == "HELP":
                help.printTemplatesUsage(BANNER, sys.argv[0])
                sys.exit(0)
            else:
                mpSession.template = arg
        elif opt=="-q" or opt=="--quiet": 
            logLevel = "ERROR"
        elif opt=="-v" or opt=="--vba-output": 
            mpSession.vbaFilePath = os.path.abspath(arg)
            mpSession.fileOutput = True
        elif opt == "--dde":
            if sys.platform == "win32":
                mpSession.ddeMode = True
        elif opt in ("-G", "--generate"): 
            # Document generation enabled only on windows
            if sys.platform == "win32":
                mpSession.outputFilePath = os.path.abspath(arg)
        elif opt=="-h" or opt=="--help": 
            help.printUsage(BANNER, sys.argv[0], mpSession)                         
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
                elif opt == "T" or opt=="--trojan":
                    # Document generation enabled only on windows
                    if sys.platform == "win32":
                        mpSession.outputFilePath = os.path.abspath(arg)
                        mpSession.trojan = True
                elif opt == "--stealth":
                    mpSession.stealth = True
                elif opt == "--dcom":
                    mpSession.dcom = True
                    mpSession.dcomTarget = arg 
                else:
                    help.printUsage(BANNER, sys.argv[0], mpSession)                         
                    sys.exit(0)
            else:
                help.printUsage(BANNER, sys.argv[0], mpSession)                          
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
        #else:
        #    logging.error("   [!] ERROR: No input provided")                        
        #    sys.exit(2)
    else:
        if not os.path.isfile(mpSession.vbaInput):
            logging.error("   [!] ERROR: Could not find %s!" % mpSession.vbaInput)
            sys.exit(2)
        else:
            logging.info("   [-] Input file path: %s" % mpSession.vbaInput)
    
    if mpSession.trojan==False:
        # verify that output file does not already exist
        if mpSession.outputFilePath is not None:
            if os.path.isfile(mpSession.outputFilePath):
                logging.error("   [!] ERROR: Output file %s already exist!" % mpSession.outputFilePath)
                sys.exit(2)
    else:
        # In trojan mod, file are tojane if they already exist and created if they dont.
        # except for vba output which is not concerned by trojan feature
        if mpSession.outputFilePath is not None and mpSession.outputFileType == MSTypes.VBA:
            if mpSession.outputFilePath is not None:
                if os.path.isfile(mpSession.outputFilePath):
                    logging.error("   [!] ERROR: Output file %s already exist!" % mpSession.outputFilePath)
                    sys.exit(2)
    
    
    #Create temporary folder
    logging.info("   [-] Temporary working dir: %s" % WORKING_DIR)
    if not os.path.exists(WORKING_DIR):
        os.makedirs(WORKING_DIR)
    
    try:
        
        # Create temporary work file.
        if mpSession.ddeMode or mpSession.template:
            inputFile = os.path.join(WORKING_DIR,"command.cmd")
        else:
            inputFile = os.path.join(WORKING_DIR,utils.randomAlpha(9))+".vba"
        if mpSession.stdinContent is not None: 
            logging.info("   [-] Store std input in file..." )
            f = open(inputFile, 'w')
            f.writelines(mpSession.stdinContent)
            f.close()    
        else:
            # Create temporary work file
            if mpSession.vbaInput is not None:
                logging.info("   [-] Store input file..." )
                shutil.copy2(mpSession.vbaInput, inputFile)
        if os.path.isfile(inputFile):
            logging.info("   [-] Temporary input file: %s" %  inputFile)
        # Check output file format
        if mpSession.outputFilePath:
            logging.info("   [-] Target output format: %s" %  mpSession.outputFileType)
        
         
               
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
                if MSTypes.XL in mpSession.outputFileType:
                    generator = ExcelGenerator(mpSession)
                    generator.run()
                if MSTypes.WD in mpSession.outputFileType:
                    generator = WordGenerator(mpSession)
                    generator.run()
                if MSTypes.PPT in mpSession.outputFileType:
                    generator = PowerPointGenerator(mpSession)
                    generator.run()
            else:
                if MSTypes.XL in mpSession.outputFileType:
                    if os.path.isfile(mpSession.outputFilePath):
                        generator = ExcelTrojan(mpSession)
                        generator.run()
                    else:
                        generator = ExcelGenerator(mpSession)
                        generator.run()
                if MSTypes.WD in mpSession.outputFileType:
                    if os.path.isfile(mpSession.outputFilePath):
                        generator = WordTrojan(mpSession)
                        generator.run()
                    else:
                        generator = WordGenerator(mpSession)
                        generator.run()
                if MSTypes.PPT in mpSession.outputFileType:
                    if os.path.isfile(mpSession.outputFilePath):
                        generator = PptTrojan(mpSession)
                        generator.run()
                    else:
                        generator = PowerPointGenerator(mpSession)
                        generator.run()

            if mpSession.stealth == True:
                obfuscator = Stealth(mpSession)
                obfuscator.run()
                
            if mpSession.ddeMode: # DDE Attack mode
                if MSTypes.WD in mpSession.outputFileType:
                    generator = WordDDE(mpSession)
                    generator.run()
                else:
                    logging.warn(" [!] Word and Word97 are only format supported for DDE attacks.")
                
            if mpSession.dcom: #run dcom attack
                generator = DcomGenerator(mpSession)
                generator.run()
    
        if mpSession.outputFileType == MSTypes.VBA or mpSession.outputFilePath == None:
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
    
    