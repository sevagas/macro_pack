#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
import time
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport

import logging
from modules.mp_module import MpModule
from common.utils import MSTypes


class ComGenerator(MpModule):
    """ 
    Copy and play macro on local PC using COM
    """ 
    def __init__(self,mpSession):
        self.comTarget =  mpSession.runTarget 
        super().__init__(mpSession)
    
    
    def run(self):
        # Extract information
        
        logging.info(" [+] Running %s on local PC..." % self.comTarget)
        if not os.path.isfile(self.comTarget):
            logging.error("   [!] Could not find %s " % self.comTarget)
            return
        
        targetApp = MSTypes.guessApplicationType(self.comTarget)
        if MSTypes.XL in targetApp or MSTypes.SYLK in targetApp:
            comApp = "Excel.Application"
        elif MSTypes.WD in targetApp:
            comApp = "Word.Application"
        elif MSTypes.PPT in targetApp:
            comApp = "PowerPoint.Application"
        elif MSTypes.VSD in targetApp:
            comApp = "Visio.InvisibleApp"
        elif MSTypes.ACC in targetApp:
            comApp = "Access.Application"
        elif MSTypes.MPP in targetApp:
            comApp = "MSProject.Application"
        else:
            logging.error("   [!] Could not recognize file extension for %s" % self.comTarget)
            return
        
        
        try:
            logging.info("   [-] Create handler...")
            comObj = win32com.client.Dispatch(comApp)
        except:
            logging.exception("   [!] Cannot access COM!")
            return 
    
        # We need to force run macro if it is not triggered at document open
        if self.startFunction and self.startFunction not in self.potentialStartFunctions:
            logging.info("   [-] Run macro %s..." % self.startFunction)
        else:
            logging.info("   [-] No specific start function, running auto open macro...")

        # do the operation in background without actually opening Excel
        try:
            if MSTypes.XL in targetApp or MSTypes.SYLK in targetApp:
                if self.mpSession.runVisible:
                    comObj.Visible = True
                document = comObj.Workbooks.Open(self.comTarget)
            elif MSTypes.WD in targetApp or MSTypes.VSD in targetApp:
                if self.mpSession.runVisible:
                    comObj.Visible = True
                document = comObj.Documents.Open(self.comTarget)
            elif MSTypes.PPT in targetApp:
                document = comObj.Presentations.Open(self.comTarget)
            elif MSTypes.ACC in targetApp:
                comObj.OpenCurrentDatabase(self.comTarget)
                #comObj.DoCmd.RunMacro(self.startFunction)
            elif MSTypes.MPP in targetApp:
                document = comObj.FileOpen(self.comTarget, True)
            if self.startFunction and self.startFunction not in self.potentialStartFunctions:
                document = comObj.run(self.startFunction)
        except Exception:
            logging.exception("   [!] Problem detected!")
        
        time.sleep(3.5) # need to have app alive to launch async call with --background option

                
        logging.info("   [-] Cleanup...")
        try:
            document.close()
            comObj.Application.Quit()
        except:
            pass
        try:
            comObj.FileClose ()
        except:
            pass
        try:
            comObj.Quit()
        except:
            pass
        # garbage collection
        del comObj
        
        logging.info("   [-] OK!") 
        

         
        
        