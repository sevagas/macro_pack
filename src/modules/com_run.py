#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
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
        if MSTypes.XL in targetApp:
            dcomApp = "Excel.Application"
        elif MSTypes.WD in targetApp:
            dcomApp = "Word.Application"
        elif MSTypes.PPT in targetApp:
            dcomApp = "PowerPoint.Application"
        elif MSTypes.VSD in targetApp:
            dcomApp = "Visio.InvisibleApp"
        else:
            logging.error("   [!] Could not recognize file extension for %s" % self.comTarget)
            return
        
        
        try:
            logging.info("   [-] Create handler...")
            comObj = win32com.client.Dispatch(dcomApp)
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
            if MSTypes.XL in targetApp:
                document = comObj.Workbooks.Open(self.comTarget)
            elif MSTypes.WD in targetApp or MSTypes.VSD in targetApp:
                document = comObj.Documents.Open(self.comTarget)
            elif MSTypes.PPT in targetApp:
                document = comObj.Presentations.Open(self.comTarget)
            if self.startFunction and self.startFunction not in self.potentialStartFunctions:
                document = comObj.run(self.startFunction)
        except Exception:
            logging.exception("   [!] Problem detected!")
        
        logging.info("   [-] Cleanup...")
        try:
            document.close()
            comObj.Application.Quit()
        except:
            pass
        try:
            comObj.Quit()
        except:
            pass
        # garbage collection
        del comObj
        
        logging.info("   [-] OK!") 
        

         
        
        