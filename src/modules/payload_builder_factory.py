
import sys
from modules.excel_gen import ExcelGenerator
from modules.word_gen import WordGenerator
from modules.ppt_gen import PowerPointGenerator
from modules.msproject_gen import MSProjectGenerator
from modules.vba_gen import VBAGenerator
from modules.vbs_gen import VBSGenerator
from modules.hta_gen import HTAGenerator
from modules.sct_gen import SCTGenerator
from modules.wsf_gen import WSFGenerator
from modules.visio_gen import VisioGenerator
from modules.access_gen import AccessGenerator
from modules.scf_gen import SCFGenerator
from modules.xsl_gen import XSLGenerator
from modules.url_gen import UrlShortcutGenerator
from modules.glk_gen import GlkGenerator
from modules.lnk_gen import LNKGenerator
from modules.settingsms_gen import SettingsShortcutGenerator
from modules.libraryms_gen import LibraryShortcutGenerator
from modules.inf_gen import InfGenerator
from modules.csproj_gen import CsProjGenerator
from modules.iqy_gen import IqyGenerator
from common.utils import MSTypes



class PayloadBuilderFactory():
    """ Used to provide payload generators """
    
    def _handleOfficeFormats(self, mpSession):
        """
        Handle MS Office output formats generation
        """    
        if MSTypes.XL in mpSession.outputFileType:
            generator = ExcelGenerator(mpSession)
        elif MSTypes.WD in mpSession.outputFileType:
            generator = WordGenerator(mpSession)
        elif MSTypes.PPT in mpSession.outputFileType:
            generator = PowerPointGenerator(mpSession)
        elif MSTypes.MPP == mpSession.outputFileType:
            generator = MSProjectGenerator(mpSession)
        elif MSTypes.VSD in mpSession.outputFileType:
            generator = VisioGenerator(mpSession)
        elif MSTypes.ACC in mpSession.outputFileType:
            generator = AccessGenerator(mpSession)
        
        return generator
    
    
    
    def getPayloadBuilder(self, mpSession):
        """ Build and return a PayloadGenerator  object """
        # MS Office generation/trojan is only enabled on windows
        payloadBuilder = None
        if sys.platform == "win32" and mpSession.outputFileType in MSTypes.MS_OFFICE_FORMATS:
            payloadBuilder = self._handleOfficeFormats(mpSession)
            
        if mpSession.outputFileType == MSTypes.VBS:
            payloadBuilder = VBSGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.HTA:
            payloadBuilder = HTAGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.SCT:
            payloadBuilder = SCTGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.WSF:
            payloadBuilder = WSFGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.XSL:
            payloadBuilder = XSLGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.LNK:
            payloadBuilder = LNKGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.VBA:
            payloadBuilder = VBAGenerator(mpSession)

        if mpSession.outputFileType == MSTypes.SCF:
            payloadBuilder = SCFGenerator(mpSession)

        if mpSession.outputFileType == MSTypes.URL:
            payloadBuilder = UrlShortcutGenerator(mpSession)

        if mpSession.outputFileType == MSTypes.GLK:
            payloadBuilder = GlkGenerator(mpSession)


        if mpSession.outputFileType == MSTypes.SETTINGS_MS:
            payloadBuilder = SettingsShortcutGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.LIBRARY_MS:
            payloadBuilder = LibraryShortcutGenerator(mpSession)

            
        if mpSession.outputFileType == MSTypes.INF:
            payloadBuilder = InfGenerator(mpSession)
        if mpSession.outputFileType == MSTypes.CSPROJ:
            payloadBuilder = CsProjGenerator(mpSession)
            
        if mpSession.outputFileType == MSTypes.IQY:
            payloadBuilder = IqyGenerator(mpSession)
        
        return payloadBuilder
       
       
            
            
            