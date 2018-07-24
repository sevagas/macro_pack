
#!/usr/bin/python3.4
# encoding: utf-8


from cheroot import wsgi
from wsgidav.dir_browser import WsgiDavDirBrowser
from wsgidav.debug_filter import WsgiDavDebugFilter
from wsgidav.error_printer import ErrorPrinter
from wsgidav.http_authenticator import HTTPAuthenticator
from wsgidav.request_resolver import RequestResolver
from wsgidav.wsgidav_app import WsgiDAVApp
import os
# Import Needed modules

import logging

from modules.mp_module import MpModule
from common.utils import getRunningApp

class WListenServer(MpModule):


    def __init__(self, mpSession):
        self.WRoot = mpSession.WRoot
        MpModule.__init__(self, mpSession)
    
    def run(self):
        """ Starts listening server"""

        logging.info (" [+] Starting Macro_Pack WebDAV server...")
        logging.info ("   [-] Files in current folder are accessible using http://<hostname>:%s/u/" % self.WRoot)
        logging.info ("   [-] Listening on port 80 (ctrl-c to exit)...")
        
        # Run web server in another thread
        config = {
            "host": "localhost",
            "port": 8080,
            "provider_mapping": {
            "/": self.WRoot,
            },
            "verbose": 1,
            "dir_browser": {
                "enable": True,
            }
        }

        app = WsgiDAVApp(config)

        server_args = {
            "bind_addr": (config["host"], config["port"]),
        "wsgi_app": app,
        }
        server = wsgi.Server(**server_args)
        server.start()


