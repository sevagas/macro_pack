
#!/usr/bin/python3.4
# encoding: utf-8

from cheroot import wsgi
from wsgidav.dir_browser import WsgiDavDirBrowser
from wsgidav.error_printer import ErrorPrinter
from wsgidav.http_authenticator import HTTPAuthenticator
from wsgidav.request_resolver import RequestResolver
from wsgidav.wsgidav_app import WsgiDAVApp
from common.utils import getHostIp
# Import Needed modules

import logging
from modules.mp_module import MpModule



class WListenServer(MpModule):


    def __init__(self, mpSession):
        self.WRoot = mpSession.WRoot
        self.listenPort = mpSession.listenPort
        MpModule.__init__(self, mpSession)



    def run(self):
        """ Starts listening server"""

        logging.info (" [+] Starting Macro_Pack WebDAV server...")
        logging.info ("   [-] Files in \"" + self.WRoot + r"\" folder are accessible using url http://{ip}:{port}  or \\{ip}@{port}\DavWWWRoot".format(ip=getHostIp(), port=self.listenPort))
        logging.info ("   [-] Listening on port %s (ctrl-c to exit)...", self.listenPort)

        # Prepare WsgiDAV config
        config = {

            'verbose': 3,

            'add_header_MS_Author_Via': True,
            "hotfixes": {
                "emulate_win32_lastmod": False,  # True: support Win32LastModifiedTime
                "re_encode_path_info": True,  # (See issue #73)
                "unquote_path_info": False,  # (See issue #8, #228)
                "win_accept_anonymous_options": True,
            },
            'host': '0.0.0.0',

            'port': self.listenPort,    # Specifying listening port
            'provider_mapping': {'/': self.WRoot},  #Specifying root folder
            "middleware_stack": [
                # WsgiDavDebugFilter,
                ErrorPrinter,
                #HTTPAuthenticator,
                WsgiDavDirBrowser,  # configured under dir_browser option (see below)
                RequestResolver,  # this must be the last middleware item
            ],
            # HTTP Authentication Options
            "http_authenticator": {
                # None: dc.simple_dc.SimpleDomainController(user_mapping)
                "domain_controller": None,
                "accept_anonymous": True,  # Allow basic authentication, True or False
                #"accept_digest": True,  # Allow digest authentication, True or False
                #"default_to_digest": True,  # True (default digest) or False (default basic)
                # Name of a header field that will be accepted as authorized user
                "trusted_auth_header": None,

            },
            'dir_browser': {
                'davmount': False,
                'enable': True,  # Enabling directory browsing on dir_browser
                # List of fnmatch patterns:
                "ignore": [
                    ".DS_Store",  # macOS folder meta data
                    "._*",  # macOS hidden data files
                    "Thumbs.db",  # Windows image previews
                ],
                'ms_mount': False,
                'show_user': True,
                'ms_sharepoint_support': True,
                'ms_sharepoint_urls': False,
                'response_trailer': True,
            },
        }

        app = WsgiDAVApp(config)

        server_args = {
            "bind_addr": (config["host"], config["port"]),
            "wsgi_app": app,
        }
        server = wsgi.Server(**server_args)
        
        try:
            log = logging.getLogger('wsgidav')
            log.raiseExceptions = False # hack to avoid false exceptions
            log.propagate = True
            log.setLevel(logging.INFO)
            server.start()
        except KeyboardInterrupt:
            logging.info("  [!] Ctrl + C detected, closing WebDAV sever")
            server.stop()
