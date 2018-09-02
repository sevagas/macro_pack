
#!/usr/bin/python3.4
# encoding: utf-8

from cheroot import wsgi
from wsgidav.dir_browser import WsgiDavDirBrowser
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
            'middleware_stack' : {
                WsgiDavDirBrowser,  #Enabling dir_browser middleware
            },
            'verbose': 3,

            'add_header_MS_Author_Via': True,
            'unquote_path_info': False,
            're_encode_path_info': None,
            'host': '0.0.0.0',
            'dir_browser': {'davmount': False,
                'enable': True, #Enabling directory browsing on dir_browser
                'ms_mount': False,
                'show_user': True,
                'ms_sharepoint_plugin': True,
                'ms_sharepoint_urls': False,
                'response_trailer': False},
            'port': self.listenPort,    # Specifying listening port
            'provider_mapping': {'/': self.WRoot}   #Specifying root folder
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
