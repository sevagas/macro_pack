
import logging

from modules.mp_module import MpModule
from http.server import HTTPServer, SimpleHTTPRequestHandler, HTTPStatus
from common.utils import getHostIp
from functools import partial
import urllib

class WebServer(SimpleHTTPRequestHandler):

    def do_GET(self):
        super().do_GET()
    
    def do_POST(self):
        content_len = int(self.headers.get("Content-Length"), 0)
        raw_body = self.rfile.read(content_len)
        parsed_input = urllib.parse.parse_qs(raw_body.decode('utf-8'))
        try:
            clientId = parsed_input['id'][0]
            cmdOutput = parsed_input['cmdOutput'][0]
            logging.info("   [-] From %s(%s) received:\n %s " % (self.address_string(),clientId,cmdOutput))
            self.send_response_only(HTTPStatus.OK,"OK")
        except Exception:
            self.send_error(HTTPStatus.INTERNAL_SERVER_ERROR, "Error")
        
        

class ListenServer(MpModule):


    def __init__(self, mpSession):
        self.listenPort = mpSession.listenPort
        self.listenRoot = mpSession.listenRoot
        MpModule.__init__(self, mpSession)

    def run(self):
        """ Starts listening server"""

        logging.info (" [+] Starting Macro_Pack web server...")
        logging.info ("   [-] Files in \"" + self.listenRoot + "\" folder are accessible via http://{ip}:{port}/".format(ip=getHostIp(), port=self.listenPort))
        logging.info ("   [-] Listening on port %s (ctrl-c to exit)..." % self.listenPort)
        handler_class = partial(WebServer, directory=self.listenRoot)
        httpdServer = HTTPServer(("0.0.0.0", self.listenPort), handler_class)
        httpdServer.serve_forever()

