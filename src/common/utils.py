#!/usr/bin/env python
# encoding: utf-8


from random import choice
import string
import logging
from termcolor import colored
import os

def randomAlpha(length):
    """ Returns a random alphabetic string of length 'length' """
    key = ''
    for i in range(length): # @UnusedVariable
        key += choice(string.ascii_lowercase)
    return key


def guessApplicationType(documentPath):
    """ Guess MS office application type based on extension """
    result = ""
    extension = os.path.splitext(documentPath)[1]
    if "xls" in extension:
        result = "Excel"
    elif "doc" in  extension:
        result = "Word"
    elif "ppt" in extension:
        result = "PowerPoint"
    return result


class ColorLogFiler(logging.StreamHandler):
    """ Override logging class to enable terminal colors """
    def emit(self, record):
        try:
            msg = self.format(record)
            msg = msg.replace("[+]",colored("[+]", "green"))
            msg = msg.replace("[-]",colored("[-]", "green"))
            msg = msg.replace("[!]",colored("[!]", "red"))
            stream = self.stream
            stream.write(msg)
            stream.write(self.terminator)
            self.flush()
        except Exception:
            self.handleError(record)
    