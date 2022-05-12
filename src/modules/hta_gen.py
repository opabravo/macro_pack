#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.vbs_gen import VBSGenerator

HTA_TEMPLATE = \
r"""
<!DOCTYPE html>
<html>
<head>
<HTA:APPLICATION icon="#" WINDOWSTATE="normal" SHOWINTASKBAR="no" SYSMENU="no"  CAPTION="no" BORDER="none" SCROLL="no" />
<script type="text/vbscript">
<<<VBS>>>

window.resizeTo 0,0
<<<MAIN>>>
Close
</script>
</head>
<body>
</body>
</html>

"""

class HTAGenerator(VBSGenerator):
    """ Module used to generate HTA file from working dir content"""
        
    def vbScriptConvert(self):
        super().vbScriptConvert()
        with open(f"{self.getMainVBAFile()}.vbs") as f:
            vbsContent = f.read()
        logging.info("   [-] Convert VBScript to HTA...")
        vbsContent = vbsContent.replace("WScript.Echo ", "MsgBox ")
        vbsContent = vbsContent.replace('WScript.Sleep(1000)','CreateObject("WScript.Shell").Run "cmd /c ping localhost -n 1",0,True')
        vbsContent = vbsContent.replace('Wscript.Quit 0', 'Self.Close')
        vbsContent = vbsContent.replace('Wscript.ScriptFullName', 'self.location.pathname')

        with open(f"{self.getMainVBAFile()}.vbs", 'w') as f:
            f.writelines(vbsContent)
        
        
    def generate(self):
        logging.info(f" [+] Generating {self.outputFileType} file...")
        self.vbScriptConvert()
        with open(f"{self.getMainVBAFile()}.vbs") as f:
            vbsContent = f.read()
        # Write VBS in template
        htaContent = HTA_TEMPLATE
        htaContent = htaContent.replace("<<<VBS>>>", vbsContent)
        htaContent = htaContent.replace("<<<MAIN>>>", self.startFunction)
        with open(self.outputFilePath, 'w') as f:
            f.writelines(htaContent)
        logging.info(f"   [-] Generated HTA file: {self.outputFilePath}")
        logging.info("   [-] Test with : \nmshta %s\n" % self.outputFilePath)
        

        
        
        