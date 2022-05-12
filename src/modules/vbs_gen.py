#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.vba_gen import VBAGenerator
import re, os


VBS_TEMPLATE = \
r"""
<<<VBS>>>
<<<MAIN>>>
"""

class VBSGenerator(VBAGenerator):
    """ Module used to generate VBS file from working dir content"""
    
    def check(self):
        logging.info("   [-] Check if VBA->VBScript is possible...")
        # Check nb of source file
        vbaFiles = self.getVBAFiles()
        if len(vbaFiles)>1:
            logging.warning("   [!] This module has more than one source file. They will be concatenated into a single VBS file.")


        for vbaFile in vbaFiles:
            with open(vbaFile) as f:
                content = f.readlines()
            # Check there are no call to Application object
            for line in content:
                if line.find("Application.Run") != -1:
                    logging.error("   [-] You cannot access Application object in VBScript. Abort!")
                    logging.error(f"   [-] Line: {line}")
                    return False  

            # Check there are no DLL import
            for line in content:
                if matchObj := re.match(
                    r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"\s*.*',
                    line,
                    re.M | re.I,
                ):
                    logging.error("   [-] VBScript does not support DLL import. Abort!")
                    logging.error(f"   [-] Line: {line}")
                    return False
        return True
    
    
    
    def printFile(self):
        """ Display generated code on stdout """
        if os.path.isfile(self.outputFilePath):
            logging.info(" [+] Generated file content:\n") 
            with open(self.outputFilePath,'r') as f:
                print(f.read())
    
    
    def vbScriptConvert(self):
        logging.info("   [-] Convert VBA to VBScript...")
        translators = [
            ("Val(", "CInt("),
            (" Chr$", " Chr"),
            (" Mid$", " Mid"),
            ("On Error GoTo", "'//On Error GoTo"),
            ("byebye:", ""),
            ("Next ", "Next '//"),
            *[("() As String", " "), ("CVar", "")],
            *[
                (" As String", " "),
                (" As Object", " "),
                (" As Long", " "),
                (" As Integer", " "),
                (" As Variant", " "),
                (" As Boolean", " "),
                (" As Byte", " "),
                (" As Excel.Application", " "),
                (" As Word.Application", " "),
            ],
            *[
                ("MsgBox ", "WScript.Echo "),
                (
                    'Application.Wait Now + TimeValue("0:00:01")',
                    'WScript.Sleep(1000)',
                ),
            ],
            *[('ChDir ', 'createobject("WScript.Shell").currentdirectory =  ')],
        ]

        content = []
        for vbaFile in self.getVBAFiles():  
            with open(vbaFile) as f:
                content.extend(f.readlines())
            if not content[-1].endswith("\n"): # Add \n to avoid overlap
                content[len(content) - 1] += "\n"
        isUsingEnviron = False
        for n,line in enumerate(content):
            # Do easy translations
            for translator in translators:
                line = line.replace(translator[0],translator[1])

            if line.strip() == "End":
                line = "Wscript.Quit 0 \n"

            # Check if ENVIRON is used
            if line.find("Environ(")!= -1:
                isUsingEnviron = True
                line = re.sub('Environ\("([A-Z_]+)"\)',r'wshShell.ExpandEnvironmentStrings( "%\1%" )', line, flags=re.I)
            content[n] = line
            # ENVIRON("COMPUTERNAME") ->   
            #Set wshShell = CreateObject( "WScript.Shell" )
            #strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

        with open(f"{self.getMainVBAFile()}.vbs", 'a') as f:
            if isUsingEnviron:
                f.write('Set wshShell = CreateObject( "WScript.Shell" )\n')
            f.writelines(content)
        
    
    def generate(self):
                
        logging.info(f" [+] Generating {self.outputFileType} file...")
        self.vbScriptConvert()

        with open(f"{self.getMainVBAFile()}.vbs") as f:
            vbsTmpContent = f.read()
        # Write VBS in template
        vbsContent = VBS_TEMPLATE
        vbsContent = vbsContent.replace("<<<VBS>>>", vbsTmpContent)
        vbsContent = vbsContent.replace("<<<MAIN>>>", self.startFunction)

        with open(self.outputFilePath, 'w') as f:
            f.writelines(vbsContent)
        logging.info(f"   [-] Generated VBS file: {self.outputFilePath}")
        logging.info("   [-] Test with : \nwscript %s\n" % self.outputFilePath)
        

        
        
        