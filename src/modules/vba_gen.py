#!/usr/bin/env python
# encoding: utf-8

import logging
from modules.payload_builder import PayloadBuilder
import shutil
import os
from modules.obfuscate_names import ObfuscateNames
from modules.obfuscate_form import ObfuscateForm
from modules.obfuscate_strings import ObfuscateStrings
from modules.uac_bypass import UACBypass


class VBAGenerator(PayloadBuilder):
    """ Module used to generate VBA file from working dir content"""
        
        
    def transformAndObfuscate(self):
        """ 
        Call this method to apply transformation and obfuscation on the content of temp directory 
        This method does obfuscation for all VBA and VBA like types
        
        """
        
        # Enable UAC bypass
        if self.mpSession.uacBypass:
            uacBypasser = UACBypass(self.mpSession)
            uacBypasser.run()
        
        # Macro obfuscation
        if self.mpSession.obfuscateNames:
            obfuscator = ObfuscateNames(self.mpSession)
            obfuscator.run()
        # Mask strings
        if self.mpSession.obfuscateStrings:
            obfuscator = ObfuscateStrings(self.mpSession)
            obfuscator.run()
        # Macro obfuscation
        if self.mpSession.obfuscateForm:
            obfuscator = ObfuscateForm(self.mpSession)
            obfuscator.run() 


    
    def check(self):
        return True
    
    
    def printFile(self):
        """ Display generated code on stdout """
        logging.info(" [+] Generated VB code:\n")
        if len(self.getVBAFiles())==1: 
            vbaFile = self.getMainVBAFile() 
            with open(vbaFile,'r') as f:
                print(f.read())
        else:
            logging.info("   [!] More then one VB file generated")
            for vbaFile in self.getVBAFiles():
                with open(vbaFile,'r') as f:
                    print(f" =======================  {vbaFile}  ======================== ")
                    print(f.read())
                    
    
    def generate(self):
        if len(self.getVBAFiles()) <= 0:
            return
        logging.info(" [+] Analyzing generated VBA files...")
        if len(self.getVBAFiles())==1:
            shutil.copy2(self.getMainVBAFile(), self.outputFilePath)
            logging.info(f"   [-] Generated VBA file: {self.outputFilePath}")
        else:
            logging.info(
                f"   [!] More then one VBA file generated, files will be copied in same dir as {self.outputFilePath}"
            )

            for vbaFile in self.getVBAFiles():
                if vbaFile != self.getMainVBAFile():
                    shutil.copy2(vbaFile, os.path.join(os.path.dirname(self.outputFilePath),os.path.basename(vbaFile)))
                    logging.info(
                        f"   [-] Generated VBA file: {os.path.join(os.path.dirname(self.outputFilePath), os.path.basename(vbaFile))}"
                    )

                else:
                    shutil.copy2(self.getMainVBAFile(), self.outputFilePath)
                    logging.info(f"   [-] Generated VBA file: {self.outputFilePath}")
                        
    
    def getAutoOpenVbaFunction(self):
        return "AutoOpen"

    def getAutoOpenVbaSignature(self):
        return "Sub AutoOpen()"
            
    def resetVBAEntryPoint(self):
        """
        If macro has an autoopen like mechanism, this will replace the entry_point with what is given in newEntrPoin param
        Ex for Excel it will replace "Sub AutoOpen ()" with "Sub Workbook_Open ()"
        """
        mainFile = self.getMainVBAFile()
        if (
            mainFile != ""
            and self.startFunction is not None
            and self.startFunction != self.getAutoOpenVbaFunction()
            and self.startFunction in self.potentialStartFunctions
        ):
            logging.info(
                f"   [-] Changing auto open function from {self.startFunction} to {self.getAutoOpenVbaFunction()}..."
            )

            with open(mainFile) as f:
                content = f.readlines()
            for n,line in enumerate(content):
                if line.find(f" {self.startFunction}") != -1:  
                    #logging.info("     -> %s becomes %s" %(content[n], self.getAutoOpenVbaSignature()))  
                    content[n] = self.getAutoOpenVbaSignature() + "\n"
            with open(mainFile, 'w') as f:
                f.writelines(content)
            # 2 Change  current module start function
            self._startFunction = self.getAutoOpenVbaFunction()
                            
    