#!/usr/bin/env python
# encoding: utf-8

import os, mmap, logging,re
from common import utils
from common.utils import MSTypes
import shlex

class MpModule:
    def __init__(self,mpSession):
        self.mpSession = mpSession
        self.workingPath = mpSession.workingPath
        self._startFunction = mpSession.startFunction
        self.outputFilePath = mpSession.outputFilePath
        self.outputFileType = mpSession.outputFileType
        self.template = mpSession.template

        self.reservedFunctions = []
        if self._startFunction is not None:
            self.reservedFunctions.append(self._startFunction)
        self.reservedFunctions.append("AutoOpen")
        self.reservedFunctions.append("AutoExec")
        self.reservedFunctions.append("Workbook_Open")
        self.reservedFunctions.append("Document_Open")
        self.reservedFunctions.append("Auto_Open")
        self.reservedFunctions.append("Document_DocumentOpened")
        self.potentialStartFunctions = [
            "AutoOpen",
            "AutoExec",
            "Workbook_Open",
            "Document_Open",
            "Auto_Open",
            "Document_DocumentOpened",
        ]    
        
    @property
    def startFunction(self):
        """ Return start function, attempt to find it in vba files if _startFunction is not set """
        result = None
        if self._startFunction is not None:
            result =  self._startFunction
        else:
             
            vbaFiles = self.getVBAFiles()
            for vbaFile in vbaFiles:
                if os.stat(vbaFile).st_size != 0:
                    with open(vbaFile, 'rb', 0) as file, mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
                        for potentialStartFunction in self.potentialStartFunctions:
                            if s.find(potentialStartFunction.encode()) != -1:
                                self._startFunction = potentialStartFunction
                                if self._startFunction not in self.reservedFunctions:
                                    self.reservedFunctions.append(self._startFunction)
                                result = potentialStartFunction
                                break                
        return result
    
    
    def getVBAFiles(self):
        """ Returns path of all vba files in working dir """
        vbaFiles = []
        vbaFiles += [os.path.join(self.workingPath,each) for each in os.listdir(self.workingPath) if each.endswith('.vba')]
        return vbaFiles
    

    
    def getCMDFile(self):
        """ Return command line file (for DDE mode)"""
        if os.path.isfile(os.path.join(self.workingPath,"command.cmd")):
            return os.path.join(self.workingPath,"command.cmd")
        else:
            return ""
        

    def fillInputParams(self, paramArray):
        """ 
        Fill parameters dictionary using given input. If input is missing, ask for input to user.
        """
        mandatoryParamLen = sum(
            not paramArray[i].optional for i in range(len(paramArray))
        )

        allMandatoryParamFilled = False
        if mandatoryParamLen == 0:
            allMandatoryParamFilled = True

        cmdFile = self.getCMDFile()
        if cmdFile is not None and cmdFile != "":
            with open(cmdFile, 'r') as f:
                valuesFileContent = f.read()
                valuesFileContent = valuesFileContent.replace(r'\"', r'\\"') # This is necessary fo avoid schlex bux if string ends with \"
                logging.debug("    -> CMD file content: \n%s" % valuesFileContent)

            os.remove(cmdFile)
            if self.mpSession.fileInput is None or len(paramArray) > 1:# if values where passed by input pipe or in a file but there are multiple params
                inputValues = shlex.split(valuesFileContent)# split on space but preserve what is between quotes
            else: 
                inputValues = [valuesFileContent] # value where passed using -f option


            if len(inputValues) >= mandatoryParamLen: 
                i = 0  
                # Fill entry parameters
                while i < len(inputValues):
                    paramArray[i].value = inputValues[i]
                    i += 1
                    allMandatoryParamFilled = True

        if not allMandatoryParamFilled:
            # if input was not provided
            logging.warning("   [!] Could not find some mandatory input parameters. Please provide the next values:")
            for param in paramArray:
                
                if param.value is None or param.value == "" or param.value.isspace():
                    newValue = None
                    while newValue is None or newValue == "" or newValue.isspace():
                        newValue = input(f"    {param.name}:")
                    param.value = newValue
     

     
    
    def getMainVBAFile(self):
        """ return main vba file (the one containing macro entry point) """
        result = ""
        vbaFiles = self.getVBAFiles()
        if len(vbaFiles)==1:
            result = vbaFiles[0]
        elif self.startFunction is not None:
            for vbaFile in vbaFiles:
                if os.stat(vbaFile).st_size != 0:
                    with open(vbaFile, 'rb', 0) as file, mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
                        if s.find(self.startFunction.encode()) != -1:
                            result = vbaFile
                            break
        logging.debug(f"    [*] Start function:{self.startFunction}")
        logging.debug(f"    [*] Main VBA file:{result}")
        if not self.startFunction:
            logging.error("   [!] Error: Could not find a start function, please use --start-function=FUNCTION_NAME")
        return result
    
    
    def getFileContainingString(self, strToFind):
        """ Returns fist VB file containing the string to find"""
        result = ""
        vbaFiles = self.getVBAFiles()

        for vbaFile in vbaFiles:
            if os.stat(vbaFile).st_size != 0:
                with open(vbaFile, 'rb', 0) as file, mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
                    if s.find(strToFind.encode()) != -1:
                        result = vbaFile
                        break
                            
        return result
    
    
    def addVBAModule(self, moduleContent, moduleName=None):
        """ 
        Add a new VBA module file containing moduleContent and with random name
        Returns name of new VBA file
        """
        if moduleName is None:
            moduleName = utils.randomAlpha(9)
            modulePath = os.path.join(self.workingPath, f"{moduleName}.vba")
        else:
            modulePath = os.path.join(self.workingPath, f"{utils.randomAlpha(9)}.vba")
        if moduleName in self.mpSession.vbModulesList:
            logging.debug(f"    [,] {moduleName} module already loaded")
        else:
            logging.debug(f"    [,] Adding module: {moduleName}")
            self.mpSession.vbModulesList.append(moduleName)
            with open(modulePath, 'w') as f:
                f.write(moduleContent)
        return modulePath
    
    
    def addVBLib(self, vbaLib):
        """ 
        Add a new VBA Library module depending on the current context 
        """
        if self.outputFileType in MSTypes.MS_OFFICE_FORMATS:
            if MSTypes.WD in self.outputFileType and hasattr(vbaLib, 'VBA_WD'):
                moduleContent = vbaLib.VBA_WD
            elif MSTypes.XL in self.outputFileType and hasattr(vbaLib, 'VBA_XL'): 
                moduleContent = vbaLib.VBA_XL
            elif MSTypes.PPT in self.outputFileType and hasattr(vbaLib, 'VBA_PPT'): 
                moduleContent = vbaLib.VBA_PPT
            else:
                moduleContent = vbaLib.VBA
        elif self.outputFileType in [MSTypes.HTA, MSTypes.SCT] and hasattr(vbaLib, 'VBS_HTA'):
            moduleContent = vbaLib.VBS_HTA
        else:
            moduleContent = vbaLib.VBS if hasattr(vbaLib, 'VBS') else vbaLib.VBA
        return self.addVBAModule(moduleContent, vbaLib.__name__)


    def getVBLibContent(self, vbaLib):
        """
        Return VBA code Library module depending on the current context
        """
        moduleContent = ''
        if self.outputFileType in MSTypes.MS_OFFICE_FORMATS:
            if MSTypes.WD in self.outputFileType and hasattr(vbaLib, 'VBA_WD') and not self.mpSession.runInExcel:
                moduleContent = vbaLib.VBA_WD
            elif hasattr(vbaLib, 'VBA_XL') and (MSTypes.XL in self.outputFileType or self.mpSession.runInExcel):
                logging.debug(f"     [,] Add Excel version of module {vbaLib} ")
                moduleContent = vbaLib.VBA_XL
            elif MSTypes.PPT in self.outputFileType and hasattr(vbaLib, 'VBA_PPT'):
                moduleContent = vbaLib.VBA_PPT
            else:
                moduleContent = vbaLib.VBA
        elif self.mpSession.runInExcel:
            moduleContent = vbaLib.VBA
        elif self.outputFileType in [MSTypes.HTA, MSTypes.SCT] and hasattr(vbaLib, 'VBS_HTA'):
            logging.debug(f"     [,] Add HTA version of module {vbaLib} ")
            moduleContent = vbaLib.VBS_HTA
        else:
            moduleContent = vbaLib.VBS if hasattr(vbaLib, 'VBS') else vbaLib.VBA
        return moduleContent
    
    
    def insertVbaCode(self, targetModule, targetFunction,targetLine, vbaCode):
        """
        Insert some code at targetLine (number) at targetFunction in targetModule
        """
        logging.debug(
            f"     [,] Opening {targetFunction} in {targetModule} to inject code..."
        )

        with open(targetModule) as f:
            content = f.readlines()
        for n,line in enumerate(content):
            if matchObj := re.match(
                r'.*(Sub|Function)\s+%s\s*\(.*\).*' % targetFunction,
                line,
                re.M | re.I,
            ):
                logging.debug(f"     [,] Found {targetFunction} ")
                content[n+targetLine] = content[n+targetLine]+ vbaCode+"\n"
                break


        logging.debug("     [,] New content: \n " + "".join(content) + "\n\n ")
        with open(targetModule, 'w') as f:
            f.writelines(content)
    
    
    def getAutoOpenFunction(self):
        """ Return the VBA Function/Sub name which triggers auto open for the current outputFileType """
        if MSTypes.WD in self.outputFileType:
            return "AutoOpen"
        elif MSTypes.XL in self.outputFileType:
            return "Workbook_Open"
        elif MSTypes.PPT in self.outputFileType:
            return "AutoOpen"
        elif MSTypes.MPP in self.outputFileType:
            return "Auto_Open"
        elif MSTypes.VSD in self.outputFileType:
            return "Document_DocumentOpened"
        elif MSTypes.ACC in self.outputFileType:
            return "AutoExec"
        elif MSTypes.PUB in self.outputFileType:
            return "Document_Open"
        else:
            return "AutoOpen"
            
    
        
    def run(self):
        """ Run the module """
        raise NotImplementedError
    


    

    