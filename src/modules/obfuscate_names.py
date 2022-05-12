#!/usr/bin/env python
# encoding: utf-8

import re
from modules.mp_module import MpModule
from common.utils import extractStringsFromText, progressBar
import logging
from random import randint

class ObfuscateNames(MpModule):

    vbaFunctions = []
    win32Functions = []

    def _findAllFunctions(self):
        
        for vbaFile in self.getVBAFiles():
            if self.mpSession.obfOnlyMain and vbaFile != self.getMainVBAFile():
                continue
            with open(vbaFile) as f:
                content = f.readlines()
            for line in content:
                if matchObj := re.match(
                    r'.*(Sub|Function)\s+([a-zA-Z0-9_]+)\s*\(.*\).*',
                    line,
                    re.M | re.I,
                ):
                    keyword = matchObj.groups()[1]
                    if keyword not in self.reservedFunctions:
                        self.vbaFunctions.append(keyword)
                        self.reservedFunctions.append(keyword)
                elif matchObj := re.match(
                    r'.*(Sub|Function)\s+([a-zA-Z0-9_]+)\s*As\s*.*',
                    line,
                    re.M | re.I,
                ):
                    keyword = matchObj.groups()[1]
                    if keyword not in self.reservedFunctions:
                        self.vbaFunctions.append(keyword)
                        self.reservedFunctions.append(keyword)

            with open(vbaFile, 'w') as f:
                f.writelines(content)
    


    def _replaceFunctions(self):
        # Identify function, subs and variables names
        self._findAllFunctions()

        if len(self.vbaFunctions) > 0:

            # Different situation surrounding variables
            varDelimitors=[(" "," "),("\t"," "),("\t","("),("\t"," ="),(" ","("), (",","("),(", ",")"),("AddressOf ",")"),("(","("),(" ","\n"),(" ",","),(" "," ="),("."," "),(".","\"")]

            disableProgressBar = self.mpSession.logLevel == "WARN"
            for keyWord in progressBar(self.vbaFunctions, prefix='Progress:', suffix='Complete', length=50, disableProgressBar=disableProgressBar):

                keyTmp = self._generateRandomVbaName() # Generate random names with random size
                self.reservedFunctions.append(keyTmp)

                for vbaFile in self.getVBAFiles():
                    if (
                        self.mpSession.obfOnlyMain
                        and vbaFile != self.getMainVBAFile()
                    ):
                        continue
                    with open(vbaFile) as f:
                        if keyWord not in f.read(): # optimisation, do not obfuscate if function is not in file
                            f.close()
                            continue

                    with open(vbaFile) as f:
                        content = f.readlines()
                    for varDelimitor in varDelimitors:
                        newKeyWord = varDelimitor[0] + keyTmp + varDelimitor[1]
                        keywordTmp = varDelimitor[0] + keyWord + varDelimitor[1]

                        for n,line in enumerate(content):
                            if "\"" in line:
                                extractedStrings=extractStringsFromText(line)
                                if (
                                    keyWord in extractedStrings
                                    and (
                                        "Application.Run" in line
                                        or "Application.OnTime" in line
                                    )
                                    or keyWord not in extractedStrings
                                ): # dynamic function call detected
                                    content[n] = line.replace(keywordTmp, newKeyWord)
                            else:
                                content[n] = line.replace(keywordTmp, newKeyWord)

                    with open(vbaFile, 'w') as f:
                        f.writelines(content)


    def _generateRandomVbaName(self):
        """
        Generate random names with random size
        :return:
        """
        #logging.info(vbaName)
        return self.mpSession.nameObfuscationCallback(
            randint(
                self.mpSession.obfuscatedNamesMinLen,
                self.mpSession.obfuscatedNamesMaxLen,
            ),
            self.mpSession.obfuscatedNamesCharset,
        )


    def _findAllWin32Api(self):

        for vbaFile in self.getVBAFiles():
            if self.mpSession.obfOnlyMain and vbaFile != self.getMainVBAFile():
                continue
            with open(vbaFile) as f:
                content = f.readlines()
            for line in content:
                if matchObj := re.match(
                    r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"\s*.*',
                    line,
                    re.M | re.I,
                ):
                    keyword = matchObj.groups()[1]
                    if keyword not in self.reservedFunctions:
                        self.win32Functions.append(keyword)
                        self.reservedFunctions.append(keyword)



    def _replaceLibImports(self):

        self._findAllWin32Api()

        # Replace functions and function calls by random string
        for keyWord in self.win32Functions:
            
            keyTmp = self._generateRandomVbaName()
            self.reservedFunctions.append(keyTmp)
            for vbaFile in self.getVBAFiles():
                if self.mpSession.obfOnlyMain and vbaFile != self.getMainVBAFile():
                    continue
                with open(vbaFile) as f:
                    macroLines = f.readlines()
                #logging.info("Keyword:%s,keyTmp:%s "%(keyWord,keyTmp))
                for n,line in enumerate(macroLines):
                    if "Lib " in line and f"{keyWord} " in line: # take care of declaration
                        if "Alias " in line: # if fct already has an alias we can change the original keyword
                            #logging.debug(line)
                            macroLines[n] = line.replace(f" {keyWord} ", f" {keyTmp} ", 1)
                                                #logging.debug(macroLines[n])
                        else:
                            # We have to create a new alias
                            matchObj = re.match(r'.*(Sub|Function)\s*([a-zA-Z0-9_]+)\s*Lib\s*"(.+)"(\s*).*', line, re.M|re.I)
                            #logging.debug(line)
                            line = line.replace(f" {keyWord} ", f" {keyTmp} ")
                            #logging.debug(line+"\n")
                            macroLines[n] = line.replace(matchObj.groups()[2],matchObj.groups()[2] + "\" Alias \"%s" % keyWord)
                    elif matchObj := re.match(
                        r'.*".*%s.*".*' % keyWord, line, re.M | re.I
                    ):
                        if "Application.Run" in line: # dynamic function call detected
                            macroLines[n] = line.replace(keyWord, keyTmp)

                    elif f"{keyWord} " in line or f"{keyWord}(" in line:
                        #logging.info(line)
                        macroLines[n] = line.replace(keyWord, keyTmp)

                with open(vbaFile, 'w') as f:
                    f.writelines(macroLines)



    def _replaceVariables(self,macroLines):
        
        #  will contain variables names
        keyWords = []
        # format something As ...
        for line in macroLines:
            if findList := re.findall(
                r'([a-zA-Z0-9_]+)(\(\))?\s+As\s+(String|Integer|Long|Object|Byte|Variant|Range|Boolean|Single|Any|Collection|Word.Application|Excel.Application|VbVarType)',
                line,
                re.I,
            ):
                for keyWord in findList:
                    if keyWord[0] not in self.reservedFunctions: # prevent erase of previous variables and function names
                        keyWords.append(keyWord[0])
                        self.reservedFunctions.append(keyWord[0])
        # format Set <something> =  ...
        for line in macroLines:
            if findList := re.findall(r'Set\s+([a-zA-Z0-9_]+)\s+=', line, re.I):
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)
        # format Const <something> =  ...
        for line in macroLines:
            if findList := re.findall(r'Const\s+([a-zA-Z0-9_]+)\s+=', line, re.I):
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)
        # format Type <something>
        for line in macroLines:
            if findList := re.findall(r'Type\s+([a-zA-Z0-9_]+)$', line, re.I):
                for keyWord in findList:
                    if keyWord not in self.reservedFunctions:  # prevent erase of previous variables and function names
                        keyWords.append(keyWord)
                        self.reservedFunctions.append(keyWord)

        #logging.info(str(keyWords))

        # Different situation surrounding variables
        varDelimitors = [
            (" ", " "),
            (" ", "."),
            (" ", "("),
            (" ", "\n"),
            (" ", ","),
            (" ", ")"),
            (" ", " ="),
            *[
                (".", " ="),
                (".", " A"),
                (".", " O"),
                (".", ")"),
                (".", ","),
                (".", " "),
                (".", "\n"),
            ],
            *[("#", " "), ("#", ","), ("#", "\n")],
            *[
                ("\t", " "),
                ("\t", "."),
                ("\t", "("),
                ("\t", "\n"),
                ("\t", ","),
                ("\t", ")"),
                ("\t", " ="),
            ],
            *[
                ("(", ")"),
                ("(", "("),
                ("(", ","),
                ("(", " +"),
                ("(", " *"),
                ("(", " &"),
                ("(", " -"),
                ("(", " ="),
                ("(", " As"),
                ("(", " And"),
                ("(", " To"),
                ("(", " Or"),
                ("(", "."),
            ],
            *[("=", " "), ("=", ","), ("=", "\n"), ("Set ", " =")],
        ]

        # replace all keywords by random name
        for keyWord in keyWords:
            keyTmp = self._generateRandomVbaName()
            #logging.info("|%s|->|%s|" %(keyWord,keyTmp))
            for varDelimitor in varDelimitors:
                newKeyWord = varDelimitor[0] + keyTmp + varDelimitor[1]
                keywordTmp = varDelimitor[0] + keyWord + varDelimitor[1]
                firstWordNewKeyWord = keyTmp + varDelimitor[1]
                firstWordKeywordTmp = keyWord + varDelimitor[1]
                for n,line in enumerate(macroLines):
                    if line.startswith(firstWordKeywordTmp):
                        line = line.replace(firstWordKeywordTmp, firstWordNewKeyWord)
                    macroLines[n] = line.replace(keywordTmp, newKeyWord)

        return macroLines


    def _replaceConsts(self, macroLines):

        # Identify and replace constants
        constList = ["0","1", "2"]

        for constant in constList:
            # Create random string to replace constant 
            keyTmp = self._generateRandomVbaName()
            constDeclaration = f"Const {keyTmp} = {constant}" + "\n"
            macroLines.insert(0, constDeclaration)
            newKeyWord = f" {keyTmp}, "
            keywordTmp = f" {constant}, "
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord)
            newKeyWord = f",{keyTmp},"
            keywordTmp = f",{constant},"
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord)
            newKeyWord = f", {keyTmp})"
            keywordTmp = f", {constant})"
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord)
            newKeyWord = f"({keyTmp},"
            keywordTmp = f"({constant},"
            for n,line in enumerate(macroLines):
                macroLines[n] = line.replace(keywordTmp, newKeyWord)
        return macroLines


    def run(self):
        if self.mpSession.noNamesObfuscation:
            return
        logging.info(" [+] VBA names obfuscation ...")

        # Obfuscate function name
        logging.info("   [-] Rename functions...")
        self._replaceFunctions()


        logging.info("   [-] Rename variables...")

        # go through each file
        vbaFiles = self.getVBAFiles()
        disableProgressBar = self.mpSession.logLevel == "WARN"
            # for vbaFile in self.getVBAFiles():
        for vbaFile in progressBar(vbaFiles, prefix='Progress:', suffix='Complete', length=50,
                                       disableProgressBar=disableProgressBar):

            if self.mpSession.obfOnlyMain and vbaFile != self.getMainVBAFile():
                continue
            with open(vbaFile) as f:
                content = f.readlines()
            # Obfuscate variables name
            content = self._replaceVariables(content)
            if self.mpSession.ObfReplaceConstants and (
                ",0," in content or " 0," in "".join(content)
            ):
                content = self._replaceConsts(content)

            with open(vbaFile, 'w') as f:
                f.writelines(content)
        if self.mpSession.ObfReplaceConstants:
            logging.info("   [-] Rename some numeric const...")

        # replace lib imports
        if self.mpSession.obfuscateDeclares:
            logging.info("   [-] Rename API imports...")
            self._replaceLibImports()
        logging.info("   [-] OK!")
        
