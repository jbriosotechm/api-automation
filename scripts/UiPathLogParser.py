import glob
import json
import os
import sys
import subprocess

LOG_LEVELS = ["TRACE", "INFORMATION", "WARNING", "ERROR", "FATAL"]

class UiPathLogParser:
    def __init__(self, logDirectory, logLevel, isTailing):
        self.logDirectory = logDirectory
        self.logLevel = LOG_LEVELS.index(logLevel.upper())
        self.isTailing = eval(isTailing)
        self.logFile = self.getLatestFile()
        self.lastLine = ""
        sys.stdout.flush()

        print("Reading log: " + self.logFile)

        if self.isTailing:
            open(self.logFile, 'w').close() # Clears Log File

        self.file = open(self.logFile, 'r')

    def getLatestFile(self):
        listOfFiles = glob.glob(os.path.join(self.logDirectory, '*Execution.log'))

        if not listOfFiles:
            print("No Execution Log File in " + self.logDirectory)
            sys.exit(-1)

        latest = max(listOfFiles, key=os.path.getctime)
        return latest

    def isProcessRunning(self):
        isRunning = False
        getProcessCommandrobot = 'tasklist /fo csv | findstr /i "UiRobot.exe" | wc -l'

        processCount = subprocess.check_output(getProcessCommandrobot, shell=True)

        if "1" == processCount.strip():
            isRunning = True

        return isRunning

    def parseLog(self, logLine):
        if "{" in logLine:
            index = logLine.index("{")
            jsonLine = json.loads(logLine[index::])

            if self.logLevel <= LOG_LEVELS.index(jsonLine["level"].upper()):
                print ("[" + jsonLine["level"].upper() + "] " + jsonLine["message"])

    def readFile(self):
        lines = self.file.readlines()
        for line in lines:
            if self.lastLine != line:
                self.parseLog(line)
            self.lastLine = line
        self.file.close()

    def tailFile(self):
        while self.isProcessRunning():
            line = self.file.readline()
            if self.lastLine != line:
                self.parseLog(line)
            self.lastLine = line
        self.file.close()

    def start(self):
        if self.isTailing:
            self.tailFile()
        else:
            self.readFile()

    def getLogFile(self):
        return self.logFile

if __name__=="__main__":
    if 4 > len(sys.argv):
        print "\n"
        print "***********************************************************************"
        print "Usage : python <Python Path> <Log Directory> <Log Level> <Will Tail?>"
        print "***********************************************************************"
        print "\n"
        sys.exit(-1)

    logDirectory = sys.argv[1]

    if sys.argv[2].upper() not in LOG_LEVELS:
        print "Log Level should be in " + ", ".join(LOG_LEVELS)
        sys.exit(-1)
    logLevel = sys.argv[2]

    try:
        eval(sys.argv[3])
    except Exception as e:
        raise TypeError('Param 3 should be "True" or "False"')
    willTail = sys.argv[3]

    parser = UiPathLogParser(logDirectory=logDirectory, logLevel=logLevel, isTailing=willTail)
    parser.start()