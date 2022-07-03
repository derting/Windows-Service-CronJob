import win32timezone
import win32serviceutil
import win32service
import servicemanager

import time
import os
import pandas as pd
from croniter import croniter
from datetime import datetime, timedelta

configDir = 'C:\Program Files\MyWinService\CronJob';
configPath = configDir + '\config.csv'


def roundDownTime(dt=None, dateDelta=timedelta(minutes=1)):
    roundTo = dateDelta.total_seconds()
    if dt == None:
        dt = datetime.now()
    seconds = (dt - dt.min).seconds
    rounding = (seconds+roundTo/2) // roundTo * roundTo
    return dt + timedelta(0, rounding-seconds, -dt.microsecond)


def getPrevCronRunTime(schedule):
    return croniter(schedule, datetime.now()).get_prev(datetime)

def getCurrCronRunTime(schedule):
    return croniter(schedule, datetime.now()).get_current(datetime)

def getNextCronRunTime(schedule):
    return croniter(schedule, datetime.now()).get_next(datetime)

def sleepTillTopOfNextMinute():
    t = datetime.utcnow()
    sleeptime = 60 - (t.second + t.microsecond/1000000.0)
    time.sleep(sleeptime)


class ImplService:
    def stop(self):
        self.running = False
        servicemanager.LogInfoMsg("Service stoped.")

    def run(self):
        self.running = True
        servicemanager.LogInfoMsg("Service running.")

        while self.running:    
            config = getConfiguration()
            prevRunTime = getPrevCronRunTime(config.cronJob)
            currRunTime = getCurrCronRunTime(config.cronJob)
            nextRunTime = getNextCronRunTime(config.cronJob)
            roundedDownTime = roundDownTime()
            pRT = prevRunTime.strftime("%m/%d/%Y, %H:%M:%S")
            cRT = currRunTime.strftime("%m/%d/%Y, %H:%M:%S")
            nRT = nextRunTime.strftime("%m/%d/%Y, %H:%M:%S")
            rDT = roundedDownTime.strftime("%m/%d/%Y, %H:%M:%S")
            servicemanager.LogInfoMsg("The job will run in next time. (prevRunTime: " + pRT + ", currRunTime: " + cRT + ", nextRunTime: " + nRT + ", roundedDownTime: " + rDT + ")")

            if (rDT == pRT):
                servicemanager.LogInfoMsg("Time equal, run job (" + config.cronJob + ").")
                nextRunTime = getNextCronRunTime(config.cronJob)
                
            sleepTillTopOfNextMinute()

class SVCHandler(win32serviceutil.ServiceFramework):
    _svc_name_ = 'MyWinService_CronJob'
    _svc_display_name_ = 'MyWinService_CronJob'

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.service_impl.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
      self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
      self.service_impl = ImplService()
      self.ReportServiceStatus(win32service.SERVICE_RUNNING)
      self.service_impl.run()

class Config:
    def __init__(self, cronJob):
        self.cronJob = cronJob

def getConfiguration():
    if os.path.isfile(configPath) == False:
        if not os.path.exists(configDir):
            os.makedirs(configDir)

        file = open(configPath,'w+')
        file.writelines([
            "cronJob\n",
            "*/5 * * * *"]);
        file.close()
    
    config = pd.read_csv(configPath)
    cronJob = config.loc[0, "cronJob"]
    return Config(cronJob)

def ServiceInitial():
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(SVCHandler)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(SVCHandler)

if __name__ == '__main__':
    ServiceInitial()