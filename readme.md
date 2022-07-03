# Windows-Service-CronJob

Example of a windows servcie with schedule implemented in python.
The schedule using the `croniter` to parse a complex cron schedule string, And repeat a function everyday at a specific time.

The servcie should be build in a python environment that inclouds the `pywin32`, `pyinstaller`, `croniter` packages.

The command for building the service is:
```bash
pyinstaller --clean -y -n "main" .\main.py
> ...
> INFO: Building EXE from EXE-00.toc completed successfully.
```

Install Win-Service
```bash
.\dist\main\main.exe install
> Installing service MyWinService_CronJob
> Service installed
```

Run Win-Service
```bash
.\dist\main\main.exe start
> Starting service MyWinService_CronJob
```

Stop Win-Service
```bash
.\dist\main\main.exe stop
> Stoping service MyWinService_CronJob
```

Remove Win-Service
```bash
.\dist\main\main.exe stop
> Removing service MyWinService_CronJob
```

![Event_View](images/RunJob.png?raw=true "Title")
