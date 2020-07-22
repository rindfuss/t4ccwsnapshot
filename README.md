# Tools4Church / ChurchWindows Snapshot Export

## Source code highlights:
- .NET Framework  
- Multi-threaded program so that UI progress updates display while another thread blocks waiting for data from the database.
- Data cleanup and mapping built into SQL queries.
- Use of SQL Server, csv, and zip libraries.
- Log files track progress and aid in troubleshooting.

## Background
A church client was using ChurchWindows, a Microsoft Windows application for tracking donor data. The application consisted of a SQL Server Express database running on an employees' desktop PC and accessible over the LAN to client applications running on the employee's PC and others around the church office. Project was to provide a simple way to extract donation data from the database for import into [Tools4Church](https://tools4church.com), a cloud-based financial analytics application I had written. 

Export utilities built into the application were insufficient. Configuring the client's IT infrastructure for remote database access involved logistical difficulties and raised security concerns. Also, the employee desktop PC hosting the database might be turned off at unpredictable intervals! I hadn't built a native Windows application in a *long* time, so I set out to learn enough .NET to create an application the client could use to generate an export file that could be uploaded to my cloud-based analytics app. 

After some research and trial-and-error I chose to work with .NET Framework. It offered libraries for querying SQL Server, generating .csv files, and compressing the .csv files into one .zip file. It also enabled distributing .NET Framework and my app as a single .exe file to the client without requiring code-signing or asking the client to ensure .NET Framework was already installed.

> Written with [StackEdit](https://stackedit.io/).