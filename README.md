# ScanMeC
A PowerShell script that serves as a proof of concept for reporting files on workstations. If setup with a database that contains a table of extensions and a table of found files, the script will upload the metadata of all the files found, that have one of the "flagged" extension, to the database for querying. 

This is just a prototype/proof of concept so it is not mention for production implementation. To test with the script, create a scheduled task on the workstation that runs the PowerShell script.
