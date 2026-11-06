# SQL Server Multi-Connection Tester

A Windows Forms utility for testing SQL Server connectivity, opening and managing multiple connections, running queries on open connections, viewing results, and exporting results to CSV.

## Requirements

- Windows
- Windows PowerShell 5.1
- Access to a SQL Server instance and authorization to test it

The script uses .NET Framework assemblies included with Windows PowerShell and requires no additional packages.

## Run

From this directory, launch the application in an STA PowerShell process:

```powershell
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File .\SqlMultiConnect.ps1
```

Enter the server and authentication details in the connection section. Use **Test Connection** to check one connection, or set the connection count and choose **Open N Connections**. Open connections remain active until closed in the UI or when the application exits.

## Caution

Opening many connections and running queries can create significant load on the target server. Use only against servers you are authorized to test, and agree on safe connection counts before testing production systems.