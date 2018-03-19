# FileInfoUpdate

### Purpose
* Read or update PE (EXE/DLL/OCX/...) version information
* Limitation: The updated strings cannot be bigger than the existing information

### Usage
Read version information:  (output in rc format)
```powershell
.\FileInfoUpdate.exe MyExecutable.exe
```

Update version information strings
```powershell
.\FileInfoUpdate.exe MyExecutable.exe /fv 1.2.3.4 /pv 1.2.3.4
```

Update version information strings
```powershell
.\FileInfoUpdate.exe MyExecutable.exe /s FileDescription "Updated FileDescription" /s Comments "Updated comments"
```

### Return values
| Code | Explanation                 |
| ----:| --------------------------- |
| 0    | Succeeded                   |
| 1    | Argument error              |
| 2    | Could not access executable |
| 3    | Version information string too long |
| 4    | Failed to read version information from executable (invalid size) |
| 5    | Failed to read version information from executable |
| 6    | Failed to read fixed version info |
| 7    | Failed to read translations |
| 8    | Failed to set version information string, string doesn't exist in original version information |
| 9    | Failed to set version information string, string in original version information is too short |
| 10   | Failed to start update resources  |
| 11   | Failed to update resources  |
| 12   | Failed to commit resources  |
