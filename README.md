# Add-Fonts
Adds TrueType &amp; OpenType fonts to the Windows registry and "C:\Windows\Fonts" folder location based on a font's metadata properties.

## Installation
Requires an administration account to run this code in a powershell terminal; however a seperate `.bat` file can be created to execute the following code without a Powershell terminal:
```batch
powershell.exe -ExecutionPolicy UnRestricted -File %~dp0file-name.ps1
pause
```

**Note:** Ensure both `.bat` & `.ps1` files reside in the same file directory or folder.

## Usage
Script can be used in a Endpoint Configuration Manager to batch install a library of fonts in a large enterprise environment that manages many computers or if a domain-joined user needs to install locally downloaded fonts to their computer.

### Features
- Limits unnecessary writes to the system registry and "C:\Windows\Fonts" file directory by verifiying installed fonts using Microsoft's "GlyphTypeface" object instead of the font's file name.
- Accepts a "piped" input, such as a "Read-Host" command for user input to enter a file path of a font's location.
- Will automatically search a (user-provided) file directory for the existenance of `.otf` & `.ttf` file types to begin the installation process.


## Etcetera
Please make use of the "helper" code built-in to the script to learn more about syntax and view examples.
- Use the following command: 
```powershell
Get-Help C:\file-dir\file-name.ps1 -full
```
