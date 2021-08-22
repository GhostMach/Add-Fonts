
# Authored by: Adam H. Meadvin 
# Email: h3rbert@protonmail.ch
# GitHub: @GhostMach 
# Creation Date: 19 June 2021


<#
.SYNOPSIS
Function: Add-Fonts
Paramaters: 'FontPathParam'
ValueFromPipeline: True
.DESCRIPTION
- Adds fonts with file extension .OTF and .TTF, ignoring other file types that may reside in same file directory.
- Will NOT install fonts found in C:Windows\Fonts file directory based on the font's metadata properties and NOT
(cont'd) the font's file name.
.EXAMPLE
Add-Fonts "C:\Users\ameadvin\Downloads"
- Without pipe, using mandatory paramater.
.EXAMPLE
Read-Host "Enter a valid file path of fonts to be installed." | Add-Fonts
- With pipe, using a Read-Host command for user-input.
#>


#	'Add-Type -AssemblyName PresentationCore' is a required .NET class that must be invoked in order to instantiate
#	(cont'd) the 'Windows.Media.GlyphTypeface' object.
#	If working in a Powershell terminal, enter the 'Add-Type' command first and then enter the code found in the Add-Fonts function.
Add-Type -AssemblyName PresentationCore

function Add-Fonts {
	[cmdletbinding(SupportsPaging = $true)]
	Param (
		[Parameter (Mandatory = $true, position=0, ValueFromPipeline)][string]$FontPathParam
	)

#	Tests if there's a valid path, appending a backslash-astericks '\*' or a astericks '*' to the end of the file path,
#	(cont'd) which's a Microsoft requirement to search a file directory.
	if(Test-Path -Path $FontPathParam){
		if("\" -ne [regex]::Match($FontPathParam,'[\\]$').Value){
			$FontPathFormattedAstr = "$FontPathParam\*"
		} else {
			$FontPathFormattedAstr = "$FontPathParam*"
		}
	} else{
		Write-Host "Please enter a valid file path!" -Foreground Magenta
		break
	}

#	Custom class created so two instantiated objects maybe called to compare installed vs. un-installed fonts.
	Class ReadFont {
		[System.Array[]]$ReadGlyphsCollected = @()
	
		[String[]] GetGlyphName([string[]]$UnreadGlyphsPath){
			if("*" -eq [regex]::Match($UnreadGlyphsPath,'[\*]$').Value){
				$UnreadGlyphsPathAstr = $UnreadGlyphsPath
			} elseif("\" -ne [regex]::Match($UnreadGlyphsPath,'[\\]$').Value) {
				$UnreadGlyphsPathAstr = "$UnreadGlyphsPath\*"
			} else {
				$UnreadGlyphsPathAstr = "$UnreadGlyphsPath*"
			}

			$UnreadGlyphs = Get-ItemProperty -Path $UnreadGlyphsPathAstr -Include *.otf,*.ttf
		
			if ($this.ReadGlyphsCollected -ne 0) {
				$this.ReadGlyphsCollected = @()
			}

#		Loops through each font's metadata properties, capturing the following attributes in an array: name, weight, style.
		foreach ($Glyphs in $UnreadGlyphs.FullName) {
			$GlyphFace = [Windows.Media.GlyphTypeface]::new([uri]::new($Glyphs))
			$ReadFontName = $GlyphFace.FamilyNames.Value
			$ReadFontWeight = $GlyphFace.Weight
			$ReadFontStyle = $GlyphFace.Style
		
			if(($ReadFontWeight -ne "Normal") -and ($ReadFontStyle -ne "Normal")){
				$ReadGlyphName = "$ReadFontName $ReadFontWeight $ReadFontStyle"			
			} elseif(($ReadFontWeight -ne "Normal") -and ($ReadFontStyle -eq "Normal")) {
				$ReadGlyphName = "$ReadFontName $ReadFontWeight"
			} elseif(($ReadFontWeight -eq "Normal") -and ($ReadFontStyle -ne "Normal")){
				$ReadGlyphName = "$ReadFontName $ReadFontStyle"
			} else {
				$ReadGlyphName = $ReadFontName
				}
			
			$this.ReadGlyphsCollected += @([PSCUstomObject]@{[string]'GlyphName' = $ReadGlyphName})
			}
		
		return $this.ReadGlyphsCollected.GlyphName
		}
	}
	
	$InstalledFontsLocation = "C:\Windows\Fonts"

	#"FTBIC" = Fonts To Be Installed Collection
	$FTBIC = (Get-ItemProperty -Path $FontPathFormattedAstr -Include *.otf,*.ttf)
	
	#"FTBI" = Fonts To Be Installed
	$FTBI = [ReadFont]::new()
	
	#"FAI" = Fonts Already Installed
	$FAI = [ReadFont]::new()
	
	#"IFA" = Installed Fonts Array	
	$IFA = $FAI.GetGlyphName($InstalledFontsLocation)
	
	#"NIFA" = Non-Installed Fonts Array
	$NIFA = $FTBI.GetGlyphName($FontPathFormattedAstr)
	
	#"ai" = array iterator
	$ai = 0

#	Loop begins by iterating through the three arrays of the 'fonts to be installed collection' in order to build the full file path
#	(cont'd) used to compare fonts already installed in the 'C:\Windows\Fonts' file directory.
	foreach($NIF in $NIFA){
		$FontFilePath = $FTBIC[$ai].FullName
		$FontFileName = $FTBIC[$ai].Name
		$FontFileExt = $FTBIC[$ai].Extension.ToLower()
		$ai++
		if($IFA.Contains($NIF)){
			Write-Host "Verified " -NoNewLine
			Write-Host $NIF -NoNewLine -Foreground Cyan
			Write-Host " is currently installed"
		} else {
			Copy-Item $FontFilePath -Destination $InstalledFontsLocation
			if(".ttf" -eq $FontFileExt){
				Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts\" -Name "$NIF (TrueType)" -Value $FontFileName
			} else {
				Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts\" -Name "$NIF (OpenType)" -Value $FontFileName
			}
			Write-Host "Successfully installed " -NoNewLine
			Write-Host $NIF -Foreground Yellow
		}
	}

#	Conditionally executed if "Read-Host" command piped to 'Add-Fonts' function.
	if ($PSCmdlet.MyInvocation.ExpectingInput) {
		Write-Host "`nPlease log out and then login to complete font installation ONLY IF a font(s) has been installed." -Foreground Magenta
	}
}