<#
.SYNOPSIS
7-Zip unzip wrapper

.DESCRIPTION
Function that wraps 7-Zip's extract functionality.  Requires 7-Zip be installed.

.PARAMETER InputArchive
Archive to process

.PARAMETER OutputFolder
Location to place extracted files

.PARAMETER ArchivePass
Password to use while extracting files

.PARAMETER FileSpecToExtract
Filespec of file(s) to extract

.PARAMETER ExtractWithoutPaths
Don't create file paths when extracting

.EXAMPLE
se7enUnZip -InputArchive In.zip -OutputFolder 'C:\temp' -ArchivePass 'L4meP4ssw0rd!'

.NOTES
Warning: This will overwrite matching output files without prompting.
#>
function se7enUnZip {
    param (
        [Parameter(Position=0,
            Mandatory=$true)]
        [string] $InputArchive,
        [Parameter(Position=1,
            Mandatory=$true)]
        [string] $OutputFolder,
        [Parameter(Position=2)]
        [string] $ArchivePass,
        [Parameter(Position=3)]
        [string] $FileSpecToExtract,
        [switch] $ExtractWithoutPaths = $false
    )

    if ($ExtractWithoutPaths) {
        $extractCommand = 'e'
    } else {
        $extractCommand = 'x'
    }

    if (-Not (Test-Path "$env:ProgramFiles\7-Zip\7z.exe")) {
        throw "$env:ProgramFiles\7-Zip\7z.exe needed."
    }
    Set-Alias 7z "$env:ProgramFiles\7-Zip\7z.exe"

    if ($ArchivePass -eq '') {
        7z $extractCommand $InputArchive "-o$OutputFolder" -aoa $FileSpecToExtract
    } else {
        7z $extractCommand $InputArchive "-o$OutputFolder" "-p$ArchivePass" -aoa $FileSpecToExtract
    }

}

<#
.SYNOPSIS
Archive files into a self-extracting EXE

.DESCRIPTION
Function that uses 7-Zip to archive files into a self-extracting EXE.  Requires 7-Zip be installed.

.PARAMETER OutputArchive
Name of archive to create (ie: 'Output.exe')

.PARAMETER FilesToArchive
File spec of files to add to archive (ie: '*.csv')

.PARAMETER ArchivePass
Password to apply to completed archive

.EXAMPLE
se7enZipToExe -FilesToArchive 'C:\temp\*.CSV' -OutputArchive 'C:\temp\csvFiles.EXE' -ArchivePass 'L4meP4ssw0rd!'

.NOTES
Requires 7-Zip be installed.
#>
function se7enZipToExe {
    param (
        [Parameter(Position=0,
            Mandatory=$true)]
        [string] $OutputArchive,
        [Parameter(Position=1,
            Mandatory=$true)]
        [string] $FilesToArchive,
        [Parameter(Position=2)]
        [string] $ArchivePass
    )

    if (-Not (Test-Path "$env:ProgramFiles\7-Zip\7z.exe")) {
        throw "$env:ProgramFiles\7-Zip\7z.exe needed."
    }
    Set-Alias 7z "$env:ProgramFiles\7-Zip\7z.exe"

    if (!$ArchivePass.Trim()) {
        7z a $OutputArchive $FilesToArchive -sfx -mx=7
    } else {
        7z a $OutputArchive $FilesToArchive "-p$ArchivePass" -sfx -mx=7
    }
}

<#
.SYNOPSIS
Re-archive, removing provided password.

.DESCRIPTION
Uses 7-Zip to extract password-protected archive, and then creates a new archive without it.

.PARAMETER InputArchive
Archive to process

.PARAMETER OutputFolder
Location to place produced archive

.PARAMETER WorkingFolder
Working folder used to hold files which re-archiving (default = "TempWork" in script's folder)

.PARAMETER ArchivePass
Password to remove

.EXAMPLE
RemoveZipPassword -InputArchive In.zip -ArchivePass 'L4meP4ssw0rd!'

.NOTES
If the output folder is not specified, then use the input file's folder, and overwrite the originals.
#>

function RemoveZipPassword {
    param (
        [Parameter(Mandatory=$true)]
        [string] $InputArchive,
        [string] $OutputFolder,
        [string] $WorkingFolder = (Join-Path $PSScriptRoot "TempWork"),
        [Parameter(Mandatory=$true)]
        [string] $ArchivePass
    )

    # If the output folder is not specified, then use the input file's folder, and overwrite the originals.
    if (!$OutputFolder) {
        $OutputFolder = Split-Path $InputArchive
    }

    if ($OutputFolder.EndsWith('\')) {
        $OutputFolder.Substring(0, $OutputFolder.Length-1)
    }

    $randomFileName = [io.path]::GetFileNameWithoutExtension([System.IO.Path]::GetRandomFileName())
    $tempWorkFolder = Join-Path $WorkingFolder $randomFileName
    if (!(Test-Path $tempWorkFolder)) {
        New-Item -Path $tempWorkFolder -ItemType Directory -Force
    }

    se7enUnZip -InputArchive $InputArchive -archivePass $ArchivePass -outputFolder $tempWorkFolder

    if ((Split-Path $InputArchive) -eq $OutputFolder) {
        Remove-Item $InputArchive -Force
    } else {
        $InputArchive = Join-Path $OutputFolder (Split-Path $InputArchive -Leaf)
    }

    $FilesToArchive = Join-Path $tempWorkFolder '*'
    se7enZipToExe -outputArchive $InputArchive -filesToArchive $FilesToArchive
    Remove-Item $tempWorkFolder -Recurse -Force
}