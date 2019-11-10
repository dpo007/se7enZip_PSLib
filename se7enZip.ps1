<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER InputArchive
Parameter description

.PARAMETER OutputFolder
Parameter description

.PARAMETER ArchivePass
Parameter description

.PARAMETER FileSpecToExtract
Parameter description

.PARAMETER ExtractWithoutPaths
Parameter description

.EXAMPLE
An example

.NOTES
General notes
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
        throw "$env:ProgramFiles\7-Zip\7z.exe needed"
    }
    Set-Alias 7z "$env:ProgramFiles\7-Zip\7z.exe"

    if ($ArchivePass -eq '') {
        7z $extractCommand $InputArchive "-o$OutputFolder" -aoa $FileSpecToExtract
    } else {
        7z $extractCommand $InputArchive "-o$OutputFolder" "-p$ArchivePass" -aoa $FileSpecToExtract
    }

}

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
        throw "$env:ProgramFiles\7-Zip\7z.exe needed"
    }
    Set-Alias 7z "$env:ProgramFiles\7-Zip\7z.exe"

    if (!$ArchivePass.Trim()) {
        7z a $OutputArchive $FilesToArchive -sfx -mx=7
    } else {
        7z a $OutputArchive $FilesToArchive "-p$ArchivePass" -sfx -mx=7
    }
}

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