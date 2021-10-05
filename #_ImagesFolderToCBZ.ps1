# ----------------------------------------------------------------------
# --
# -- Created by PonyLucky.
# --
# -- version: 0.3.
# --
# 
# -- Description
# The function of this script is to convert an image container into a
# CBZ file (used for comics or mangas, readable on Ebooks and any
# devices with the good application).
# 
# -- Prerequises
# You need to authorize the PowerShell scripts on your computer. Go to
# this URL if you don't know how: https://superuser.com/a/106363.
# 
# You will need to install 7zip if not already installed (it is faster,
# more stable than WinRAR and it is Open-Source, hence free). Go to
# this URL to install it: https://www.7-zip.org/.
# 
# -- How do this script works
# This script ZIPs all folders in the root folder of the scsript and
# changes their extension to .CBZ.
# 
# -- Notes
# If you take this script into another one or a project, credits will
# be appreciated. Though it is not an obligation.
# 
# ---------------------------------------------------------------------- 

Clear-Host

# ROOT - Script folder path
$root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
# 7Zip path
$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"

# Test if 7zip is installed
if (-not (Test-Path -Path $7zipPath -PathType Leaf)) {
  throw "7 zip file '$7zipPath' not found"
}

# Ask the user if we delete the folders after their conversion
$isSdel = Read-Host -Prompt "Do you want to delete the folders after conversion [Y/N]: "

# Write for the user
Write-Host "`r`nConverted:"

# Get the files inside the INPUT path and forEach children
Get-ChildItem "$root" | ForEach-Object {    
  # If the child is a folder/container
  if ($true -eq $_.psiscontainer) {
    # Get the full path of the file
    $pathFolder = $_.FullName

    # Zip the content of the folder
    & $7zipPath a "$pathFolder.zip" "$pathFolder\*" > $null

    # Change the extension to CBZ
    $newName = [System.IO.Path]::ChangeExtension("$pathFolder.zip",".cbz")
    Move-Item -Path "$pathFolder.zip" -Destination $newName -Force

    # If the user asked for deletion of folders
    if ("Y" -eq $isSdel.ToUpper()) {
      # Remove the folder
      $pathFolder | Remove-Item -Force -Recurse
    }

    # Tells the user this one is finished converting
    Write-Host "--" -ForegroundColor DarkGray -NoNewline
    Write-Host " $_.cbz"
  }
}

# Tells the user it's finished
Write-Host "`r`nFinished`r`n" -ForegroundColor Green

# Pause to let us see the result
Pause
