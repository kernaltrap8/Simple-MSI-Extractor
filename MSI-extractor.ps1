# ==================================== #
# Simple MSI Extractor // Source file  #
# kernaltrap                           #
# Version 1.3                          #
# ==================================== #

Add-Type -AssemblyName System.Windows.Forms

$Shell = New-Object -ComObject "WScript.Shell" # Setup for Wscript.Shell usage in Powershell

Write-Output ("What MSI do you want me to extract?")

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Windows Packages (*.msi)|*.msi'
}

$FileBrowser.ShowDialog() # Display the dialog

# Select output directory

Write-Output ("What folder do you want me to extract the content to?")

$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    Description = 'Select the directory you would like to extract to. Hint: use Make New Folder to organize the install.'
}

$FolderBrowser.ShowDialog()

# ===================================================================== #
# msiexec.exe // Calls the Win32 MSI program                            #
# /a // Administrative install, passive                                 #
# $FileBrowser.FileName // filename taken from FileBrowser.ShowDialog() #
# TARGETDIR // Use directory taken from FolderBrowser.SelectedPath      #
# ===================================================================== #

msiexec.exe /a  $FileBrowser.FileName TARGETDIR="$($FolderBrowser.SelectedPath)" /qb

Get-Process | Where-Object{$_.path -eq $path} | Out-Null # Setup for Get-Process, this has a lot of output so it is directed to Out-Null (same as >$null)

# If statement to check if msiexec is running, if not, end the program and display the Shell popup

if(Get-Process | Where-Object{$_.path -eq "C:\Windows\System32\msiexec.exe"}){
  Write-Output ("Extracting...")
} else {
  Write-Output ("Done! Go to the path you provided to see the contents.")
  $Shell.Popup("MSI extracted.", 0, "Thank you for using MSI Extractor", 0)
}