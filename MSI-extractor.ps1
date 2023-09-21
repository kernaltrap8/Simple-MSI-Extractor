# ==================================== #
# Simple MSI Extractor // Source file  #
# kernaltrap                           #
# Version 1.2                          #
# ==================================== #

Add-Type -AssemblyName System.Windows.Forms

# Use Windows Forms to open a file select dialog

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

$FolderBrowser.ShowDialog() # Display the dialog

Write-Output ("Extracting...")

$FolderBrowser.SelectedPath

msiexec.exe /a  $FileBrowser.FileName TARGETDIR="$($FolderBrowser.SelectedPath)" /qb

# ===================================================================== #
# msiexec.exe // Calls the Win32 MSI program                            #
# /a // Administrative install, passive                                 #
# $FileBrowser.FileName // filename taken from FileBrowser.ShowDialog() #
# TARGETDIR // Use directory taken from FolderBrowser.SelectedPath      #
# ===================================================================== #

Write-Output ("Done! Go to the path you provided to see the contents.")

#A helpful message

$Shell = New-Object -ComObject "WScript.Shell"

$Shell.Popup("MSI extracted.", 0, "Thank you for using MSI Extractor", 0)