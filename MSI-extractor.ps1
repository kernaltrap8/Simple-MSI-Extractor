# Simple MSI Extractor // Source file
# Written by JamesIsWack // kernaltrap
# Version 1.0.2.1

Add-Type -AssemblyName System.Windows.Forms

#Use Windows Forms to open a file select dialog

Write-Output ("What MSI do you want me to extract?")

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Windows Packages (*.msi)|*.msi'
}

$Out = $FileBrowser.ShowDialog() #Display the dialog

#Select output directory

Write-Output ("What folder do you want me to extract the content to?")

$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    Description = 'Select the directory you would like to extract to. Hint: use Make New Folder to organize the install.'
}

$Out = $FolderBrowser.ShowDialog() #Display the dialog

Write-Output ("Extracting...")

$FolderBrowser.SelectedPath #Variable stuff

Start-Process msiexec.exe /a $FileBrowser.FileName /qb TARGETDIR=$($FolderBrowser.SelectedPath) # This uses the built in Windows tool to extract the MSI

Write-Output ("Done! Go to the path you provided to see the contents.")

#A helpful message

$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Once you install the MSI using this PowerShell script, please add any programs that run from a shell (i.e. CMD, PowerShell) be added to Path. 
To add a program to path, search for Control Panel in Windows Search, and open it. Once in Control Panel,
select User Accounts, then User Accounts again. On the side bar, select Change my Enviorment Variables.
Select the Path variable, and then Edit. Select a unfilled box, and type the path to the program (for most, it can be just the root folder, some may need to be bin) and then Ok, and Ok again. 
You WILL need to restart any open shells.", 0, "Thank you for using MSI-Extractor", 0)
