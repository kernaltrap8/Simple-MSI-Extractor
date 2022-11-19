Add-Type -AssemblyName System.Windows.Forms

#Use Windows Forms to open a file select dialog

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Windows Packages (*.msi)|*.msi'
}

$Out = $FileBrowser.ShowDialog() #Display the dialog

#Select output directory

$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    Description = 'Output'
}

$Out = $FolderBrowser.ShowDialog() #Display the dialog

$FolderBrowser.SelectedPath #Variable stuff

msiexec /a $FileBrowser.FileName /qb TARGETDIR=$($FolderBrowser.SelectedPath) # This uses the built in Windows tool to extract the MSI

#A helpful message

$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Once you install the MSI using this PowerShell script, please add any programs that run from a shell (i.e. CMD, PowerShell) be added to Path. 
To add a program to path, search for Control Panel in Windows Search, and open it. Once in Control Panel,
select User Accounts, then User Accounts again. On the side bar, select Change my Enviorment Variables.
Select the Path variable, and then Edit. Select a unfilled box, and type the path to the program (for most, it can be just the root folder, some may need to be bin) and then Ok, and Ok again. 
You WILL need to restart any open shells.", 0, "Thank you for using MSI-Extractor", 0)