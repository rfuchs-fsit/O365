#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Param (
    [switch]$SuppressReboot = $true,
    [switch]$UseSetupRemoval = $true
)
#----------------------------------------------------------[Declarations]----------------------------------------------------------
$SaRA_URL = "https://aka.ms/SaRA_CommandLineVersionFiles"
$SaRA_ZIP = "$env:TEMP\SaRA.zip"
$SaRA_DIR = "$env:TEMP\SaRA"
$SaRA_EXE = "$SaRA_DIR\SaRAcmd.exe"
$Office365Setup_URL = "https://github.com/rfuchs-fsit/O365/raw/main/Install-Files"
#-----------------------------------------------------------[Functions]------------------------------------------------------------
Function Invoke-OfficeUninstall {
    if (-Not (Test-Path "$SaRA_DIR")) {
        New-Item "$SaRA_DIR" -ItemType Directory | Out-Null
    }
    if ($UseSetupRemoval) {
        Write-Host "Invoking default setup method ..."
        Start-BitsTransfer -Source "$Office365Setup_URL/setup.exe" -Destination "$SaRA_DIR\setup.exe"
        Start-BitsTransfer -Source "$Office365Setup_URL/purge.xml" -Destination "$SaRA_DIR\purge.xml"
        Start-Process -FilePath "$SaRA_DIR\setup.exe" -ArgumentList "/configure $SaRA_DIR\purge.xml" -Wait
    }
}
 
Function Stop-OfficeProcess {
    Write-Host "Stopping running Office applications ..."
    $OfficeProcessesArray = "lync", "winword", "excel", "msaccess", "mstore", "infopath", "setlang", "msouc", "ois", "onenote", "outlook", "powerpnt", "mspub", "groove", "visio", "winproj", "graph", "teams"
    foreach ($ProcessName in $OfficeProcessesArray) {
        if (get-process -Name $ProcessName -ErrorAction SilentlyContinue) {
            if (Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue) {
                Write-Output "Process $ProcessName was stopped."
            }
            else {
                Write-Warning "Process $ProcessName could not be stopped."
            }
        } 
    }
}

Function Invoke-SetupOffice365($Office365ConfigFile) {
        Write-Host "Downloading Office365 Installer ..."
        Start-BitsTransfer -Source "$Office365ConfigFile" -Destination "$SaRA_DIR\config.xml"
        Write-Host "Executing Office365 Setup ..."
        $OfficeSetup = Start-Process -FilePath "$SaRA_DIR\setup.exe" -ArgumentList "/configure $SaRA_DIR\config.xml" -Wait
        switch ($OfficeSetup.ExitCode) {
            0 {
                Write-Host "Install successful!"
                Break
            }

            1 {
                Write-Error "Install failed!"
                Break
            }
        }
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------

Stop-OfficeProcess
Invoke-OfficeUninstall
Invoke-SetupOffice365 "$Office365Setup_URL/install.xml"