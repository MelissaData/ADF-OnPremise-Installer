# =======================================
# Start PowerShell As Administrator
# =======================================
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Restarting PowerShell as Administrator..."
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# =======================================
# Configuration - MAKE MODIFICATION HERE
# =======================================

# Blob storage information
$storageAccount     = "<your_storage_account_name>"
$containerName      = "<your_container_name>"

# SAS token for accessing the container (include the leading '?')
$sasToken           = "<your_blob_container_SAS>" # This should look like "?sp=...&sig=..."

# User Product License
$productLicense     = "<your_melissa_license_key>"


# Azure File Share connection script block (paste the full snippet here)
#
# Note:                         For your ease of use, this variable stores the entire script block as a string and extracts required values using regular expressions.
#                               If parsing fails or is inaccurate, you can skip this and manually set the variables in the section "Variable - Parse necessary variables from the Azure File Share script block".
#
# Full Script Block Example:
#                               $azureFileShareScript = @'
#                               $connectTestResult = Test-NetConnection -ComputerName melissadatafileshare.file.core.windows.net -Port 445
#                               if ($connectTestResult.TcpTestSucceeded) {
#                                   # Save the password so the drive will persist on reboot
#                                   cmd.exe /C "cmdkey /add:`"melissadatafileshare.file.core.windows.net`" /user:`"localhost\melissadatafileshare`" /pass:`"m+sGyFrCecd0LabS...hy2gA==`""
#                                   # Mount the drive
#                                   New-PSDrive -Name Z -PSProvider FileSystem -Root "\\melissadatafileshare.file.core.windows.net\ssis-data-files" -Persist
#                               } else {
#                                   Write-Error -Message "Unable to reach the Azure storage account via port 445. Check to make sure your organization or ISP is not blocking port 445, or use Azure P2S VPN, Azure S2S VPN, or Express Route to tunnel SMB traffic over a different port."
#                               }
#                               '@

$azureFileShareScript = @'
<your_Azure_File_Share_full_script_block>
'@

# Mapping subfolders within the root Azure File Share – DO NOT MODIFY unless your directory structure differs from the default.
$folderMapping = @{
    'EVC_DataPath'         = 'object-data-files'
    'EVC_GeoLoggingPath'   = 'contact-verify-geologging'
    'MatchUP'              = 'matchup-data-files'
    'Profiler'             = 'object-data-files'
    'Cleanser'             = 'object-data-files'
}

# ======================================================= #
# ======================================================= #
# PLEASE DO NOT MAKE ANY MODIFICATION FROM THIS PART DOWN #
# ======================================================= #
# ======================================================= #

# ============================================================================
# Variable - Parse necessary variables from the Azure File Share script block
# ============================================================================
$driveLetter    = ([regex]::Match($azureFileShareScript, '-Name\s+(?<d>\w)').Groups['d'].Value)
$azureHost      = ([regex]::Match($azureFileShareScript, '-ComputerName\s+(?<h>[\w\.]+)').Groups['h'].Value)
$azurePort      = ([regex]::Match($azureFileShareScript, '-Port\s+(?<p>\d+)').Groups['p'].Value)
$azureUsername  = ([regex]::Match($azureFileShareScript, '/user:`"(?<u>[^`]+)`"').Groups['u'].Value)
$azurePassword  = ([regex]::Match($azureFileShareScript, '/pass:`"(?<p>[^`]+)`"').Groups['p'].Value)
$pattern        = 'Root\s+"(?<unc>\\\\[^"]+)"'
$uncRoot        = [regex]::Match($azureFileShareScript, $pattern).Groups['unc'].Value

# ============================================================================
# Variable - Timestamp for Logging
# ============================================================================
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$timestampFileName = $timestamp -replace "[: ]", "_"

# ============================================================================
# Variable - Web Service Installer Name
# ============================================================================
$webInstallerName = "ADF-DQC-Web"

# ============================================================================
# Local Paths
# ============================================================================
$setupDir                       = "C:\SetupFiles"
$logDir                         = "C:\SetupLogs"
$dotnetFramework35Dir           = "$setupDir\dotnetFramework35\"

$customSetupLogFileName         = "$timestampFileName-custom_setup.log"
$customSetupLogFile             = "$logDir\$customSetupLogFileName"

$installerLogName               = "$timestampFileName-$webInstallerName-Install.log"
$installerLogPath               = "$logDir\$installerLogName"
$tempInstallerLog               = "$logDir\temp-$webInstallerName-Install.log"

$dqcWebLocalPath                = Join-Path $setupDir "$webInstallerName.exe"
$dotnetFramework35CabLocalPath  = Join-Path $dotnetFramework35Dir "microsoft-windows-netfx3-ondemand-package~31bf3856ad364e35~amd64~~.cab"

$smartMoverExportsDir           = "C:\SmartMoverExports"

# ============================================================================
# Create Local Directories
# ============================================================================
New-Item -ItemType Directory -Force -Path $setupDir                 | Out-Null
New-Item -ItemType Directory -Force -Path $logDir                   | Out-Null
New-Item -ItemType Directory -Force -Path $dotnetFramework35Dir     | Out-Null
New-Item -ItemType Directory -Force -Path $smartMoverExportsDir     | Out-Null

# ============================================================================
# Blob Paths
# ============================================================================
$dqcWebUrl                          = "https://$storageAccount.blob.core.windows.net/$containerName/melissa-web-installer/$webInstallerName.exe$sasToken"
$dotnetFramework35Url               = "https://$storageAccount.blob.core.windows.net/$containerName/dotnet-framework-35/microsoft-windows-netfx3-ondemand-package~31bf3856ad364e35~amd64~~.cab$sasToken"
$customSetupLogsUploadUrl           = "https://$storageAccount.blob.core.windows.net/$containerName/custom-setup-logs/$customSetupLogFileName$sasToken"
$installerLogsUploadUrl             = "https://$storageAccount.blob.core.windows.net/$containerName/installer-logs/$installerLogName$sasToken"

# ============================================================================
# Logging Function with Immediate Upload (for main log)
# ============================================================================
Function Write-Log {
    param ([string]$Message)
    $logTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$logTime - $Message"
    $logEntry | Out-File -Append -FilePath $customSetupLogFile

    try {
        Invoke-RestMethod -Uri $customSetupLogsUploadUrl -Method Put -InFile $customSetupLogFile -ContentType "text/plain" -Headers @{ "x-ms-blob-type" = "BlockBlob" }
    }
    catch {
        Write-Host "Failed to upload main log file: $($_.Exception.Message)" -ForegroundColor Red
        Exit 1
    }
}

# ============================================================================
# Helper function to upload a log file via a temporary copy
# ============================================================================
Function Upload-Installer-LogFile {
    try {
        Copy-Item -Path $installerLogPath -Destination $tempInstallerLog -ErrorAction Stop
        Invoke-RestMethod -Uri $installerLogsUploadUrl -Method Put -InFile $tempInstallerLog -ContentType "text/plain" -Headers @{ "x-ms-blob-type" = "BlockBlob" }
        Remove-Item $tempInstallerLog -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        Write-Log "Failed to upload installer log from temporary file: $($_.Exception.Message)"
        return $false
        Exit 1
    }
}

# ============================================================================
# Main Execution Block
# ============================================================================
try {
    # ------------------------------------------------------------------------
    # Step 1: Ensure .NET Framework 3.5 is installed 
    # ------------------------------------------------------------------------
    Try {
        if ((Get-WindowsFeature -Name Net-Framework-Core).Installed) {
            Write-Log ".NET framework 3.5 has already been installed."
        }
        else {
            Write-Log ".NET Framework 3.5 is not enabled. Attempting to install..."
            try {
                Write-Log "Downloading .NET 3.5 CAB from $dotnetFramework35Url"
                Start-BitsTransfer -Source $dotnetFramework35Url -Destination $dotnetFramework35CabLocalPath
                Write-Log ".NET Framework 3.5 CAB downloaded to $dotnetFramework35CabLocalPath"
            }
            catch {
                Write-Log "Failed to download .NET framework 3.5 - $_"
                Exit 1
            }
            if ((Install-WindowsFeature -Name Net-Framework-Core -Source $dotnetFramework35Dir -LogPath $installerLogPath).Success) {
                Write-Log ".NET framework 3.5 installed successfully"
            }
            else {
                Write-Log "Failed to install .NET framework 3.5"
                Exit 1
            }
        }
    }
    Catch {
        Write-Log "ERROR: Checking .NET 3.5 status failed - $_"
        Exit 1
    }

    # ------------------------------------------------------------------------
    # Step 2: Connect to Azure File Share dynamically
    # ------------------------------------------------------------------------
    Write-Log "Checking for existing mapping to drive $driveLetter..."
    try {
        if (Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue) {
            Write-Log "Drive $driveLetter is already mapped. Removing..."
            net use "${driveLetter}:" /delete /yes | Out-Null
            Start-Sleep -Seconds 2
            Write-Log "Drive $driveLetter unmapped successfully."
        } else {
            Write-Log "Drive $driveLetter is not currently mapped."
        }
    }
    catch {
        Write-Log "ERROR while checking/removing existing drive mapping - $($_.Exception.Message)"
        Exit 1
    }

    try {
        $connectTestResult = Test-NetConnection -ComputerName $azureHost -Port $azurePort
        if ($connectTestResult.TcpTestSucceeded) {
            Write-Log "Port $azurePort is open. Proceeding with Azure File Share mount..."
            cmd.exe /C "cmdkey /add:`"$azureHost`" /user:`"$azureUsername`" /pass:`"$azurePassword`""
            Write-Log "Azure File Share credentials saved via cmdkey."
            try {
                New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $uncRoot -Persist -ErrorAction Stop
                Write-Log "Drive $driveLetter mapped successfully to $uncRoot"
            } catch {
                Write-Log "ERROR: Failed to map drive $driveLetter - $($_.Exception.Message)"
                Exit 1
            }
        } else {
            Write-Log "ERROR: Port $azurePort is not reachable. Aborting Azure File Share connection."
            Exit 1
        }
    }
    catch {
        Write-Log "ERROR during Azure File Share connection test - $($_.Exception.Message)"
        Exit 1
    }

    # ------------------------------------------------------------------------
    # Step 3: Download Web Service Installer from Blob
    # ------------------------------------------------------------------------
    try {
        Write-Log "Downloading $webInstallerName.exe from Blob: $dqcWebUrl"
        Start-BitsTransfer -Source $dqcWebUrl -Destination $dqcWebLocalPath -ErrorAction Stop
        if (-Not (Test-Path $dqcWebLocalPath)) {
            throw "Download failed: $dqcWebLocalPath not found."
        }
        Write-Log "$webInstallerName.exe downloaded to $dqcWebLocalPath"
    }
    catch {
        Write-Log "ERROR: Failed to download $webInstallerName.exe - $($_.Exception.Message)"
        Exit 1
    }

    # ------------------------------------------------------------------------
    # Step 4: Install Web Service Components Silently
    # ------------------------------------------------------------------------
    Write-Log "Installing $webInstallerName.exe in silent mode..."
    $global:InstallerProcess = Start-Process -FilePath $dqcWebLocalPath `
        -ArgumentList "/VERYSILENT","/NORESTART","/ForceSSIS2017x64","/LOG=$installerLogPath","-License",$productLicense,"/NoPopUp" -PassThru

    while (-not (Test-Path $installerLogPath)) { Start-Sleep -Seconds 1 }
    $prevLineCount = (Get-Content $installerLogPath | Measure-Object -Line).Lines
    if ($prevLineCount -gt 0 -and (Upload-Installer-LogFile)) {
        Write-Log "Uploaded initial installer log with $prevLineCount lines."
    }

    while (-not $global:InstallerProcess.HasExited) {
        Start-Sleep -Seconds 30
        if (Test-Path $installerLogPath) {
            $currentLineCount = (Get-Content $installerLogPath | Measure-Object -Line).Lines
            if ($currentLineCount -gt $prevLineCount -and (Upload-Installer-LogFile)) {
                Write-Log "Uploaded installer log update: $currentLineCount lines (was $prevLineCount)."
                $prevLineCount = $currentLineCount
            }
        }
    }

    if ((Test-Path $installerLogPath) -and (Upload-Installer-LogFile)) {
        Write-Log "Final installer log upload complete with $prevLineCount lines."
    }
    Write-Log "$webInstallerName.exe installation completed with Exit Code: $($global:InstallerProcess.ExitCode)"

    # ------------------------------------------------------------------------
    # Step 5: Prepare Contact Verify Config File
    # ------------------------------------------------------------------------
    Write-Log "----- Preparing Contact Verify Config File -----"
    $configPath                         = "C:\ProgramData\Melissa DATA\EVC\EVC.SSIS.Config"
    $contactVerifyTargetDataPath        = [string](Join-Path $uncRoot $folderMapping['EVC_DataPath'])
    $contactVerifyTargetGeoLoggingPath  = [string](Join-Path $uncRoot $folderMapping['EVC_GeoLoggingPath'])

    try {
        if (-Not (Test-Path $configPath)) {
            Write-Log "Contact Verify - ERROR: Config file not found at $configPath"
            throw "Config file not found"
        }

        $configXml = [xml](Get-Content $configPath -ErrorAction Stop)
        Write-Log "Contact Verify - Loaded config file successfully from $configPath"

        # --- Update ProcessingMode ---
        if ($configXml.EVC.ProcessingMode) {
            $currentMode = $configXml.EVC.ProcessingMode

            if ($currentMode -eq "Dlls") {
                Write-Log "Contact Verify - No change needed. ProcessingMode already set to Dlls."
            } else {
                Write-Log "Contact Verify - Current ProcessingMode is '$currentMode'. Updating to 'Dlls'..."
                $configXml.EVC.ProcessingMode = "Dlls"
                $configXml.Save($configPath)
                Write-Log "Contact Verify - Updated ProcessingMode to 'Dlls' and saved the file."
            }
        }
        else {
            Write-Log "Contact Verify - ProcessingMode tag not found. Creating new <ProcessingMode>Dlls</ProcessingMode> node..."
            $newNode = $configXml.CreateElement("ProcessingMode")
            $newNode.InnerText = "Dlls"
            $configXml.EVC.AppendChild($newNode) | Out-Null
            $configXml.Save($configPath)
            Write-Log "Contact Verify - Added new ProcessingMode node with value 'Dlls' and saved the file."
        }

        # --- Update MailboxLookupMode ---
        if ($configXml.EVC.MailboxLookupMode) {
            $currentMode = $configXml.EVC.MailboxLookupMode

            if ($currentMode -eq "Express") {
                Write-Log "Contact Verify - No change needed. MailboxLookupMode already set to Express."
            } else {
                Write-Log "Contact Verify - Current MailboxLookupMode is '$currentMode'. Updating to 'Express'..."
                $configXml.EVC.MailboxLookupMode = "Express"
                $configXml.Save($configPath)
                Write-Log "Contact Verify - Updated MailboxLookupMode to 'Express' and saved the file."
            }
        }
        else {
            Write-Log "Contact Verify - MailboxLookupMode tag not found. Creating new <MailboxLookupMode>Express</MailboxLookupMode> node..."
            $newNode = $configXml.CreateElement("MailboxLookupMode")
            $newNode.InnerText = "Express"
            $configXml.EVC.AppendChild($newNode) | Out-Null
            $configXml.Save($configPath)
            Write-Log "Contact Verify - Added new MailboxLookupMode node with value 'Express' and saved the file."
        }

        # --- Update DataPath ---
        if ($configXml.EVC.DataPath) {
            $currentDataPath = $configXml.EVC.DataPath

            if ($currentDataPath -eq $contactVerifyTargetDataPath) {
                Write-Log "Contact Verify - No change needed. DataPath already set to '$contactVerifyTargetDataPath'."
            } else {
                Write-Log "Contact Verify - Current DataPath is '$currentDataPath'. Updating to '$contactVerifyTargetDataPath'..."
                $configXml.EVC.DataPath = $contactVerifyTargetDataPath
                $configXml.Save($configPath)
                Write-Log "Contact Verify - Updated DataPath to '$contactVerifyTargetDataPath' and saved the file."
            }
        }
        else {
            Write-Log "Contact Verify - DataPath tag not found. Creating new <DataPath>$contactVerifyTargetDataPath</DataPath> node..."
            $newNode = $configXml.CreateElement("DataPath")
            $newNode.InnerText = $contactVerifyTargetDataPath
            $configXml.EVC.AppendChild($newNode) | Out-Null
            $configXml.Save($configPath)
            Write-Log "Contact Verify - Added new DataPath node with value '$contactVerifyTargetDataPath' and saved the file."
        }

        # --- Update GeoLoggingPath ---
        if ($configXml.EVC.GeoLoggingPath) {
            $currentGeoLoggingPath = $configXml.EVC.GeoLoggingPath

            if ($currentGeoLoggingPath -eq $contactVerifyTargetGeoLoggingPath) {
                Write-Log "Contact Verify - No change needed. GeoLoggingPath already set to '$contactVerifyTargetGeoLoggingPath'."
            }
            else {
                Write-Log "Contact Verify - Current GeoLoggingPath is '$currentGeoLoggingPath'. Updating to '$contactVerifyTargetGeoLoggingPath'..."
                $configXml.EVC.GeoLoggingPath = $contactVerifyTargetGeoLoggingPath
                $configXml.Save($configPath)
                Write-Log "Contact Verify - Updated GeoLoggingPath to '$contactVerifyTargetGeoLoggingPath' and saved the file."
            }
        }
        else {
            Write-Log "Contact Verify - GeoLoggingPath tag not found. Creating new <GeoLoggingPath>$contactVerifyTargetGeoLoggingPath</GeoLoggingPath> node..."
            $newNode = $configXml.CreateElement("GeoLoggingPath")
            $newNode.InnerText = $contactVerifyTargetGeoLoggingPath
            $configXml.EVC.AppendChild($newNode) | Out-Null
            $configXml.Save($configPath)
            Write-Log "Contact Verify - Added new GeoLoggingPath node with value '$contactVerifyTargetGeoLoggingPath' and saved the file."
        }
    }
    catch {
        Write-Log "Contact Verify - ERROR: Modifying EVC config failed - $_"
        Exit 1
    }

    # ------------------------------------------------------------------------
    # Step 6: Prepare MatchUp Config File
    # ------------------------------------------------------------------------
    Write-Log "----- Preparing MatchUp Config File -----"
    $matchupConfigPath      = "C:\ProgramData\Melissa DATA\MatchUP\MatchUP.SSIS.Config"
    $matchupTargetDataPath  = [string](Join-Path $uncRoot $folderMapping['MatchUP'])

    try {
        if (-Not (Test-Path $matchupConfigPath)) {
            Write-Log "MatchUp - ERROR: MatchUp config file not found at $matchupConfigPath"
            throw "MatchUp - MatchUp config file not found"
        }

        $matchupXml = [xml](Get-Content $matchupConfigPath -ErrorAction Stop)
        Write-Log "MatchUp - Loaded MatchUp config file successfully from $matchupConfigPath"

        if ($matchupXml.MatchUP.DataPath) {
            $currentPath = $matchupXml.MatchUP.DataPath

            if ($currentPath -eq $matchupTargetDataPath) {
                Write-Log "No change needed. MatchUp DataPath already set to '$matchupTargetDataPath'."
            } else {
                Write-Log "MatchUp - Current MatchUp DataPath is '$currentPath'. Updating to '$matchupTargetDataPath'..."
                $matchupXml.MatchUP.DataPath = $matchupTargetDataPath
                $matchupXml.Save($matchupConfigPath)
                Write-Log "MatchUp - Updated MatchUp DataPath to '$matchupTargetDataPath' and saved the file."
            }
        }
        else {
            Write-Log "MatchUp - MatchUp DataPath tag not found. Creating new <DataPath>$matchupTargetDataPath</DataPath> node..."
            $newNode = $matchupXml.CreateElement("DataPath")
            $newNode.InnerText = $matchupTargetDataPath
            $matchupXml.MatchUP.AppendChild($newNode) | Out-Null
            $matchupXml.Save($matchupConfigPath)
            Write-Log "MatchUp - Added new MatchUp DataPath node with value '$matchupTargetDataPath' and saved the file."
        }
    }
    catch {
        Write-Log "MatchUp - ERROR: Failed to modify MatchUp config file - $($_.Exception.Message)"
        Exit 1
    }

    # ------------------------------------------------------------------------
    # Step 7: Prepare Profiler Config File
    # ------------------------------------------------------------------------
    Write-Log "----- Preparing Profiler DataPath -----"
    $profilerConfigPath     = "C:\ProgramData\Melissa DATA\Profiler\Profiler.SSIS.Config"
    $profilerTargetDataPath = [string](Join-Path $uncRoot $folderMapping['Profiler'])

    try {
        if (-Not (Test-Path $profilerConfigPath)) {
            Write-Log "Profiler - ERROR: Profiler config file not found at $profilerConfigPath"
            throw "Profiler config file not found"
        }

        $profilerXml = [xml](Get-Content $profilerConfigPath -ErrorAction Stop)
        Write-Log "Profiler - Loaded Profiler config file successfully from $profilerConfigPath"

        if ($profilerXml.Profiler.DataPath) {
            $currentPath = $profilerXml.Profiler.DataPath

            if ($currentPath -eq $profilerTargetDataPath) {
                Write-Log "Profiler - No change needed. Profiler DataPath already set to '$profilerTargetDataPath'."
            } else {
                Write-Log "Profiler - Current Profiler DataPath is '$currentPath'. Updating to '$profilerTargetDataPath'..."
                $profilerXml.Profiler.DataPath = $profilerTargetDataPath
                $profilerXml.Save($profilerConfigPath)
                Write-Log "Profiler - Updated Profiler DataPath to '$profilerTargetDataPath' and saved the file."
            }
        }
        else {
            Write-Log "Profiler - Profiler DataPath tag not found. Creating new <DataPath> node with '$profilerTargetDataPath'..."
            $newNode = $profilerXml.CreateElement("DataPath")
            $newNode.InnerText = $profilerTargetDataPath
            $profilerXml.Profiler.AppendChild($newNode) | Out-Null
            $profilerXml.Save($profilerConfigPath)
            Write-Log "Profiler - Added new Profiler DataPath node with value '$profilerTargetDataPath' and saved the file."
        }
    }
    catch {
        Write-Log "Profiler - ERROR: Failed to modify Profiler config file - $($_.Exception.Message)"
        Exit 1
    }

    # ------------------------------------------------------------------------
    # Step 8: Prepare Cleanser Config File
    # ------------------------------------------------------------------------
    Write-Log "----- Preparing Cleanser Config File -----"
    $cleanserConfigPath      = "C:\ProgramData\Melissa DATA\Cleanser\Cleanser.SSIS.Config"
    $cleanserTargetDataPath  = [string](Join-Path $uncRoot $folderMapping['Cleanser'])

    try {
        if (-Not (Test-Path $cleanserConfigPath)) {
            Write-Log "Cleanser - ERROR: Cleanser config file not found at $cleanserConfigPath"
            throw "Cleanser config file not found"
        }

        $cleanserXml = [xml](Get-Content $cleanserConfigPath -ErrorAction Stop)
        Write-Log "Cleanser - Loaded Cleanser config file successfully from $cleanserConfigPath"

        if ($cleanserXml.Cleanser.DataPath) {
            $currentPath = $cleanserXml.Cleanser.DataPath

            if ($currentPath -eq $cleanserTargetDataPath) {
                Write-Log "Cleanser - No change needed. Cleanser DataPath already set to '$cleanserTargetDataPath'."
            } else {
                Write-Log "Cleanser - Current Cleanser DataPath is '$currentPath'. Updating to '$cleanserTargetDataPath'..."
                $cleanserXml.Cleanser.DataPath = $cleanserTargetDataPath
                $cleanserXml.Save($cleanserConfigPath)
                Write-Log "Cleanser - Updated Cleanser DataPath to '$cleanserTargetDataPath' and saved the file."
            }
        }
        else {
            Write-Log "Cleanser - Cleanser DataPath tag not found. Creating new <DataPath> node with '$cleanserTargetDataPath'..."
            $newNode = $cleanserXml.CreateElement("DataPath")
            $newNode.InnerText = $cleanserTargetDataPath
            $cleanserXml.Cleanser.AppendChild($newNode) | Out-Null
            $cleanserXml.Save($cleanserConfigPath)
            Write-Log "Cleanser - Added new Cleanser DataPath node with value '$cleanserTargetDataPath' and saved the file."
        }
    }
    catch {
        Write-Log "Cleanser - ERROR: Failed to modify Cleanser config file - $($_.Exception.Message)"
        Exit 1
    }

    # ------------------------------------------------------------------------
    # Step 9: Final upload of main log
    # ------------------------------------------------------------------------
    Write-Log "Setup Script Completed. Final log upload..."
    try {
        Invoke-RestMethod -Uri $customSetupLogsUploadUrl -Method Put -InFile $customSetupLogFile -ContentType "text/plain" -Headers @{ "x-ms-blob-type" = "BlockBlob" }
        Write-Log "Final main log uploaded successfully."
    }
    catch {
        Write-Log "ERROR: Failed final log upload - $_"
        Exit 1
    }

}
catch {
    Write-Log "FATAL ERROR: $_"
    throw $_
    Exit 1
}

Write-Host "Setup Script Completed Successfully!"
