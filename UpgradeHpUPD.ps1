# Script to detect and upgrade HP Universal Print Driver (UPD) for PCL 6 and PostScript
# Designed for GPO deployment - runs silently, with optional interactive output and version skip

param (
    [switch]$SkipVersionCheck
)

# Define variables
$minVersion = [version]"61.315.1.25959"
$drivers = @{
    "HP Universal Printing PCL 6" = @{
        Url = "https://ftp.hp.com/pub/softlib/software13/printers/UPD/upd-pcl6-x64-7.4.0.25959.zip"
        InfPattern = "hpcu*.inf"
        SearchPattern = '"HP Universal Printing PCL 6"'
        DriverPath = "C:\Windows\System32\DriverStore\FileRepository\hpcu315u.inf_amd64_*\hpmdp315.dll"
    }
    "HP Universal Printing PS" = @{
        Url = "https://ftp.hp.com/pub/softlib/software13/printers/UPD/upd-ps-x64-7.4.0.25959.zip"
        InfPattern = "hpcu*.inf"
        SearchPattern = '"HP Universal Printing PS"'
        DriverPath = "C:\Windows\System32\DriverStore\FileRepository\hpcu315v.inf_amd64_*\hpmdp315.dll"
    }
}
$tempBasePath = "$env:TEMP\HP_UPD"
$logFile = "$env:SYSTEMROOT\Temp\HP_UPD_Upgrade.log"

# Function to log messages
function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    $logMessage | Out-File -FilePath $logFile -Append
    Write-Host $message
}

# Check drivers unless -SkipVersionCheck is specified
$driversToProcess = @()
if (-not $SkipVersionCheck) {
    Write-Log "Checking installed drivers in Print Management..."
    $installedDrivers = Get-PrinterDriver | Where-Object { $drivers.Keys -contains $_.Name }

    if (-not $installedDrivers) {
        Write-Log "No supported drivers ('HP Universal Printing PCL 6' or 'HP Universal Printing PS') found in Print Management. Aborting script."
        exit 1
    }

    foreach ($driver in $installedDrivers) {
        Write-Log "Found driver '${driver.Name}' in Print Management: $($driver | Format-List | Out-String)"
        
        # Check version from specific driver file in DriverStore
        $driverPath = $drivers[$driver.Name].DriverPath
        $driverFile = Get-Item $driverPath -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($driverFile) {
            $installedVersion = [version]$driverFile.VersionInfo.FileVersion
            Write-Log "Detected version for '${driver.Name}' from file '${driverFile.FullName}': $installedVersion"
            if ($installedVersion -ge $minVersion) {
                Write-Log "Version ($installedVersion) for '${driver.Name}' is equal to or newer than $minVersion. Skipping upgrade."
                continue
            } else {
                Write-Log "Version ($installedVersion) for '${driver.Name}' is older than $minVersion. Proceeding with upgrade."
                $driversToProcess += $driver.Name
            }
        } else {
            Write-Log "No driver file found for '${driver.Name}' at '$driverPath'. Proceeding with upgrade."
            $driversToProcess += $driver.Name
        }
    }

    if (-not $driversToProcess) {
        Write-Log "All detected drivers are up to date. No action needed."
        exit 0
    }
} else {
    Write-Log "Skipping version and driver checks due to -SkipVersionCheck parameter. Processing all drivers."
    $driversToProcess = $drivers.Keys
}

# Process each driver
foreach ($driverName in $driversToProcess) {
    $driverInfo = $drivers[$driverName]
    $tempPath = "$tempBasePath_$($driverName -replace ' ', '_').zip"
    $extractPath = "$tempBasePath_$($driverName -replace ' ', '_')_Extracted"

    # Download the ZIP
    try {
        Write-Log "Downloading ${driverName} from $($driverInfo.Url)"
        Invoke-WebRequest -Uri $driverInfo.Url -OutFile $tempPath -ErrorAction Stop
        Write-Log "Download completed successfully for ${driverName}."
    } catch {
        Write-Log "Failed to download ZIP for ${driverName}: $_"
        continue
    }

    # Extract the ZIP
    try {
        Write-Log "Extracting ZIP for ${driverName} to $extractPath"
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        Expand-Archive -Path $tempPath -DestinationPath $extractPath -Force -ErrorAction Stop
        Write-Log "Extraction completed successfully for ${driverName}."
    } catch {
        Write-Log "Failed to extract ZIP for ${driverName}: $_"
        continue
    }

    # Find the INF file
    $infFile = Get-ChildItem -Path $extractPath -Recurse -Filter $driverInfo.InfPattern | 
        Where-Object { (Get-Content $_.FullName) -match $driverInfo.SearchPattern } | 
        Select-Object -First 1
    if (-not $infFile) {
        Write-Log "No INF file found for ${driverName} in extracted folder. Skipping."
        continue
    }
    Write-Log "Found INF file for ${driverName}: $($infFile.FullName)"

    # Log DriverVer from INF for reference (no comparison)
    try {
        $infContent = Get-Content $infFile.FullName
        $driverVerLine = ($infContent | Select-String "DriverVer\s*=").Line
        if ($driverVerLine -match "DriverVer\s*=\s*(\d+/\d+/\d+),([\d\.]+)") {
            $driverDate = $matches[1]
            $driverVersion = [version]$matches[2]
            Write-Log "Extracted DriverVer for ${driverName} from INF (for reference): $driverDate, $driverVersion"
        } else {
            Write-Log "Could not parse DriverVer for ${driverName} from INF. Proceeding anyway."
        }
    } catch {
        Write-Log "Failed to extract DriverVer for ${driverName} from INF: $_"
    }

    # Stage the driver with pnputil
    try {
        Write-Log "Staging driver ${driverName} with pnputil..."
        $pnputilOutput = & "pnputil.exe" /add-driver "$($infFile.FullName)" /install 2>&1
        Write-Log "pnputil output for ${driverName}: $pnputilOutput"
        if ($pnputilOutput -match "Failed") {
            Write-Log "pnputil reported a failure for ${driverName}. Check if all driver files are present."
            continue
        }

        # Use INF name for fallback
        Write-Log "Searching DriverStore for staged INF for ${driverName}"
        $stagedInfPath = Get-ChildItem -Path "C:\Windows\System32\DriverStore\FileRepository" -Recurse -Filter $infFile.Name -ErrorAction SilentlyContinue | 
            Select-Object -First 1 -ExpandProperty FullName
        if ($stagedInfPath) {
            Write-Log "Found staged INF for ${driverName} in DriverStore: $stagedInfPath"
        } else {
            Write-Log "No staged INF found in DriverStore for ${driverName}. Falling back to original path."
            $stagedInfPath = $infFile.FullName
        }
    } catch {
        Write-Log "Failed to stage driver or find staged INF for ${driverName}: $_"
        $stagedInfPath = $infFile.FullName
    }

    # Register the driver
    try {
        Write-Log "Driver registration phase for '${driverName}'."
        $existingDriver = Get-PrinterDriver | Where-Object { $_.Name -eq $driverName }
        if ($existingDriver) {
            Write-Log "Driver '${driverName}' already registered in Print Management: $($existingDriver | Format-List | Out-String)"
            #continue
        }

        Write-Log "Registering driver '${driverName}' using INF: $stagedInfPath"
        Add-PrinterDriver -Name $driverName -InfPath $stagedInfPath -ErrorAction Stop
        Write-Log "Driver '${driverName}' registered successfully."
    } catch {
        Write-Log "Failed to register driver '${driverName}': $_"
        $existingDriver = Get-PrinterDriver | Where-Object { $_.Name -eq $driverName }
        if ($existingDriver) {
            Write-Log "Driver '${driverName}' still registered in Print Management: $($existingDriver | Format-List | Out-String)"
        } else {
            Write-Log "Driver '${driverName}' not found in Print Management after failure."
        }
    }

    # Clean up for this driver
    if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
    if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
    Write-Log "Cleaned up temporary files for ${driverName}."
}

Write-Log "Script execution completed."
exit 0