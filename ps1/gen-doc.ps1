# Record the start time and format it as DD-MM-YYYY
$startTime = Get-Date
$startTimeFormatted = $startTime.ToString("dd-MM-yyyy")

# Get basic data
$hostname    = $env:COMPUTERNAME
$currentUser = $env:USERNAME
$os          = (Get-CimInstance Win32_OperatingSystem).Caption

# --- Gather Physical Disk Inventory (not partitions) ---
$disks = Get-WmiObject -Class Win32_DiskDrive
$physicalDiskText = ""
$diskIndex = 1
foreach ($disk in $disks) {
    $sizeGB = "{0:N2}" -f ($disk.Size / 1GB)
    $diskType = $disk.InterfaceType
    $physicalDiskText += "Disk $diskIndex : $sizeGB GB $diskType`n"
    $diskIndex++
}

# --- Gather Hardware Details ---
$computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
$computerModel  = $computerSystem.Model
$totalMemoryGB  = "{0:N2}" -f ($computerSystem.TotalPhysicalMemory / 1GB)

$cpuInfo = Get-CimInstance -ClassName Win32_Processor
$cpuCount = $cpuInfo.Count
$coresPerProcessor = $cpuInfo[0].NumberOfCores
$cpuModels = ($cpuInfo | Select-Object -ExpandProperty Name | Sort-Object -Unique) -join ", "

# --- Gather Partition (Disk) Inventory ---
$partitions = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
$partitionText = ""
foreach ($partition in $partitions) {
    $deviceID   = $partition.DeviceID
    $label      = $partition.VolumeName
    $fileSystem = $partition.FileSystem
    $totalSizeGB = "{0:N2}" -f ($partition.Size / 1GB)
    $freeSpaceGB = "{0:N2}" -f ($partition.FreeSpace / 1GB)
    $usedGB = "{0:N2}" -f ((($partition.Size - $partition.FreeSpace) / 1GB))
    
    $partitionText += "DeviceID: $deviceID`n"
    $partitionText += "Label: $label`n"
    $partitionText += "File System: $fileSystem`n"
    $partitionText += "Total Size GB: $totalSizeGB GB`n"
    $partitionText += "Used GB: $usedGB GB`n"
    $partitionText += "Free GB: $freeSpaceGB GB`n"
    $partitionText += "`n"
}

# --- Gather Network Parameter Inventory (for all NICs) ---
$networkText = ""
$nicConfigs = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True"
foreach ($nic in $nicConfigs) {
    $description    = $nic.Description
    $ipaddress      = if ($nic.IPAddress) { $nic.IPAddress -join ", " } else { "N/A" }
    $defaultGateway = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway -join ", " } else { "N/A" }
    $subnet         = if ($nic.IPSubnet) { $nic.IPSubnet -join ", " } else { "N/A" }
    $dns            = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder -join ", " } else { "N/A" }
    $winsPrimary    = if ($nic.WINSPrimaryServer) { $nic.WINSPrimaryServer } else { "N/A" }
    $winsSecondary  = if ($nic.WINSSecondaryServer) { $nic.WINSSecondaryServer } else { "N/A" }
    
    $networkText += "Description: $description`n"
    $networkText += "IPAddress: $ipaddress`n"
    $networkText += "DefaultIPGateway: $defaultGateway`n"
    $networkText += "Mask: $subnet`n"
    $networkText += "DNS: $dns`n"
    $networkText += "WINS Primary Server: $winsPrimary`n"
    $networkText += "WINS Secondary Server: $winsSecondary`n"
    $networkText += "`n"
}

# --- Gather Time Zone Details ---
$timeZone       = (Get-TimeZone).DisplayName
$cultureInfo    = Get-Culture
$locale         = $cultureInfo.Name
$languageDetail = $cultureInfo.DisplayName

$timeZoneText = @"
Current Time Zone : $timeZone
Locale            : $locale
Language Details  : $languageDetail
"@

# --- Gather Installed Applications ---
$appList = @()
$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

foreach ($regPath in $registryPaths) {
    if (Test-Path $regPath) {
        $apps = Get-ChildItem -Path $regPath | ForEach-Object {
            $app = Get-ItemProperty $_.PSPath
            if ($app.DisplayName) {
                [PSCustomObject]@{
                    Name          = $app.DisplayName
                    Version       = $app.DisplayVersion
                    Size          = if ($app.EstimatedSize) { "{0:N2} MB" -f ($app.EstimatedSize/1024) } else { "N/A" }
                    Location      = if ($app.InstallLocation) { $app.InstallLocation } else { "N/A" }
                    InstalledDate = if ($app.InstallDate -and ($app.InstallDate -match "^\d{8}$")) {
                                        ([datetime]::ParseExact($app.InstallDate, "yyyyMMdd", $null)).ToString("dd-MM-yyyy")
                                     }
                                     elseif ($app.InstallDate) {
                                        $app.InstallDate
                                     }
                                     else {
                                        "N/A"
                                     }
                }
            }
        }
        $appList += $apps
    }
}

$appList = $appList | Sort-Object -Property Name -Unique
$appTable = $appList | Format-Table -Property Name, Version, Size, Location, InstalledDate -Wrap -AutoSize | Out-String -Width 500

# --- Gather Installed Microsoft Patches ---
$patches = Get-HotFix
$patchTable = $patches | Format-Table -Property HotFixID, Description, InstalledOn, InstalledBy -Wrap -AutoSize | Out-String -Width 500

# --- Gather Installed Services ---
$services = Get-CimInstance Win32_Service
$serviceTable = $services | Format-Table `
    @{Label="Name"; Expression={$_.Name}}, `
    @{Label="Description"; Expression={
        if ($_.Description) {
            if ($_.Description.Length -gt 50) {
                $_.Description.Substring(0,50) + "..."
            }
            else {
                $_.Description
            }
        }
        else { "" }
    }}, `
    @{Label="Status"; Expression={$_.State}}, `
    @{Label="Startup Type"; Expression={$_.StartMode}}, `
    @{Label="Log On As"; Expression={$_.StartName}}, `
    @{Label="Path to Executable"; Expression={$_.PathName}} -Wrap -AutoSize | Out-String -Width 500

# --- Gather Test Cases ---

# Initialize an array to hold test case objects.
$testCases = @()

# Initialize a counter for Test Case IDs.
$tcCounter = 1

# ----- Test Case: Check if "Administrator" is disabled -----
$tc1Start = Get-Date
$adminUser = Get-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
if ($adminUser) {
    $actualResult1 = -not $adminUser.Enabled
} else {
    $actualResult1 = "Administrator account not found"
}
$tc1End = Get-Date

$expectedResult1 = $true
$status1 = if ($actualResult1 -eq $expectedResult1) { "Pass" } else { "Fail" }

$testCase1 = [PSCustomObject]@{
    "Test ID"          = ("TC{0:D3}" -f $tcCounter)
    "Start Time"       = $tc1Start.ToString("dd-MM-yyyy HH:mm:ss")
    "End Time"         = $tc1End.ToString("dd-MM-yyyy HH:mm:ss")
    "Rule Group"       = "Disable Administrator"
    "Test Description" = "Check if Windows User 'Administrator' is disabled."
    "Expected Result"  = $expectedResult1
    "Actual Result"    = $actualResult1
    "Passing Criteria" = $true
    "Pass/Fail Status" = $status1
}
$testCases += $testCase1
$tcCounter++

# ----- Test Case: Check if "DeltaVADM" is in Administrators group -----
$tc2Start = Get-Date
$adminMembers = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue
$actualResult2 = $false
if ($adminMembers) {
    foreach ($member in $adminMembers) {
        if ($member.Name -eq "DeltaVADM") {
            $actualResult2 = $true
            break
        }
    }
}
$tc2End = Get-Date

$expectedResult2 = $true
$status2 = if ($actualResult2 -eq $expectedResult2) { "Pass" } else { "Fail" }

$testCase2 = [PSCustomObject]@{
    "Test ID"          = ("TC{0:D3}" -f $tcCounter)
    "Start Time"       = $tc2Start.ToString("dd-MM-yyyy HH:mm:ss")
    "End Time"         = $tc2End.ToString("dd-MM-yyyy HH:mm:ss")
    "Rule Group"       = "DeltaVADM as Administrator"
    "Test Description" = "Check if Windows User 'DeltaVADM' is part of Administrators Local group."
    "Expected Result"  = $expectedResult2
    "Actual Result"    = $actualResult2
    "Passing Criteria" = $true
    "Pass/Fail Status" = $status2
}
$testCases += $testCase2
$tcCounter++

# ----- Test Cases for NIC Power On: Check if "Allow Computer to turn off this device" is disabled for each NIC -----
$nicAdapters = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true -and $_.NetConnectionStatus -ne $null }
foreach ($nic in $nicAdapters) {
    $tcNICStart = Get-Date

    # Retrieve the NIC's NetCfgInstanceId for registry lookup.
    $netCfgId = $nic.NetCfgInstanceId
    $regBase = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"
    $matchedKey = Get-ChildItem $regBase | Where-Object {
        (Get-ItemProperty $_.PSPath -Name "NetCfgInstanceId" -ErrorAction SilentlyContinue).NetCfgInstanceId -eq $netCfgId
    } | Select-Object -First 1

    if ($matchedKey) {
        $pnpcap = (Get-ItemProperty $matchedKey.PSPath -Name "PnPCapabilities" -ErrorAction SilentlyContinue).PnPCapabilities
        # Bit 0x20 (32 decimal) indicates that the power management option "Allow the Computer to Turn off this device" is disabled.
        if ($pnpcap -and (($pnpcap -band 32) -eq 32)) {
            $actualResultNIC = "Disabled"
        }
        else {
            $actualResultNIC = "Enabled"
        }
    }
    else {
        $actualResultNIC = "Unknown"
    }
    
    $tcNICEnd = Get-Date
    $expectedResultNIC = "Disabled"
    $statusNIC = if ($actualResultNIC -eq $expectedResultNIC) { "Pass" } else { "Fail" }
    
    $nicTestCase = [PSCustomObject]@{
        "Test ID"          = ("TC{0:D3}" -f $tcCounter)
        "Start Time"       = $tcNICStart.ToString("dd-MM-yyyy HH:mm:ss")
        "End Time"         = $tcNICEnd.ToString("dd-MM-yyyy HH:mm:ss")
        "Rule Group"       = "NIC Power On"
        "Test Description" = "Allow Computer to turn off this device " + $nic.Name
        "Expected Result"  = $expectedResultNIC
        "Actual Result"    = $actualResultNIC
        "Passing Criteria" = "Disabled"
        "Pass/Fail Status" = $statusNIC
    }
    $testCases += $nicTestCase
    $tcCounter++
}




# Format the test cases array into a table.
$testCaseTable = $testCases | Format-Table -Wrap -AutoSize | Out-String -Width 500


# Record the end time and format it as DD-MM-YYYY
$endTime = Get-Date
$endTimeFormatted = $endTime.ToString("dd-MM-yyyy")

# Prepare the document content with the new Test Cases section
$docText = @"
General Document Data

Document Title : Report for $hostname
Descriptive Purpose: SAMPLE TEXT
Classification : General
FileName: Windows-Report - $hostname

Execution Data 
Tester : $currentUser
Start Date : $startTimeFormatted
End Date : $endTimeFormatted

Device Inventory 
Operating System : $os

Hardware Inventory 
$physicalDiskText
Model : $computerModel
Total Physical Memory : $totalMemoryGB GB
Number of CPU : $cpuCount
Cores per Processor : $coresPerProcessor
CPU MODELS : $cpuModels

Disk Inventory 
$partitionText

Network Parameter Inventory 
$networkText

Time Zone Details 
$timeZoneText

Installed Applications 
$appTable

Installed Microsoft Patches 
$patchTable

Installed Services 
$serviceTable

Test Cases
$testCaseTable
"@

# Save the content to the report file in the current directory
$docText | Out-File -FilePath ".\Report.txt" -Encoding UTF8
