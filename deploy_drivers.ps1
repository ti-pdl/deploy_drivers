#######################
# deploy_drivers.ps1
# David Carré @ 2024
#######################

#####################
# script parameters #
#####################

param (
    [string]$srv_path = "", # unc path to "pilotes.md" directory
    [string]$srv_username = "", # srv_path unc share username
    [string]$srv_password = "", # srv_path unc share password
    [string]$search = "", # search a driver on "https://catalog.update.microsoft.com/"
    [switch]$scan = $false, # scan for missing drivers
    [switch]$force = $false, # use "-force" script argument to force execution even if "$log_file" exists
    [switch]$init = $false, # use "-init" script argument to download "data" (pilote table and all drivers on server)
    [switch]$use_mirror = $false, # download from mirror links
    [string]$log_file = "c:\deploy_drivers.log", # log file on client computer
    [string]$db_url = "https://github.com/ti-pdl/wiki/raw/refs/heads/master/system/windows/pilotes.md" # url to database
)

##############################
# custom functions / helpers #
##############################

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$LogLevel = 'Info'
    )

    # Get the current timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Format the log message with a timestamp and log level
    $logMessage = "[$timestamp] [$LogLevel] $Message"

    # Write to console
    if ($LogLevel.Equals("Error")) {
        Write-Error $logMessage
    }
    elseif ($LogLevel.Equals("Warning")) {
        Write-Warning $logMessage
    }
    else {
        Write-Host $logMessage
    }

    # append to log file
    Add-Content -Path $log_file -Value $logMessage > $null 2>&1
}

function GetComputerModel {
    $computer = Get-WmiObject Win32_ComputerSystem

    # hum hum...
    if ($computer.Manufacturer.Equals("LENOVO")) {
        return  $computer.SystemFamily
    }

    return $computer.Model
}

function GetWindowsVersion {
    $winver = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion

    # windows update catlog seems to use 22H2 or 24H2 (we assume drivers for 22H2 are the best for 23H2)
    if ($winver -eq "23H2") {
        $winver = "22H2"
    }

    return $winver
}

function GetDeviceName {
    param (
        [string]$Id
    )

    try {
        # try to find the device and retrieve needed props
        $device = Get-PnpDeviceProperty -InstanceId $Id -KeyName DEVPKEY_Device_DeviceDesc `
            -ErrorAction Stop 2>$null | Select-Object -Property *
        if ($null -ne $device) {
            return $device.Data
        }
    }
    catch {
    }

    return "N/A"
}

function GetDeviceDriver {
    param (
        [string]$Id
    )

    try {
        # try to find the device and retrieve needed props
        $device = Get-PnpDeviceProperty -InstanceId $Id `
            -KeyName DEVPKEY_Device_Manufacturer, DEVPKEY_Device_Class, DEVPKEY_Device_DriverVersion `
            -ErrorAction Stop 2>$null | Select-Object -Property *
        # if success, return a "windows update catalog" compatible "title" (https://catalog.update.microsoft.com/)
        if ($null -ne $device -and $device.Count -eq 3) {
            Write-Output $device.Data[0] "-" $device.Data[1] "-" $device.Data[2]
            return
        }
    }
    catch {
    }

    return $false
}

function QueryMsCatalog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [string]$Name
    )

    $rows = @()

    # get windows product version for later use
    $winver = GetWindowsVersion
    if ($winver -eq "22H2") {
        Write-Host "QueryMsCatalog: searching windows 11 (22H2/23H2) driver for `"$Id`" ($Name)"
    }
    else {
        Write-Host "QueryMsCatalog: searching windows $winver driver for `"$Id`" ($Name)"
    }

    # perform the search by sending an HTTP request
    $uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$([System.Uri]::EscapeDataString($Id))"
    $response = Invoke-WebRequest -Uri $uri
    if ($response.StatusCode -ne "200") {
        return $null
    }

    # find all rows in the HTML table
    $rows = $response.ParsedHtml.getElementsByTagName('tr') | Where-Object {
        $_.getElementsByTagName('td').length -ge 7 -and
        $_.getElementsByTagName('a').length -eq 1
    }

    # nothing found...
    if ($rows.Count -eq 0) {
        return $null
    }

    # prepare an array to store the results
    $results = @()

    # loop through each filtered row
    foreach ($row in $rows) {
        $title = $row.getElementsByTagName('td')[1].getElementsByTagName('a')[0].InnerText.Trim()
        $products = $row.getElementsByTagName('td')[2].InnerText.Trim()
        $classification = $row.getElementsByTagName('td')[3].InnerText.Trim()
        $lastUpdated = $row.getElementsByTagName('td')[4].InnerText.Trim()
        $size = $row.getElementsByTagName('td')[6].InnerText.Trim()

        # add data to results array
        $results += [PSCustomObject]@{
            Title          = $title
            Products       = $products
            Classification = $classification
            LastUpdated    = $lastUpdated
            Size           = $size
            Link           = $uri
        }
    }

    # return driver matching windows product version if any
    foreach ($driver in $results) {
        if ($driver.Products.Contains($winver)) {
            Write-Host "QueryMsCatalog: found windows 11 ($winver) driver"
            return $driver
        }
    }

    # try to find a windows 10 (1903) compatible driver
    foreach ($driver in $results) {
        if ($driver.Products.Contains("1903")) {
            Write-Host "QueryMsCatalog: found windows 10 (1903) compatible driver"
            return $driver
        }
    }

    return $null
}

function FindDriver {
    param (
        [string]$Id
    )

    # get device name...
    $devname = GetDeviceName $Id

    <# DEBUG
    $ids = Get-WmiObject Win32_PnPEntity | Select-Object -ExpandProperty DeviceID
    foreach($i in $ids) {
        if ($i.StartsWith("ACPI")) {
            Write-Host $i
            $i = $i.Substring(0, $i.LastIndexOf(("\")))
            $driver = QueryMsCatalog $i
        }
    }
    exit
    #>

    if (!$id.StartsWith("PCI") -and !$id.StartsWith("ACPI") -and !$id.StartsWith("USB")) {
        Write-Host "FindDriver: skipping `"$id`" (not supported yet)"
        continue # TODO: handle other classes ?
    }

    # handle PCI/APCI id
    if ($id.StartsWith("PCI")) {
        # "PCI\VEN_0000&DEV_0000&SUBSYS_00000000&REV_00\0&00" > "PCI\VEN_0000&DEV_0000&SUBSYS_00000000"
        $id = $id.Substring(0, $id.IndexOf("&REV"))
    }
    else {
        # "ACPI\INTC0000\0&AAAAAAA&0" > "ACPI\INTC0000"
        # "USB\VID_0000&PID_00CD&MI_00\0&00000000&0&0000 > USB\VID_0000&PID_00CD&MI_00"
        $id = $id.Substring(0, $id.LastIndexOf(("\")))
    }

    # query ms catalog
    $driver = QueryMsCatalog $id -Name $devname
    if ($null -eq $driver -and $id.StartsWith("PCI")) {
        # try PCI device id without SUBSYS
        # "PCI\VEN_0000&DEV_0000&SUBSYS_00000000" > "PCI\VEN_0000&DEV_0000"
        $id = $id.Substring(0, $id.IndexOf("&SUBSYS"))
        $driver = QueryMsCatalog $id -Name $devname
    }

    return $driver
}

function FindMissingDrivers {
    Get-PnpDevice -PresentOnly | Where-Object { 
        ($_.Status -ne "OK" -or $_.Description -like "Carte vid*" -or $_.Description -eq "Contrôleur vidéo") -and
        ($_.DeviceID.StartsWith("PCI") -or $_.DeviceID.StartsWith("USB\V") -or $_.DeviceID.StartsWith("ACPI\"))
    } | Select-Object Status, Manufacturer, Description, DeviceID | ForEach-Object {
        FindDriver $_.DeviceID
    }
}

function MapDrive {
    param (
        [string]$unc_path,
        [string]$drive
    )

    Write-Log -Message "MapDrive: mapping $unc_path to ${drive}:" -LogLevel Info

    # extract server address from unc path
    $start = $unc_path.IndexOf("\\") + 2;
    $end = $unc_path.IndexOf("\", $start);
    if ($start -lt 2 -or $end -lt 3 -or $end -gt $unc_path.Length - 1) {
        Write-Log -Message "MapDrive: could not extract server address from $unc_path" -LogLevel Error
        return $false
    }

    # test connexion to server
    $address = $unc_path.Substring($start, $end - $start)
    Write-Log -Message "MapDrive: checking connexion to $address" -LogLevel Info
    if (!(Test-Connection $address)) {
        Write-Log -Message "MapDrive: could not map $unc_path to $drive (connexion to $address failed)" -LogLevel Error
        return $false
    }

    # actually map the drive (TODO: use New-PSDrive ?)
    $net = New-Object -ComObject WScript.Network
    $net.MapNetworkDrive("${drive}:", $unc_path, $false, $srv_username, $srv_password)

    # return true if the drive was correctly mapped
    return Test-Path "${drive}:"
}

function UnMapDrive {
    param (
        [string]$drive
    )

    Write-Log -Message "UnMapDrive: unmapping ${drive}:" -LogLevel Info
    $net = New-Object -ComObject WScript.Network
    $net.RemoveNetworkDrive("${drive}:")
}

# load driver table from markdown table
function LoadDriverDb {
    param (
        [string]$DbPath
    )

    # download driver database if needed
    if (!(Test-Path $DbPath -PathType Leaf)) {
        Write-Log -Message "LoadDriverDb: downloading database from $db_url" -LogLevel Info
        try {
            Invoke-WebRequest -Uri $db_url -OutFile $DbPath -TimeoutSec 5 -ErrorAction Stop
        }
        catch {
            $msg = $_.Exception.Message
            Write-Log -Message "LoadDriverDb: download failed: $msg" -LogLevel Error
            exit 1
        }
    }
    else {
        Write-Log -Message "LoadDriverDb: loading database from $DbPath" -LogLevel Info
    }

    # load markdown page
    $markdown = Get-Content -Path "$DbPath" -Raw

    # "extract" drivers table
    $startIndex = $markdown.IndexOf("| MODEL")
    $endIndex = $markdown.IndexOf("{.dense}")
    if ($startIndex -lt 0 -or $endIndex -lt 0) {
        Write-Log -Message "LoadDriverDb: could not extract driver table from markdown page..." -LogLevel Error
        exit 1
    }

    $markdown = $markdown.Substring($startIndex, $endIndex - $startIndex).Trim()

    # split the markdown content into lines
    $lines = $markdown -split "`n"

    # Extract the header row and remove the column separator row
    $headerLine = $lines[0].Trim('|').Trim()
    $headers = $headerLine -split '\s*\|\s*'

    # Initialize an array to store parsed rows
    $table = @()

    # Parse each row (starting from the third line because second line is separator)
    for ($i = 2; $i -lt $lines.Length; $i++) {
        $line = $lines[$i].Trim('|').Trim()

        if ($line -ne "") {
            $columns = $line -split '\s*\|\s*'
            # Create a hashtable or custom object
            $object = @{}
            for ($j = 0; $j -lt $headers.Length; $j++) {
                if ($headers[$j] -eq "MIRROR" -or $headers[$j] -eq "DDL") {
                    # handle mirror item
                    $start = $columns[$j].IndexOf("http")
                    $end = $columns[$j].Length - 1
                    $object[$headers[$j]] = $columns[$j].Substring($start, $end - $start)
                }
                elseif ($headers[$j] -eq "DRIVER") {
                    # handle driver item
                    $start = 1
                    $end = $columns[$j].IndexOf("]")
                    $object[$headers[$j]] = $columns[$j].Substring($start, $end - $start)
                }
                else {
                    $object[$headers[$j]] = $columns[$j]
                }
            }

            # Convert to a custom object and add to the table array
            $table += [pscustomobject]$object
        }
    }

    Write-Log -Message "LoadDriverDb: loaded $($table.Count) drivers" -LogLevel Info

    return $table
}

# download/cache required drivers on server
# TODO: add md5 check for downloaded drivers/cab?
function InitDriverDb {
    $driversPath = "$PSScriptRoot\drivers"

    # load database
    Remove-Item -Path "$PSScriptRoot\pilotes.md" -Force > $null 2>&1
    $db = LoadDriverDb("$PSScriptRoot\pilotes.md")

    # create driver directory
    $null = New-Item -Path $driversPath -ItemType Directory -Force

    # loop through each drivers
    foreach ($driver in $db) {
        # set download url
        if ($use_mirror) {
            $url = $driver.MIRROR
        }
        else {
            $url = $driver.DDL
        }
        
        if ([string]::IsNullOrEmpty($url)) {
            continue
        }

        # define the URL and the output file path
        $filename = Split-Path $url -Leaf
        $outPath = "$driversPath\$filename"
        if (!(Test-Path $outPath -PathType Leaf)) {
            Write-Log -Message "InitDriverDb: please wait... downloading $url to $outPath..." -LogLevel Info
            # Download the file (disable progress for speed)
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $url -OutFile $outPath
            $ProgressPreference = 'Continue'
        }
        else {
            Write-Log "InitDriverDb: skipping $outPath (file exists)..."
        }
    }
}

# install local drivers
function GetLocalDrivers {
    $model = GetComputerModel
    $path = "C:\drivers\$model"

    if (Test-Path -Path $path) {
        # install all drivers (*.inf)
        Write-Log "GetLocalDrivers: installing all drivers in $path..."
        Start-Process -FilePath "C:\Windows\System32\pnputil.exe" -ArgumentList "/add-driver `"$path\*.inf`" /subdirs /install" -Wait
        # cleanup
        Remove-Item "$path" -Recurse -Force > $null 2>&1
    }
    else {
        Write-Log "GetLocalDrivers: skipping local drivers, directory not found ($path)"
    }
}

# install remote drivers
function GetRemoteDrivers {
    param (
        [string]$Path
    )

    # load database
    $db = LoadDriverDb("$Path\pilotes.md")

    # get local computer model
    $model = GetComputerModel

    # set local drivers path (temp)
    $local_driver_path = ([System.IO.Path]::GetTempPath()) + "deploy_drivers"
    $null = New-Item -Path "$local_driver_path" -ItemType Directory -Force
    Write-Log -Message "GetDrivers: local_driver_path set to $local_driver_path" -LogLevel Info

    # loop through each drivers
    foreach ($driver in $db) {
        if ([string]::IsNullOrEmpty($driver.MODEL) -or [string]::IsNullOrEmpty($driver.ID)) {
            Write-Log -Message "GetDrivers: skipping invalid row..." -LogLevel Warning
            continue
        }

        # check if the driver is for our "model" 
        if ($driver.MODEL -ne "ANY" -and $driver.MODEL -ne $model) {
            Write-Log -Message "GetDrivers: skipping $($driver.DRIVER) ($($driver.MODEL) != $model)" -LogLevel Info
            continue
        }

        # check if device exists on host
        $id = $driver.ID
        $db_drv = $driver.DRIVER
        $host_drv = GetDeviceDriver($id)
        if (!$host_drv) {
            Write-Log -Message "GetDrivers: skipping $db_drv (device not found: $id)" -LogLevel Info
            continue
        }

        # check if driver match, if not process it
        if ("$host_drv" -ne "$db_drv") {
            # setup paths
            $filename = Split-Path $driver.DDL -Leaf
            $remote_file = "$Path\drivers\$filename"
            $tmp_file = "$local_driver_path\$filename"

            # if drivers was already copied locally for another device, skip it
            if (Test-Path $tmp_file -PathType Leaf) {
                Write-Log -Message "GetDrivers: skipping $db_drv (duplicate driver)..." -LogLevel Info
                continue
            }

            # copy the file locally
            Write-Log -Message "GetDrivers: downloading $db_drv ($remote_file)..." -LogLevel Info
            Copy-Item -Path "$remote_file" -Destination "$tmp_file"

            # process driver file (cab/exe)
            if ("$db_drv".Contains("NVIDIA - Display") -and $filename.EndsWith(".exe")) {
                # special case: NVIDIA package
                Write-Log -Message "GetDrivers: installing $db_drv..." -LogLevel Info
                Start-Process -FilePath "$tmp_file" -ArgumentList "-s -noreboot" -Wait
            }
            elseif ("$db_drv".Contains("Advanced Micro Devices, Inc. - Display") -and $filename.EndsWith(".exe")) {
                # special case: AMD package
                Write-Log -Message "GetDrivers: installing $db_drv..." -LogLevel Info
                # extract (and install for recent amd packages)
                Start-Process -FilePath "$tmp_file" -ArgumentList "-install" -Wait
                # legacy installer doesn't seems to install itself
                if ($filename.StartsWith("radeon-software-adrenalin")) {
                    # install from extracted installer
                    Write-Log -Message "GetDrivers: detected legacy amd driver..." -LogLevel Info
                    if (Test-Path "C:\AMD") {
                        $setup = Get-ChildItem -Path "C:\AMD" -Depth 1 -Filter "Setup.exe" | Select-Object -First 1
                        Write-Log -Message "GetDrivers: installing legacy driver from $($setup.FullName)..." -LogLevel Info
                        Start-Process -FilePath "$($setup.FullName)" -ArgumentList "-uninstall" -Wait
                        Start-Process -FilePath "$($setup.FullName)" -ArgumentList "-install" -Wait
                    }
                    else {
                        Write-Log -Message "GetDrivers: could not install $db_drv (path not found: C:\AMD)" -LogLevel Error
                    }
                }
                # cleanup
                Remove-Item "C:\AMD" -Recurse -Force > $null 2>&1
            }
            else {
                # extract cab content to local drivers path
                Write-Log -Message "GetDrivers: extracting $db_drv from $tmp_file..." -LogLevel Info
                $tmp_path = "$local_driver_path\$db_drv"
                $null = New-Item -Path "$tmp_path" -ItemType Directory -Force
                Start-Process -FilePath "C:\Windows\System32\expand.exe" -ArgumentList "`"$tmp_file`" -F:* `"$tmp_path`"" -Wait
            }
        }
        else {
            Write-Log -Message "GetDrivers: skipping $db_drv (already installed)..." -LogLevel Info
        }
    }

    # install all extracted drivers (*.inf)
    Write-Log -Message "GetDrivers: installing all drivers in $local_driver_path..." -LogLevel Info
    Start-Process -FilePath "C:\Windows\System32\pnputil.exe" -ArgumentList "/add-driver `"$local_driver_path\*.inf`" /subdirs /install" -Wait

    # cleanup
    Remove-Item "$local_driver_path" -Recurse -Force > $null 2>&1
}

####################
# main entry point #
####################

# search for a device id on ms catalog
if ($search.Length -gt 0) {
    $drv = FindDriver "$search"
    if ($drv) {
        $model = GetComputerModel
        $name = GetDeviceName $search
        # output driver info
        Write-Output $drv
        # output markdown row for the wiki database (http://wiki.mydedibox.fr/system/windows/pilotes)
        Write-Output "Markdown:`n| $model | $name | $search | [$($drv.Title)]($($drv.Link)) | [:floppy_disk:](TODO) | [:floppy_disk:](TODO) | NON |"
    }
    return
}

# scan for missing drivers
if ($scan) {
    return FindMissingDrivers
}

# init db...
if ($init) {
    # change log location and cleanup logs
    $log_file = "$PSScriptRoot\init.log"
    Remove-Item -Path "$log_file" -Force > $null 2>&1
    Write-Log "Downloading drivers..."
    InitDriverDb
    Write-Log "All done..."
    return
}

# safety checks
if ([string]::IsNullOrEmpty($srv_path) `
        -or [string]::IsNullOrEmpty($srv_username) `
        -or [string]::IsNullOrEmpty($srv_password)) {
    Write-Log "One or more argments are missing, exiting..." -LogLevel Error
    return 1
}

# stop here if script was already executed (log file exists) and "-force" parameter is not set
if (!$force -and (Test-Path $log_file -PathType Leaf)) {
    Write-Host "Drivers already installed, skipping (use -force to... force)"
    return
}

# cleanup logs
Remove-Item -Path "$log_file" -Force > $null 2>&1

# first look for drivers on local computer ("C:\drivers")
GetLocalDrivers

# now the real deal
if (!(MapDrive -unc_path $srv_path -drive "r")) {
    Write-Log "MapDrive: failed to map drive, exiting..." -LogLevel Error
    return 1
}

# main stuff
GetRemoteDrivers("r:")

# unmap drive
UnMapDrive -drive "r"

# great
Write-Log "All done..."
