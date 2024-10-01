param (
    [switch]$init_db = $false,
    [switch]$use_mirror = $false,
    [string]$db_url = "https://github.com/ti-pdl/wiki/raw/refs/heads/master/serveurs/windows/pilotes.md",
    #[string]$db_path = "$PSScriptRoot\pilotes.md",
    [string]$server_path = "\\srv-applis\Drivers"
    #[string]$username = "",
    #[string]$password = ""
)

#######################
# deploy_drivers.ps1
# David Carré @ 2024
#######################

##############################
# custom functions / helpers #
##############################

function GetComputerModel {
    (Get-WmiObject Win32_ComputerSystem).Model
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
        #Write-Warning "GetDeviceDriver($Id): $($_.Exception.Message)"
    }

    return $false
}

# load driver table from markdown table
function LoadDriverDb {
    param (
        [string]$DbPath
    )

    # download driver database if needed
    if (!(Test-Path $DbPath -PathType Leaf)) {
        Write-Host "LoadDriverDb: downloading database from $db_url"
        try {
            Invoke-WebRequest -Uri $db_url -OutFile $DbPath -TimeoutSec 5 -ErrorAction Stop
        }
        catch {
            $msg = $_.Exception.Message
            Write-Error "LoadDriverDb: download failed: $msg"
            exit 1
        }
    }
    else {
        Write-Host "LoadDriverDb: loading database from $DbPath"
    }

    # load markdown page
    $markdown = Get-Content -Path "$DbPath" -Raw

    # "extract" drivers table
    $startIndex = $markdown.IndexOf("| MODEL")
    $endIndex = $markdown.IndexOf("{.dense}")
    if ($startIndex -lt 0 -or $endIndex -lt 0) {
        Write-Error "LoadDriverDb: could not extract driver table from markdown page..."
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

    Write-Host "LoadDriverDb: loaded $($table.Count) drivers"

    return $table
}

# download/cache required drivers on server
# TODO: add md5 check for downloaded drivers/cab?
function InitDriverDb {
    $driversPath = "$PSScriptRoot\drivers"

    # load database
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
            Write-Host "InitDriverDb: please wait... downloading $url to $outPath..."
            # Download the file (disable progress for speed)
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $url -OutFile $outPath
            $ProgressPreference = 'Continue'
        }
        else {
            Write-Host "InitDriverDb: skipping $outPath (file exists)..."
        }
    }
}

function GetDrivers {
    param (
        [string]$Path
    )

    # convert the password to a SecureString and create creds
    #$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    #$creds = New-Object System.Management.Automation.PSCredential ($username, $securePassword)

    # load database
    $db = LoadDriverDb("$Path\pilotes.md")

    # get local computer model
    $model = GetComputerModel

    # set local drivers path
    $local_driver_path = ([System.IO.Path]::GetTempPath()) + "deploy_drivers"
    $null = New-Item -Path "$local_driver_path" -ItemType Directory -Force

    # loop through each drivers
    foreach ($driver in $db) {
        if ([string]::IsNullOrEmpty($driver.MODEL) -or [string]::IsNullOrEmpty($driver.ID)) {
            Write-Waring "GetDrivers: skipping invalid row..."
            continue
        }

        if ($driver.MODEL -ne $model) {
            Write-Host "GetDrivers: skipping $($driver.DRIVER) ($($driver.MODEL) != $model)"
            continue
        }

        # check if device exists on host
        $id = $driver.ID
        $db_drv = $driver.DRIVER
        $host_drv = GetDeviceDriver($id)
        if (!$host_drv) {
            Write-Host "GetDrivers: skipping $db_drv (device not found: $id)"
            continue
        }

        # check if driver match, if not process it
        if ("$host_drv" -ne "$db_drv") {
            $filename = Split-Path $driver.DDL -Leaf
            $cab_path = "$Path\drivers\$filename"

            # copy cab file to a temp folder and extract it's content
            Write-Host "GetDrivers: downloading $db_drv ($cab_path)..."
            $tmp_file = ([System.IO.Path]::GetTempPath()) + $filename
            Copy-Item -Path "$cab_path" -Destination "$tmp_file"

            # extract cab content to local drivers path
            Write-Host "GetDrivers: extracting $db_drv ($cab_path)..."
            $tmp_path = "$local_driver_path\$db_drv"
            $null = New-Item -Path "$tmp_path" -ItemType Directory -Force
            expand "$tmp_file" -F:* "$tmp_path" > $null

            # cleanup temp cab file
            $null = Remove-Item "$tmp_file" -Force
        }
        else {
            Write-Host "GetDrivers: skipping $db_drv (already installed)..."
        }
    }

    # install all extracted drivers (*.inf)
    pnputil /add-driver "$local_driver_path\*.inf" /subdirs /install

    # cleanup
    $null = Remove-Item "$local_driver_path" -Recurse -Force
}

####################
# main entry point #
####################

if ($init_db) {
    Write-Host "Downloading drivers..."
    InitDriverDb
    Write-Host "All done..."
}
else {
    GetDrivers($server_path)
}
