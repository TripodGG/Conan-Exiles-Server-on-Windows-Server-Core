# Name: Conan Exiles Dedicated Server setup
# Author: TripodGG
# Purpose: Download and install all necessary files to host a Conan Exiles Dedicated Server
# License: MIT License, Copyright (c) 2024 TripodGG




# Clear the screen
Clear-Host

########################
# Checks and Functions #
########################

# Check the version of Windows Server to confirm it is a supported version
$osVersion = (Get-CimInstance Win32_OperatingSystem).Version

# Check if the OS is Windows Server 2016/2019
if ($osVersion -match '10\.0\.(14393|17763)') {
    Write-Host "Windows Server 2016/2019 detected." -ForegroundColor Yellow
}
# Check if the OS Windows Server 2022
elseif ($osVersion -match '10\.0\.(20348)') {
    Write-Host "Windows Server 2022 detected." -ForegroundColor Yellow
}
else {
    Write-Host "Unsupported Windows Server version. Please use a supported version of Windows Server." -ForegroundColor Red
	exit
}

# Function for error logging
function Log-Error {
    param (
        [string]$ErrorMessage
    )

    $LogPath = Join-Path $env:USERPROFILE -ChildPath "ConanExiles\errorlog.txt"

    try {
        # Create the logs directory if it doesn't exist
        $LogsDirectory = Join-Path $env:USERPROFILE -ChildPath "ConanExiles"
        if (-not (Test-Path $LogsDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $LogsDirectory -Force | Out-Null
        }

        # Append the error message with timestamp to the log file
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogMessage = "[$Timestamp] $ErrorMessage"
        $LogMessage | Out-File -Append -FilePath $LogPath

        # Inform the user about the error and the log location
        Write-Host "An error occurred. Check the log file for details: $LogPath" -ForegroundColor Red
    } catch {
        Write-Host "Failed to log the error. Please check the log file manually: $LogPath" -ForegroundColor Red
    }
}

# Function to check for an active Internet connection
function Test-InternetConnection {
    $pingResult = Test-Connection -ComputerName "www.google.com" -Count 1 -ErrorAction SilentlyContinue

    if ($pingResult -eq $null) {
        Log-Error "No active internet connection detected. Please ensure that you are connected to the internet before running this script."
        exit 1
    } else {
        Write-Host "Internet connection detected. Proceeding with the script..." -ForegroundColor Cyan
    }
}

# Call the function to check for an active internet connection
Test-InternetConnection

# Function to alert the user to the text by blinking
function Blink-Message {
    param(
        [String]$Message,
        [int]$Delay,
        [int]$Count,
        [ConsoleColor[]]$Colors
    )

    $startColor = [Console]::ForegroundColor
    $colorCount = $Colors.Length

    for ($i = 0; $i -lt $Count; $i++) {
        [Console]::ForegroundColor = $Colors[$($i % $colorCount)]
        [Console]::WriteLine($Message)
        Start-Sleep -Milliseconds $Delay
    }

    [Console]::ForegroundColor = $startColor
}

# Function to prompt user for install path
function Get-ValidDirectory {
    $attempts = 0

    do {
        # Prompt user for the installation directory
		Start-Sleep -Seconds 5
        $installPath = Read-Host "Enter the directory where you would like to install the dedicated server. (i.e. 'C:\ConanServer')"

        # Check if the path is valid
        if (Test-Path $installPath -IsValid) {
            # Check if the path is a container (directory)
            if (-not (Test-Path $installPath -PathType Container)) {
                # Prompt user to create the directory
                $createDirectory = Read-Host "The directory does not exist. Would you like to create it? (yes/no)"
                if ($createDirectory -eq 'y' -or $createDirectory -eq 'yes') {
                    try {
                        # Attempt to create the directory
                        New-Item -ItemType Directory -Path $installPath -Force | Out-Null
                        Write-Host "Directory created successfully at $installPath" -ForegroundColor Green
                    } catch {
                        # Handle unexpected error during directory creation
                        $errorMessage = "An unexpected error occurred while creating the installation directory. Error: $_" 
                        Log-Error $errorMessage
                        exit 1
                    }
                } elseif ($createDirectory -eq 'n' -or $createDirectory -eq 'no') {
                    # User chose not to create the directory
                    Write-Host "Installation directory not created. Please choose a valid directory." -ForegroundColor Red
                    continue
                } else {
                    # Invalid input for createDirectory
                    Write-Host "Invalid input. Please enter 'yes' or 'no'." -ForegroundColor Red
                    continue
                }
            }
            # Return the validated installation path
            return $installPath
        } else {
            # Invalid path entered by the user
            Write-Host "Invalid path. Please enter a valid directory path." -ForegroundColor Red
        }

        # Increment attempts and exit if reached the limit
        $attempts++
        if ($attempts -eq 3) {
            Write-Host "Failed after 3 attempts. Exiting script." -ForegroundColor Red
            exit 1
        }
    } while ($true)
}

# Success message
try {
    # Call the Get-ValidDirectory function
    $installDirectory = Get-ValidDirectory
    Write-Host "Success! The dedicated server will be installed to $installDirectory" -ForegroundColor Green
} catch {
    # Handle unexpected error during the installation directory prompt
    $errorMessage = "An unexpected error occurred during the installation directory prompt. Error: $_"
    Log-Error $errorMessage
    exit 1
}

# Function to check if a firewall rule exists
function Test-FirewallRule {
    param([string]$DisplayName)
    return (Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction SilentlyContinue)
}

# Function to create a scheduled backup
function Create-ServerBackupScheduledTask {
    param (
        [string]$installDirectory,
        [string]$defaultBackupDirectory
    )

    # Prompt for backup directory location
    do {
        $backupDirectory = Read-Host "Enter the location for the backup directory (default: $defaultBackupDirectory)" 
        if (-not $backupDirectory) {
            $backupDirectory = $defaultBackupDirectory
        }

        # Check if the directory exists
        if (Test-Path $backupDirectory) {
            $createNewDirectory = $null
            # Directory exists, ask for confirmation to use it
            while ($createNewDirectory -notin @('Y', 'N')) {
                $createNewDirectory = Read-Host "The directory '$backupDirectory' already exists. Do you want to use it? (Y/N)" 
            }
            if ($createNewDirectory -eq 'N') {
                continue
            }
        } else {
            $createNewDirectory = $null
            # Directory doesn't exist, ask for confirmation to create it
            while ($createNewDirectory -notin @('Y', 'N')) {
                $createNewDirectory = Read-Host "The directory '$backupDirectory' does not exist. Do you want to create it? (Y/N)" 
            }
            if ($createNewDirectory -eq 'Y') {
                New-Item -ItemType Directory -Path $backupDirectory | Out-Null
            } else {
                continue
            }
        }
    } while (-not (Test-Path $backupDirectory))

    # Ensure $backupDirectory is not inside $installDirectory
    if ($backupDirectory -like "$installDirectory*") {
        Write-Host "Error: Backup directory cannot be inside the install directory." -ForegroundColor Red
        return
    }

    # Prompt for task frequency
    $taskFrequency = Read-Host "How often should the backup run? (Daily, Weekly, Bi-weekly, Monthly)" 

    # Prompt for task execution time
    $taskExecutionTime = $null
    while (-not $taskExecutionTime) {
        $taskExecutionTime = Read-Host "What time should the backup run? (24-hour format, e.g., 14:30)" 
    }

    # Create scheduled task with repetition to run indefinitely
    $taskAction = New-ScheduledTaskAction -Execute "robocopy" -Argument "$installDirectory $backupDirectory /MIR /R:0 /W:0 /NFL /NDL /NP"

    # Create a string representing the repetition interval
    $repetitionInterval = 'P99999DT23H59M59S'

    # Create a string representing the repetition duration
    $repetitionDuration = 'P99999DT23H59M59S'

    # Create the scheduled task using schtasks.exe
    schtasks.exe /Create /TN "Server Backup - $(Get-Date -Format 'yyyyMMddHHmmss')" /TR "robocopy '$installDirectory' '$backupDirectory' /MIR /R:0 /W:0 /NFL /NDL /NP" /SC ONCE /ST $taskExecutionTime /RI $repetitionInterval /DU $repetitionDuration /F

    Write-Host "Scheduled backup task created successfully." -ForegroundColor Green
}

# Function to check if Visual C++ Redistributable 2005 is installed
function Check-VCRedist2005Installed {
    $redistVersion2005 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\8.0\VC\Runtimes\x64' -ErrorAction SilentlyContinue

    if ($redistVersion2005 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable 2008 is installed
function Check-VCRedist2008Installed {
    $redistVersion2008 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{9A25302D-30C0-39D9-BD6F-21E6EC160475}' -ErrorAction SilentlyContinue

    if ($redistVersion2008 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable 2010 is installed
function Check-VCRedist2010Installed {
    $redistVersion2010 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\10.0\VC\VCRedist\x64' -ErrorAction SilentlyContinue

    if ($redistVersion2010 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable 2012 is installed
function Check-VCRedist2012Installed {
    $redistVersion2012 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\{ca67548a-5ebe-413a-b50c-4b9ceb6d66c6}' -ErrorAction SilentlyContinue

    if ($redistVersion2012 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable 2013 is installed
function Check-VCRedist2013Installed {
    $redistVersion2013 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Classes\Installer\Dependencies\{050d4fc8-5d48-4b8f-8972-47c82c46020f}' -ErrorAction SilentlyContinue

    if ($redistVersion2013 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable 2015 is installed
function Check-VCRedist2015Installed {
    $redistVersion2015 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64' -ErrorAction SilentlyContinue

    if ($redistVersion2015 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable 2017 is installed
function Check-VCRedist2017Installed {
    $redistVersion2017 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64' -ErrorAction SilentlyContinue

    if ($redistVersion2017 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable 2022 is installed
function Check-VCRedist2022Installed {
    $redistVersion2022 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64' -ErrorAction SilentlyContinue

    if ($redistVersion2022 -eq $null) {
        return $false
    } else {
        return $true
    }
}

# Function to check if Visual C++ Redistributable is installed
function Check-VCRedist {
    $vcRedistInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft Visual C++ % Redistributable%'" -ErrorAction SilentlyContinue

    if ($vcRedistInstalled) {
        Write-Host "Visual C++ Redistributable is already installed."
    } else {
        Install-VCRedist
    }
}

# Function to install Visual C++ Redistributable using Chocolatey
function Install-VCRedist {
    try {
        Write-Host "Installing Visual C++ Redistributable using Chocolatey..."  -ForegroundColor Cyan
        choco install vcredist2005 vcredist2008 vcredist2010 vcredist2012 vcredist2013 vcredist140 -y
        Write-Host "Visual C++ Redistributable installed successfully." -ForegroundColor Green
    } catch {
        # Handle unexpected error during installation
        $errorMessage = "An unexpected error occurred while installing Visual C++ Redistributable. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
}

# Function to check if a command is available
function CommandExists($command) {
    Get-Command $command -ErrorAction SilentlyContinue
}



########################
# Install updates and  #
#    dependencies      #
########################

# Check if NuGet is installed
if (-not (Get-Module -ListAvailable -Name NuGet)) {
    try {
        # NuGet is not installed, so install it silently
        Install-PackageProvider -Name NuGet -Force -ForceBootstrap -Scope CurrentUser -Confirm:$false
        Install-Module -Name NuGet -Force -Scope CurrentUser -Confirm:$false
    } catch {
        $errorMessage = "Failed to install NuGet. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
}

# Make sure all Windows updates have been applied - This can also be done from sconfig under option 6
# Install the PSWindowsUpdate module
Install-Module -Name PSWindowsUpdate -Force

# Import the module
Import-Module PSWindowsUpdate

# Set the execution policy
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

# Check for and install KB5009608 (required to run  Windows desktop application compatibility)
$kbNumber = 'KB5009608'
$installedUpdate = Get-HotFix | Where-Object {$_.HotFixID -eq $kbNumber}

if ($installedUpdate) {
    Write-Host "$kbNumber is already installed on this system." -ForegroundColor Green
} else {
    # Prompt the user to install KB5009608
    $userChoice = Read-Host -Prompt "$kbNumber is required to run this game server. Do you want to install $kbNumber? (Y/N)" 

    if ($userChoice -eq 'Y' -or $userChoice -eq 'Yes') {
        Get-WindowsUpdate -KBArticleID $kbNumber -Install -AcceptAll
        Write-Host "Installing $kbNumber..." -ForegroundColor Cyan
    } else {
        Write-Host "Installation of $kbNumber canceled." -ForegroundColor Red
    }
}

# Check for and install any additional updates
Get-WindowsUpdate -Install -AcceptAll

# Check if .NET Framework 3.5 is installed
$dotNet35 = Get-WindowsFeature -Name "NET-Framework-Features" | Where-Object {$_.Name -eq "NET-Framework-Core"}

if ($dotNet35.Installed) {
    Write-Host ".NET Framework 3.5 is already installed." -ForegroundColor Green
} else {
    # Install .NET Framework 3.5
    Write-Host "Installing .NET Framework 3.5..." -ForegroundColor Cyan

    Install-WindowsFeature -Name NET-Framework-Core -IncludeAllSubFeature -Restart

    Write-Host ".NET Framework 3.5 has been installed." -ForegroundColor Green
}

# Function to check if a reboot is pending due to updates
function Check-And-RebootIfNeeded {
    $pendingReboot = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ErrorAction SilentlyContinue).RebootRequired

    if ($pendingReboot) {
        # Reboot is required, prompt the user and proceed with reboot if agreed
        $rebootChoice = Read-Host -Prompt "A reboot is required to install updates. Would you like to reboot now? (Y/N)" 

        if ($rebootChoice -eq 'Y' -or $rebootChoice -eq 'y') {
            Write-Host "Rebooting the machine. Please run the script again after the server reboot." -ForegroundColor Green
            Restart-Computer -Force
        } else {
            Write-Host "You chose not to reboot. The script will now exit. Please manually reboot the server and run the script again." -ForegroundColor Cyan
            return $false
        }
    } else {
        # No reboot is required
        Write-Host "No reboot is required at the moment." -ForegroundColor Cyan
    }

    return $true
}

# Check and reboot if needed
$rebootCompleted = Check-And-RebootIfNeeded

# Check the version of Windows Server and add the correct Windows desktop application compatibility files
$osVersion = (Get-CimInstance Win32_OperatingSystem).Version

# Check if the OS is Windows Server 2016/2019
if ($osVersion -match '10\.0\.(14393|17763)') {
    Write-Host "Installing App Compatibility Tools for Windows Server 2016/2019." -ForegroundColor Cyan
    # Run the command for Server 2016 or 2019
    Add-WindowsCapability -Online -Name ServerCore.AppCompatibility
}
# Check if the OS Windows Server 2022
elseif ($osVersion -match '10\.0\.(20348)') {
    Write-Host "Installing App Compatibility Tools for Windows Server 2022." -ForegroundColor Cyan
    # Run the command for Server 2022
    Add-WindowsCapability -Online -Name ServerCore.AppCompatibility~~~~0.0.1.0
}
else {
    Write-Host "Continuing with installation..." -ForegroundColor Green
}

# Install DirectX Configuration Database
Write-Host "Installing DirectX Configuration Database." -ForegroundColor Cyan
Add-WindowsCapability -Online -Name DirectX.Configuration.Database~~~~0.0.1.0

# Download and install the DirectX development kit
#Write-Host "DirectX Software Development Kit..." -ForegroundColor Cyan
#Write-Host "Downloading..." -ForegroundColor Blue
#$exePath = "$env:temp\DXSDK_Jun10.exe" # set the temporary download path
#(New-Object Net.WebClient)
#Invoke-WebRequest -Uri 'https://download.microsoft.com/download/A/E/7/AE743F1F-632B-4809-87A9-AA1BB3458E31/DXSDK_Jun10.exe' -OutFile $exePath
#Write-Host "Installing..." -ForegroundColor Cyan
#$installPath = "C:\Program Files (x86)\Microsoft DirectX SDK"
#cmd /c start /wait $exePath /P $installPath /U

#Remove-Item $exePath
#Write-Host "Installed!" -ForegroundColor Green

# Download the DirectX End-User Runtimes (June 2010) offline installer
# Invoke-WebRequest -Uri https://download.microsoft.com/download/8/4/A/84A35BF1-DAFE-4AE8-82AF-AD2AE20B6B14/directx_Jun2010_redist.exe -UseBasicParsing -OutFile ~\Downloads\directx_june2010_redist.exe


# Check if Scoop is installed
if (-not (CommandExists 'scoop')) {
    try {
        # Scoop is not installed, so install it
        Write-Host "Scoop is not installed. Installing Scoop..." -ForegroundColor Cyan
        
        # Run the Scoop installation command with elevated privileges
        iex "& {$(irm get.scoop.sh)} -RunAsAdmin"

        # Check if Scoop installation was successful
        if (CommandExists 'scoop') {
            Write-Host "Scoop installed successfully." -ForegroundColor Green
        } else {
            $errorMessage = "Failed to install Scoop. Please check your internet connection and try again."
            Log-Error $errorMessage
            exit 1
        }
    } catch {
        $errorMessage = "An unexpected error occurred during Scoop installation. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
} else {
    # Scoop is already installed
    Write-Host "Scoop is already installed. Skipping installation." -ForegroundColor Green
}

# Check if Chocolatey is installed
if (-not (CommandExists 'choco')) {
    try {
        # Chocolatey is not installed, so install it
        Write-Host "Chocolatey is not installed. Installing Chocolatey..." -ForegroundColor Cyan
        
        # Run the Chocolatey installation command with elevated privileges
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

        # Check if Chocolatey installation was successful
        if (CommandExists 'choco') {
            Write-Host "Chocolatey installed successfully." -ForegroundColor Green
        } else {
            $errorMessage = "Failed to install Chocolatey. Please check your internet connection and try again."
            Log-Error $errorMessage
            exit 1
        }
    } catch {
        $errorMessage = "An unexpected error occurred during Chocolatey installation. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
} else {
    # Chocolatey is already installed
    Write-Host "Chocolatey is already installed. Skipping installation." -ForegroundColor Green
}

# Check if Git is installed
if (-not (CommandExists 'git')) {
    try {
        # Git is not installed, so install it using Chocolatey
        Write-Host "Git is not installed. Installing Git..." -ForegroundColor Cyan
        scoop install git

        # Check if Git installation was successful
        if (CommandExists 'git') {
            Write-Host "Git installed successfully." -ForegroundColor Green
        } else {
            $errorMessage = "Failed to install Git. Please check your internet connection and try again."
            Log-Error $errorMessage
            exit 1
        }
    } catch {
        $errorMessage = "An unexpected error occurred during Git installation. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
} else {
    # Git is already installed
    Write-Host "Git is already installed. Skipping installation." -ForegroundColor Green
}

# Check if Scoop Extras bucket is added
if (-not (Test-Path "$env:USERPROFILE\scoop\buckets\extras")) {
    try {
        # Scoop Extras bucket is not added, so add it
        Write-Host "Scoop Extras bucket is not added. Adding Scoop Extras bucket..." -ForegroundColor Cyan
        
        # Run the Scoop command to add the Extras bucket
        scoop bucket add extras

        # Check if Scoop Extras bucket addition was successful
        if (Test-Path "$env:USERPROFILE\scoop\buckets\extras") {
            Write-Host "Scoop Extras bucket added successfully."
        } else {
            $errorMessage = "Failed to add Scoop Extras bucket. Please check your internet connection and try again."
            Log-Error $errorMessage
            exit 1
        }
    } catch {
        $errorMessage = "An unexpected error occurred during Scoop Extras bucket addition. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
} else {
    # Scoop Extras bucket is already added
    Write-Host "Scoop Extras bucket is already added. Skipping addition." -ForegroundColor Green
}

# Check if Nano for Windows is installed
if (-not (CommandExists 'nano')) {
    try {
        # Nano for Windows is not installed, so install it using Scoop
        Write-Host "Nano for Windows is not installed. Installing Nano for Windows..." -ForegroundColor Cyan
        scoop install nano

        # Check if Nano for Windows installation was successful
        if (CommandExists 'nano') {
            Write-Host "Nano for Windows installed successfully." -ForegroundColor Green
        } else {
            $errorMessage = "Failed to install Nano for Windows. Please check your internet connection and try again."
            Log-Error $errorMessage
            exit 1
        }
    } catch {
        $errorMessage = "An unexpected error occurred during Nano for Windows installation. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
} else {
    # Nano for Windows is already installed
    Write-Host "Nano for Windows is already installed. Skipping installation." -ForegroundColor Green
}

# Check if NSSM is installed
if (-not (Get-Command nssm -ErrorAction SilentlyContinue)) {
    try {
        # NSSM is not installed, so install it using Chocolatey
        Write-Host "NSSM is not installed. Installing NSSM..." -ForegroundColor Cyan
        choco install nssm -y

        # Check if NSSM installation was successful
        if (Get-Command nssm -ErrorAction SilentlyContinue) {
            Write-Host "NSSM installed successfully." -ForegroundColor Green
        } else {
            $errorMessage = "Failed to install NSSM. Please check the Chocolatey logs for more information."
            Log-Error $errorMessage
            exit 1
        }
    } catch {
        $errorMessage = "An unexpected error occurred during NSSM installation. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
} else {
    # NSSM is already installed
    Write-Host "NSSM is already installed. Skipping installation." -ForegroundColor Green
}

# Check if SteamCMD is installed
if (-not (CommandExists 'steamcmd')) {
    try {
        # SteamCMD is not installed, so install it using Scoop
        Write-Host "SteamCMD is not installed. Installing SteamCMD..." -ForegroundColor Cyan
        scoop install steamcmd

        # Check if SteamCMD installation was successful
        if (CommandExists 'steamcmd') {
            Write-Host "SteamCMD installed successfully." -ForegroundColor Green
        } else {
            $errorMessage = "Failed to install SteamCMD. Please check your internet connection and try again."
            Log-Error $errorMessage
            exit 1
        }
    } catch {
        $errorMessage = "An unexpected error occurred during SteamCMD installation. Error: $_"
        Log-Error $errorMessage
        exit 1
    }
} else {
    # SteamCMD is already installed
    Write-Host "SteamCMD is already installed. Skipping installation." -ForegroundColor Green
}

# Check if Visual C++ Redistributable is installed
Check-VCRedist

# Install Conan dedicated server 
steamcmd +force_install_dir $installDirectory +login anonymous +app_update 443030 validate +quit

# Run the Conan server executable to generate necessary files and folders
Write-Host "Generating Conan Exiles files and folders.  Please wait..."  -ForegroundColor Cyan
start $installDirectory\ConanSandboxServer.exe

# Pause to allow the server to fully start
Start-Sleep -Seconds 60

# Stop the Conan server
taskkill /F /IM ConanSandboxServer.exe



####################
# User prompts for # 
#  configurations  #
####################

# Create the Conan server config files and write the contents to the files
# Prompt the user for server name
$serverName = Read-Host -Prompt "Enter the name you would like to use for your game server"

# Prompt the user for server password
$serverPassword = Read-Host -Prompt "Enter the password you would like to use for your game server"

# Prompt the user for the game port with default value 7777
$defaultGamePort = 7777
$gamePort = Read-Host -Prompt "Enter the game port to be used (press enter to use default port: $defaultGamePort)"
if (-not $gamePort) {
    $gamePort = $defaultGamePort
}

# Prompt the user for the peer port with default value 7778
$defaultPeerPort = 7778
$peerPort = Read-Host -Prompt "Enter the peer port to be used (press enter to use default port: $defaultPeerPort)"
if (-not $peerPort) {
    $peerPort = $defaultPeerPort
}

# Ensure that the peer port is not more than 1 higher than the game port
while ($peerPort -gt ($gamePort + 1)) {
    Write-Host "Error: The peer port cannot be more than one port number higher than the game port." -ForegroundColor Red
    $gamePort = Read-Host -Prompt "Enter the game port (press enter to use default port: $defaultGamePort)"
    if (-not $gamePort) {
        $gamePort = $defaultGamePort
    }
    $peerPort = Read-Host -Prompt "Enter the query port (press enter to use default port: $defaultPeerPort)"
    if (-not $peerPort) {
        $peerPort = $defaultPeerPort
    }
}

# Prompt the user for the query port with default value 27015
$defaultQueryPort = 27015
$queryPort = Read-Host -Prompt "Enter the query port to be used (press enter to use default port: $defaultQueryPort)"
if (-not $queryPort) {
    $queryPort = $defaultQueryPort
}

# Prompt the user for the RCON port with a default value
do {
    $rconPort = Read-Host "Enter the port for RCON (Press Enter for default: 25575)"
    if ($rconPort -eq "") {
        $rconPort = 25575  # Set default value
        break
    }
    elseif (-not ([int]::TryParse($rconPort, [ref]$null)) -or $rconPort -lt 1 -or $rconPort -gt 65535) {
        Write-Host "Invalid input. Please enter a valid port number (1-65535) or press Enter for default." -ForegroundColor Red
    }
} until ([int]::TryParse($rconPort, [ref]$null) -and $rconPort -ge 1 -and $rconPort -le 65535)

# Prompt the user for the number of players allowed (with a maximum of 70)
$maxPlayers = Read-Host -Prompt "Enter the maximum number of players (up to 70)"
$maxPlayers = [math]::Min(70, [math]::Max(1, [int]$maxPlayers)) # Ensure the value is between 1 and 70

# Prompt the user for the level of nudity that should be allowed
do {
    # Prompt the user for the level of nudity
    $maxNudity = Read-Host "What level of nudity do you want on the server? (0=None, 1=Partial, 2=Full)"

    # Validate the input
    if ($maxNudity -notin '0','1','2') {
        Write-Host "Invalid input. Please enter 0, 1, or 2." -ForegroundColor Red
    }
} until ($maxNudity -in '0','1','2')

# Prompt the user for enabling Blitz mode
$enableBlitz = Read-Host "Would you like to enable Blitz mode? (accelerated progression) [y/n]"

# Check the user's input
if ($enableBlitz -eq 'y' -or $enableBlitz -eq 'yes') {
    $pvpBlitzServer = $true
}
elseif ($enableBlitz -eq 'n' -or $enableBlitz -eq 'no') {
    $pvpBlitzServer = $false
}
else {
    Write-Host "Invalid input. Please enter 'y' or 'n'." -ForegroundColor Red
}

# Prompt the user for enabling PVP
$enablePVP = Read-Host "Would you like to enable PVP? [y/n]"

# Check the user's input
if ($enablePVP -eq 'y' -or $enablePVP -eq 'yes') {
    $pvpEnabled = $true
}
elseif ($enablePVP -eq 'n' -or $enablePVP -eq 'no') {
    $pvpEnabled = $false
}
else {
    Write-Host "Invalid input. Please enter 'y' or 'n'." -ForegroundColor Red
}

# Prompt the user for the server region
do {
    $serverRegion = Read-Host "What region do you want to list your server in? (0=EU, 1=NA, 2=Asia) [Press Enter for default: 1]"
    if ($serverRegion -eq "") {
        $serverRegion = "1"  # Set default value as NA
        break
    }
    elseif ($serverRegion -notin '0','1','2') {
        Write-Host "Invalid input. Please enter 0, 1, or 2, or press Enter for default." -ForegroundColor Red
    }
} until ($serverRegion -in '0','1','2')

# Prompt the user for the server's play style
do {
    $serverCommunity = Read-Host "What is the server's play style? (0=None, 1=Purist, 2=Relaxed, 3=Hard Core, 4=Role Playing, 5=Experimental)"

    if ($serverCommunity -notin '0','1','2','3','4','5') {
        Write-Host "Invalid input. Please enter a valid option (0, 1, 2, 3, 4, 5)." -ForegroundColor Red
    }
} until ($serverCommunity -in '0','1','2','3','4','5')

# Prompt the user to enable BattleEye
$enableBattleEye = Read-Host "Would you like to enable BattleEye? [y/n]"

# Check the user's input
if ($enableBattleEye -eq 'y' -or $enableBattleEye -eq 'yes') {
    $isBattlEyeEnabled = $true
}
elseif ($enableBattleEye -eq 'n' -or $enableBattleEye -eq 'no') {
    $isBattlEyeEnabled = $false
}
else {
    Write-Host "Invalid input. Please enter 'y' or 'n'." -ForegroundColor Red
}




#######################
# Create config files #
#######################

# Set the file path for the Engine.ini file
$engineFilePath = Join-Path -Path $installDirectory -ChildPath "ConanSandbox\Saved\Config\Windows\"

# Set the content for the Engine.ini file
$engineContent = @"
[OnlineSubsystemSteam]
ServerName=$serverName
ServerPassword=$serverPassword
AsyncTaskTimeout=360
GameServerQueryPort=$queryPort

[url]
Port=$gamePort
PeerPort=$peerPort

[/script/onlinesubsystemutils.ipnetdriver]
NetServerMaxTickRate=30
"@

# Write the content to the file
$engineContent | Out-File $engineFilePath\Engine.ini 

Write-Host "Engine Configuration saved to $engineFilePath" -ForegroundColor Cyan

# Set the file path for the Game.ini file
$gameFilePath = Join-Path -Path $installDirectory -ChildPath "ConanSandbox\Saved\Config\Windows\"

# Set the content for the Game.ini file
$gameContent = @"
[/script/engine.gamesession]
MaxPlayers=$MaxPlayers

[Rcon Plugin]
RconPort=$rconPort
"@

# Write the content to the file
$gameContent | Out-File $gameFilePath\Game.ini 

Write-Host "Game Configuration saved to $gameFilePath" -ForegroundColor Cyan

# Set the file path for the ServerSettings.ini file
$serverSettingsFilePath = Join-Path -Path $installDirectory -ChildPath "ConanSandbox\Saved\Config\Windows\"

# Set the content for the ServerSettings.ini file
$serverSettingsContent = @"
[ServerSettings]
AdminPassword=$serverPassword
MaxNudity=$maxNudity
PVPBlitzServer=$pvpBlitzServer
PVPEnabled=$pvpEnabled
serverRegion=$serverRegion
ServerCommunity=$serverCommunity
IsBattlEyeEnabled=$isBattlEyeEnabled
"@

# Write the content to the file
$serverSettingsContent | Out-File $serverSettingsFilePath\ServerSettings.ini 

Write-Host "Server Settings Configuration saved to $serverSettingsFilePath" -ForegroundColor Cyan

# Create a batch file that launches the server
$batchFilePath = "$installDirectory\ConanServer.bat"
$batchContent = "start ConanSandboxServer.exe -QueryPort=$queryPort"

# Write the content to the batch file
$batchContent | Out-File -FilePath $batchFilePath -Encoding ASCII -Append




############################
# Create rules, shortcuts, #
#    services and tasks    #
############################

# Create firewall rules if they don't exist
if (-not (Test-FirewallRule "Allow TCP and UDP for Game Port")) {
    New-NetFirewallRule -DisplayName "Allow TCP and UDP for Game Port" -Direction Inbound -Action Allow -Protocol TCP,UDP -LocalPort $gamePort
}

if (-not (Test-FirewallRule "Allow UDP for Peer Port")) {
    New-NetFirewallRule -DisplayName "Allow UDP for Peer Port" -Direction Inbound -Action Allow -Protocol UDP -LocalPort $peerPort
}

if (-not (Test-FirewallRule "Allow UDP for Query Port")) {
    New-NetFirewallRule -DisplayName "Allow UDP for Query Port" -Direction Inbound -Action Allow -Protocol UDP -LocalPort $queryPort
}

if (-not (Test-FirewallRule "Allow TCP for RCON Port")) {
    New-NetFirewallRule -DisplayName "Allow TCP for RCON Port" -Direction Inbound -Action Allow -Protocol TCP -LocalPort $rconPort
}

# Create a shortcut link to the Conan server application in the home directory. This will allow you to run the server at logon by typing '.\conanserver.lnk' 
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$Home\conanserver.lnk")
$Shortcut.TargetPath = "$batchFilePath"
$Shortcut.Save()

# Create the scheduled task
$defaultBackupDirectory = "C:\ConanBackups"
Create-ServerBackupScheduledTask -installDirectory $installDirectory -defaultBackupDirectory $defaultBackupDirectory

# Create a new service to auto-start the Conan server
# Define the executable path
$executablePath = Join-Path $installDirectory "$batchFilePath"

# Define the service name
$serviceName = "ConanServer"

# Use NSSM to create a new service
$arguments = @(
    "install", $serviceName, $executablePath,
    "-DelayedAutoStart", # Enable delayed startup
    "-DisplayName", "Conan Server", # Display name for the service
    "-Description", "Service for the Conan Server application" # Description for the service
)

try {
    Start-Process "nssm.exe" -ArgumentList $arguments -Wait -NoNewWindow
    Write-Host "Service 'Conan Server' created successfully." -ForegroundColor Green
} catch {
    $errorMessage = "An error occurred while creating the service. Error: $_"
    Log-Error $errorMessage
    exit 1
}

# Echo the completion of the script and provide the command to start the server app.
Write-Output "Conan dedicated server has successfully been installed. Use '.\conanserver.lnk' to start the game server app." -ForegroundColor Green

# Prompt the user to reboot the server
$userChoice = Read-Host -Prompt "A reboot is required for changes to take effect. The Conan server will not run until a reboot is completed. Do you want to reboot now? (Y/N)"

if ($userChoice -eq 'Y' -or $userChoice -eq 'Yes') {
    Write-Host "Rebooting the server..." -ForegroundColor Cyan
    # Invoke the restart command
    Restart-Computer -Force
} else {
    Write-Host "The Conan server will not run until a reboot is completed. Please reboot at your earliest convenience." -ForegroundColor Red
}

Start-Sleep 10

Clear-Host