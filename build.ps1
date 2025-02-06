$VSVersion = 70
$PythonVersionMajor = 3
$PythonVersionMid = 12
$PythonVersionMinor = 8

Write-Output "#################################################"
Write-Output "### VapourSynth portable FATPACK build script ###"
Write-Output "#################################################"
Write-Output ""

Write-Output "Download files..."

$root = $PSScriptRoot
$vsFolder = Join-Path $root "VapourSynth64Portable\VapourSynth64"
$vsFolderWobbly = Join-Path $vsFolder "wobbly-win64"
$vsFolderFull = $vsFolder

# Create folder if it does not exist
if (-Not (Test-Path -Path $vsFolder)) {
    Write-Output "Create folder $vsFolder"
    New-Item -Path $vsFolder -ItemType Directory -Force | Out-Null
}

# Download URLs
$urls = @{
    "7zr"        = "https://github.com/ip7z/7zip/releases/download/24.08/7zr.exe"
    "python"     = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip"
    "vs"         = "https://github.com/vapoursynth/vapoursynth/releases/download/R70/VapourSynth64-Portable-R70.zip"
    "pip"        = "https://bootstrap.pypa.io/get-pip.py"
    "vseditor"   = "https://github.com/YomikoR/VapourSynth-Editor/releases/download/R19-mod-6.8/VapourSynth.Editor-r19-mod-6.8.zip"
    #"mveditor"   = "https://github.com/mysteryx93/VapourSynthViewer.NET/releases/download/v0.9.3/VapourSynthMultiViewer-v0.9.3.zip"
    "wobbly"     = "https://github.com/Jaded-Encoding-Thaumaturgy/Wobbly/releases/download/v8/Wobbly-win64.zip"
    "d2vwitch"   = "https://github.com/dubhater/D2VWitch/releases/download/v4/D2VWitch-v4-win64.7z"
    "d2vwitchv5" = "https://github.com/dubhater/D2VWitch/releases/download/v5/D2VWitch-v5-win64.7z"
    "vsrepogui"  = "https://github.com/theChaosCoder/VSRepoGUI/releases/download/v0.9.8/VSRepoGUI-0.9.8.zip"
    "pedeps"     = "https://github.com/brechtsanders/pedeps/releases/download/0.1.15/pedeps-0.1.15-win64.zip"
}

# Output of Downloaded file
$outputs = @{
    "7zr"        = Join-Path $root "7zr.exe"
    "python"     = Join-Path $root (Split-Path $urls["python"] -Leaf)
    "vs"         = Join-Path $root (Split-Path $urls["vs"] -Leaf)
    "vseditor"   = Join-Path $root (Split-Path $urls["vseditor"] -Leaf)
    "pip"        = Join-Path $vsFolder "get-pip.py"
    #"mveditor"   = Join-Path $root (Split-Path $urls["mveditor"] -Leaf)
    "wobbly"     = Join-Path $root (Split-Path $urls["wobbly"] -Leaf)
    "d2vwitch"   = Join-Path $root (Split-Path $urls["d2vwitch"] -Leaf)
    "d2vwitchv5" = Join-Path $root (Split-Path $urls["d2vwitchv5"] -Leaf)
    "vsrepogui"  = Join-Path $root (Split-Path $urls["vsrepogui"] -Leaf)
    "pedeps"     = Join-Path $root (Split-Path $urls["pedeps"] -Leaf)
}

# Download File function
function Download-File {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,

        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-Not (Test-Path -Path $Destination)) {
        Write-Output "Download of $Name from $Url"
        try {
            Invoke-WebRequest -Uri $Url -OutFile $Destination -ErrorAction Stop
            Write-Output "$Name downloaded with success"
        }
        catch {
            Write-Error "Error while downloading $Name from $Url : $_"
            exit 1
        }
    }
    else {
        Write-Output "$Name already downloaded, ignoring."
    }
}

# Extraction of ZIPs with Expand-Archive
function Extract-Zip {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    Write-Output "Extraction of $(Split-Path $ArchivePath -Leaf) in $DestinationPath..."
    try {
        Expand-Archive -Path $ArchivePath -DestinationPath $DestinationPath -Force
        Write-Output "Extraction done."
    }
    catch {
        Write-Error "Error while extracting $ArchivePath : $_"
        exit 1
    }
}

# Extraction of 7z archive with 7zr.exe
function Extract-7z {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    Write-Output "Extraction of $(Split-Path $ArchivePath -Leaf) in $DestinationPath with 7zr.exe..."
    try {
        & $outputs["7zr"] x $ArchivePath -o"$DestinationPath" -y
        Write-Output "Extraction ok."
    }
    catch {
        Write-Error "Error while extracting $ArchivePath with 7zr.exe : $_"
        exit 1
    }
}

# Download file
foreach ($key in $urls.Keys) {
    Download-File -Url $urls[$key] -Destination $outputs[$key] -Name $key
}

# Extraction of file
$zipArchives = @("vs", "vseditor", "python", "vsrepogui", "wobbly")
foreach ($archive in $zipArchives) {
	if ($archive -eq "wobbly") {
		Extract-Zip -ArchivePath $outputs[$archive] -DestinationPath $vsFolderWobbly
	} else {
		Extract-Zip -ArchivePath $outputs[$archive] -DestinationPath $vsFolder
	}
}

# Verification if python is extracted
$pythonExePath = Join-Path $vsFolder "python.exe"
if (-Not (Test-Path -Path $pythonExePath)) {
    Write-Error "python.exe not found after extraction."
    Write-Output "folder containing $vsFolder :"
    Get-ChildItem -Path $vsFolder
    exit 1
}
else {
    Write-Output "python.exe found."
}

$sevenZArchives = @("d2vwitch", "d2vwitchv5")
foreach ($archive in $sevenZArchives) {
							 
	Extract-7z -ArchivePath $outputs[$archive] -DestinationPath $vsFolder 
}

Write-Output "Extract pedeps (only bin folder contenant)..."
$tempPedepsFolder = Join-Path $root "temp_pedeps"

# Delete temp folder if exist already
if (Test-Path $tempPedepsFolder) {
    Remove-Item $tempPedepsFolder -Recurse -Force
}

# Create temp folder
New-Item -ItemType Directory -Path $tempPedepsFolder | Out-Null

try {
    # Full exact of archive in temp folder
    Expand-Archive -Path $outputs["pedeps"] -DestinationPath $tempPedepsFolder -Force

    # Path of bin folder in exacted archive
    $binFolder = Join-Path $tempPedepsFolder "bin"
    if (Test-Path $binFolder) {
        # Copy full of bin folder to $vsFolder
        Copy-Item -Path (Join-Path $binFolder "*") -Destination $vsFolder -Recurse -Force
        Write-Output "Extraction done."
    }
    else {
        Write-Error "The 'bin' folder doesn't exist in extract archive."
        exit 1
    }
}
catch {
    Write-Error "Error while extracting pedeps : $_"
    exit 1
}
finally {
    # Clean temp folder
    if (Test-Path $tempPedepsFolder) {
        Remove-Item $tempPedepsFolder -Recurse -Force
    }
}

$sevenZPath = Join-Path $vsFolder "7z.exe"
if (-Not (Test-Path -Path $sevenZPath)) {
    Write-Error "7z.exe not found after vapoursynth extraction."
    exit 1
}
else {
    Write-Output "7z.exe found."
}

# Path of python 3.12
$sourcePth = Join-Path $root "python312._pth"
$destinationPth = Join-Path $vsFolder "python312._pth"

if (Test-Path -Path $sourcePth) {
    Copy-Item -Path $sourcePth -Destination $destinationPth -Force
    Write-Output "python312._pth copied."
}
else {
    Write-Error "python312._pth not found in root directory."
    exit 1
}

Write-Output ""
Write-Output "Download / installation of Python packages wih pip..."
try {
    & $pythonExePath $outputs["pip"]
    & $pythonExePath -m pip install --upgrade pip
    & $pythonExePath -m pip install tqdm numpy
    Write-Output "Python packages installed."
}
catch {
    Write-Error "Error while installing Python packages : $_"
    exit 1
}

Write-Output "Installation of VapourSynth..."
$wheelPath = Join-Path $vsFolder "wheel\VapourSynth-$VSVersion-cp${PythonVersionMajor}${PythonVersionMid}-cp${PythonVersionMajor}${PythonVersionMid}-win_amd64.whl"
if (Test-Path -Path $wheelPath) {
    try {
        & $pythonExePath -m pip install $wheelPath
        Write-Output "VapourSynth installed."
    }
    catch {
        Write-Error "Error while installing VapourSynth : $_"
        exit 1
    }
}
else {
    Write-Error "Wheel file of VapourSynth not found : $wheelPath"
    exit 1
}

Write-Output ""

# Escaping regex patern for python
$pattern = [regex]::Escape("#!$vsFolder\python.exe")
$replacement = "#!python.exe"

Get-ChildItem -Path (Join-Path $vsFolder "Scripts") -Filter *.exe -Recurse | ForEach-Object {
    try {
        (Get-Content -Path $_.FullName -Raw) -replace $pattern, $replacement | Set-Content -Path $_.FullName -NoNewline
        Write-Output "Modified : $($_.FullName)"
    }
    catch {
        Write-Error "Error while modifying $($_.FullName) : $_"
    }
}

Write-Output ""
Write-Output "Optimize / Cleaning..."
$cleanupFiles = @(
    #"VapourSynthMultiViewer-x86.exe",
    #"VapourSynthMultiViewer-x86.exe.config",
    "get-pip.py"
)

foreach ($file in $cleanupFiles) {
    $filePath = Join-Path $root $file
    if (Test-Path -Path $filePath) {
        Remove-Item -Path $filePath -Force
        Write-Output "Deleted : $filePath"
    }
}

#Move the configuration file in the sub-folder for replace the default config.
Move-Item -Path "$root\VapourSynth64Portable\vsedit.config" -Destination "$vsFolder\vsedit.config" -Force

Write-Output ""
Write-Output "Done."
Write-Output "MANUAL TASK : COPY/PASTE x264.exe, x265.exe in the bin folder and all plugin in the plugin folder."
Write-Output ""

Pause
