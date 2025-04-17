<#
.SYNOPSIS
Automates the creation of audio sample archives, extracts metadata, and generates a markdown file documenting the samples.

.DESCRIPTION
The UpdateGit.ps1 script processes audio sample directories, creates .zip archives, extracts metadata from .json files, 
and generates a README.md file with a table documenting the samples. It is designed to streamline the management and 
documentation of audio samples.

.FEATURES
- Creates .zip files for audio samples organized by language, region, and voice name.
- Extracts metadata from .json files, including style, speed multiplier, trailing silence, and leading silence.
- Handles missing or incomplete metadata gracefully by defaulting to "Not Specified."
- Generates a markdown file with a table summarizing the samples and their metadata.
- Supports additional metadata from .csv files to document audio file details.
- Maps language and region codes to their full names using predefined dictionaries.
- Cleans up temporary directories after processing.

.PARAMETERS
None. The script uses predefined directories:
- `out`: Contains the audio sample directories to be processed.
- `samples`: Stores the generated .zip files and README.md file.
- `in`: Contains .csv files with additional metadata.

.INPUTS
- Audio sample directories in the `out` directory.
- .json files in the audio sample directories for metadata.
- .csv files in the `in` directory for additional metadata.

.OUTPUTS
- .zip files for each audio sample in the `samples` directory.
- A README.md file in the `samples` directory documenting the samples.

.NOTES
- The script ensures the `samples` directory exists and clears its contents before processing.
- Temporary directories are used for extracting .zip files and are cleaned up after processing.
- Language and region codes are mapped to their full names using predefined dictionaries.
- If metadata fields are missing in the `.json` file, the script defaults to "Not Specified."
- The script explicitly handles `0` values for `leadingSilence` to ensure they are not treated as missing.

.EXAMPLE
To run the script, simply execute it in PowerShell:
    .\UpdateGit.ps1

This will process the audio samples in the `out` directory, generate `.zip` files as sample archives, extract settings 
from `.json` files in `out` if present, and create the `README.md` file in the `samples` directory.

#>

$baseDir = "$PSScriptRoot\out"
$samplesDir = "$PSScriptRoot\samples"

# Ensure the samples directory exists and clear its contents if it does
if (Test-Path -Path $samplesDir) {
    Remove-Item -Path "$samplesDir\*" -Recurse -Force
} else {
    New-Item -ItemType Directory -Path $samplesDir
}

$subDirs = Get-ChildItem -Path $baseDir -Directory -Recurse | Where-Object { $_.FullName.Split('\').Count -eq ($baseDir.Split('\').Count + 3) }

Add-Type -AssemblyName System.IO.Compression.FileSystem

foreach ($dir in $subDirs) {
    $parentDir = Split-Path -Parent $dir.FullName
    $grandParentDir = Split-Path -Parent $parentDir
    $languageCode = Split-Path -Leaf $grandParentDir
    $regionCode = Split-Path -Leaf $parentDir
    $zipFileName = "$($languageCode)-$($regionCode)-$($dir.Name).zip"
    $zipFilePath = Join-Path -Path $baseDir -ChildPath $zipFileName
    Compress-Archive -Path "$($dir.FullName)\*" -DestinationPath $zipFilePath -Force
    Write-Host "Created zip file: $zipFilePath"

    # Move the zip file to the samples directory
    Move-Item -Path $zipFilePath -Destination $samplesDir
    Write-Host "Moved zip file to $samplesDir/$zipFileName"
}

$csvDir = "$PSScriptRoot\in"
$markdownDir = "$PSScriptRoot\samples"

# Ensure the markdown directory exists
if (-not (Test-Path -Path $markdownDir)) {
    New-Item -ItemType Directory -Path $markdownDir
}

# Get all CSV files from the input directory
$csvFiles = Get-ChildItem -Path $csvDir -Filter *.csv

foreach ($csvFile in $csvFiles) {
    $csvContent = Import-Csv -Path $csvFile.FullName
    $markdownFileName = "README.md"
    $markdownFilePath = Join-Path -Path $markdownDir -ChildPath $markdownFileName

    # Create the markdown content
    $markdownContent = @()
    $markdownContent += "# Samples"

    # Load language and region mappings from voices.json
    $voicesJsonPath = Join-Path -Path $PSScriptRoot -ChildPath "data/voices.json"
    if (-not (Test-Path -Path $voicesJsonPath)) {
        throw "voices.json file not found at $voicesJsonPath"
    }

    $voicesData = Get-Content -Path $voicesJsonPath -Raw | ConvertFrom-Json

    # Initialize dictionaries for language and region mappings
    $languageMap = @{}
    $regionMap = @{}

    # Extract unique language and region codes with their names
    foreach ($voice in $voicesData) {
        # Extract language and region codes from the Locale property
        $localeParts = $voice.Locale -split "-"
        $languageCode = $localeParts[0]
        $regionCode = if ($localeParts.Count -gt 1) { $localeParts[1] } else { $null }

        # Add language mapping if not already added
        if (-not $languageMap.ContainsKey($languageCode)) {
            $languageMap[$languageCode] = $voice.LocaleName -split "\(" | Select-Object -First 1
        }

        # Add region mapping if not already added and regionCode exists
        if ($regionCode -and -not $regionMap.ContainsKey($regionCode)) {
            $regionMap[$regionCode] = $voice.LocaleName -replace ".*\(|\)", ""
        }
    }

    # Get all zip files in the samples directory
    $zipFiles = Get-ChildItem -Path $samplesDir -Filter *.zip
    $markdownContent += "| File Link | Region | Voice | Style | Speed Multiplier | Trailing Silence | Leading Silence |"
    $markdownContent += "| --- | --- | --- | --- | --- | --- | --- |"

    foreach ($zipFile in $zipFiles) {
        $fileNameParts = $zipFile.BaseName -split "-"
        $languageCode = $fileNameParts[0]
        $regionCode = $fileNameParts[1]
        $voiceName = $fileNameParts[2]

        $languageFullName = $languageMap[$languageCode]
        $regionFullName = $regionMap[$regionCode]

        # Extract the zip file to a temporary directory
        $tempDir = Join-Path -Path $samplesDir -ChildPath "temp"
        if (-not (Test-Path -Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir | Out-Null
        }
        Expand-Archive -Path $zipFile.FullName -DestinationPath $tempDir -Force

        # Look for the .json file in the extracted contents
        $jsonFile = Get-ChildItem -Path $tempDir -Filter *.json -Recurse | Select-Object -First 1
        $style = "Not Specified"
        $speedMultiplier = "Not Specified"
        $trailingSilence = "Not Specified"
        $leadingSilence = "Not Specified"

        if ($jsonFile) {
            $jsonContent = Get-Content -Path $jsonFile.FullName -Raw | ConvertFrom-Json
            $style = if ($null -ne $jsonContent.style) { $jsonContent.style } else { $style }
            $speedMultiplier = if ($null -ne $jsonContent.multiplier) { $jsonContent.multiplier } else { $speedMultiplier }
            $trailingSilence = if ($null -ne $jsonContent.trailingSilence) { $jsonContent.trailingSilence } else { $trailingSilence }
            $leadingSilence = if ($jsonContent.PSObject.Properties.Name -contains 'leadingSilence') { $jsonContent.leadingSilence } else { $leadingSilence }
        }

        # Clean up the temporary directory
        Remove-Item -Path $tempDir -Recurse -Force

        if ($languageFullName -and $regionFullName) {
            $markdownContent += "| [**$($zipFile.Name)**]($($zipFile.Name)) | $languageFullName ($regionFullName) | $voiceName | $style | $speedMultiplier | $trailingSilence | $leadingSilence |"
        } else {
            $markdownContent += "| [**$($zipFile.Name)**]($($zipFile.Name)) | Language or region code not recognized | $voiceName | $style | $speedMultiplier | $trailingSilence | $leadingSilence |"
        }
    }

    $markdownContent += ""
    $markdownContent += "Due to being uploaded prior to this README including the information, some of the above sample items may say 'not specified'.<br>"
    $markdownContent += "For those, the following typically values were used:<br>"
    $markdownContent += ""
    $markdownContent += "Voice Style: Default or Narration (Narration preferred but not always available)<br>"
    $markdownContent += "Speed Multiplier: 1.10-1.20 (Varies per voice)<br>"
    $markdownContent += "Trailing Silence: 25ms<br>"
    $markdownContent += "Leading Silence: 0ms"
    $markdownContent += ""
    $markdownContent += "The above table will be updated with the information when the samples are updated in the future.<br>"
    $markdownContent += ""

    $markdownContent += "## Audio File List"
    $markdownContent += "The following table contains the list of audio files within the samples. The table is generated from the input CSV file."
    $header = "| " + ($csvContent[0].PSObject.Properties.Name -join " | ") + " |"
    $separator = "| " + (($csvContent[0].PSObject.Properties.Name | ForEach-Object { "---" }) -join " | ") + " |"
    $markdownContent += $header
    $markdownContent += $separator
    foreach ($row in $csvContent) {
        $markdownContent += "| " + ($row.PSObject.Properties.Value -join " | ") + " |"
    }

    # Write the markdown content to the file
    $markdownContent | Out-File -FilePath $markdownFilePath -Encoding utf8
    Write-Host "Created markdown file: $markdownFilePath"
}
