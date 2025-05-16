Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

Write-Host "Starting Advanced Audio Suite..."

#### Check Credentials ####
Write-Host "Checking Azure credentials..." -NoNewline

$configDirectoryPath = "$PSScriptRoot/config"
$credentialsFilePath = "$configDirectoryPath/credentials.json"

# Ensure the config directory exists
if (!(Test-Path $configDirectoryPath)) {
    New-Item -ItemType Directory -Path $configDirectoryPath -Force | Out-Null
}

if (!(Test-Path $credentialsFilePath)) {
    Write-Host "`nCredentials file not found.  Creating..."
    $defaultCredentials = @{
        Key = "yourkey"
        Region = "yourregion"
    }
    $defaultCredentials | ConvertTo-Json | Set-Content -Path $credentialsFilePath

    $key = Read-Host "Please enter your Azure Speech resource group key"
    $region = Read-Host "Please enter your Azure Speech resource group region"
    $credentialsData = @{
        Key = $key
        Region = $region
    }
    $credentialsData | ConvertTo-Json | Set-Content -Path $credentialsFilePath
} else {
    $credentialsData = Get-Content -Raw -Path $credentialsFilePath | ConvertFrom-Json
    $key = $credentialsData.Key
    $region = $credentialsData.Region
    Write-Host " OK"
}

if (!(Test-Path (Join-Path $PSScriptRoot "key")) -or !(Test-Path (Join-Path $PSScriptRoot "region"))) {
    Write-Host "`nspx key and/or region files not found. Setting them based on info in 'config/credentials.json'"
    $commandKey = "spx --% config @key --set $key"
    $commandRegion = "spx --% config @region --set $region"
    Invoke-Expression $commandKey > $null
    Invoke-Expression $commandRegion > $null
    Write-Host "Azure Key and Region set."
}

$voicesJsonPath = Join-Path $PSScriptRoot "data/voices.json"

# Check if voices.json exists
if (-not (Test-Path $voicesJsonPath)) {
    Write-Host "Getting voices from spx..." -NoNewline
    $output = spx synthesize --voices
    $path = "data/voices.json"
    # Ensure the data directory exists
    $dataDirectoryPath = Join-Path $PSScriptRoot "data"
    if (-not (Test-Path $dataDirectoryPath)) {
        New-Item -ItemType Directory -Path $dataDirectoryPath -Force | Out-Null
    }

    # Combine the output into a single string
    $outputText = $output -join "`n"

    # Extract the JSON portion of the output, ignoring the preceding text (spx response)
    $jsonStartIndex = $outputText.IndexOf('[')
    if ($jsonStartIndex -ge 0) {
        $jsonOutput = $outputText.Substring($jsonStartIndex)
    }

    # Delete the existing file if it exists
    if (Test-Path -Path $path) {
        Remove-Item -Path $path
    }

    # Write the JSON output to voices.json
    $jsonOutput | Out-File -FilePath $path -Encoding utf8

    if (Test-Path -Path $path) {
        Write-Host " OK"
    } else {
        Write-Host " Failed"
        return
    }
} else {
    Write-Host "Reading voices.json..." -NoNewline
    if (Test-Path -Path $voicesJsonPath) {
    Write-Host " OK"
    } else {
        Write-Host " Failed"
        return
    }
}

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Advanced Audio Suite" Height="600" Width="700">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Title -->
            <RowDefinition Height="Auto"/> <!-- Language -->
            <RowDefinition Height="Auto"/> <!-- Region -->
            <RowDefinition Height="Auto"/> <!-- Voice -->
            <RowDefinition Height="Auto"/> <!-- Voice Style -->
            <RowDefinition Height="Auto"/> <!-- Speed Multiplier -->
            <RowDefinition Height="Auto"/> <!-- Trailing Silence -->
            <RowDefinition Height="Auto"/> <!-- Leading Silence -->
            <RowDefinition Height="Auto"/> <!-- Buttons -->
            <RowDefinition Height="*"/>   <!-- Output TextBox -->
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/> <!-- Labels -->
            <ColumnDefinition Width="*"/>   <!-- Input Fields -->
        </Grid.ColumnDefinitions>

        <!-- Title -->
        <Label Content="Advanced Audio Suite" FontSize="20" HorizontalAlignment="Center" Grid.ColumnSpan="2" Margin="10" Grid.Row="0"/>

        <!-- Language -->
        <Label Content="Language:" FontSize="14" Grid.Row="1" Grid.Column="0" VerticalAlignment="Center"/>
        <ComboBox Name="CmbLanguage" Width="200" Height="25" Grid.Row="1" Grid.Column="1" HorizontalAlignment="Right" Margin="0,5,10,5"/>

        <!-- Region -->
        <Label Content="Region:" FontSize="14" Grid.Row="2" Grid.Column="0" VerticalAlignment="Center"/>
        <ComboBox Name="CmbRegion" Width="200" Height="25" Grid.Row="2" Grid.Column="1" HorizontalAlignment="Right" Margin="0,5,10,5"/>

        <!-- Voice -->
        <Label Content="Voice:" FontSize="14" Grid.Row="3" Grid.Column="0" VerticalAlignment="Center"/>
        <ComboBox Name="CmbVoice" Width="200" Height="25" Grid.Row="3" Grid.Column="1" HorizontalAlignment="Right" Margin="0,5,10,5"/>

        <!-- Voice Style -->
        <Label Content="Voice Style:" FontSize="14" Grid.Row="4" Grid.Column="0" VerticalAlignment="Center"/>
        <ComboBox Name="CmbStyle" Width="200" Height="25" Grid.Row="4" Grid.Column="1" HorizontalAlignment="Right" Margin="0,5,10,5"/>

        <!-- Speed Multiplier -->
        <Label Content="Speed Multiplier:" FontSize="14" Grid.Row="5" Grid.Column="0" VerticalAlignment="Center"/>
        <TextBox Name="TxtSpeed" Width="200" Height="25" Grid.Row="5" Grid.Column="1" HorizontalAlignment="Right" Margin="0,5,10,5"/>

        <!-- Trailing Silence -->
        <Label Content="Trailing Silence (ms):" FontSize="14" Grid.Row="6" Grid.Column="0" VerticalAlignment="Center"/>
        <TextBox Name="TxtPostSilence" Width="200" Height="25" Grid.Row="6" Grid.Column="1" HorizontalAlignment="Right" Margin="0,5,10,5"/>

        <!-- Leading Silence -->
        <Label Content="Leading Silence (ms):" FontSize="14" Grid.Row="7" Grid.Column="0" VerticalAlignment="Center"/>
        <TextBox Name="TxtPreSilence" Width="200" Height="25" Grid.Row="7" Grid.Column="1" HorizontalAlignment="Right" Margin="0,5,10,5"/>

        <!-- Buttons -->
        <StackPanel Orientation="Horizontal" Grid.Row="8" Grid.Column="1" HorizontalAlignment="Right" Margin="0,10,10,0">
            <Button Name="BtnPlaySample" Content="Play Sample" FontSize="14" Width="100" Height="25" Margin="0,0,10,0"/>
            <Button Name="BtnStartSynthesis" Content="Start Synthesis" FontSize="14" Width="120" Height="25"/>
        </StackPanel>

        <!-- Output TextBox -->
        <TextBox Name="TxtOutput" Grid.Row="9" Grid.ColumnSpan="2" Margin="10" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" AcceptsReturn="True" IsReadOnly="True"/>
    </Grid>
</Window>
"@

# Load the XAML
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get the ComboBoxes, TextBoxes, and buttons by their names
$cmbLanguage = $window.FindName("CmbLanguage")
$cmbRegion = $window.FindName("CmbRegion")
$cmbVoice = $window.FindName("CmbVoice")
$cmbStyle = $window.FindName("CmbStyle")
$txtSpeed = $window.FindName("TxtSpeed")
$txtPostSilence = $window.FindName("TxtPostSilence")
$txtPreSilence = $window.FindName("TxtPreSilence")
$btnPlaySample = $window.FindName("BtnPlaySample")
$btnStartSynthesis = $window.FindName("BtnStartSynthesis")
$txtOutput = $window.FindName("TxtOutput")

Write-Host "Starting GUI..."

# Function to write info to the GUI textbox
function Write-OutputToTextBox {
    param (
        [string]$message
    )
    $window.Dispatcher.Invoke([action]{
        $TxtOutput.AppendText("$message`n")
        $TxtOutput.ScrollToEnd()
    })
}

# Populate ComboBoxes with options
$matchingStyles = @("Default", "Chat", "Narration") # Example data
$cmbStyle.ItemsSource = $matchingStyles

# Load voices
$voicesJsonPath = Join-Path $PSScriptRoot "data/voices.json"
if (-not (Test-Path $voicesJsonPath)) {
    Write-Host "Getting voices from spx..."
    $output = spx synthesize --voices
    $outputText = $output -join "`n"
    $jsonStartIndex = $outputText.IndexOf('[')
    if ($jsonStartIndex -ge 0) {
        $jsonOutput = $outputText.Substring($jsonStartIndex)
    }
    $jsonOutput | Out-File -FilePath $voicesJsonPath -Encoding utf8
}

$voicesJson = Get-Content -Raw -Path $voicesJsonPath | ConvertFrom-Json

# Extract available languages and their full names
$languageMap = @{}
$voicesJson | ForEach-Object {
    $locale = $_.Locale.Split('-')[0]
    if (-not $languageMap.ContainsKey($locale)) {
        $languageMap[$locale] = $_.LocaleName.Split('(')[0].Trim()
    }
}

# Extract available languages from the language map
$availableLanguages = $languageMap.Values | Sort-Object -Unique

# Populate ComboBoxes with options
$cmbLanguage.ItemsSource = $availableLanguages

# Function to filter voices based on selected language and region
function Update-Voices {
    param (
        [string]$language,
        [string]$region
    )
    $languageCode = $languageMap.Keys | Where-Object { $languageMap[$_] -eq $language }
    $filteredVoices = $voicesJson | Where-Object { $_.Locale -like "$languageCode-$region*" } | Select-Object -ExpandProperty ShortName
    $filteredVoicesDisplay = $filteredVoices | ForEach-Object { $_.Split('-')[-1] }
    $cmbVoice.ItemsSource = $filteredVoicesDisplay
    if ($filteredVoicesDisplay.Count -gt 0) {
        $cmbVoice.SelectedItem = $filteredVoicesDisplay[0]
    }
    if ($filteredVoices.Count -gt 0) {
    Write-OutputToTextBox "Voices updated for $language - $region"
    }
}

# Function to filter regions based on selected language
function Update-Regions {
    param (
        [string]$language
    )
    $languageCode = $languageMap.Keys | Where-Object { $languageMap[$_] -eq $language }
    $filteredRegions = $voicesJson | Where-Object { $_.Locale -like "$languageCode-*" } | ForEach-Object {
        $_.Locale.Split('-')[1]
    } | Sort-Object -Unique
    $cmbRegion.ItemsSource = $filteredRegions
    if ($filteredRegions.Count -gt 0) {
        $cmbRegion.SelectedItem = $filteredRegions[0]
    }
}

# Update regions and voices when language changes
$cmbLanguage.Add_SelectionChanged({
    $Language = $cmbLanguage.SelectedItem
    Update-Regions -language $Language
    Update-Voices -language $Language -region $cmbRegion.SelectedItem
})

# Update voices when region changes
$cmbRegion.Add_SelectionChanged({
    $Language = $cmbLanguage.SelectedItem
    $Region = $cmbRegion.SelectedItem
    Update-Voices -language $Language -region $Region
})

# Set default values from config.json if it exists
$configFilePath = Join-Path $PSScriptRoot "config/config.json"
if (Test-Path $configFilePath) {
    $configData = Get-Content -Raw -Path $configFilePath | ConvertFrom-Json

    # Ensure the selected language exists in the language map
    $languageCode = $languageMap.Keys | Where-Object { $languageMap[$_] -eq $configData.Language }
    if ($languageCode) {
        $Language = $configData.Language
        $cmbLanguage.SelectedItem = $Language
    } else {
        Write-Host "Warning: Selected language from config not found in available languages. Using default."
        $Language = "English"
        $cmbLanguage.SelectedItem = $Language
    }

    $Region = $configData.Region
    $cmbRegion.SelectedItem = $Region

    $Voice = $configData.Voice.Split('-')[-1]
    $cmbVoice.SelectedItem = $Voice

    $Style = if ($null -eq $configData.Style) { "Default" } else { $configData.Style }
    $cmbStyle.SelectedItem = $Style

    $Speed = [double]$configData.Speed
    $txtSpeed.Text = [string]$Speed

    $PostSilence = [int]$configData.PostSilence
    $txtPostSilence.Text = [string]$PostSilence

    $PreSilence = [int]$configData.PreSilence
    $txtPreSilence.Text = [string]$PreSilence
} else {
    $Language = "English"
    $cmbLanguage.SelectedItem = $Language

    $Region = "AU"
    $cmbRegion.SelectedItem = $Region

    $Voice = "ElsieNeural"
    $cmbVoice.SelectedItem = $Voice

    $Style = "Default"
    $cmbStyle.SelectedItem = $Style

    $Speed = 1.25
    $txtSpeed.Text = [string]$Speed

    $PostSilence = 25
    $txtPostSilence.Text = [string]$PostSilence

    $PreSilence = 0
    $txtPreSilence.Text = [string]$PreSilence
}

# Add event handler for the start synthesis button
$btnStartSynthesis.Add_Click({

    # Read the current values from the GUI
    $Language = $cmbLanguage.SelectedItem
    $Region = $cmbRegion.SelectedItem
    $Voice = $cmbVoice.SelectedItem
    $Style = $cmbStyle.SelectedItem
    $Speed = [double]$txtSpeed.Text
    $PostSilence = [int]$txtPostSilence.Text
    $PreSilence = [int]$txtPreSilence.Text

    # Save the current configuration to config.json
    Save-Config

    # Map the selected language back to its code
    $LanguageCode = $languageMap.Keys | Where-Object { $languageMap[$_] -eq $Language }
    $ShortName = "$LanguageCode-$Region-$Voice"
    
    Write-OutputToTextBox "Starting synthesis..."

    # Perform synthesis operation here
    Start-Synthesis -Language $LanguageCode -Region $Region -ShortName $ShortName -Style $Style -Speed $Speed -PostSilence $PostSilence -PreSilence $PreSilence
    
    # Print Complete message
    $LanguageCode = $languageMap.Keys | Where-Object { $languageMap[$_] -eq $Language }
    Write-OutputToTextBox "Speech synthesis complete. Synthesized audio can be found in 'out/$LanguageCode/$Region'."
    
    # Clear logs after running
    Clear-Logs
})

# Define a global variable to keep track of the current sample text index
if (-not $global:currentSampleIndex) {
    $global:currentSampleIndex = 0
}

# Add event handler for the play sample button
$btnPlaySample.Add_Click({
    # Read the current config from the GUI
    $Language = $cmbLanguage.SelectedItem
    $Region = $cmbRegion.SelectedItem
    $Voice = $cmbVoice.SelectedItem
    $ShortName = "$LanguageCode-$Region-$Voice"
    $Style = $cmbStyle.SelectedItem
    $Speed = [double]$txtSpeed.Text
    $PostSilence = [int]$txtPostSilence.Text
    $PreSilence = [int]$txtPreSilence.Text
    $sampleTextToPlays = @("25 Percent", "Welcome to Ethos", "Flaps Up", "Throttle Hold", "RSSI Low Warning", 
                        "RSSI Critical Warning", "VFR Low Warning", "Landing Gear Up", "Landing Gear Down", "Idle Up 1", "Ignition Off")

    # Save the current configuration to config.json
    Save-Config

    # Get the current text to generate
    $textToGenerate = $sampleTextToPlays[$global:currentSampleIndex]
    Write-OutputToTextBox "Playing Sample..."

    spx synthesize --ssml   "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='$language-$region'>
                                    <voice name='$ShortName'>
                                        <mstts:express-as style='$Style' styledegree='2'>
                                            <lang xml:lang='$language-$region'>
                                                <prosody rate='$Speed'>
                                                    <mstts:silence type='Leading-exact' value='$PreSilence'/>
                                                        $textToGenerate
                                                    <mstts:silence type='Tailing-exact' value='$PostSilence'/>
                                                </prosody>
                                            </lang>
                                        </mstts:express-as>
                                    </voice>
                                </speak>"

    Clear-Logs

    # Increment the counter and wrap around if it exceeds the numebr of samples
    $global:currentSampleIndex += 1
    if ($global:currentSampleIndex -ge $sampleTextToPlays.Count) {
        $global:currentSampleIndex = 0
    }
})

# Define the Start-Synthesis function
function Start-Synthesis {
    param (
        [string]$Language,
        [string]$Region,
        [string]$ShortName,
        [string]$Style,
        [double]$Speed,
        [int]$PostSilence,
        [int]$PreSilence
    )

    $baseFilePath = "out/$Language/$Region/$Voice/"

    # get full CSV path and name
    $csvFilePath = Join-Path $PSScriptRoot "in\*.csv"

    # Read the CSV file
    Get-Content -Path $csvFilePath | Out-Null

    $csvData = Import-Csv -Path $csvFilePath 
    # Ensure the 'old' directory exists
    $oldDirPath = Join-Path $PSScriptRoot "data"
    if (-not (Test-Path $oldDirPath)) {
        New-Item -ItemType Directory -Path $oldDirPath | Out-Null
    }

    # Export csvData to lastCsvData.csv only if lastCsvData doesn't already exist
    if (-not (Test-Path "$oldDirPath/lastCsvData.csv")) {
        $csvData | Export-Csv -Path "$oldDirPath/lastCsvData.csv" -NoTypeInformation
    }

    #### Check csvData.csv and compare to $csvData to determine if any changes have been made ####
    Write-OutputToTextBox "Checking for changes in .csv file from last run..."
    Start-Sleep -Seconds 0.5

    $lastCsvData = Import-Csv -Path "data/lastCsvData.csv"
    # Check if the 'text to play' row for a given 'path' row has changed from the last run
    $changedRows = @()
    foreach ($row in $csvData) {
        $rowPath = $row.Path
        $textToPlay = $row.'text to play'
        $lastRowPath = $lastCsvData | Where-Object { $_.Path -eq $rowPath }
        if ($lastRowPath -and $lastRowPath.'text to play' -ne $textToPlay) {
            $changedRows += $row
            Write-OutputToTextBox "$rowPath text to play has changed.  Removing old audio file."
            if (Test-Path "$baseFilePath/$rowPath") {
                Remove-Item -Path "$baseFilePath/$rowPath" -Force
            }
        }
    }   

    # Check if the 'text to play' row for a given 'path' row is changed
    $newRows = @()
    if ($lastCsvData) {
        $lastCsvData = $lastCsvData | Select-Object -ExpandProperty Path
    } else {
        $lastCsvData = @()
    }

    # Check if any new rows have been added to the .csv file since last run
    foreach ($row in $csvData) {
        $rowPath = $row.Path
        $textToPlay = $row.'text to play'
        if (-not $lastCsvData.Contains($rowPath)) {
            $newRows += $row
            Write-OutputToTextBox "$rowPath is a new row."
        }
    }

    # Figure out what to do with $csvData (what files to generate)
    if (($changedRows) -and ($newRows)) {
        $csvData = $changedRows + $newRows
        Write-OutputToTextBox "Changes detected in .csv file.  Synthesizing new and changed rows."
    }
    elseif ($changedRows.Count -gt 0) {
        $csvData = $changedRows
        Write-OutputToTextBox "Changes detected in .csv file.  Synthesizing changed rows."
    }
    elseif ($newRows.Count -gt 0) {
        $csvData = $newRows
        Write-OutputToTextBox "Changes detected in .csv file.  Synthesizing new rows."
    }
    else {
        Write-OutputToTextBox "No changes detected in .csv file. Only deleted or never-created files will be synthesized."
        $csvData = Import-Csv -Path $csvFilePath 
    }

    
    # Set synthesisFailed to false for first run
    [boolean] $synthesisFailed = $false

    # Main loop to iterate through each synthesis cycle
    for ($i = 0; $i -lt $csvData.Count; $i++) {
        # Set filepath to base file path for each iteration
        $filePath = $baseFilePath
        $filePath += $csvData[$i].PSObject.Properties.Value[0]

        # Replace invalid characters in file path
        $filePath = $filePath -replace ':', '_'

        # If audio file already exists for this iteration, skip it.
        if (Test-Path $filePath) {
            Write-OutputToTextBox "File $filePath already exists. Skipping synthesis."
            continue
        }

        Write-OutputToTextBox "Synthesizing $filePath..." -NoNewline

        # Get the subfolder path from the file path
        $folderPath = Split-Path -Path $filePath -Parent

        # Check if the subfolder path exists, if not, create it
        if (-not (Test-Path $folderPath)) {
            New-Item -ItemType Directory -Path $folderPath | Out-Null
        }

        $textToGenerate = $csvData[$i].PSObject.Properties.Value[1]

        # Check if the generated file size is greater than 1KB and retry if it is not (up to 3 retries)
        $retryCount = 0
        while (-not ($synthesisFailed)) {
            if ($retryCount -eq 3) {
                Write-OutputToTextBox "Audio synthesis failed. Retry limit reached. Re-run script to try again."
                Remove-Item -Path $filePath -Force
                break
            }
            spx synthesize --ssml   "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='$language-$region'>
                                    <voice name='$ShortName'>
                                        <mstts:express-as style='$Style' styledegree='2'>
                                            <lang xml:lang='$language-$region'>
                                                <prosody rate='$Speed'>
                                                    <mstts:silence type='Leading-exact' value='$PreSilence'/>
                                                        $textToGenerate
                                                    <mstts:silence type='Tailing-exact' value='$PostSilence'/>
                                                </prosody>
                                            </lang>
                                        </mstts:express-as>
                                    </voice>
                                </speak>" --audio output $filePath > $null 

            # Check if the file size is greater than 1KB (to detect if there was an error during synthesis)
            if ((Get-Item $filePath).Length -lt 1024) {
                $retryCount++
                Write-OutputToTextBox "File generated is invalid. There was an error during synthesis.  Retrying..."
                Write-OutputToTextBox "Retry Count: $retryCount" 
                Start-Sleep -Seconds 0.25
            }
            else {
                $synthesisFailed = $false
                $retryCount = 0
                break
            }
        }
        # Reset $synthesisFailed for the next iteration
        $synthesisFailed = $false 
        Write-OutputToTextBox " Done"
    }
    
    # Export the csvData.csv to lastCsvData.csv for examination on next run
    $lastCsvDataPath = "data/lastCsvData.csv"
    Copy-Item -Path $csvFilePath -Destination $lastCsvDataPath -Force

    # Create the settings.json file
    $settings = @{
        style = $Style
        multiplier = $Speed
        trailingSilence = $PostSilence
        leadingSilence = $PreSilence
    }

    $settingsFilePath = Join-Path -Path $baseFilePath -ChildPath "settings.json"
    $settings | ConvertTo-Json -Depth 10 | Set-Content -Path $settingsFilePath -Encoding utf8

    Write-OutputToTextBox "Settings saved to $settingsFilePath"
}

# Function to save the current configuration to config.json
function Save-Config {
    $configData = [ordered]@{
        Language = $cmbLanguage.SelectedItem
        Region = $cmbRegion.SelectedItem
        Voice = $cmbVoice.SelectedItem
        Speed = [double]$txtSpeed.Text
        PostSilence = [int]$txtPostSilence.Text
        PreSilence = [int]$txtPreSilence.Text
    }
    $configFilePath = Join-Path $PSScriptRoot "config/config.json"
    $configData | ConvertTo-Json | Set-Content -Path $configFilePath
}

# Function to clean up log files
function Clear-Logs {
    Write-OutputToTextBox "Cleaning up..."
    $logFiles = Get-ChildItem -Path $PSScriptRoot -Filter "log-*" -File
    foreach ($file in $logFiles) {
        Remove-Item -Path $file.FullName -Force
    }
}

# Show the window
$window.ShowDialog() | Out-Null