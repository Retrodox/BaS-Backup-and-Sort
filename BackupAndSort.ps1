# Welcome message
Write-Host "Welcome to Backup and Sort which was made by Jared!"
Write-Host "Scanning Hard Drive.."

# Define the source directories
$sourceDirectories = @("$env:USERPROFILE\Desktop", "$env:USERPROFILE\Downloads", "$env:USERPROFILE\Documents", "$env:USERPROFILE\Pictures")

# Define the Word document file extensions
$wordDocExtensions = @('.doc', '.docx', '.dot', '.dotx', '.docm', '.dotm')

# Initialize a list to hold unique file paths and a hashtable to track files by name
$filePathsByName = @{}

# Count Word documents in each directory
foreach ($sourceDirectory in $sourceDirectories) {
    $filesInDirectory = Get-ChildItem -Path $sourceDirectory -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in $wordDocExtensions }
    $count = $filesInDirectory.Count
    foreach ($file in $filesInDirectory) {
        $fileNameWithExtension = $file.Name
        if (-not $filePathsByName.ContainsKey($fileNameWithExtension)) {
            $filePathsByName[$fileNameWithExtension] = @($file.FullName)
        } else {
            $existingPaths = $filePathsByName[$fileNameWithExtension]
            if (-not $existingPaths.Contains($file.FullName)) {
                $filePathsByName[$fileNameWithExtension] += $file.FullName
            }
        }
    }
    $wordDocCounts += $count
    Write-Host "$($sourceDirectories.IndexOf($sourceDirectory) + 1). $(Split-Path $sourceDirectory -Leaf) - $count Word documents"
}


# Prompt the user to skip any folders
Write-Host "Would you like to skip over any folders that are listed above? Please type in the number associated above to each respective folder that you would like to skip over followed by enter. After finishing please type in 'r'"

$skipFolders = @()
while ($true) {
    $input = Read-Host "Enter the number of the folder to skip or 'r' to run the backup"
    if ($input -eq 'r') {
        break
    } elseif ($input -match '^\d+$' -and [int]$input -gt 0 -and [int]$input -le $sourceDirectories.Count) {
        $skipFolders += [int]$input - 1
    } else {
        Write-Host "Invalid input. Please enter a valid folder number or 'r' to run the backup."
    }
}

# Filter out the folders to be skipped
$foldersToProcess = $sourceDirectories | Where-Object { $skipFolders -notcontains $sourceDirectories.IndexOf($_) }

# Prompt user for selection if duplicates are found
foreach ($fileName in $keysCopy) {
    if ($filePathsByName[$fileName].Count -gt 1) {
        Write-Host "Multiple instances of the same file detected!"
        $i = 1
        foreach ($filePath in $filePathsByName[$fileName]) {
            if (Test-Path $filePath) {
                $file = Get-Item -Path $filePath
                Write-Host "$i. $filePath - Last modified $($file.LastWriteTime.ToString('MM/dd/yyyy - hh:mm tt'))"
                $i++
            }
        }
        $selection = Read-Host "Please Select 1 out of the following files to save by entering the number associated"
        if ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $i - 1) {
            $selectedFilePath = $filePathsByName[$fileName][$selection - 1]
            # Update filePathsByName to only include the selected file
            $filePathsByName[$fileName] = @($selectedFilePath)
        } else {
            Write-Host "Invalid selection. Please enter a number between 1 and $($i - 1)."
        }
    }
}

# Flatten the list of unique file paths
$uniqueFilePaths = $filePathsByName.Values | ForEach-Object { $_ } | Select-Object -Unique


# Define the backup directory on the desktop with today's date in month_day_year format
$backupDirectory = "$env:USERPROFILE\Desktop\Backup_$(Get-Date -Format 'MM_dd_yyyy')"

# Define the Word Docs directory
$wordDocsDirectory = "$backupDirectory\Word Docs"

# Create the backup and Word Docs directories if they don't exist
if (-not (Test-Path $backupDirectory)) {
    New-Item -Path $backupDirectory -ItemType Directory
}
if (-not (Test-Path $wordDocsDirectory)) {
    New-Item -Path $wordDocsDirectory -ItemType Directory
}

# Copy Word documents to the Word Docs directory
foreach ($filePath in $uniqueFilePaths) {
    $file = Get-Item $filePath
    $destinationPath = Join-Path -Path $wordDocsDirectory -ChildPath $file.Name
    if ($file.FullName -ne $destinationPath) {
        Copy-Item -Path $file.FullName -Destination $destinationPath
    }
}

Write-Host "Backup completed."
pause