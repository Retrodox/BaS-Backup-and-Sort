# Welcome message
Write-Host "Welcome to Backup and Sort which was made by Jared!"
Write-Host "Scanning Hard Drive.."

# Define the source directories
$sourceDirectories = @("$env:USERPROFILE\Desktop", "$env:USERPROFILE\Downloads", "$env:USERPROFILE\Documents", "$env:USERPROFILE\Pictures")

# Define the Word document file extensions
$wordDocExtensions = @('.doc', '.docx', '.dot', '.dotx', '.docm', '.dotm')

# Initialize an array to hold the count of Word documents in each directory
$wordDocCounts = @()

# Count Word documents in each directory
foreach ($sourceDirectory in $sourceDirectories) {
    $count = 0
    foreach ($extension in $wordDocExtensions) {
        $count += (Get-ChildItem -Path $sourceDirectory -Filter "*$extension" -Recurse -ErrorAction SilentlyContinue).Count
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
foreach ($sourceDirectory in $foldersToProcess) {
    foreach ($extension in $wordDocExtensions) {
        $files = Get-ChildItem -Path $sourceDirectory -Filter "*$extension" -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $destinationPath = Join-Path -Path $wordDocsDirectory -ChildPath $file.Name
            if ($file.FullName -ne $destinationPath) {
                Copy-Item -Path $file.FullName -Destination $destinationPath
            }
        }
    }
}

Write-Host "Backup completed."
pause
