# BaS-Backup-and-Sort
Simple Explaination: A powershell script that copies all word documents from your desktop, documents, downloads, and pictures folders, (including all subdirectories), to a condensed folder on your desktop. I will eventually expand this to encompass multiple file types in the future.


# Backup and Sort PowerShell Script

## Description
This PowerShell script, is designed to backup and sort Word documents from the user's Desktop, Downloads, Documents, and Pictures folders. The script creates a backup folder on the desktop with the current date and a subfolder named "Word Docs" where all the Word documents are copied.

## Features
- Scans the user's Desktop, Downloads, Documents, and Pictures folders for Word documents.
- Counts and displays the number of Word documents in each folder.
- Allows the user to skip specific folders from the backup process.
- Creates a backup folder on the desktop with the format "Backup_MM_dd_yyyy".
- Copies all Word documents, including those in subdirectories, into the "Word Docs" subfolder in the backup folder.

## Supported File Extensions
The script supports the following Word document file extensions:
- .doc
- .docx
- .dot
- .dotx
- .docm
- .dotm

## Usage
1. change this line "@echo off PowerShell -ExecutionPolicy Bypass -File "C:\Path\To\BaS\BackupAndSort.ps1"" to your actual path
2. click launch.bat or just run with powershell

## Requirements
- PowerShell 1.0

## Author
Jared Mingle

## Disclaimer
This script does not delete or modify any existing files. However, it is always recommended to test the script in a safe environment before using it on important data.
