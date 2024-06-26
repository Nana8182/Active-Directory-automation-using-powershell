# Active-Directory-automation-using-powershell

Overview

This PowerShell script automates the management of Active Directory (AD) users by syncing user data from a CSV file. It supports creating new users, updating existing users, and disabling users not in the CSV file.

Features

Import Users from CSV: Reads user data from a CSV file and imports it into AD.

Compare Users: Compares existing AD users with CSV data to identify new, updated, and removed users.

Create Users: Automatically creates new users in specified organizational units.

Update Users: Syncs existing users' data in AD with the CSV data.

Disable Users: Disables users not present in the current CSV file for a specified number of days.

Prerequisites

Windows PowerShell,
Active Directory module for Windows PowerShell,
Administrative privileges to modify AD,

Usage

Prepare the CSV File: Ensure your CSV file contains the required user data fields.

Modify the Script: Update the script parameters such as $CSVFilePath, $Delimiter, $Domain, $SyncFieldMap, and $OUProperty.

Run the Script: Execute the script in PowerShell with appropriate permissions.
