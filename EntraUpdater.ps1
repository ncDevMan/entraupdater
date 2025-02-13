# Get the inputs from the commandline
param (
    [string]$InputFile,
    [string]$OutputFile
)

# Import necessary module
Import-Module ImportExcel

# Column Mappings: We will use the keys as the InputFile column names and the values as the OutputFile column names
$columnMappings = @{
    "Work Email" = "UserPrincipalName"
    "Manager" = "ManagerUserPrincipalName"
    "Job Title" = "JobTitle"
    "Employee Code" = "EmployeeID"
    "Contract" = "Department"
    "Hire Date" = "ExtensionAttribute1"
    "Birth Date" = "ExtensionAttribute2"
    "Supervisor" = "ExtensionAttribute3"
    "BUD" = "ExtensionAttribute4"
}

# These InputFile columns are skipped when the OutputFile is created.
$skippedColumns = @(
    "Employee Name"
    "Employee Status"
)


function Process-File {

    param (
        [string]$InputFile,
        [string]$OutputFile
    )

    #Determine what File Type we are inputting
    $fileExtension = [System.IO.Path]::GetExtension($InputFile)


    if ($fileExtension -eq ".xlsx") {
        # Read Excel file
        $data = Import-Excel -Path $InputFile

        Write-Host "Reading File: $InputFile"
    } elseif ($fileExtension -eq ".csv") {
        # Read CSV file
        $data = Import-Csv -Path $InputFile

        Write-Host "Reading File: $InputFile"
    } else {
        Write-Host "Unsupported file type: $fileExtension. Use XLSX or CSV."
        return
    }

    # Check if the column exists and rename it
    $columnNames = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

     # Loop through column mappings and rename matching columns
     foreach ($oldColumn in $columnMappings.Keys) {
        $newColumn = $columnMappings[$oldColumn]

        # Find matching column (case-insensitive)
        $matchedColumn = $columnNames | Where-Object { $_ -match [regex]::Escape($oldColumn) }

        if ($matchedColumn) {
            $data | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name $newColumn -Value $_.$matchedColumn -Force
                $_.PSObject.Properties.Remove($matchedColumn)  # Remove old column
            }
            Write-Output "Renamed '$matchedColumn' -> '$newColumn'"
        }
    }

    # Remove skipped columns
    foreach ($col in $skippedColumns) {
        if ($columnNames -contains $col) {
            $data | ForEach-Object { $_.PSObject.Properties.Remove($col) }
            Write-Output "Skipped column: $col"
        }
    }

    # Export the modified data to CSV
    $data | Export-Csv -Path $OutputFile -NoTypeInformation
    Write-Output "Processed file saved as: $OutputFile"

    
}


# Call the Process-File Function above
Process-File -InputFile $InputFile -OutputFile $OutputFile