# Get the inputs from the commandline
param (
    [string]$InputFile,
    [string]$OutputFile
)

# Import necessary module
Import-Module ImportExcel

# Column Mappings: We will use the keys as the InputFile column names and the values as the OutputFile column names. 
# We will also use this as an "Ordered List" for the output file.
$columnMappings = [ordered]@{
    "Work Email" = "UserPrincipalName"
    "Job Title" = "JobTitle"
    "Manager" = "ManagerUserPrincipalName"
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

    Write-Output "Processing file: $InputFile"

    # Determine file type
    $fileExtension = [System.IO.Path]::GetExtension($InputFile).ToLower()

    # Read the input file
    if ($fileExtension -eq ".xlsx") {
        $data = Import-Excel -Path $InputFile
    } elseif ($fileExtension -eq ".csv") {
        $data = Import-Csv -Path $InputFile
    } else {
        Write-Output "Unsupported file type: $fileExtension. Use XLSX or CSV."
        return
    }

    # Get current column names
    $columnNames = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

    # Create a transformed dataset
    $transformedData = @()

    foreach ($row in $data) {
        # Create an empty ordered hashtable for each row
        $newRow = [ordered]@{}

        foreach ($oldColumn in $columnMappings.Keys) {
            $newColumn = $columnMappings[$oldColumn]

            # Find matching column (case-insensitive)
            $matchedColumn = $columnNames | Where-Object { $_ -match [regex]::Escape($oldColumn) }

            if ($matchedColumn) {
                $newRow[$newColumn] = $row.$matchedColumn
            } else {
                # If column is missing, keep it empty
                $newRow[$newColumn] = $null  
            }
        }

        $transformedData += [PSCustomObject]$newRow
    }

    # Force column order explicitly before exporting
    $orderedColumns = @($columnMappings.Values)

    # Export the modified data with enforced column order
    $transformedData | Select-Object -Property $orderedColumns | Export-Csv -Path $OutputFile -NoTypeInformation

    Write-Output "Processed file saved as: $OutputFile"
}


# Call the Process-File Function above
Process-File -InputFile $InputFile -OutputFile $OutputFile