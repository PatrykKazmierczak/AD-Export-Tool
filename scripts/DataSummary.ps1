# Import the required module
Import-Module ImportExcel

# Get all CSV files in the directory
$csvFiles = Get-ChildItem -Path "C:\ITOA" -Filter "*.csv"

# Create a new Excel package
$excel = New-ExcelPackage -Path "C:\ITOA\DataSummary.xlsx"

# Loop through each CSV file
foreach ($csvFile in $csvFiles) {
    # Import the CSV data
    $csvData = Import-Csv -Path $csvFile.FullName

    # Add a new worksheet to the Excel package with the CSV data
    Add-Worksheet -ExcelPackage $excel -WorksheetName $csvFile.BaseName -DataTable $csvData
}

# Save and close the Excel package
Close-ExcelPackage $excel