<# Invoice-Analysis.ps1 for Pax8 Invoice CSVs for Nerdio, Azure, and CloudJumper (retired) cost analysis

# HOW TO USE:
## Prerequisites
- Script should be in its own folder
- Inside the folder, create a subfolder named "invoices" and put a CSV file from Pax8 (Billing, Invoices tab, "CSV" download button in Actions column for each month) into the folder.
- If you need to noramlize/transform one company/ID to another from the source Pax8 files to your reports, copy `SAMPLE-Company-Specific-Functions.ps1` to remove the `SAMPLE-` prefix and 
    edit with your changes.
- When running the script, it should have permissions to create folders in the current folder, and write files into the subfolder.
- This script has only been tested with PowerShell 7.4, though it may work with other versions.

The reason for the `Company-Specific-Functions.ps1` script is because we've encountered the following cases where we need to normalize data for our reports:
- Pax8 wasn't set up to assign certain licenses from our internal company to the correct one initially and the reporting will forever be 
    incorrect in the Pax8 file, so we adjust.
- Two companies merged into one and we want the report to reflect totals as if all historical billing were for one of the two companies.

If you don't need to make adjustments like this, you can ignore this file as the script checks to see if it exists, and also checks to confirm 
that the `Normalize-Row-Data` function exists (from that file) before calling it, so it will run fine without. The function, if it does exist, 
directly accesses the current $row values from the main script for comparisons, and then updates the original variables $companyName and $companyID 
from that loop iteration (by using $script:companyName and $script:companyID to access the variables from the script scope); there are no variables 
passed to or from the script. This isn't neceessarily the cleanest code, but it was the simplest way to extract the code with customer names out 
of the main script to make it sharable, providing an example for you to use, and making its use optional if you don't need it.

## Running The Script
Once all prerequisites are good to go, inside the script folder from a pwsh PowerShell prompt, run the script:

`./Invoice-Anslysis.ps1`

You can optionally provide command-line arguments to override some of the settings, if desired, with these named arguments (all optional with solid defaults):
- CsvFolderPath
- OutputPath
- OutputFile
- RawOutputFile

You can also pass the -Verbose switch to substantially increase the detail output to the screen during the run for diagnostic purposes.

By default, the script will validate that the "invoices" subfolder exists, will create (if it doesn't exist) a subfolder named yyyy-MM based on the date you're 
running the script, will loop through all the invoice .csv files in the "invoices" subfolder, perform it's analysis, and create two output files in the 
yyyy-MM folder, one named "analysis_raw.csv" and the other named "analysis.csv". If you pass different values to the parameters, the outputs will change.

See the param() block at the start of the code for the default values and the formatting of the argument inputs.

The `analysis.csv` file will contain one line per client-month with these columns (definitions in parentheses):
- company_name (the client's name from the Pax8 file, or as modified/transformed in your normalization)
- subtotal (the retail price subtotal of matching lines for the billing period from the Pax8 file)
- cost_total (same as subtotal but the sum of cost_total instead rather than price)
- margin (the difference between subtotal and cost_total columns)
- start_period (always the FIRST day of the month for the period in which all charges in that line were pulled from in the Pax8 .csv files)

The `analysis_raw.csv` file should basically have the lines from the original Pax8 .csv invoice files, but filtered to be only the ones that were used in 
calculating the analysis.csv data.

The items that are included in the analysis should be, if any:
- Azure arrears-billable consumption items, including Reserved Instances
- Microsoft subscription Windows licensing
- Nerdio licensing
- CloudJumper licenses (this service, purchased by NetApp, no longer exists but a client used to use it before moving to AVD)

All lines with a `subtotal` field equal to $0 will be skipped/excluded since it doesn't affect the totals. You can find the line of the script where 
`$filteredData` is defined to adjust your own version of which SKUs from the Pax8 invoices file will be included or not if you need to see or 
customize this for your purposes.

## How We Analyze Further

We take the output and open the two resulting .csv files from the subfolder in Excel and Save As .xlsx files in the same folder, to make saving formatting 
and updates easier.

It's not yet provided with sample data, but we take the resulitng `analysis.csv` file and use it to build a pivot table in Excel, where we add new lines each month 
and have the following additional columns we add to the five in `analysis.csv` which we then save as the file `new_analysis-thru-yyyy-MM.xlsx` in the same folder:
- retail_price (we manually update the price these items were sold to the client from their invoice to this column of the spreadsheet)
- margin_from_price_vs_cost (calculated value of the retail_price column minus the cost_total column, for us the Excel formula is `=F2-C2`)
- margin_percent (calculated value of margin_from_price_vs_cost divided by the retail_price column (for us the Excel formula is `=G2/F2`))
- invoice_month (calculated MONTH function value of the start_period value, for us the Excel formula is `=MONTH(E2)`)
- invoice_year (calculated YEAR function value of the start_period value, for us the Excel formula is `=YEAR(E2)`)
- invoice_yearmonth (calculated zero-padded yyyy-MM value string from the invoice_year and invoice_month columns, for us the Excel formula is `=J2 & "-" & TEXT(I2,"00")`)

From here, we can create or update two pivot tables with the above data as the source. The outcome is a pivot table with columns for client name, 
margin % average, total margin, total cost, and invoiced price, and a tree of rows we can expand and subtotal by client, year, or month. There are 
not currently samples of these additional fields or the pivot tables with sample data that have been created, so they're left as an exercise to 
the reader in the initial release. You can do whatever other analysis you wish, this is just how we use it so we can have a repeatable/updated 
process to get updated information regularly.

# VERSION HISTORY:
## V9 is a complete refactor of V8 on 2023-12-05 that changes the following (original version details not tracked):
- Reads in all invoice*.csv files at once to one large array.
- Filters the data using a hash that combines the company name and the Start Period date (SP_ID) as the key.
- Moves the company_name to be a separate field in the output since the key is no longer just company name.
- All final analysis output is based on the actual month in which an individual charge was for, regardless of invoice.
- Loses the name of the invoice file in the output but it isn't relevant for the final analysis.
- Data can be copied to an Excel spreadsheet with pivot table that pulls some date info out and collects invoice amounts.
- Removes the need for the Invoice-More-Summary.ps1 script to do further processing since the above replaces it.
- Removes the redundant assignment of $results to $allResults since there's no looping through CSVs now to collect.
- Removes multiple += array operator usages and increases performance.

## V10 is a minor tweak of V9 on 2023-12-06 that changes the following:
- Removes old debug comments to clean up code.
- Sets defaults for parameters after moving invoices around, also changes parameter names.
- Fixes parsing of company name for Nerdio licenses and ignores parent company free licenses.

## V11 is an adjustment on 2024-06-13 that changes the following:
- Add third company to the list of companies with Nerdio licenses.
- Refactors the code that updates the company for Nerdio licenses based on the Description field to properly output it in the summary.
- Adds some Verbose output for debugging the Nderio code activated if the -Verbose flag is passed.

## V12 is an adjustment on 2024-07-15 that changes the following for the first public release:
- Moves any client names and client-specific normalization into external Company-Specific-Functions.ps1 file to enable sharing without compromising privacy.
- Documenteation added for how to use the script to to the top in markdown format in order to create sharable version.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType 'Container' })]
    [string]$CsvFolderPath = '.\invoices',

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (".\" + (Get-Date -Format "yyyy-MM")),

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = 'analysis.csv',

    [Parameter(Mandatory = $false)]
    [string]$RawOutputFile = 'analysis_raw.csv'
)

# Verify if the folder to check exists
if (-Not (Test-Path -Path $CsvFolderPath)) {
    Write-Error "The folder '$CsvFolderPath' does not exist."
    exit 1
}

# Get all CSV files in the specified folder that start with "invoice"
$csvFiles = Get-ChildItem -Path $CsvFolderPath -Filter "invoice*.csv" -File
if ($csvFiles.Count -eq 0) {
    Write-Error "No CSV files found in the folder '$CsvFolderPath'."
    exit 1
}

# Verify if the output path exists, create it if not
if (-Not (Test-Path -Path $outputPath)) {
    Write-Host "The output path '$outputPath' does not exist. Creating it now."
    New-Item -ItemType Directory -Path $outputPath | Out-Null
}

# Join the output path with the filename
$OutputFilePath = Join-Path -Path $outputPath -ChildPath $OutputFile
$RawOutputFilePath = Join-Path -Path $outputPath -ChildPath $RawOutputFile

Write-Host "Reading files from $CsvFolderPath"
Write-Host "Output file: $OutputFilePath"
Write-Host "Raw output file: $RawOutputFilePath"
Write-Host "These files will be overwritten if they already exist."

# Import all CSV file lines into the $data variable
$data = Get-ChildItem -Path $CsvFolderPath -Filter "invoice*.csv" -File | ForEach-Object { 
    Import-Csv -Path $_
}

# Create an empty array list object to store the raw results
# Create a List to store the raw data
$rawData = New-Object System.Collections.Generic.List[object]

# Filter the data based on the sku field to only include Azure, Microsoft subscription Windows licensing, Nerdio, and CloudJumper (old) licenses
$filteredData = $data |
    Where-Object { $_.sku -like "MST-AZR*" -or $_.sku -like "MST-ARI*" -or $_.sku -like "AZR-ARR-*" -or $_.sku -like "NIO-INF*" -or $_.sku -like "CJR-*" } |
    Sort-Object -Property company_name, start_period

# Create a hashtable to store the results
$companySums = @{}

# Load in any Company-Specific-Functions file, if it exists, to allow for normalization of company name/ID specific to organization:
if (Test-Path -PathType Leaf -Path "./Company-Specific-Functions.ps1") {
    . "./Company-Specific-Functions.ps1"
}
else {
    Write-Host "No Company-Specific-Functions.ps1 file found in current folder, skipping these normalization/transformation calls."
}

# Iterate over each row in the filtered data
foreach ($row in $filteredData) {
    $companyName = $row.company_name
    $companyID = $row.company_id
    $subtotal = [decimal]$row.subtotal
    $costTotal = [decimal]$row.cost_total
    $startPeriod = [datetime]$row.start_period

    # First skip the line if the subtotal is $0 because it doesn't matter to the result.
    if ($script:subtotal -eq 0) {
        continue
    }

    # Call function in Comapany-Specific-Functions.ps1 script to normalize some company names/IDs by updating $script:companyName and $script:companyID to 
    # new values based on the values of $row.details. If the function doesn't exist, skip the call.
    if (Get-Command "Normalize-Row-Data-General" -ErrorAction SilentlyContinue) {
        Normalize-Row-Data
    }

    $row.company_name = $companyName
    $row.company_id = $companyID

    # Set the date to the first day of the month and make it the $sp_id
    $startPeriod = $startPeriod.Date.AddDays(1 - $startPeriod.Day)
    $sp_id = $startPeriod.ToString("yyyy-MM-dd")

    # If the company name is not already in the hashtable, add it
    if (-not $companySums.ContainsKey("$companyName-$sp_id")) {
        $companySums["$companyName-$sp_id"] = @{
            company_name = $companyName
            subtotal     = 0
            costTotal    = 0
            margin       = 0
            start_period = $startPeriod
        }
    }

    # Update the sums for the company
    $companySums["$companyName-$sp_id"].subtotal += $subtotal
    $companySums["$companyName-$sp_id"].costTotal += $costTotal
    $companySums["$companyName-$sp_id"].margin += ($subtotal - $costTotal)

    # Add the raw data to the $rawData array
    $rawData.Add($row)
}

# Convert the hashtable to an array of objects
$results = foreach ($key in $companySums.Keys) {
    [PSCustomObject]@{
        key          = $key
        company_name = $companySums[$key].company_name
        subtotal     = $companySums[$key].subtotal
        cost_total   = $companySums[$key].costTotal
        margin       = $companySums[$key].margin
        start_period = $companySums[$key].start_period.ToString("yyyy-MM-dd")
    }
}

# Sort the results alphabetically by company_name, and then by oldest-first date order by start_period
$results = $results | Sort-Object -Property @{Expression = "company_name"; Ascending = $true }, @{Expression = "start_period"; Ascending = $true }

# Export the results to a CSV file
$results | Select-Object company_name, subtotal, cost_total, margin, start_period | Export-Csv -Path $OutputFilePath -NoTypeInformation

# Export the raw results to a CSV file
$rawData | Export-Csv -Path $RawOutputFilePath -NoTypeInformation
