#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = 'Path to the data folder containing the HTML reports.')]
    [string]
    $DataPath = (Join-Path $PSScriptRoot 'data'),

    [Parameter(Mandatory = $false, HelpMessage = 'Path to the analysis folder where summary reports will be saved.')]
    [string]
    $AnalysisPath = (Join-Path $PSScriptRoot 'analysis'),

    [Parameter(Mandatory = $false, HelpMessage = 'Path to the client mapping CSV file.')]
    [string]
    $ClientMappingPath = (Join-Path $PSScriptRoot 'data' 'clientmapping.csv')
)

# --- Dependency Management: HtmlAgilityPack ---
try {
    Add-Type -Path (Join-Path $PSScriptRoot 'bin' 'HtmlAgilityPack.dll') -ErrorAction Stop
    Write-Verbose "HtmlAgilityPack.dll loaded from local bin directory."
} catch {
    Write-Verbose "Local HtmlAgilityPack.dll not found or failed to load. Attempting to download..."
    $binDir = Join-Path $PSScriptRoot 'bin'
    if (-not (Test-Path $binDir)) {
        New-Item -Path $binDir -ItemType Directory | Out-Null
    }
    $dllPath = Join-Path $binDir 'HtmlAgilityPack.dll'
    $packageUrl = 'https://www.nuget.org/api/v2/package/HtmlAgilityPack/1.11.46'
    $zipPath = Join-Path $binDir 'hap.zip'

    try {
        Invoke-WebRequest -Uri $packageUrl -OutFile $zipPath -UseBasicParsing
        Microsoft.PowerShell.Archive\Expand-Archive -Path $zipPath -DestinationPath $binDir -Force
        $hapDll = Get-ChildItem -Path (Join-Path $binDir 'lib' 'netstandard2.0') -Filter 'HtmlAgilityPack.dll' | Select-Object -First 1
        if ($hapDll) {
            Copy-Item -Path $hapDll.FullName -Destination $dllPath -Force
            Add-Type -Path $dllPath
            Write-Verbose "Successfully downloaded and loaded HtmlAgilityPack.dll"
        } else {
            throw "Could not find HtmlAgilityPack.dll in the downloaded package."
        }
    } catch {
        Write-Error "Failed to download or load HtmlAgilityPack. A working internet connection is required for the first run. Error: $_"
        return
    } finally {
        if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    }
}

# --- Helper Functions ---
function Get-TableByHeading($htmlDoc, $headingText) {
    $headingNode = $htmlDoc.DocumentNode.SelectSingleNode("//h2[normalize-space(.)='${headingText}']")
    if ($headingNode) {
        # The table is often wrapped in a div that is the sibling of the h2
        return $headingNode.SelectSingleNode("following-sibling::div[1]/table[1]")
    }
    return $null
}

# Helper function to determine background color based on percentage and column type
function Get-BackgroundColor {
    param(
        [string]$PercentageText,
        [string]$ColumnType
    )
    
    # Extract numeric percentage from text (e.g., "94%" -> 94)
    if ($PercentageText -match '(\d+(?:\.\d+)?)%') {
        $percentage = [double]$matches[1]
    } else {
        return "" # No background color for non-percentage values
    }
    
    if ($ColumnType -eq "DKIMAligned") {
        # DKIM Aligned: Higher is better
        if ($percentage -ge 80) {
            return "background: #e3fbe3; background: rgba(144,238,144,0.25); border-left: 4px solid #90ee90;"
        } elseif ($percentage -ge 50) {
            return "background: #fbf9ea; background: rgba(238,232,170,0.25); border-left: 4px solid #eee8aa;"
        } else {
            return "background: #ffd8d1; background: rgba(255,99,71,0.25); border-left: 4px solid #ff6347;"
        }
    } elseif ($ColumnType -eq "Rejected") {
        # Rejected: Lower is better
        if ($percentage -lt 10) {
            return "background: #e3fbe3; background: rgba(144,238,144,0.25); border-left: 4px solid #90ee90;"
        } elseif ($percentage -lt 50) {
            return "background: #fbf9ea; background: rgba(238,232,170,0.25); border-left: 4px solid #eee8aa;"
        } else {
            return "background: #ffd8d1; background: rgba(255,99,71,0.25); border-left: 4px solid #ff6347;"
        }
    }
    
    return "" # Default: no background color
}

function Get-CellText {
    param(
        [Parameter(Mandatory = $true)]
        [HtmlAgilityPack.HtmlNode]$Cell,
        [string]$ColumnType = "default"
    )
    
    if ($null -eq $Cell) {
        return ""
    }
    
    # Special handling for Rating column - extract both letter grade and score
    if ($ColumnType -eq "rating") {
        $textNode = $Cell.SelectSingleNode(".//text()[normalize-space()]")
        $letterGrade = if ($textNode) { $textNode.InnerText.Trim() } else { "" }
        $scoreSpan = $Cell.SelectSingleNode(".//span[contains(@style, 'color: #555')]")
        if ($scoreSpan) {
            $score = $scoreSpan.InnerText.Trim()
            return "$letterGrade $score"
        }
        return $letterGrade
    }
    

    # Special handling for DKIM Aligned and Rejected columns - preserve arrows and secondary percentages
    if ($ColumnType -eq "dkim_metric") {
        $htmlContent = $Cell.InnerHtml
        
        # Extract primary percentage
        $primarySpan = $Cell.SelectSingleNode(".//span[1]")
        $primaryPercent = if ($primarySpan) { $primarySpan.InnerText.Trim() } else { "" }
        
        # Extract arrow with color
        $arrowSpan = $Cell.SelectSingleNode(".//span[contains(@style, 'font-size: 0.85em')]")
        $arrow = if ($arrowSpan) { $arrowSpan.InnerText.Trim() } else { "" }
        $arrowColor = ""
        if ($arrowSpan) {
            $style = $arrowSpan.GetAttributeValue("style", "")
            if ($style -match "color:\s*(\w+)") {
                $arrowColor = $matches[1]
            }
        }
        
        # Extract secondary percentage
        $secondarySpan = $Cell.SelectSingleNode(".//span[contains(@style, 'font-size: 0.7em')]")
        $secondaryPercent = if ($secondarySpan) { $secondarySpan.InnerText.Trim() } else { "" }
        
        # Format the result with color-coded arrow and newline for secondary percentage
        if ($arrow -and $secondaryPercent) {
            if ($arrowColor) {
                return "$primaryPercent <span style='color: $arrowColor;'>$arrow</span>`n$secondaryPercent"
            } else {
                return "$primaryPercent $arrow`n$secondaryPercent"
            }
        } elseif ($arrow) {
            if ($arrowColor) {
                return "$primaryPercent <span style='color: $arrowColor;'>$arrow</span>"
            } else {
                return "$primaryPercent $arrow"
            }
        } elseif ($secondaryPercent) {
            # No arrow but secondary percentage exists - show in parentheses
            return "$primaryPercent ($secondaryPercent)"
        } else {
            return $primaryPercent
        }
    }
    
    # Default handling - replace <br> tags with actual newlines for proper multiline parsing
    $htmlContent = $Cell.InnerHtml
    $htmlContent = $htmlContent -replace '<br\s*/?>', "`n"
    
    # Create a temporary HTML document to parse the modified content
    $tempDoc = [HtmlAgilityPack.HtmlDocument]::new()
    $tempDoc.LoadHtml("<div>$htmlContent</div>")
    
    # Get the text content with newlines preserved
    $text = $tempDoc.DocumentNode.InnerText
    
    return $text.Trim()
}

function Format-ComplexField($rawText) {
    $lines = $rawText.Trim() -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($lines.Count -eq 0) { return '' }

    $status = $lines[0]
    $detail = if ($lines.Count -gt 1) { $lines[1..($lines.Count - 1)] -join ' ' } else { '' }

    # Convert "n" to "no" for MTA-STS and TLSRPT columns
    if ($status -eq 'n') {
        $status = 'no'
    }

    if ($status -match '^(yes|no|unknown|pass|fail|testing)$' -and -not [string]::IsNullOrWhiteSpace($detail)) {
        return "$($status) ($($detail))"
    }
    
    return $status
}

function Get-DmarcDataForDomain($domain, $dmarcReportsData) {
    if ($dmarcReportsData.ContainsKey($domain)) {
        return $dmarcReportsData[$domain]
    }

    $parts = $domain.Split('.')
    for ($i = 1; $i -lt $parts.Length - 1; $i++) {
        $parentDomain = $parts[$i..($parts.Length - 1)] -join '.'
        if ($dmarcReportsData.ContainsKey($parentDomain)) {
            return $dmarcReportsData[$parentDomain]
        }
    }

    return $null
}

# --- Main Script Logic ---
# Ensure data directory exists
if (-not (Test-Path $DataPath)) {
    Write-Host "Data directory '$DataPath' does not exist." -ForegroundColor Yellow
    Write-Host "This directory is needed to store MailHardener HTML reports for processing." -ForegroundColor Yellow
    Write-Host ""
    $response = Read-Host "Would you like to create the data directory now? (y/N)"
    if ($response -eq 'y' -or $response -eq 'Y') {
        New-Item -ItemType Directory -Path $DataPath -Force | Out-Null
        Write-Host "Created data directory: $DataPath" -ForegroundColor Green
        Write-Host ""
        Write-Host "NEXT STEPS:" -ForegroundColor Cyan
        Write-Host "1. Download your MailHardener HTML reports" -ForegroundColor White
        Write-Host "2. Place them in the '$DataPath' directory" -ForegroundColor White
        Write-Host "3. Run this script again" -ForegroundColor White
        Write-Host "4. See README.md for additional configuration options" -ForegroundColor White
        Write-Host ""
        Write-Host "Exiting. Please add HTML reports to '$DataPath' and run the script again." -ForegroundColor Yellow
        return
    } else {
        Write-Host "Data directory creation cancelled. Exiting." -ForegroundColor Red
        Write-Host "Please create the '$DataPath' directory manually and add your MailHardener HTML reports." -ForegroundColor Yellow
        Write-Host "See README.md for more information." -ForegroundColor Yellow
        return
    }
}

# Ensure analysis directory exists
if (-not (Test-Path $AnalysisPath)) {
    New-Item -ItemType Directory -Path $AnalysisPath -Force | Out-Null
}

# Load client mapping if it exists
$clientMapping = @{}
if (Test-Path $ClientMappingPath) {
    Write-Verbose "Loading client mapping from $ClientMappingPath"
    $mappingData = Import-Csv -Path $ClientMappingPath
    foreach ($entry in $mappingData) {
        if (-not [string]::IsNullOrWhiteSpace($entry.Domain)) {
            $clientMapping[$entry.Domain] = $entry.ClientName
        }
    }
    Write-Verbose "Loaded $($clientMapping.Count) client mappings"
}

$reportFiles = Get-ChildItem -Path $DataPath -Filter '*.html' -File
if (-not $reportFiles) {
    Write-Host ""
    Write-Host "No HTML report files found in '$($DataPath)'." -ForegroundColor Red
    Write-Host ""
    Write-Host "TO RESOLVE THIS ISSUE:" -ForegroundColor Cyan
    Write-Host "1. Download your MailHardener HTML reports from the MailHardener dashboard" -ForegroundColor White
    Write-Host "2. Place the HTML files in the '$DataPath' directory" -ForegroundColor White
    Write-Host "3. Run this script again" -ForegroundColor White
    Write-Host ""
    Write-Host "EXPECTED FILE STRUCTURE:" -ForegroundColor Cyan
    Write-Host "$DataPath/" -ForegroundColor White
    Write-Host "├── report April 2025.html" -ForegroundColor Gray
    Write-Host "├── report May 2025.html" -ForegroundColor Gray
    Write-Host "└── clientmapping.csv (optional)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "For more information, see README.md in the project directory." -ForegroundColor Yellow
    Write-Host ""
    return
}

# Initialize collection for index.html generation
$reportIndex = @()

# Process each file individually
foreach ($file in $reportFiles) {
    Write-Verbose "Processing file: $($file.Name)"
    $fileContent = Get-Content -Path $file.FullName -Raw
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($fileContent)

    # Extract summary info from current file
    $aggregatedPeriod = $doc.DocumentNode.SelectSingleNode("//p[starts-with(normalize-space(.), 'Aggregated period:')] ")?.InnerText.Trim()
    $dmarcReportInfo = $doc.DocumentNode.SelectSingleNode("//p[starts-with(normalize-space(.), 'DMARC reports represent')] ")?.InnerText.Trim() -replace '\(in subtext\)', '(current vs. last delta)'
    $domainRatingDate = $doc.DocumentNode.SelectSingleNode("//p[starts-with(normalize-space(.), 'The domain rating below is based on')] ")?.InnerText.Trim()

    # Extract DMARC reports data from current file
    $dmarcReportsData = @{}
    $dmarcReportsTable = Get-TableByHeading -htmlDoc $doc -headingText 'DMARC reports'
    if ($dmarcReportsTable) {
        $headers = $dmarcReportsTable.SelectNodes(".//th") | ForEach-Object { $_.InnerText.Trim() }
        Write-Verbose "DMARC table headers found: $($headers -join ', ')"
        $dkimAlignedIndex = [array]::IndexOf($headers, 'DKIM aligned')
        # Handle superscript numbers in headers (e.g., "Rejected2", "Volume1")
        $rejectedIndex = -1
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i] -match '^Rejected\d*$') {
                $rejectedIndex = $i
                break
            }
        }
        Write-Verbose "DKIM aligned index: $dkimAlignedIndex, Rejected index: $rejectedIndex"

        if ($dkimAlignedIndex -ne -1 -and $rejectedIndex -ne -1) {
            $rows = $dmarcReportsTable.SelectNodes(".//tr[td]")
            if ($rows) {
                foreach ($row in $rows) {
                    $cells = $row.SelectNodes(".//td")
                    if ($cells.Count -gt $dkimAlignedIndex -and $cells.Count -gt $rejectedIndex) {
                        $domain = $cells[0].InnerText.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($domain)) {
                            $dkimAlignedText = Get-CellText -Cell $cells[$dkimAlignedIndex] -ColumnType "dkim_metric"
                            $rejectedText = Get-CellText -Cell $cells[$rejectedIndex] -ColumnType "dkim_metric"
                            # Use the formatted text directly from Get-CellText
                            $dkimAlignedValue = $dkimAlignedText.Trim()
                            $rejectedValue = $rejectedText.Trim()
                            
                            Write-Verbose "Domain: $domain, DKIM Aligned: '$dkimAlignedValue', Rejected: '$rejectedValue'"
                            
                            $dmarcReportsData[$domain] = @{
                                DKIMAligned = $dkimAlignedValue
                                Rejected    = $rejectedValue
                            }
                        }
                    }
                }
            }
        }
    }

    # Process "Domain rating" table and merge with DMARC data from current file
    $domainData = @()
    $domainRatingTable = Get-TableByHeading -htmlDoc $doc -headingText 'Domain rating'
    if ($domainRatingTable) {
        $rows = $domainRatingTable.SelectNodes(".//tr[td]")
        if ($rows) {
            foreach ($row in $rows) {
                $cells = $row.SelectNodes(".//td")
                if ($cells.Count -ge 7) {
                    $domain = $cells[0].InnerText.Trim()
                    $dmarcText = $cells[2].InnerText.Trim()
                    $dmarcValue = 'no'
                    $dmarcDetail = 'N/A'

                    if ($dmarcText -match '^(yes|no|unknown)(.*)') {
                        $dmarcValue = $matches[1]
                        $dmarcDetail = $matches[2].Trim()
                        if ([string]::IsNullOrWhiteSpace($dmarcDetail)) {
                            $dmarcDetail = 'ok'
                        }
                    }

                    $dmarcData = Get-DmarcDataForDomain -domain $domain -dmarcReportsData $dmarcReportsData

                    # Get DKIM Aligned and Rejected values with background colors
                    $dkimAlignedValue = if ($dmarcData) { $dmarcData.DKIMAligned.Trim() } else { 'N/A' }
                    $rejectedValue = if ($dmarcData) { $dmarcData.Rejected.Trim() } else { 'N/A' }
                    
                    $dkimAlignedBgColor = Get-BackgroundColor -PercentageText $dkimAlignedValue -ColumnType "DKIMAligned"
                    $rejectedBgColor = Get-BackgroundColor -PercentageText $rejectedValue -ColumnType "Rejected"

                    # Get client name from mapping if available
                    $clientName = ''
                    if ($clientMapping.ContainsKey($domain)) {
                        $clientName = $clientMapping[$domain]
                    }

                    $domainData += [PSCustomObject]@{
                        ClientName  = $clientName
                        Domain      = $domain
                        Rating      = Get-CellText -Cell $cells[1] -ColumnType "rating"
                        DMARC       = $dmarcValue
                        DMARCStatus = $dmarcDetail
                        SPF         = Format-ComplexField -rawText (Get-CellText -Cell $cells[3])
                        DKIM        = Format-ComplexField -rawText (Get-CellText -Cell $cells[4])
                        MTASTS      = Format-ComplexField -rawText (Get-CellText -Cell $cells[5])
                        TLSRPT      = Format-ComplexField -rawText (Get-CellText -Cell $cells[6])
                        DKIMAligned = $dkimAlignedValue
                        Rejected    = $rejectedValue
                        DKIMAlignedBgColor = $dkimAlignedBgColor
                        RejectedBgColor = $rejectedBgColor
                    }
                }
            }
        }
    }

    # --- Report Generation for current file ---
    $domainData = $domainData | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_.Domain) }
    
    # Sort domains by DMARC Status first, then by Client Name (if available), then by Domain
    # This ensures domains are grouped by client within each DMARC status group
    if ($clientMapping.Count -gt 0) {
        $domainData = $domainData | Sort-Object DMARCStatus, @{Expression = {if ([string]::IsNullOrWhiteSpace($_.ClientName)) { "zzz_NoClient" } else { $_.ClientName }}}, Domain
    } else {
        $domainData = $domainData | Sort-Object DMARCStatus, Domain
    }
    
    $noneCount = ($domainData | Where-Object { $_.DMARCStatus -eq 'none' }).Count
    $otherCount = $domainData.Count - $noneCount
    $jsonData = $domainData | ConvertTo-Json -Compress

    # Generate output filename and extract month for heading
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $outputPath = Join-Path -Path $AnalysisPath -ChildPath "$baseName-Summary.html"
    
    # Extract period and year from filename for dynamic heading
    $reportMonth = "Unknown"
    if ($baseName -match "report\s+(week\s+\d+\s+of\s+\d{4})") {
        # Weekly report pattern: "report week 27 of 2025"
        $reportMonth = $matches[1]
    } elseif ($baseName -match "report\s+(\w+\s+\d{4})") {
        # Monthly report pattern: "report April 2025"
        $reportMonth = $matches[1]
    }

    $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>MailHardener Summary Report - {REPORT_MONTH} (Company Confidential)</title>
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/rowgroup/1.4.1/css/rowGroup.dataTables.min.css">
    <script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.7.0.js"></script>
    <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/rowgroup/1.4.1/js/dataTables.rowGroup.min.js"></script>
    <style>
        body { font-family: sans-serif; padding: 2em; font-size: 0.9em; background-color: #f5f5f5; }
        h1 { color: #3f51b5; }
        table.dataTable thead th { background-color: #3f51b5; color: white; }
        tr.dtrg-start.dtrg-level-0 > td { font-weight: bold; background-color: #e8eaf6 !important; }
        .dataTables_wrapper .dataTables_filter input, .dataTables_wrapper .dataTables_length select {
            border: 1px solid #ccc;
            border-radius: 4px;
            padding: 5px;
        }
        #reportTable { background-color: white; }
    </style>
</head>
<body>
    <h1>MailHardener Summary Report - {REPORT_MONTH} (Company Confidential)</h1>
    <div style="margin-bottom: 1em; text-align: left;">
        <a href="index.html" style="color: #1976d2; text-decoration: none; font-weight: bold; padding: 0.5em 1em; border: 1px solid #1976d2; border-radius: 4px; background-color: #f5f5f5; transition: background-color 0.3s;" onmouseover="this.style.backgroundColor='#e3f2fd'" onmouseout="this.style.backgroundColor='#f5f5f5'">← Back to Reports Index</a>
    </div>
    <div class="summary" style="margin-bottom: 2em; padding: 1em; background-color: #e8eaf6; border-radius: 4px;">
        <h2>Report Information</h2>
        <p><em>NOTE: Mailhardener does not store the history of domain properties. Please be aware that the 'domain rating' section in the report reflects current known values for the domain at the time the report was generated, not historical data. The DMARC reports and SMTP TLS reports are historic.</em></p>
        <p>{AGGREGATED_PERIOD}</p>
        <p>{DMARC_REPORT_INFO}</p>
        <p>{DOMAIN_RATING_DATE}</p>
        <p><em>Report Generated: {REPORT_GENERATION_TIME}</em></p>
        <h2 style="margin-top: 1.5em;">DMARC Status Summary</h2>
        <p style="font-size: 1.1em;">Domains with DMARC status 'none': <strong style="color:rgb(164, 0, 0); font-size: 1.3em;">{NONE_COUNT}</strong></p>
        <p style="font-size: 1.1em;">Domains with other DMARC statuses (e.g., quarantine, reject): <strong style="color:rgb(0, 134, 0); font-size: 1.1em;">{OTHER_COUNT}</strong></p>
        <h2 style="margin-top: 1.5em;">Desired Outcome</h2>
        <p>As a baseline, all domain names should have their DMARC settings configured and set to 'yes'. The goal is to change 
        all DMARC Status values from 'none' to 'quarantine' or 'reject' while ensuring only impersonated emails are 
        marked as such, and all legitimate email from legitimate sources are DMARC-aligned. Clicking the values will take you 
        to the MailHardener dashboard for that domain in the correct area for more details.</p>
    </div>
    <table id="reportTable" class="display" style="width:100%">
        <thead>
            <tr>
                <th>Client Name</th>
                <th>Domain</th>
                <th>Rating</th>
                <th>DMARC</th>
                <th>DMARC Status</th>
                <th>SPF</th>
                <th>DKIM</th>
                <th style="color: #888;">MTA-STS</th>
                <th style="color: #888;">TLSRPT</th>
                <th>DKIM Aligned</th>
                <th>Rejected</th>
            </tr>
        </thead>
        <tbody>
        </tbody>
    </table>

    <script>
    $(document).ready(function() {
        var dataSet = {JSON_DATA};
        var table = $('#reportTable').DataTable({
            data: dataSet,
            columns: [
                { data: 'ClientName', title: 'Client Name', visible: false },
                { 
                    data: 'Domain',
                    render: function(data, type, row) {
                        // Add left padding to align with client heading
                        return '<span style="padding-left: 28px;">' + data + '</span>';
                    }
                },
                { data: 'Rating', render: function(data, type, row) { 
                    return '<a href="https://www.mailhardener.com/dashboard/overview/' + row.Domain + '" target="_blank">' + data + '</a>'; 
                } },
                { data: 'DMARC', render: function(data, type, row) { 
                    return '<a href="https://www.mailhardener.com/dashboard/dmarc/' + row.Domain + '" target="_blank">' + data + '</a>'; 
                } },
                { data: 'DMARCStatus', render: function(data, type, row) { 
                    return '<a href="https://www.mailhardener.com/dashboard/dmarc-aggregate-reports/' + row.Domain + '" target="_blank">' + data + '</a>'; 
                } },
                { data: 'SPF', render: function(data, type, row) { 
                    return '<a href="https://www.mailhardener.com/dashboard/spf/' + row.Domain + '" target="_blank">' + data + '</a>'; 
                } },
                { data: 'DKIM', render: function(data, type, row) { 
                    return '<a href="https://www.mailhardener.com/dashboard/dkim/' + row.Domain + '" target="_blank">' + data + '</a>'; 
                } },
                { data: 'MTASTS', render: function(data, type, row) { 
                    return '<a href="https://www.mailhardener.com/dashboard/mtasts/' + row.Domain + '" target="_blank"><span style="color: #888;">' + data + '</span></a>'; 
                } },
                { data: 'TLSRPT', render: function(data, type, row) { 
                    var linkUrl = data.toLowerCase().startsWith('no') ? 
                        'https://www.mailhardener.com/dashboard/tlsrpt/' + row.Domain : 
                        'https://www.mailhardener.com/dashboard/smtp-tls-reports/' + row.Domain; 
                    return '<a href="' + linkUrl + '" target="_blank"><span style="color: #888;">' + data + '</span></a>'; 
                } },
                { data: 'DKIMAligned', 
                    createdCell: function(td, cellData, rowData, row, col) {
                        if (rowData.DKIMAlignedBgColor) {
                            $(td).attr('style', rowData.DKIMAlignedBgColor);
                        }
                    }
                },
                { data: 'Rejected', 
                    createdCell: function(td, cellData, rowData, row, col) {
                        if (rowData.RejectedBgColor) {
                            $(td).attr('style', rowData.RejectedBgColor);
                        }
                    }
                }
            ],
            pageLength: 100,
            order: [[4, 'asc']], // Update to use DMARC Status column (index 4 with Client Name added),
            rowGroup: {
                // Enable multi-level grouping
                multiLevelGrp: true,
                // Define data sources for each level
                dataSrc: [
                    // Level 0: DMARC Status
                    'DMARCStatus',
                    // Level 1: Client Name
                    function(row) {
                        return row.ClientName || 'Unknown Client';
                    }
                ],
                // Customize the rendering of each level
                startRender: function(rows, group, level) {
                    if (level === 0) {
                        // Level 0: DMARC Status
                        var status = group.toLowerCase().trim();
                        var style = '';
                        
                        if (status === 'quarantine' || status === 'reject') {
                            style = 'color: #006400; font-weight: bold;';
                        } else if (status === 'none') {
                            style = 'color: #8B4513; font-weight: bold;';
                        }
                        
                        return '<span style="' + style + '">' + group + ' (' + rows.count() + ' domains)</span>';
                    } else {
                        // Level 1: Client Name
                        return group + ' (' + rows.count() + ' domains)';
                    }
                },
                // Enable both levels of grouping
                enabled: true
            },
            createdRow: function(row, data, dataIndex) {
                // DMARC column (index 2 with Client Name hidden)
                if (data.DMARC && data.DMARC.toLowerCase() !== 'yes') {
                    $('td', row).eq(2).attr('style', 'background: #ffd8d1; background: rgba(255,99,71,0.25); border-left: 4px solid #ff6347;');
                }

                // DMARC Status column (index 3 with Client Name hidden)
                if (data.DMARCStatus) {
                    var status = data.DMARCStatus.toLowerCase().trim();
                    if (status === 'quarantine' || status === 'reject') {
                        $('td', row).eq(3).attr('style', 'background: #e3fbe3; background: rgba(144,238,144,0.25); border-left: 4px solid #90ee90;');
                    } else if (status === 'none') {
                        $('td', row).eq(3).attr('style', 'background: #fbf9ea; background: rgba(238,232,170,0.25); border-left: 4px solid #eee8aa;');
                    }
                }

                // SPF column (index 4 with Client Name hidden)
                if (data.SPF && !data.SPF.toLowerCase().startsWith('yes')) {
                    $('td', row).eq(4).attr('style', 'background: #ffd8d1; background: rgba(255,99,71,0.25); border-left: 4px solid #ff6347;');
                }

                // DKIM column (index 5 with Client Name hidden)
                if (data.DKIM && !data.DKIM.toLowerCase().startsWith('yes')) {
                    var dkimStatus = data.DKIM.toLowerCase();
                    var tlsrptStatus = data.TLSRPT ? data.TLSRPT.toLowerCase() : "";
                    // Don't color red if DKIM starts with 'unknown' and TLSRPT includes 'null mx'
                    if (!(dkimStatus.startsWith('unknown') && tlsrptStatus.includes('null mx'))) {
                        $('td', row).eq(5).attr('style', 'background: #ffd8d1; background: rgba(255,99,71,0.25); border-left: 4px solid #ff6347;');
                    }
                }
            }
        });

        // Make row groups collapsible
        $('#reportTable tbody').on('click', 'tr.dtrg-start', function() {
            var level = $(this).data('dtrg-level');
            if (level === 0) {
                // For DMARC Status level, toggle all rows until the next DMARC Status group
                $(this).nextUntil('tr.dtrg-level-0').toggle();
            } else {
                // For Client Name level, toggle only rows until the next group of any level
                $(this).nextUntil('tr.dtrg-start').toggle();
            }
        });
        
        // Add cursor pointer to group headers to indicate they're clickable
        $('#reportTable').on('draw.dt', function() {
            $('tr.dtrg-start').css('cursor', 'pointer');
        });

    });
    </script>

</body>
</html>
'@

    $reportGenerationTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $htmlContent = $htmlTemplate -replace '{REPORT_MONTH}', $reportMonth -replace '{AGGREGATED_PERIOD}', $aggregatedPeriod -replace '{DMARC_REPORT_INFO}', $dmarcReportInfo -replace '{DOMAIN_RATING_DATE}', $domainRatingDate -replace '{REPORT_GENERATION_TIME}', $reportGenerationTime -replace '{NONE_COUNT}', $noneCount -replace '{OTHER_COUNT}', $otherCount -replace '{JSON_DATA}', $jsonData

    $htmlContent | Out-File -FilePath $outputPath -Encoding UTF8

    # Collect report metadata for index.html generation
    $reportInfo = @{
        OriginalFile = $file.Name
        SummaryFile = "$baseName-Summary.html"
        BaseName = $baseName
        ReportMonth = $reportMonth
        IsWeekly = $baseName -match "week\s+\d+"
        IsMonthly = $baseName -match "^report\s+\w+\s+\d{4}$"
    }
    
    # Parse additional date information for sorting
    if ($reportInfo.IsWeekly -and $baseName -match "week\s+(\d+)\s+of\s+(\d{4})") {
        $reportInfo.WeekNumber = [int]$matches[1]
        $reportInfo.Year = [int]$matches[2]
        # Calculate approximate month for week (week 1-4 = Jan, 5-8 = Feb, etc.)
        $reportInfo.ApproxMonth = [Math]::Min(12, [Math]::Max(1, [Math]::Ceiling($reportInfo.WeekNumber / 4.33)))
    } elseif ($reportInfo.IsMonthly -and $baseName -match "^report\s+(\w+)\s+(\d{4})$") {
        $reportInfo.MonthName = $matches[1]
        $reportInfo.Year = [int]$matches[2]
        # Convert month name to number for sorting
        $monthMap = @{
            'January' = 1; 'February' = 2; 'March' = 3; 'April' = 4; 'May' = 5; 'June' = 6;
            'July' = 7; 'August' = 8; 'September' = 9; 'October' = 10; 'November' = 11; 'December' = 12
        }
        $reportInfo.MonthNumber = $monthMap[$reportInfo.MonthName]
        if (-not $reportInfo.MonthNumber) { $reportInfo.MonthNumber = 99 } # Unknown month goes to end
    }
    
    $reportIndex += $reportInfo

    Write-Host "Successfully generated summary report: $outputPath" -ForegroundColor Green
}

# --- Generate Index.html ---
Write-Verbose "Generating index.html..."

# Sort reports by year and month for proper ordering
$sortedReports = $reportIndex | Sort-Object {
    if ($_.Year) { $_.Year } else { 9999 }
}, {
    if ($_.MonthNumber) { $_.MonthNumber } 
    elseif ($_.ApproxMonth) { $_.ApproxMonth }
    else { 99 }
}, {
    if ($_.WeekNumber) { $_.WeekNumber } else { 0 }
}

# Group reports by year and month for hierarchical display
$groupedReports = @{}
foreach ($report in $sortedReports) {
    $year = if ($report.Year) { $report.Year } else { "Unknown" }
    $monthKey = if ($report.IsMonthly -and $report.MonthNumber) {
        "$year-$($report.MonthNumber.ToString('00'))-$($report.MonthName)"
    } elseif ($report.IsWeekly -and $report.ApproxMonth) {
        "$year-$($report.ApproxMonth.ToString('00'))-Month$($report.ApproxMonth)"
    } else {
        "$year-99-Unknown"
    }
    
    if (-not $groupedReports[$monthKey]) {
        $groupedReports[$monthKey] = @{
            MonthlyReport = $null
            WeeklyReports = @()
            Year = $year
            MonthNumber = if ($report.MonthNumber) { $report.MonthNumber } elseif ($report.ApproxMonth) { $report.ApproxMonth } else { 99 }
            MonthName = if ($report.MonthName) { $report.MonthName } else { "Month $($report.ApproxMonth)" }
        }
    }
    
    if ($report.IsMonthly) {
        $groupedReports[$monthKey].MonthlyReport = $report
    } elseif ($report.IsWeekly) {
        $groupedReports[$monthKey].WeeklyReports += $report
    }
}

# Generate HTML content for index
$indexHtml = @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MailHardener Reports Index (Company Confidential)</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 2em;
            background-color: #f5f5f5;
            color: #333;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 2em;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #3f51b5;
            border-bottom: 3px solid #3f51b5;
            padding-bottom: 0.5em;
            margin-bottom: 1em;
        }
        .report-list {
            list-style: none;
            padding: 0;
        }
        .month-item {
            margin-bottom: 1em;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            overflow: hidden;
        }
        .month-header {
            background: #e8eaf6;
            padding: 1em;
            font-weight: bold;
            color: #3f51b5;
        }
        .month-content {
            padding: 0;
        }
        .report-link {
            display: block;
            padding: 0.75em 1em;
            text-decoration: none;
            color: #1976d2;
            border-bottom: 1px solid #f0f0f0;
            transition: background-color 0.2s;
        }
        .report-link:hover {
            background-color: #f8f9fa;
            text-decoration: underline;
        }
        .report-link:last-child {
            border-bottom: none;
        }
        .weekly-report {
            padding-left: 2em;
            background-color: #fafafa;
            font-size: 0.9em;
        }
        .weekly-report .report-link {
            color: #666;
        }
        .orphaned-weekly {
            background-color: #fff3e0;
        }
        .orphaned-weekly .month-header {
            background-color: #ffe0b2;
            color: #f57c00;
        }
        .generated-info {
            margin-top: 2em;
            padding: 1em;
            background-color: #f8f9fa;
            border-radius: 4px;
            font-size: 0.9em;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>MailHardener Reports Index (Company Confidential)</h1>
        <ul class="report-list">
'@

# Sort month keys for display
$sortedMonthKeys = $groupedReports.Keys | Sort-Object {
    $parts = $_ -split '-'
    [int]$parts[0] # Year
}, {
    $parts = $_ -split '-'
    [int]$parts[1] # Month number
}

foreach ($monthKey in $sortedMonthKeys) {
    $monthGroup = $groupedReports[$monthKey]
    $hasMonthlyReport = $null -ne $monthGroup.MonthlyReport
    $hasWeeklyReports = $monthGroup.WeeklyReports.Count -gt 0
    
    if ($hasMonthlyReport -or $hasWeeklyReports) {
        $isOrphaned = $hasWeeklyReports -and -not $hasMonthlyReport
        $cssClass = if ($isOrphaned) { "month-item orphaned-weekly" } else { "month-item" }
        
        $indexHtml += "            <li class=`"$cssClass`">`n"
        $indexHtml += "                <div class=`"month-header`">`n"
        
        if ($hasMonthlyReport) {
            $indexHtml += "                    $($monthGroup.MonthlyReport.ReportMonth)`n"
        } else {
            # Orphaned weekly reports - show year and approximate month
            $indexHtml += "                    $($monthGroup.Year) - Week(s) in $($monthGroup.MonthName)`n"
        }
        
        $indexHtml += "                </div>`n"
        $indexHtml += "                <div class=`"month-content`">`n"
        
        # Add monthly report link if it exists
        if ($hasMonthlyReport) {
            $indexHtml += "                    <a href=`"$($monthGroup.MonthlyReport.SummaryFile)`" class=`"report-link`">Monthly Report - $($monthGroup.MonthlyReport.ReportMonth)</a>`n"
        }
        
        # Add weekly report links
        foreach ($weeklyReport in ($monthGroup.WeeklyReports | Sort-Object WeekNumber)) {
            $cssClass = if ($hasMonthlyReport) { "report-link weekly-report" } else { "report-link" }
            $indexHtml += "                    <a href=`"$($weeklyReport.SummaryFile)`" class=`"$cssClass`">Weekly Report - $($weeklyReport.ReportMonth)</a>`n"
        }
        
        $indexHtml += "                </div>`n"
        $indexHtml += "            </li>`n"
    }
}

$indexHtml += @'
        </ul>
        <div class="generated-info">
            <p><strong>Generated:</strong> {GENERATION_TIME}</p>
            <p><strong>Total Reports:</strong> {TOTAL_REPORTS}</p>
        </div>
    </div>
</body>
</html>
'@

# Replace placeholders
$generationTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$totalReports = $reportIndex.Count
$indexHtml = $indexHtml -replace '{GENERATION_TIME}', $generationTime -replace '{TOTAL_REPORTS}', $totalReports

# Write index.html to analysis folder
$indexPath = Join-Path -Path $AnalysisPath -ChildPath "index.html"
$indexHtml | Out-File -FilePath $indexPath -Encoding UTF8

Write-Host "Generated index.html: $indexPath" -ForegroundColor Green
Write-Host "All reports generated successfully!" -ForegroundColor Green

