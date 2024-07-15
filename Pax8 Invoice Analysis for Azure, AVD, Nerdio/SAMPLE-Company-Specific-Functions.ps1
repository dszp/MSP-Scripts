# Company-Specific-Functions.ps1

function Normalize-Row-Data () {
    # Normalize company name and ID from one to another to correct some Pax8 billing companies to desired reporting company
    
    # Normalize company name and ID from one to another if the SKU is Nerdio and thus starts with NIO-INF:
    if ($row.sku -like "NIO-INF*") {
        Write-Verbose "ORIGINAL: $script:companyName name and $script:companyID ID using separate script!"
        if ($row.details -like '*Original Company Name 1*') {
            $script:companyName = "New Company 1 To Use"
            $script:companyID = "1234567"
        }
        elseif ($row.details -like '*Original Company Name 2*') {
            $script:companyName = "New Company 2 To Use"
            $script:ompanyID = "2345678"
        }
        Write-Verbose "NEW:      $script:companyName name and $script:companyID ID using separate script!"

        Write-Verbose "---------------------------------------------------------------"
        Write-Verbose "NERDIO SKU:          $($row.sku)"
        Write-Verbose "NERDIO COMPANY ID:   $($row.company_id)"
        Write-Verbose "NERDIO ORIG COMPANY: $($row.company_name)"
        Write-Verbose "NERDIO COMPANY NAME: $companyName"
        Write-Verbose "NERDIO SUBTOTAL:     $subtotal"
        Write-Verbose "NERDIO COST:         $costTotal"
        Write-Verbose "NERDIO PERIOD:       $startPeriod"
        Write-Verbose "NERDIO DESC:         $($row.details)"
        Write-Verbose "NERDIO DESC:         $($row.description)"
    }

    # Normalize company name and ID from one to another for any CloudJumpber (CJR-) SKU licenses:
    if ($row.sku -like "CJR-*") {
        Write-Verbose "ORIGINAL: $script:companyName name and $script:companyID ID using separate script!"
        $script:companyName = "Original Company Name 3"
        $script:companyID = "3456789"
        Write-Verbose "NEW:      $script:companyName name and $script:companyID ID using separate script!"
    }
}
