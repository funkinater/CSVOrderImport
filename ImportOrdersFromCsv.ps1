param(
    $WorkingDir = ".",
    $Path = "$WorkingDir\ImportExample.csv",
    $OutputFile = "$WorkingDir\Output.csv",
    $OrderErrors = "$WorkingDir\ERRORS.csv",
    $LogFile="$WorkingDir\log.txt",
    $SettingsFile = "$WorkingDir\settings.txt"
    )


Write-Host "Script executed. Settings file is $SettingsFile"

if(-not($Path) -or -not($OutputFile) -or -not($SettingsFile)) {Throw "You must supply a valid import, output and settings file."}

$createOrderUrl = "https://api.statcourierservice.com/api/Orders"

$settings = Get-Content $SettingsFile -Raw | ConvertFrom-Json

$apiKey = $settings.apiKey
$priceSet = $settings.priceSet


class STATOrder {
    [string]$RequestedBy
    [string]$Description
    [string]$Comments
    [string]$ReferenceNumber
    [string]$PriceSet
    [Location]$CollectionLocation
    [Location]$DeliveryLocation
    [string]$CollectionSignatureRequired
    [string]$DeliverySignatureRequired
    [string]$SuppressUserNotification
}

class Location {
    [string]$ContactName
    [string]$CompanyName
    [string]$AddressLine1
    [string]$AddressLine2
    [string]$City
    [string]$State
    [string]$PostalCode
    [string]$Country
    [string]$Comments
    [string]$Email
    [string]$Phone
    [string]$Category
    [string]$LatitudeLongitude
}

class OrderReturnData {
    [string]$TrackingNumber
    [string]$TrackingURL
    [string]$LabelURL
    [string]$Description
    [string]$ContactName
    [string]$CompanyName
    [string]$AddressLine1
    [string]$AddressLine2
    [string]$City
    [string]$State
    [string]$PostalCode
    [string]$Country
    [string]$Comments
    [string]$Email
    [string]$Phone

}

class OrderError {
    [STATOrder]$Order
    [string]$ErrorMessage
}

[System.Collections.ArrayList]$Orders = @{}
[System.Collections.ArrayList]$STATOrderResponses = @{}
[System.Collections.ArrayList]$ERRORS = @{}

Function AddToLog {
    param(
        [string]$level,
        [string]$message
    )

    Add-Content $logFile "$(date)  $level  $message`r`n"

    if($level -eq "INFO") {
        Write-Host "$level $message" -ForegroundColor Green
    } else {
        Write-Host "$level $message" -ForegroundColor Red
    }
}

Function AddOrderToArray {
    param(
        [Parameter(Mandatory)]
        $OrderDetails
    )

    try {

        $Order = [STATOrder]::new()
        $CLoc = [Location]::new()
        $DLoc = [Location]::new()

        $Order.RequestedBy = $OrderDetails.RequestedBy
        $Order.Description = $OrderDetails.Description
        $Order.Comments = $OrderDetails.Comments
        $Order.PriceSet = $priceSet
        $Order.SuppressUserNotification = "false"
        $Order.CollectionSignatureRequired = "false"
        $Order.ReferenceNumber = $OrderDetails.ReferenceNumber
        $Order.DeliverySignatureRequired = if($OrderDetails.DeliverySignatureRequired -eq "yes") {"true"} Else {"false"}

        $Cloc.ContactName = $settings.CollectionLocation.ContactName
        $Cloc.CompanyName = $settings.CollectionLocation.CompanyName
        $Cloc.AddressLine1 = $settings.CollectionLocation.AddressLine1
        $Cloc.AddressLine2 = $settings.CollectionLocation.AddressLine2
        $Cloc.City = $settings.CollectionLocation.City
        $Cloc.State = $settings.CollectionLocation.State
        $Cloc.PostalCode = $settings.CollectionLocation.PostalCode
        $Cloc.Country = $settings.CollectionLocation.Country
        $Cloc.Comments = $settings.CollectionLocation.Comments
        $Cloc.Email = $settings.CollectionLocation.Email
        $Cloc.Phone = $settings.CollectionLocation.Phone
        $Cloc.Category = $settings.CollectionLocation.Category

        $DLoc.ContactName = $OrderDetails.DeliveryContactName
        $DLoc.CompanyName = $OrderDetails.DeliveryCompanyName
        $DLoc.AddressLine1 = $OrderDetails.DeliveryAddressLine1
        $DLoc.AddressLine2 = $OrderDetails.DeliveryAddressLine2
        $DLoc.City = $OrderDetails.DeliveryCity
        $DLoc.State = $OrderDetails.DeliveryState
        $DLoc.PostalCode = $OrderDetails.DeliveryPostalCode
        $DLoc.Country = $OrderDetails.DeliveryCountry
        $DLoc.Comments = $OrderDetails.DeliveryComments
        $DLoc.Email = $OrderDetails.DeliveryEmail
        $DLoc.Phone = $OrderDetails.DeliveryPhone
        $DLoc.Category = $OrderDetails.DeliveryCategory

        $Order.CollectionLocation = $CLoc
        $Order.DeliveryLocation = $DLoc

        $Orders.Add($Order)
    }
    catch {
        AddToLog -level "ERROR" -message "Unable to add $($OrderDetails.DeliveryContactName) ($($OrderDetails.DeliveryCompanyName)) order to collection`r`n ($_.Exception)`r`n"
    }
}

Function PostOrderToSTAT {
    param(
        [Parameter(Mandatory)]
        $Order
    )

    $orderContactDescription = ""

    if (($Order.DeliveryLocation.ContactName) -eq "" -or ($null -eq $Order.DeliveryLocation.ContactName)) {
        $orderContactDescription = $Order.DeliveryLocation.CompanyName
    } else {
        $orderContactDescription = $Order.DeliveryLocation.ContactName
    }
    
    if (($Order.Description) -eq "" -or ($null -eq $Order.Description)) {
        $orderContactDescription += "($($Order.Description))"
    }

    AddToLog -level "INFO" -message "Posting order to STAT for $orderContactDescription..."

    $OrderJson = $Order | ConvertTo-Json -Depth 10
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

    $headers.Add("ApiKey", $apiKey)

    try {

        $response = Invoke-WebRequest -Uri $createOrderUrl -Method POST -Body $OrderJson -ContentType "application/json" -Headers $headers

        if($response.StatusCode -eq 200) {
        
            $content = [OrderReturnData]::new()
            $rValue = $response.Content | ConvertFrom-Json
            $location = $rValue.DeliveryLocation

            $content.TrackingNumber = $rValue.TrackingNumber
            $content.TrackingURL = $rValue.TrackingURL
            $content.LabelURL = $rValue.BarcodeURL
            $content.Description = $rValue.Description
            $content.ContactName = $location.ContactName
            $content.CompanyName = $location.CompanyName
            $content.AddressLine1 = $location.AddressLine1
            $content.AddressLine2 = $location.AddressLine2
            $content.City = $location.City
            $content.State = $location.State
            $content.PostalCode = $location.PostalCode
            $content.Country = $location.Country
            $content.Comments = $location.Comments
            $content.Email = $location.Email
            $content.Phone = $location.Phone

            $STATOrderResponses.Add($content)

            AddToLog -level "INFO" -message "Tracking number $($content.TrackingNumber) created for $orderContactDescription."
        } else {
            AddToLog -level "WARNING" -message "Unable to submit order. Response received from server: $($response.StatusCode) $($response.StatusDescription)"
        }
    }
    catch {
        [string]$errorMessage
        if($null -ne $_.ErrorDetails.Message) {
            $errorMessage = ($_.ErrorDetails.Message | ConvertFrom-Json).Message
        } else {
            $errorMessage = "Unknown error"
        } 
        
        $error = [OrderError]::new()
        $error.Order = $Order
        $error.ErrorMessage = $errorMessage
        $ERRORS.Add($error)
        
        AddToLog -level "ERROR" -message "Could not create order for $($orderContactDescription) -- $errorMessage"
    }
}


$logStartText = @"


**************************************************

IMPORT SESSION STARTED

"@

$logStartText += "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))`r`n`r`n**************************************************"

Write-Host $logStartText
Add-Content -Path $LogFile -Value $logStartText

AddToLog -level "INFO" -message "Gathering order details..."

Import-Csv -Path $Path | ForEach-Object { AddOrderToArray -OrderDetails $_ } | Out-Null

$Orders | ForEach-Object { PostOrderToSTAT -Order $_ } | Out-Null

if($STATOrderResponses.Count -gt 0) {
    AddToLog -level "INFO" -message "Exporting results data to $Outputfile..."
    $STATOrderResponses | Select-Object * | Export-Csv -Path $OutputFile -NoTypeInformation
}

If($ERRORS.Count -gt 0) {
    AddToLog -level "INFO" -message "Exporting ERROR data to results data to $OrderErrors..."
    ($ERRORS | ConvertTo-Json -Compress) | ConvertFrom-Json | Export-Csv -Path $OrderErrors -NoTypeInformation
}

AddToLog -level "INFO" -message "Script completed with $($ERRORS.Count) error(s)."


$folderName = (Get-Date).ToString("yyyyMMdd_hh-mm-ss")
$CompletedFolder = New-Item -ItemType Directory -Path "$WorkingDir\JobHistory\Completed" -Name $folderName -Force

try {
    Move-Item -Path $Path -Destination $CompletedFolder -Force
    Move-Item -Path $LogFile -Destination $CompletedFolder -Force
    Copy-Item -Path $OutputFile -Destination $CompletedFolder -Force
    Move-Item -Path $OrderErrors -Destination $CompletedFolder -Force
}
catch {
    AddToLog("ERROR","Unable to move/copy files to completed folder")
}

