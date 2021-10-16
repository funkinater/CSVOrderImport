﻿param(
    $Path = ".\ImportExample.csv",
    $OutputFile = ".\Output.csv",
    $LogFile=".\log.txt",
    $SettingsFile = ".\settings.txt"
    )


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
}

[System.Collections.ArrayList]$Orders = @{}
[System.Collections.ArrayList]$STATOrderResponses = @{}

Function AddToLog {
    param(
        [string]$level,
        [string]$message
    )

    Add-Content $logFile "$(date)  $level  $message"

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

    if (($Order.DeliveryLocation.ContactName) -eq "" -or ($Order.DeliveryLocation.ContactName -eq $null)) {
        $orderContactDescription = $Order.DeliveryLocation.CompanyName
    } else {
        $orderContactDescription = $Order.DeliveryLocation.ContactName
    }
    
    if (($Order.Description) -eq "" -or ($Order.Description -eq $null)) {
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

            $content.TrackingNumber = $rValue.TrackingNumber
            $content.TrackingURL = $rValue.TrackingURL
            $content.LabelURL = $rValue.BarcodeURL
            $content.Description = $rValue.Description

            $STATOrderResponses.Add($content)

            AddToLog -level "INFO" -message "Tracking number $($content.TrackingNumber) created for $orderContactDescription."
        } else {
            AddToLog -level "WARNING" -message "Unable to submit order. Response received from server: $($response.StatusCode) $($response.StatusDescription)"
        }
    }
    catch {
        AddToLog -level "ERROR" -message "Exception occurred trying to post request to STAT server: $($_.Exception)"
    }
}

AddToLog -level "INFO" -message "Gathering order details..."

Import-Csv -Path $Path | ForEach-Object { AddOrderToArray -OrderDetails $_ } | Out-Null

$Orders | ForEach-Object { PostOrderToSTAT -Order $_ } | Out-Null

AddToLog -level "INFO" -message "Exporting results data to $Outputfile..."
$STATOrderResponses | Select-Object * | Export-Csv -Path $OutputFile -NoTypeInformation

AddToLog -level "INFO" -message "Script complete."

ii $OutputFile