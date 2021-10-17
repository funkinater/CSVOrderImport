# ImportOrdersFromCSV
Create STAT Orders with a CSV File

This script was created for tempoary use for submitting orders to STAT Overnight Delivery while a permanent integration is being developed. NOTE: This script is delivered as-is and is intended for limited use only.

## Parameter Definitions

- Path: Location of CSV file containing order data to post
- OutputFile: Location of CSV file returned by the script containing tracking number, tracking URL, label URL and description field of each order
- LogFile: File containing info and error logging data
- SettingsFile: JSON-formatted file containing configuration information required by the script, including:
-- ApiKey: STAT-Provided API key (required to place orders)
-- PriceSet: STAT-Provided service level identifier
-- CollectionLocation: Address that will be used as the collection address for each order placed to the API

## Set API Key and Priceset Identifier

Before you can start using this script, you must edit the settings file by adding the API key and priceset identifier. Both values may be obtained upon request from STAT.

## Usage

```sh
.\ImportOrdersFromCSV.ps1 -Path ".\ImportExample.csv" -OutputFile ".\Output.csv" -LogFile ".\log.txt" -SettingsFile ".\settings.txt"
```

## Script Behavior

Orders included in import file (the "Path" parameter) will be placed to STAT via REST API. After the script runs, the resulting output file will contain tracking numbers for successfully submitted orders, as well as tracking/label URLs and address fields for the order. 

>Note: The included merge file (MergeForm.docx) may be used for creating and printing shipping labels. This form is included for convenience and may be edited as desired (or you may use another method entirely for creating labels).

## Screenshots

### Import CSV File

![Alt text](/img/ImportFile.png?raw=true "Optional Title")

### Settings File

![Alt text](/img/SettingsFile.png?raw=true "Optional Title")

### Script Output / Merge Data File

![Alt text](/img/OrderOutput.png?raw=true "Optional Title")

### Merge Template

![Alt text](/img/MergeDocImage2.png?raw=true "Optional Title")
![Alt text](/img/MergeDocImage.png?raw=true "Optional Title")



