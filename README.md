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

## Usage

```sh
.\ImportOrdersFromCSV.ps1 -Path ".\ImportExample.csv" -OutputFile ".\Output.csv" -LogFile ".\log.txt" -SettingsFile ".\settings.txt"
```

## Script Behavior

Orders included in the CSV file specified in the "Path" parameter will be placed to STAT. After the script runs, the resulting output file will contain tracking numbers for successfully submitted orders, as well as tracking and label URLs.
