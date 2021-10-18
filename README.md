![Alt text](/img/logo.jpg?raw=true "STAT Overnight Delivery")

# ImportOrdersFromCSV

Create STAT Orders using a CSV File

This script was created for temporary use for submitting orders to STAT Overnight Delivery while a permanent integration is being developed. NOTE: This script is delivered as-is and is intended for limited use only.

## Key Files and Folders

* ImportOrdersFromCsv.ps1: The PowerShell script that performs most core functions
* (folder)Watch: Target folder for dragging/dropping CSV files for import using the "watched folder" method (read more below)
* settings.txt: Contains API key, price set and collection location that the script needs in order to run
* WatchFolder.ps1: Creates and registers an event handler using an instance of the FileSystemWatcher class; called from a scheduled task (only needed for using "watched-folder" method of importing orders; see "Usage > Option 2" section below)

## Script Parameters

These are parameters that can be used in running the import script. Note: These parameters are not available when using the "watched-folder" method.

* Path: Location of CSV file containing order data to post
* OutputFile: (OPTIONAL) Location of CSV file returned by the script containing tracking number, tracking URL, label URL and description field of each order. If this parameter is excluded, the file will be saved as "Output.csv" in the current working directory.
* LogFile: (OPTIONAL) File containing info and error logging data. If this parameter is excluded, the file will write events to "log.txt" in the current working directory.
* SettingsFile: (OPTIONAL) JSON-formatted file containing configuration information required by the script, including API key, price set identifier and collection location. If this parameter is excluded, the script will use the "settings.txt" file in the current working directory.

## Set API Key, Priceset Identifier and Collection Location in settings.txt!!

IMPORTANT! Before you can use this script, you must edit the settings file by adding the API key and priceset identifier. Both values may be obtained upon request from STAT. Additionally, set CollectionLocation as appropriate.

## Instructions for Use
---

### Option 1 - Manual import from command line

Use the following syntax to run the script manually:

```sh
.\ImportOrdersFromCSV.ps1 -Path ".\ImportExample.csv" -OutputFile ".\Output.csv" -LogFile ".\log.txt" -SettingsFile ".\settings.txt"
```

The OutputFile, LogFile and SettingsFile parameters are optional. 

  **- OR -**

### Option 2

You can generate orders by dropping your import CSV into a watched folder. This method requires slightly more setup on the front-end, but once configured it is easier than the manual method.

**Configure and Initiate the Watched Folder (only needed for Option 2)**

**NOTE** -- Since the watched-folder method does not support script arguments, we recommend using **C:\STATUtilities\CSVOrderImport** as the root location for the script and other project files/folders. This will prevent having to change a number of variables in the script files to reference a different location.

In order to use this method for importing data from a CSV, you must first create a task using Task Scheduler. The process for that change is detailed below.

Follow these steps to create a scheduled task:

1.  Open Windows Task Scheduler

2.  Create a new Task

![Alt text](/img/Task1.png?raw=true)

3.  On the General tab, name the task, then select "Run whether user is logged on or not" and "Run with highest privileges." 

![Alt text](/img/Task2.png?raw=true)

4.  Click the Triggers tab and select New...

5.  Under "Begin the task," select "At log on"

![Alt text](/img/Task4.png?raw=true)

6.  Click OK

7.  Click the Actions tab and select New...

8.  Under Action, select "Start a Program", then:
  (a)  Enter **powershell.exe** in the Program/Script field;
  (b)  Enter **-executionpolicy bypass -File "C:\STATUtilities\CSVOrderImport\WatchFolder.ps1"** in the Add Arguments field; and
  (c)  Click OK

![Alt text](/img/Task7.png?raw=true)

9.  Click the Settings tab and DESELECT "Stop the task if it runs longer than..."

![Alt text](/img/Task8.png?raw=true)

10.  Click OK

11.  If prompted, enter the password for the Windows user account and click OK

![Alt text](/img/Task10.png?raw=true)

12.  The scheduled task is now ready and will start automatically on the next reboot. You can start it manually by right-clicking the task and selecting "Run."

![Alt text](/img/Task9.png?raw=true)


## Script Behavior

The script will attempt to post a new order to STAT for each line in the import file. After the script is finished, the resulting output file will contain tracking numbers for successfully submitted orders, as well as tracking/label URLs and address fields for the order. *If there are any problems submitting one or more orders, the script will create a second CSV file (default name is ERRORS.csv) containing information on the orders that were rejected, along with a detailed error message for each.*

## Creating Shipping Labels
---

### Option 1: Word Mail Merge

The included merge file (MergeForm.docx) may be used for creating and printing shipping labels by using the output CSV file as the merge data source. This template is included for convenience and may be edited as needed (or you may use another merge template entirely). It is critical that, in addition to containing complete address information, each label MUST include a barcode to be a valid package. The included sample merge template demonstrates how that is accomplished using merge field codes.

### Option 2: LabelURL Field in Output File

The output file that the script generates includes a column called "LabelURL." This URL points to a PDF file that can be downloaded and used as the shipping label for that specific order.

### Option 3: Develop a Custom Script

You can use any number of tools available to read the contents of the output CSV file to generate labels. It is critical that, in addition to containing complete address information, each label MUST include a barcode to be a valid package. The included sample merge template demonstrates how that is accomplished using merge field codes.

## Sample Screenshots
---

### Import CSV File

![Alt text](/img/ImportFile.png?raw=true)

### Settings File

![Alt text](/img/Settings.png?raw=true)

### Script Output / Merge Data File

![Alt text](/img/OrderOutput.png?raw=true)

### Merge Template

![Alt text](/img/MergeDocImage2.png?raw=true)
![Alt text](/img/MergeDocImage.png?raw=true)



