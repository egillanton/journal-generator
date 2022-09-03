# Journal Generator

.NET 6 Console Application that converts statement for international payment from Íslandsbanki online bank to a Payday journal import file.

## Setup

Make sure you have [.NET 6 SDK](https://docs.microsoft.com/en-us/dotnet/core/tutorials/with-visual-studio-code?pivots=dotnet-6-0#:~:text=The-,.NET%206%20SDK.,-Create%20the%20app) installed.


To run the app run in .NET CLI:
```
dotnet run
```

[Further installation guide](https://docs.microsoft.com/en-us/dotnet/core/tutorials/with-visual-studio-code?pivots=dotnet-6-0)


## How to export Islandsbanki international statement for international payment

### Step 1. Table Capture browser extension
Make sure you have Table Capture browser extension installed.
 * [Google Chrome](https://chrome.google.com/webstore/detail/table-capture/iebpjdmgckacbodjpijphcplhebcmeop?hl=en)
 * [Firefox](https://addons.mozilla.org/en-US/firefox/addon/table-capture-firefox/)
 * [Microsoft Edge](https://microsoftedge.microsoft.com/addons/detail/table-capture/cjlemjohnmihejeecaoaglgejaokmclj)


### Step 2. Export the statement using Table Capture
Right click the table and select *Table Capture > Table Capture - Launch workshop* 
![Screen Capture from Islandsbanki !](docs/images/capture_1.png  "Screen Capture from Islandsbanki")


### Step 3 Select Google Spreadsheets
![Screen Capture from Table Capture !](docs/images/capture_2.png  "Screen Capture from Table Capture")


### Step 4. Paste the table in Google Spreadsheets
Paste the table that has been automatically been saved in your clipboard into cell *A1*
![Screen Capture from Google Spreadsheets !](docs/images/capture_3.png  "Screen Capture from Google Spreadsheets")

### Step 5. Delete Column A
![Screen Capture from Google Spreadsheets !](docs/images/capture_4.png  "Screen Capture from Google Spreadsheets")


### Step 6. Make sure you are using European standards in numbers and dates.
Best way to ensure you are using the date and number standards that is supported in Iceland is to go to the file settings and select Germany as your locality.
![Screen Capture from Google Spreadsheets Settings !](docs/images/capture_5.png  "Screen Capture from Google Spreadsheets Settings")


### Example file
You can see how the input file should look like if you open up [Journal.xlsx](./example%20files/Journal.xlsx). 

## Using the application

Run the console application and select the first application to run.

```
Applications
  1: Generate Journal from Islandsbanki Foreign Payments

Enter the number of application to run or press 'x' to exit:
1
Enter file path:
C:\Users\egillanton\Desktop\Journal.xlsx
Successfully generated 24 journal entries to file: C:\Users\egill\Desktop\Journal_ouput.xlsx
```

* Input file: [Journal.xlsx](./example%20files/Journal.xlsx). 
* Exported file: [Journal_output.xlsx](./example%20files/Journal_ouput.xlsx). 

Note that 5 entry lines are generate for each entry.


## Future Improvements
* Have a separate config JSON file for all accounts (as seen in [PaydayJournalUtils](./Utils/PaydayJournalUtils.cs)).
* Add application to export it directly to Payday API when their API allows for it.
* Handle better the accounts that don't appear in the settings  