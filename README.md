# ExcelDataDump

## Package main features
 + Creates an Excel file from SQL quiry structured data sets.
 + Allows users to create custom column names, by replacing the quirie's key name with a custom string.
 + Easily create and name multiple sheets in a single workbook.
 + Easy to use, Simplest way to create excel reports from SQL quiries with minimal code.

## Getting Started

    Download free beta package and add reference to the [ExcellDataDump.dll](.\release\netstandard2.0\ExcellDataDump.dll) in the release folder - .\release\netstandard2.0\
##### >>>OR 
    Install [ExcelDataDumpv1.0.0-beta](https://www.nuget.org/packages/ExcelDataDump/1.0.0-beta) 

##### Code Sample

###### creating excel sheets

>>Because an excel workbook is esentially a collection of sheet, we adopt the same concept here
>>we will require a list (collection)  of sheets.

      ExcelReports report = new ExcelReports();                  // Instatiate a Workbook
      List<WorkSheet> workbook = new List<WorkSheet>();          // Instatatiate sheet list

>>Defining a single worksheet properties.

    ~~~~
      WorkSheet sheet = new WorkSheet{
        Type = "report",                                         // Keep this as is for future use. when adding other report types.
        Title = "Add title of your report here",                 // Replace with your report title
        Author = "Your name or organisation's name",             // Replace with your or your organisation's name
        SheetName = "Name of sheet tab",                         // Sheet tab name
        SheetData = allDevices,                                  // SQL quiry data/object
        Headers = null,                                          // Refer to Headers section** (optional).
        Description = "Report/sheet description",                // Replace with your report description
     };
     workbook.Add(sheet);
    ~~~~

###### ** sheet.Headers

>>> Headers property is of type <pre>List<sheetHeaders></pre> , where sheetHeaders is of type (string, string)
>>> sheetHeaders - (column_Key, replacement_name) 
>>>> - "column_key" is the keyname/ table column name from your data set or sql quiry
>>>> - "replacement_name" is the new text value to replace the column_key

>>> Creating Headers
>>>     - Replacing two column name with custom,readable and "report friendly" names
~~~
     List<sheetHeaders> headers = new List<sheetHeaders>();            
     headers.Add(new sheetHeaders("CustName", "Customer Name"));       // Original text "CustName" will be replaced with "Customer Name"
     headers.Add(new sheetHeaders("CustID", "Account Reference No.")); // Original text "CustID" will be replaced with "Account Reference No."
~~~
###### result

~~~
    MemoryStream results = report.CreateExcelDocument(workbook);  // Download - handle the return results accordingly (return to UI). 
~~~

###### Final code sample (API Content)

~~~
        // Instatiate a Workbook and sheets
        ExcelReports report = new ExcelReports();
        List<WorkSheet> workbook = new List<WorkSheet>();

        // creating sheet data and meta data
        WorkSheet sheet = new WorkSheet{
            Type = "report",
            Title = "Add title of your report here",
            Author = "Your name or organisation's name",
            SheetName = "Name of sheet tab",
            SheetData = allDevices,
            Headers = null,
            Description = "Report/sheet description",
        };
        workbook.Add(sheet);
        
        // Adding custom headers
        List<sheetHeaders> headers = new List<sheetHeaders>();
        headers.Add(new sheetHeaders("CustName", "Customer Name"));
        headers.Add(new sheetHeaders("CustID", "Account Reference No."));

        // Handle results (return back to UI)
        MemoryStream results = report.CreateExcelDocument(workbook);
        return results;
~~~

###### UI (Client side code - Javascript example)

Client side code required to download the generated file.

~~~
    function exportToExcel(){

        var xhr = new XMLHttpRequest();
        xhr.open("POST", "<https://replaceWithYourApiURL>", true);                // Replace <https://replaceWithYourApiURL> with API URL
        xhr.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
        //Handle Error
        xhr.onerror = function() {
            console.log(`Error during the upload: ${xhr.status}`);
        };
        //Handle Success
        xhr.responseType = "blob";
        xhr.onload = function(e) {
            if (this.status == 200) {
                var blob = this.response;
                saveAs(blob, "<MyExcelFile>.xlsx");                               // Replace <MyExcelFile> with file name.
            }
        };
        xhr.send(JSON.stringify(self.feedbackData));
    }
~~~


