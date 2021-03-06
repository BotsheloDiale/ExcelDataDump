# ExcelDataDump
***The Easiest and fast way of creating spreadsheets from your "SQL query" structured data.***

## C# Excel / Spreadsheet Library

### Package main features
 + Creates an Excel file from SQL query structured data sets.
 + Create custom column names, by replacing the quirie's key name with a custom text.
 + Add multiple sheets, sheet meta data in a single workbook.
 + Easy to use and requires minimal code.



### Getting Started

Install [ExcelDataDumpv1.0.0-beta](https://www.nuget.org/packages/ExcelDataDump/1.0.0-beta) and remember to add the namespace.


##### Code Sample

###### creating excel sheets

>Because an excel workbook is esentially a collection of sheet, we adopt the same concept here
>we will require a list (collection)  of sheets.

      ExcelReports report = new ExcelReports();                  // Instatiate a Workbook
      List<WorkSheet> workbook = new List<WorkSheet>();          // Instatatiate sheet list

>Defining a single worksheet properties.

      WorkSheet sheet = new WorkSheet{
        Type = "report",                                         // Keep this as is for future use. when adding other report types.
        Title = "Add title of your report here",                 // Replace with your report title
        Author = "Your name or organisation's name",             // Replace with your or your organisation's name
        SheetName = "Name of sheet tab",                         // Sheet tab name
        SheetData = allDevices,                                  // SQL query data/object
        Headers = null,                                          // Refer to Headers section** (optional).
        Description = "Report/sheet description",                // Replace with your report description
     };
     workbook.Add(sheet);

###### ** sheet.Headers

>> Headers property is of type List\'<\'sheetHeaders\'>\' , where sheetHeaders is of type (string, string)
>> sheetHeaders - (column_Key, replacement_name) 
>>> - "column_key" is the keyname/ table column name from your data set or sql query
>>> - "replacement_name" is the new text value to replace the column_key

>> Creating Headers
>>    - Example -> Replacing two column names with custom,readable and "report friendly" names.

     List<sheetHeaders> headers = new List<sheetHeaders>();            
     headers.Add(new sheetHeaders("CustName", "Customer Name"));       // Original text "CustName" will be replaced with "Customer Name"
     headers.Add(new sheetHeaders("CustID", "Account Reference No.")); // Original text "CustID" will be replaced with "Account Reference No."

> result

    MemoryStream results = report.CreateExcelDocument(workbook);       // Download - handle the return results accordingly (return to UI). 


###### Putting it all togather - Final code sample (Sever-side code)

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

###### UI (Client-side code - Javascript example)

Example of Client side code required to download the generated file.

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


Please share your experience, comments and suggestions. Happy coding!!

