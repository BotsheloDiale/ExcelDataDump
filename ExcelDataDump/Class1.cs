using System;
using System.IO;
using System.Collections.Generic;
using ExcelDataDump.Controllers;
using ExcelDataDump.Models;
using OfficeOpenXml;

namespace ExcelDataDump
{
    public class ExcelReports
    {
        ///<summary>
        /// Class Excess point/ Entry point constructor. May be used as an API.
        ///</summary>
        public MemoryStream CreateExcelDocument(List<WorkSheet> Sheets)
        {
            try
            {
                var stream = new MemoryStream();
                ExcelPackage package = new ExcelPackage(stream);
                foreach (WorkSheet sheet in Sheets)
                {
                    //Assign Worksheet Name
                    ExcelWorksheet wsSheet = package.Workbook.Worksheets.Add(sheet.SheetName);
                    //Generates a simple report in a table structure.
                    if ((sheet.Type).ToLower() == "report")
                    {
                        Reports report = new Reports();
                        wsSheet = report.GenerateReportSheet(wsSheet, sheet);
                    }
                    //Change Page Orientation if requested.
                    if (sheet.Landscape)
                        wsSheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    // Force Sheet data to fit on a single page
                    if (sheet.FitToPage)
                        wsSheet.PrinterSettings.FitToPage = true;
                }

                package.Save();
                // stream.Close();
                stream.Position = 0;

                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }

}
