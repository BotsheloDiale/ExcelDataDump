using System;
using System.IO;
using System.Collections.Generic;
using ExcelDataDump.Models;
using OfficeOpenXml;

namespace ExcelDataDump.Controllers
{
    public class Reports
    {
        private string findColumnHeaderText(List<sheetHeaders> Headers, string KeyName)
        {
            foreach (sheetHeaders item in Headers)
            {
                if (item.ColumnValue == KeyName)
                    return item.ColumnText;
            }
            return KeyName;
        }

        private ExcelWorksheet createTableHeaders(ExcelWorksheet WrkSheet, List<sheetHeaders> headers, dynamic objct)
        {
            var propertyNames = objct[0].GetType().GetProperties();
            int columnIndex = 1, startingRow = 1, columnWidth = 20;
            //Handle column data
            foreach (var key in propertyNames)
            {
                int rowIndex = 1;

                //Make the column wider if the column name is or has 'description'.
                if ((key.Name.ToLower()).Contains("description"))
                    WrkSheet.Column(columnIndex).Width = columnWidth * 4;
                else
                    WrkSheet.Column(columnIndex).Width = columnWidth;
                //Align Cell content to the Center.
                WrkSheet.Column(columnIndex).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                //Make column name UpperCase
                if (headers != null && headers.Count > 0)
                    WrkSheet.Cells[rowIndex++, columnIndex].Value = findColumnHeaderText(headers, key.Name);
                else
                    WrkSheet.Cells[rowIndex++, columnIndex].Value = key.Name.ToUpper();

                //Handle row data
                foreach (var reportEntry in objct)
                {
                    var ItemValue = reportEntry.GetType().GetProperty(key.Name).GetValue(reportEntry, null);
                    WrkSheet.Cells[rowIndex++, columnIndex].Value = ItemValue;
                }
                columnIndex++;
            }
            //Style the header row
            WrkSheet.Row(startingRow).Style.Font.Bold = true;
            WrkSheet.Row(startingRow).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            WrkSheet.Row(startingRow).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            return WrkSheet;
        }

        public ExcelWorksheet GenerateReportSheet(ExcelWorksheet WrkSheet, WorkSheet ReportData)
        {
            try
            {
                return createTableHeaders(WrkSheet, ReportData.Headers, ReportData.SheetData);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private bool mangepage(int num)
        {
            if (num == 1 || num == 3)
            {
                return true;
            }
            return false;
        }

    }

}
