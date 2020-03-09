using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataDump.Models
{

    public class WorkSheet
    {
        public string Type { get; set; }
        public string Title { get; set; }
        public string Author { get; set; }
        public string SheetName { get; set; }
        public dynamic SheetData { get; set; }
        public string Description { get; set; }
        public List<sheetHeaders> Headers { get; set; }
        public bool Landscape { get; set; } = false;
        public bool FitToPage { get; set; } = false;
    }

    public class sheetHeaders
    {
        public string ColumnValue { get; set; }
        public string ColumnText { get; set; }
        public sheetHeaders(string val, string text)
        {
            this.ColumnValue = val;
            this.ColumnText = text;
        }
    }
}
