using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Streamlining_Customer_Verification
{
    internal class XLtoDT
    {
        public static DataTable ConvertXl(string excelFile)
        {
            // var excelFile = @"E:\Studies\MCA\Semester 4\Philips\Bill Columns.xlsx";
            var sheetName = "Sheet1";
            var dataTable = new DataTable();

            // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new System.IO.FileInfo(excelFile)))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];


                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text.Trim());
                }

                for (var rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    var newRow = dataTable.Rows.Add();
                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text;
                    }
                }

                // Process the data table here
                Console.WriteLine("DataTable has " + dataTable.Rows.Count + " rows.");
            }
            return dataTable;
        }
    }
}
