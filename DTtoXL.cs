using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using OfficeOpenXml;
using System.IO;

namespace ConsoleApp1
{
    internal class DTtoXL
    {
        public static void ConvertDT(DataTable dataTable, string excelFile)
        {

            // var excelFile = @"E:\Studies\MCA\Semester 4\Philips\file.xlsx";
            var sheetName = "Sheet1";
            
            using (var package = new ExcelPackage(new System.IO.FileInfo(excelFile)))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];

                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add(sheetName);
                }

                // Write the column names to the worksheet
                for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                // Write the data rows to the worksheet
                for (var rowNumber = 0; rowNumber < dataTable.Rows.Count; rowNumber++)
                {
                    for (var columnNumber = 0; columnNumber < dataTable.Columns.Count; columnNumber++)
                    {
                        worksheet.Cells[rowNumber + 2, columnNumber + 1].Value = dataTable.Rows[rowNumber][columnNumber].ToString();
                    }
                }

                package.Save();
            }

        }
    }
}
