using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Streamlining_Customer_Verification
{
    internal class crtBase
    {
        public static void createBase(string destFile)
        {
            string baseFile = ConfigurationManager.AppSettings["baseFile"];

            // Copy the source file to the destination file
            File.Copy(baseFile, destFile, true);

            // Create a new Excel file
            using (var package = new ExcelPackage(new FileInfo(destFile)))
            {
                var workbook = package.Workbook;

                // Clone each worksheet in the source file
                using (var sourcePackage = new ExcelPackage(new FileInfo(baseFile)))
                {
                    var sourceWorkbook = sourcePackage.Workbook;

                    foreach (var sourceWorksheet in sourceWorkbook.Worksheets)
                    {
                        var destinationWorksheet = workbook.Worksheets.Add(sourceWorksheet.Name);
                        destinationWorksheet.Cells[1, 1].LoadFromArrays((IEnumerable<object[]>)sourceWorksheet.Cells.Value);
                    }
                }

                // Save the changes to the new file
                package.Save();
            }
        }
    }
}
