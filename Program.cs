using ConsoleApp1;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Outlook;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Tabula;
using Exception = System.Exception;

namespace Streamlining_Customer_Verification
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                DateTime today = DateTime.Today;
                string folderPath = ConfigurationManager.AppSettings["attachDownload"] + today.ToString("MMM yyyy") + "\\" + today.ToString("dd MMM yyyy") + "\\";

                // Console.WriteLine("Meow");
                /*Mail mail = new Mail();
                if (mail.getMail())
                    Console.WriteLine("New Attachments found today");*/
                // PdfExtract pdf = new PdfExtract();
                // PdfTabula pdf = new PdfTabula();

                string tdXl = folderPath + today.ToString("dd MMM yyyy ") + "Data.xlsx";
                string tdBill = folderPath + today.ToString("dd MMM yyyy ") + "Bill.xlsx";
                string baseFile = ConfigurationManager.AppSettings["baseFoldPath"] + ConfigurationManager.AppSettings["baseFile"];
                if (!File.Exists(tdXl))
                {
                    // crtBase.createBase(tdXl);
                    File.Copy(baseFile, tdXl, true);
                }

                // Converting Eligibility file to DataTable
                string eliFile = ConfigurationManager.AppSettings["baseFoldPath"] + ConfigurationManager.AppSettings["eliFile"];
                DataTable eligDt = new DataTable();
                eligDt = XLtoDT.ConvertXl(eliFile);

                foreach (var filePath in Directory.GetFiles(folderPath, "*.pdf"))
                {
                    // Process the PDF file and storing it in Tabula's Table
                    var fileName = Regex.Match(filePath, @"[^\\/]+$").Value;
                    Console.WriteLine("\nProcessing PDF: " + fileName);
                    Tabula.Table tbl = null;
                    tbl = PdfTabula.TableCon(folderPath + fileName);

                    // Converting today's XL File to DataTable
                    DataTable dt = new DataTable();
                    dt = XLtoDT.ConvertXl(tdXl);

                    // Converting today's bill which was extracted from SAP to DataTable
                    DataTable billDt = new DataTable();
                    billDt = XLtoDT.ConvertXl(tdBill);

                    // Validating the Table and concatenating to the existing DataTable
                    dt = Validation.DataValid(tbl, dt);
                    dt = Validation.CheckElig(dt, billDt, eligDt);

                    // Convert DataTable to XL
                    DTtoXL.ConvertDT(dt, tdXl);
                }
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}

/*Alt + O, C, A
 * check if today's xl exists,
	if not,
        create today's xl by cloning the template and continue the next step
	if yes,
        convert today's xl to dt and add new rows*/
