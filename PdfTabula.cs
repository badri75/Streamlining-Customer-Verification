using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Tabula.Extractors;
using UglyToad.PdfPig;
using Tabula;
using System.Configuration;
using Tabula.Detectors;

namespace Streamlining_Customer_Verification
{
    internal class PdfTabula
    {
        public static Table TableCon(string file) 
        {
            /*DateTime today = DateTime.Today;
            string file = ConfigurationManager.AppSettings["attachDownload"] + today.ToString("MMM yyyy") + "\\" + today.ToString("dd MMM yyyy");
            file += @"\Philips Bill - 1.1.pdf";*/
            using (PdfDocument document = PdfDocument.Open(file, new ParsingOptions() { ClipPaths = true }))
            {
                ObjectExtractor oe = new ObjectExtractor(document);
                PageArea page = oe.Extract(1);

                // detect canditate table zones
                SimpleNurminenDetectionAlgorithm detector = new SimpleNurminenDetectionAlgorithm();
                var regions = detector.Detect(page);

                IExtractionAlgorithm ea = new BasicExtractionAlgorithm();
                List<Table> tables = ea.Extract(page.GetArea(regions[0].BoundingBox)); // take first candidate area
                var table = tables[0];
                var rows = table.Rows;
                // var cell = table.Cells;
                // Console.WriteLine(table.Cells[0]);
                /*foreach ( var cell in table.Cells ) 
                {
                    // Console.WriteLine( cell );
                    string st = cell.ToString();
                    if(st.Contains("Customer Name"))
                        Console.WriteLine(st.Substring(13));
                }
                Console.WriteLine(table.GetType());*/
                return table;
            }
            // Console.ReadLine();
        }
    }
}
