using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace Streamlining_Customer_Verification
{
    internal class PdfExtract
    {
        public PdfExtract() 
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                DateTime today = DateTime.Today;
                string file = ConfigurationManager.AppSettings["attachDownload"] + today.ToString("MMM yyyy") + "\\" + today.ToString("dd MMM yyyy");
                file += @"\Philips Bill.pdf";
                using (PdfReader reader = new PdfReader(file))
                {
                    for (int pageNo = 1; pageNo <= reader.NumberOfPages; pageNo++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string Text = PdfTextExtractor.GetTextFromPage(reader, pageNo, strategy);
                        Text = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(Text)));
                        sb.Append(Text);
                    }
                }
                Console.WriteLine(sb.ToString());
                string svFile = @"E:\Studies\MCA\Semester 4\Philips\Mail Docs\Apr 2023\30 Apr 2023\test.txt";
                File.WriteAllText(svFile, sb.ToString());
                // Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            
        }
    }
}
