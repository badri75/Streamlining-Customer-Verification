using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Streamlining_Customer_Verification
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Console.WriteLine("Meow");
            Mail mail = new Mail();
            if (mail.getMail())
                Console.WriteLine("New Attachments found today");
            Console.ReadLine();
        }
    }
}
