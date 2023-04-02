using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace Streamlining_Customer_Verification
{
    internal class Mail
    {
        /*public Mail() 
        {
            Console.WriteLine("Check");
        }*/
        public bool getMail()
        {
            Console.WriteLine("getMail started");
            bool retResult = false;
            // Create an instance of the Outlook application
            Application outlook = new Application();

            // Get the inbox folder
            MAPIFolder inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            // Search for a particular unread email
            string searchSubject = "Test email";
            string[] searchSender = ConfigurationManager.AppSettings["email"].Split(';');
            string searchBody = "This is a test email";

            DateTime today = DateTime.Today;
            TimeSpan customTime = new TimeSpan(10, 0, 0); // 10:00 AM
            // DateTime receivedTime = today.Add(customTime);
            DateTime receivedTime = new DateTime(2023, 3, 18, 10, 0, 0);
            // string searchDate = "18-03-2023";

            // Check for Directory
            string directoryPath = @"E:\Studies\MCA\Semester 4\Philips\Mail Docs\" + today.ToString("MMM yyyy") + "\\" + today.ToString("dd MMM yyyy");
            // Check if directory exists, if not create directory
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            foreach (string sender in searchSender)
            {
                string filter = $"[Unread] = false AND [Subject] = '{searchSubject}' AND [SenderEmailAddress] = '{sender}' AND [ReceivedTime] >= '{receivedTime.ToString("g")}'";
                // Console.WriteLine(filter);

                try
                {
                    Items results = inbox.Items.Restrict(filter);

                    // Loop through the search results
                    foreach (MailItem mail in results)
                    {
                        // Download the attachments
                        for (int i = 1; i <= mail.Attachments.Count; i++)
                        {
                            Attachment attachment = mail.Attachments[i];
                            string attachmentPath = Path.Combine(directoryPath, attachment.FileName);
                            attachment.SaveAsFile(attachmentPath);
                            // Console.WriteLine("Downloaded attachment: " + attachmentPath);
                        }

                        // Mark the email as read
                        mail.UnRead = false;
                        mail.Save();

                        retResult = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlook);
            outlook = null;
            Console.WriteLine("getMail ended");
            return retResult;
        }
    }
}
