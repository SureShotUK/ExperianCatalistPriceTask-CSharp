using System;
using System.Collections.Generic;
using System.Linq;
//using System.Net;
using PortlandGraphLibrary;
using System.Text;
using System.Threading.Tasks;
using PortlandCredentials;
//using Microsoft.Exchange.WebServices.Data;
//using System.Net.Mail;
using PortlandEmail;
using Microsoft.Graph;
using System.IO;
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8602 // Dereference of a possibly null reference.

namespace ExperianCatalistPriceTask_CSharp.Utility
{
    public class GetEmailAttachment
    {
        public static async Task<bool> DownloadExperianCatalistFileAsync()
        {
            EmailReceiverCreds creds = Credentials.GetEmailReceiverCreds();
            PortlandGraph graph = new(creds.EmailReceiverUserName, creds.EmailReceiverPassword);
            List<Message> messages = await graph.GetEmailsFromPricesAsync(DateTime.Now.AddDays(-14), "steve@portland-fuel.co.uk", "Experian Catalist Price Averages", "DESC");
            foreach (Message message in messages)
            {
                if (message.HasAttachments ??= true)
                {
                    foreach (var attachment in message.Attachments)
                    {
                        {
                            if (attachment is FileAttachment)
                            {
                                string attachmentPath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Experian Catalist Price Averages.xlsx");
                                var fileAttachment = (FileAttachment)attachment;
                                System.IO.File.WriteAllBytes(attachmentPath, fileAttachment.ContentBytes);

                                return true;
                            }
                        }

                    }
                }
            }
            return false;
        }

        public static void SendEmail(bool error)
        {
            var creds = Credentials.GetEmailCreds();

            Email email = new Email(creds.Username, creds.Password);
            email.AddTo("it@portland-fuel.co.uk");

            if (error == true)
            {
                email.WriteLine(StringBuilderPlusConsole.GetErrorLogString().ToString());
                email.AddSubject("ERROR: Experian Catalist Price Task Automator");
            }
            else
            {
                email.WriteLine(StringBuilderPlusConsole.GetLogString().ToString());
                email.AddSubject("Experian Catalist Price Task Automator");
            }
            email.SendEmail();
            
        }
    }
}