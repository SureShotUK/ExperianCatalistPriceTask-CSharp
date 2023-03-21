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
        //public static bool DownloadEmailSpreadSheet()
        //{
        //    EmailReceiverCreds creds = Credentials.GetEmailReceiverCreds();
        //    // Set up the ExchangeService object with your credentials and the URL of the EWS endpoint
        //    ExchangeService service = new()
        //    {
        //        Credentials = new WebCredentials(creds.EmailReceiverUserName, creds.EmailReceiverPassword),
        //        Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")
        //    };

        //    // Define the email subject to search for
        //    string emailSubject = "Experian Catalist Price Averages";
        //    // We want to find an Email that has been sent today - so it must be greater than the last minute/second of YESTERDAY to qualify as being TODAY. 
        //    DateTime yesterday = DateTime.Now.AddDays(-1);
        //    TimeSpan ts = new(23, 59, 59);
        //    yesterday = yesterday.Date + ts;

        //    // Construct the search filter to find emails with the specified subject
        //    SearchFilter.SearchFilterCollection searchFilterCollection = new(LogicalOperator.And)
        //    {
        //        new SearchFilter.ContainsSubstring(ItemSchema.Subject, emailSubject),
        //        new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, yesterday)
        //    };

        //    // Set up the ItemView object to retrieve only the email messages that match the search filter
        //    ItemView view = new(1)
        //    {
        //        PropertySet = new PropertySet(BasePropertySet.IdOnly)
        //    };
        //    view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

        //    // Use the FindItems method to search for emails that match the search filter
        //    FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, searchFilterCollection, view);

        //    // Loop through the messages and download any attachments
        //    foreach (Item item in findResults.Items)
        //    {
        //        // Bind the item to a new EmailMessage object so that we can access its properties and attachments
        //        EmailMessage message = EmailMessage.Bind(service, item.Id);

        //        foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in message.Attachments)
        //        {
        //            // Check if the attachment is a file attachment
        //            if (attachment is FileAttachment)
        //            {
        //                FileAttachment fileAttachment = attachment as FileAttachment;
        //                string attachmentPath = Path.Combine(Directory.GetCurrentDirectory(), "Experian Catalist Price Averages.xlsx");

        //                // Download the attachment
        //                fileAttachment.Load(attachmentPath);
        //                return true;
        //            }
        //        }
        //    }
        //    StringBuilderPlusConsole.ErrorEmailBodyBuilderSBOnly("<p>Target Destination: Portland > Prices >  <b>Pump Prices vs Platts.xlsx</b></p> <hr>");
        //    StringBuilderPlusConsole.ErrorEmailBodyBuilder("No E-mail from Experian for today's date can be found in the inbox of prices@portland-fuel.co.uk.");
        //    StringBuilderPlusConsole.ErrorEmailBodyBuilder("This is perfectly normal and happens occasionally as the Experian E-mails do not come in daily.");
        //    StringBuilderPlusConsole.ErrorEmailBodyBuilder("It is recommended to check the inbox of prices@portland-fuel.co.uk, to see if an E-mail <i>was</i> received, in which case this program is having difficulties downloading said E-mail.");
        //    return false;
        //}

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