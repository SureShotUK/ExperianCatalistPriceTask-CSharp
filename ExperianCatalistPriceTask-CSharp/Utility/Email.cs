using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using PortlandCredentials;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Mail;
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8602 // Dereference of a possibly null reference.

namespace ExperianCatalistPriceTask_CSharp.Utility
{
    public class Email
    {
        public static bool DownloadEmailSpreadSheet()
        {
            EmailReceiverCreds creds = Credentials.GetEmailReceiverCreds();
            // Set up the ExchangeService object with your credentials and the URL of the EWS endpoint
            ExchangeService service = new()
            {
                Credentials = new WebCredentials(creds.EmailReceiverUserName, creds.EmailReceiverPassword),
                Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")
            };

            // Define the email subject to search for
            string emailSubject = "Experian Catalist Price Averages";
            // We want to find an Email that has been sent today - so it must be greater than the last minute/second of YESTERDAY to qualify as being TODAY. 
            DateTime yesterday = DateTime.Now.AddDays(-1);
            TimeSpan ts = new(23, 59, 59);
            yesterday = yesterday.Date + ts;

            // Construct the search filter to find emails with the specified subject
            SearchFilter.SearchFilterCollection searchFilterCollection = new(LogicalOperator.And)
            {
                new SearchFilter.ContainsSubstring(ItemSchema.Subject, emailSubject),
                new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, yesterday)
            };

            // Set up the ItemView object to retrieve only the email messages that match the search filter
            ItemView view = new(1)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly)
            };
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            // Use the FindItems method to search for emails that match the search filter
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, searchFilterCollection, view);

            // Loop through the messages and download any attachments
            foreach (Item item in findResults.Items)
            {
                // Bind the item to a new EmailMessage object so that we can access its properties and attachments
                EmailMessage message = EmailMessage.Bind(service, item.Id);

                foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in message.Attachments)
                {
                    // Check if the attachment is a file attachment
                    if (attachment is FileAttachment)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        string attachmentPath = Path.Combine(Directory.GetCurrentDirectory(), "Experian Catalist Price Averages.xlsx");

                        // Download the attachment
                        fileAttachment.Load(attachmentPath);
                        return true;
                    }
                }
            }
            StringBuilderPlusConsole.ErrorEmailBodyBuilderSBOnly("<p>Target Destination: Portland > Prices >  <b>Pump Prices vs Platts.xlsx</b></p> <hr>");
            StringBuilderPlusConsole.ErrorEmailBodyBuilder("No E-mail from Experian for today's date can be found in the inbox of prices@portland-fuel.co.uk.");
            StringBuilderPlusConsole.ErrorEmailBodyBuilder("This is perfectly normal and happens occasionally as the Experian E-mails do not come in daily.");
            StringBuilderPlusConsole.ErrorEmailBodyBuilder("It is recommended to check the inbox of prices@portland-fuel.co.uk, to see if an E-mail <i>was</i> received, in which case this program is having difficulties downloading said E-mail.");
            return false;
        }
        public static void SendEmail(bool error)
        {
            var creds = Credentials.GetEmailCreds();
            MailMessage message = new();
            SmtpClient smtp = new();
            message.From = new MailAddress(creds.Username);
            message.To.Add(new MailAddress("it@portland-fuel.co.uk"));

            if (error == true) { message.Body = StringBuilderPlusConsole.GetErrorLogString().ToString();
                message.Subject = "ERROR: Experian Catalist Price Task Automator";
            }
            else { message.Body = StringBuilderPlusConsole.GetLogString().ToString();
                message.Subject = "Experian Catalist Price Task Automator";
            }
            
            message.IsBodyHtml = true;
            smtp.Host = "smtp-mail.outlook.com";
            smtp.Port = 587;
            smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(creds.Username, creds.Password);
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(message);
        }
    }
}