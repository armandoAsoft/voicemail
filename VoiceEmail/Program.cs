using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;

namespace VoiceEmail
{
    class Program
    {
        
        static void Main(string[] args)
        {
            ExchangeService exchange = null;
            String dir = "C:\\TempAttachment\\";
            String fileExtension = ".PDF";
            
            foreach(string email in ConfigurationSettings.AppSettings.AllKeys)
            {
                String username = email;
                String userDomain = email;
                String password = ConfigurationSettings.AppSettings.Get(email);

                try
                {
                    exchange = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                    exchange.Credentials = new WebCredentials(username, password);
                    exchange.AutodiscoverUrl(userDomain, RedirectionCallback);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error connecting " + ex);
                }

                if (exchange != null)
                {
                    SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(ItemSchema.HasAttachments, true);
                    FindItemsResults<Item> findResults = exchange.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(10));

                    foreach (Item item in findResults)
                    {
                        EmailMessage message = EmailMessage.Bind(exchange, item.Id);

                        foreach (Attachment attachment in message.Attachments)
                        {
                            if (attachment is FileAttachment)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                if (attachment.Name.ToUpper().Contains(fileExtension))
                                {
                                    Console.WriteLine("Time: " + message.DateTimeReceived.ToString());
                                    Console.WriteLine("From: " + message.From.Name.ToString());
                                    Console.WriteLine("From address: " + message.From.Address.ToString());
                                    Console.WriteLine("Subject: " + message.Subject);
                                    Console.WriteLine("Attach Name: " + attachment.Name.ToString());
                                    System.IO.Directory.CreateDirectory(dir);
                                    fileAttachment.Load(dir + fileAttachment.Name);
                                    Console.WriteLine("------------------------------------------------");
                                }
                            }
                        }
                    }
                }
            }
        }

        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }
    }
}
