using System;
using Microsoft.Exchange.WebServices.Data;

namespace EWS
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Iniciado o EWS...");

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials("mche20091990@outlook.com", "!14#38#7");
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            Mailbox mb = new Mailbox("mche20091990@outlook.com");
            FolderId fid = new FolderId(WellKnownFolderName.Inbox, mb);
            // The search filter to get unread email.

            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)); 
            ItemView view = new ItemView(100);

            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);
            

            foreach (var item in findResults)
            {
                Console.WriteLine(item.Subject);
                Console.WriteLine(item.Flag.FlagStatus);   

                EmailMessage message = EmailMessage.Bind(service, item.Id);

                foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in message.Attachments)
                {
                    if (attachment is FileAttachment) {
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        // Load the file attachment into memory and print out its file name.
                        fileAttachment.Load();
                        Console.WriteLine("Attachment name: " + fileAttachment.Name);

                    }
                }

                
                message.IsRead = true;
                message.Update(ConflictResolutionMode.AlwaysOverwrite);

            }
            Console.WriteLine("end");
            // Console.ReadLine();
        }
    }
}
