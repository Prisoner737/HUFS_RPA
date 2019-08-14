using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RPA_SummerProj.core.implement
{
    class MailManager
    {
        Outlook._Application _app;
        Outlook.MailItem mail;
        Outlook._NameSpace _ns;
        Outlook.MAPIFolder inbox;


        public MailManager()
        {
            _app = new Outlook.Application();
            mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
        }

        public void sendMail(string to, string subject, string text)
        {
            try
            {
                mail.To = to;
                mail.Subject = subject;
                mail.Body = text;
                mail.Importance = Outlook.OlImportance.olImportanceNormal;
                ((Outlook._MailItem)mail).Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            } 
        }


        public void receiveMail()
        {
            int i = 0;
            try
            {
                _ns = _app.GetNamespace("MAPI");
                inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                _ns.SendAndReceive(true);
                foreach (Outlook.MailItem item in inbox.Items)
                {
                    Console.WriteLine("Subject : " + item.Subject);
                    Console.WriteLine("Sender Name : " + item.SenderName);
                    Console.WriteLine("Body : " + item.HTMLBody);
                    Console.WriteLine("Data : " + item.SentOn.ToLongDateString() + " " + item.SentOn.ToLongTimeString());
                    i++;

                    if (i == 10)
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


    }
}
