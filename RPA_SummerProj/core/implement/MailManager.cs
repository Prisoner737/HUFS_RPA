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
            try
            {
                _ns = _app.GetNamespace("MAPI");
                inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                _ns.SendAndReceive(true);
                foreach (Outlook.MailItem item in inbox.Items)
                {
                    if(item.UnRead == true)
                    {
                        Console.WriteLine("Subject : " + item.Subject);
                        Console.WriteLine("Sender Name : " + item.SenderName);
                        Console.WriteLine("Body : " + item.HTMLBody);
                        Console.WriteLine("Data : " + item.SentOn.ToLongDateString() + " " + item.SentOn.ToLongTimeString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void deleteByName(string name)
        {
            try
            {
                _ns = _app.GetNamespace("MAPI");
                inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                _ns.SendAndReceive(true);
                foreach (Outlook.MailItem item in inbox.Items)
                {
                    if (item.SenderName == name)
                    {
                        item.Delete();
                        Console.WriteLine("Subject : " + item.Subject + "from : " + item.SenderName + "is deleted");
                    }
                        
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void deleteBySubjectKeywords(string subj)
        {
            try
            {
                _ns = _app.GetNamespace("MAPI");
                inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                _ns.SendAndReceive(true);
                foreach (Outlook.MailItem item in inbox.Items)
                {
                    if (item.Subject.Contains(subj))
                    {
                        Console.WriteLine("Subject : " + item.Subject + "from : " + item.SenderName + "is deleted");
                        item.Delete();

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void moveMailFolder(string folderName, string targetMail)
        {
            Outlook.MAPIFolder inBox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items items = (Outlook.Items)inBox.Items;
            Outlook.MailItem moveMail = null;
            items.Restrict("[UnRead] = true");
            Outlook.MAPIFolder destFolder = inBox.Folders[folderName];
            foreach (object eMail in items)
            {
                try
                {
                    moveMail = eMail as Outlook.MailItem;
                    if (moveMail != null)
                    {
                        string titleSubject = (string)moveMail.Subject;
                        if (titleSubject.IndexOf(targetMail) > 0)
                        {
                            moveMail.Move(destFolder);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
