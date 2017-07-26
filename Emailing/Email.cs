using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Emailing
{
    public class Email
    {
        public static void SendEmail(string from, string To, string Subject, string Message, string fileName)
        {
            using (MailMessage mail = new MailMessage(from, To))
            {
                mail.Subject = Subject;
                mail.Body = Message;
                if (fileName != string.Empty)
                {
                    mail.Attachments.Add(new Attachment(fileName));
                }
                mail.IsBodyHtml = false;
                SmtpClient smtp = new SmtpClient();
                smtp.Host = ConfigurationManager.AppSettings["SMTPServer"];
                NetworkCredential networkCredential = new NetworkCredential(ConfigurationManager.AppSettings["UserID"], ConfigurationManager.AppSettings["Password"]);
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = networkCredential;
                smtp.Port = int.Parse(ConfigurationManager.AppSettings["Port"]);
                smtp.Send(mail);
            }
        }
    }
}
