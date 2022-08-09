using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using ExcelConverter.Domain.Interfaces;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;

namespace MailSender
{
    public class SendMail : ISendMail
    {
        public void SendMailAttachments(string receiver, string subject)
        {
            string path = ConfigurationManager.AppSettings["SaveFileAddress"];
            string directoryName = DateTime.Now.ToString("dd-MM-yyyy");
            path = path + @"\" + directoryName;

            Attachment data = new Attachment(path);

            ContentDisposition disposition = data.ContentDisposition;
            disposition.CreationDate = System.IO.File.GetCreationTime(path);
            disposition.ModificationDate = System.IO.File.GetLastWriteTime(path);
            disposition.ReadDate = System.IO.File.GetLastAccessTime(path);

            string from = ConfigurationManager.AppSettings["e.khomasuridzemails@yahoo.com"];
            MailMessage message = new MailMessage(from, receiver);

            message.Subject = subject;

            message.Attachments.Add(data);

            SmtpClient client = new SmtpClient();

            client.Port = Int32.Parse(ConfigurationManager.AppSettings["SMTPMailServerPort"]);
            client.Host = ConfigurationManager.AppSettings["SMTPMailServerHost"];
            client.EnableSsl = bool.Parse(ConfigurationManager.AppSettings["SMTPMailServerSSL"]);
                
            client.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                throw new Exception("ar gaigzavna meili");
            }

            // TODO: Need to archive directory or send each file individually.
        }
    }
}
