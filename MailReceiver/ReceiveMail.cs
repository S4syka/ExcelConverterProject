using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using OpenPop.Mime;
using OpenPop.Pop3;
using ExcelConverter.Domain.Interfaces;

namespace MailReceiver
{
    public class ReceiveMail : IReceiveMail
    {
        public void SaveMailAttachments()
        {
            using (Pop3Client client = new Pop3Client())
            {
                client.Connect(ConfigurationManager.AppSettings["MailServerHostName"], 
                    Int32.Parse(ConfigurationManager.AppSettings["MailServerPort"]), 
                    bool.Parse(ConfigurationManager.AppSettings["MailServerSSL"]));

                client.Authenticate(ConfigurationManager.AppSettings["ReceiverEmail"], 
                    ConfigurationManager.AppSettings["ReceiverPassword"], 
                    AuthenticationMethod.UsernameAndPassword);

                string path = ConfigurationManager.AppSettings["SaveFileAddress"];

                if (client.Connected)
                {
                    int messageCount = client.GetMessageCount();
                    List<Message> allMessages = new List<Message>(messageCount);
                    for (int i = messageCount; i > 0; i--)
                    {
                        allMessages.Add(client.GetMessage(i));
                    }
                    foreach (Message msg in allMessages)
                    {
                        var att = msg.FindAllAttachments();
                        foreach (var ado in att)
                        {
                            if (IsValidAttachment(ado))
                            {
                                ado.Save(new FileInfo(Path.Combine(path, ado.FileName)));
                            }
                        }
                    }
                }
            }
        }

        // TODO: Check attachemnt names for date control.
        private bool IsValidAttachment(MessagePart attachment)
        {
            return true;
        }
    }
}