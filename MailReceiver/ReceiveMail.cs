using ExcelConverter.Domain.Interfaces;
using OpenPop.Mime;
using OpenPop.Pop3;
using System.Configuration;

namespace MailReceiver
{
    public class ReceiveMail : IReceiveMail
    {
        public void SaveMailAttachments()
        {
            using (Pop3Client client = new Pop3Client())
            {
                client.Connect(ConfigurationManager.AppSettings["POPMailServerHostName"],
                    Int32.Parse(ConfigurationManager.AppSettings["POPMailServerPort"]),
                    bool.Parse(ConfigurationManager.AppSettings["POPMailServerSSL"]));

                client.Authenticate(ConfigurationManager.AppSettings["ReceiverEmail"],
                    ConfigurationManager.AppSettings["ReceiverPassword"],
                    AuthenticationMethod.UsernameAndPassword);

                if (client.Connected)
                {
                    int messageCount = client.GetMessageCount();
                    List<Message> allMessages = new List<Message>(messageCount);
                    for (int i = messageCount; i > 0; i--)
                    {
                        allMessages.Add(client.GetMessage(i));
                        //client.DeleteMessage(i);
                    }
                    foreach (Message msg in allMessages)
                    {
                        var att = msg.FindAllAttachments();
                        foreach (var ado in att)
                        {
                            if (IsValidAttachment(ado))
                            {
                                if(ado.FileName.Contains("_1"))
                                ado.Save(new FileInfo(Path.Combine(CreateNewDirectory("TwoDaysEarly"), ado.FileName)));
                                if (ado.FileName.Contains("_2"))
                                ado.Save(new FileInfo(Path.Combine(CreateNewDirectory("OneDayEarly"), ado.FileName)));
                            }
                        }
                    }
                }
            }
        }

        // TODO: Check attachemnt names for date control.
        // TODO: Check attachment names and divide in different directories.

        private bool IsValidAttachment(MessagePart attachment)
        {
            return true;
        }

        private string CreateNewDirectory(string typeOfFile)
        {
            string directoryName = DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]);
            string path = ConfigurationManager.AppSettings["SaveFileAddress"]+ @$"\{typeOfFile}" + @"\" + directoryName;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        // TODO: Add method that returns company name.
    }
}