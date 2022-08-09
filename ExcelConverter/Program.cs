// See https://aka.ms/new-console-template for more information
using MailReceiver;
using MailSender;

Console.WriteLine("Hello, World!");

ReceiveMail receiveMail = new ReceiveMail();

receiveMail.SaveMailAttachments();

SendMail sendMail = new SendMail();

sendMail.SendMailAttachments("e.khomasuridzemails@yahoo.com","test123");