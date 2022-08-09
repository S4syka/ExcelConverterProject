// See https://aka.ms/new-console-template for more information
using MailReceiver;
using MailSender;
using ExcelReader;

Console.WriteLine("Hello, World!");

ReceiveMail receiveMail = new ReceiveMail();

receiveMail.SaveMailAttachments();
/*
SendMail sendMail = new SendMail();

sendMail.SendMailAttachments("e.khomasuridzemails@yahoo.com","test123");
*/
ReadDayOne readDayOne = new(@"C:\Users\Rabbitt\Downloads\20220809_Portfolio_BG-GEORGIAN BUILDING GROUP_1.xlsx");

var xd = readDayOne.Items();

Console.WriteLine("123");