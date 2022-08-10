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
ReadOneDayEarly readDayOne = new();

var xd = readDayOne.GetDayOneDTOs();

foreach(var items in xd)
{
    Console.WriteLine(items.ToString());
}

Console.WriteLine("123");