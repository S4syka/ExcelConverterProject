// See https://aka.ms/new-console-template for more information
using ExcelBuilder;
using ExcelConverter.Domain.DTO;
using ExcelReader;
using MailReceiver;

//Console.WriteLine("Hello, World!");

//ReceiveMail receiveMail = new ReceiveMail();

//receiveMail.SaveMailAttachments();
/*
SendMail sendMail = new SendMail();

sendMail.SendMailAttachments("e.khomasuridzemails@yahoo.com","test123");
*/

//ReadOneDayEarly readDayOne = new();

//var xd = readDayOne.GetDayOneDTOs();

////Console.WriteLine(new BuildOneDayEarly().BuildExcel(null));
//foreach(var item in xd)
//{
//    BuildOneDayEarly buildOneDayEarly = new BuildOneDayEarly();
//    buildOneDayEarly.BuildExcel(item);
//}

ReadTwoDayEarly readTwoDayEarly = new();

var xd2 = readTwoDayEarly.GetTwoDayEarlyDTOs();

foreach(var item in xd2)
{
    BuildTwoDayEarly buildTwoDayEarly = new BuildTwoDayEarly();
    buildTwoDayEarly.BuildExcel(item);
}
Console.WriteLine("123");