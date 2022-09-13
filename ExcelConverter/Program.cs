using ExcelBuilder;
using ExcelConverter.App;
using ExcelConverter.Domain;
using ExcelReader;
using MailReceiver;

Console.WriteLine("Started accessing mails");

ReceiveMail receiveMail = new ReceiveMail();

receiveMail.SaveMailAttachments();

Console.WriteLine("Downloaded mail attachments");

ReadOneDayEarly readDayOne = new();

var xd = readDayOne.GetDayOneModels();

int k = 0;

foreach (var item in xd)
{
    k++;
    Console.WriteLine($"Started building OneDayEarly excel number:{k}");
    BuildOneDayEarly buildOneDayEarly = new BuildOneDayEarly();
    buildOneDayEarly.BuildExcel(item);
    Console.WriteLine($"Finished building OneDayEarly excel number:{k}");

    Console.WriteLine($"Started inserting to database OneDayEarly excel number:{k}");
    new InsertOneDayEarlyModel(item);
    Console.WriteLine($"Finished inserting to database OneDayEarly excel number:{k}");

}

ReadTwoDayEarly readTwoDayEarly = new();

var xd2 = readTwoDayEarly.GetTwoDayEarlyModels();

k = 0;

foreach (var item in xd2)
{
    k++;
    Console.WriteLine($"Started building TwoDaysEarly excel number:{k}");
    BuildTwoDayEarly buildTwoDayEarly = new BuildTwoDayEarly();
    buildTwoDayEarly.BuildExcel(item);
    Console.WriteLine($"Finished building TwoDaysEarly excel number:{k}");

    Console.WriteLine($"Started inserting to database TwoDayEarly excel number:{k}");
    new InsertTwoDayEarlyModel(item);
    Console.WriteLine($"Finished inserting to database TwoDayEarly excel number:{k}");
}