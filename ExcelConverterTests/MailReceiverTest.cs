using MailReceiver;

namespace ExcelConverterTests
{
    [TestClass]
    public class MailReceiverTest
    {
        private ReceiveMail _receiveMail; 

        [TestMethod]
        public void TestSaveMailAttachments()
        {
            _receiveMail = new ReceiveMail();
            _receiveMail.SaveMailAttachments();

            int fCount = Directory.GetFiles(@"C:\Users\99559\Desktop\aq", "*", SearchOption.TopDirectoryOnly).Length;
            Assert.AreEqual(fCount, 0);
        }
    }
}