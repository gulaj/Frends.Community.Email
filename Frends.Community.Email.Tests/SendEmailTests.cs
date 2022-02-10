using NUnit.Framework;
using System.IO;

namespace Frends.Community.Email.Tests
{   
     /// <summary>
     /// NOTE: To run these unit tests, you need an SMTP test server. Fill in the properties below with your values.
     /// </summary>
    [TestFixture]
    [Ignore("Tests requires external SMTP server to work")]
    public class SendEmailTests
    {
        // ****************************************** FILL THESE ******************************************************
        private const string USERNAME = "";
        private const string PASSWORD = "";
        private const string SMTPADDRESS = "";
        private const string TOEMAILADDRESS = "";
        private const string FROMEMAILADDRESS = "";
        private const int PORT = 465;
        private const bool USESSL = true;
        private const bool ACCEPTALLCERTS = false;
        // ************************************************************************************************************


        private const string TEMP_ATTACHMENT_SOURCE = "emailtestattachments";
        private const string TEST_FILE_NAME = "testattachment.txt";
        private const string TEST_FILE_NOT_EXISTING = "doesntexist.txt";

        private string _localAttachmentFolder;
        private string _filepath;
        private Input _input;
        private Options _options;

        [SetUp]
        public void EmailTestSetup()
        {
            _localAttachmentFolder = Path.Combine(Path.GetTempPath(), TEMP_ATTACHMENT_SOURCE);

            if (!Directory.Exists(_localAttachmentFolder))
                Directory.CreateDirectory(_localAttachmentFolder);

            _filepath = Path.Combine(_localAttachmentFolder, TEST_FILE_NAME);

            if (!File.Exists(_filepath))
                using (File.Create(_filepath)) { }

            _input = new Input()
            {
                From = FROMEMAILADDRESS,
                To = TOEMAILADDRESS,
                Cc = "",
                Bcc = "",
                Message = "testmsg",
                IsMessageHtml = false,
                SenderName = "EmailTestSender",
                MessageEncoding = "utf-8"
            };

            _options = new Options()
            {
                UserName = USERNAME,
                Password = PASSWORD,
                SMTPServer = SMTPADDRESS,
                Port = PORT,
                UseSsl = USESSL,
                AcceptAllCerts = ACCEPTALLCERTS
            };

        }
        [TearDown]
        public void EmailTestTearDown()
        {
            if (Directory.Exists(_localAttachmentFolder))
                Directory.Delete(_localAttachmentFolder, true);
        }

        [Test]
        public void SendEmailWithPlainText()
        {
            var input = _input;
            input.Subject = "Email test - PlainText";
            
            var result = EmailTask.SendEmail(input, null, _options, new System.Threading.CancellationToken());
            Assert.IsTrue(result.EmailSent);
        }

        [Test]
        public void SendEmailWithFileAttachment()
        {
            var input = _input;
            input.Subject = "Email test - FileAttachment";
            
            var attachment = new Attachment
            {
                FilePath = _filepath,
                SendIfNoAttachmentsFound = false,
                ThrowExceptionIfAttachmentNotFound = true
            };


            var Attachments = new Attachment[] { attachment };

            var result = EmailTask.SendEmail(input, Attachments, _options, new System.Threading.CancellationToken());
            Assert.IsTrue(result.EmailSent);
        }

        [Test]
        public void SendEmailWithStringAttachment()
        {
            var input = _input;
            input.Subject = "Email test - AttachmentFromString";
            var fileAttachment = new AttachmentFromString() { FileContent = "teststring ä ö", FileName = "testfilefromstring.txt" };
            var attachment = new Attachment()
            {
                AttachmentType = AttachmentType.AttachmentFromString,
                stringAttachment = fileAttachment
            };
            var Attachments = new Attachment[] { attachment};

            var result = EmailTask.SendEmail(input, Attachments, _options, new System.Threading.CancellationToken());
            Assert.IsTrue(result.EmailSent);
        }

        [Test]
        public void TrySendingEmailWithNoFileAttachmentFound()
        {
            var input = _input;
            input.Subject = "Email test";
            
            var attachment = new Attachment
            {
                FilePath = Path.Combine(_localAttachmentFolder, TEST_FILE_NOT_EXISTING),
                SendIfNoAttachmentsFound = false,
                ThrowExceptionIfAttachmentNotFound = false
            };


            var Attachments = new Attachment[] { attachment };

            var result = EmailTask.SendEmail(input, Attachments, _options, new System.Threading.CancellationToken());
            Assert.IsFalse(result.EmailSent);
        }

        [Test]
        public void TrySendingEmailWithNoFileAttachmentFoundException()
        {
            var input = _input;
            input.Subject = "Email test";
            
            var attachment = new Attachment
            {
                FilePath = Path.Combine(_localAttachmentFolder, TEST_FILE_NOT_EXISTING),
                SendIfNoAttachmentsFound = false,
                ThrowExceptionIfAttachmentNotFound = true
            };


            var Attachments = new Attachment[] { attachment };

            Assert.Throws<FileNotFoundException>(() => EmailTask.SendEmail(input, Attachments, _options, new System.Threading.CancellationToken()));
            
        }
    }
}
