using Azure.Identity;
using Microsoft.Graph;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Directory = System.IO.Directory;
using File = System.IO.File;
using Path = System.IO.Path;

namespace Frends.Community.Email.Tests
{
    [TestFixture]
    public class ReadExchangeEmailTests
    {
        private readonly string _mailbox = "inbox";
        private readonly string _username = "frends_exchange_test_user@frends.com";
        private readonly string _password = Environment.GetEnvironmentVariable("Exchange_User_Password");
        private readonly string _applicationID = Environment.GetEnvironmentVariable("Exchange_Application_ID");
        private readonly string _tenantID = Environment.GetEnvironmentVariable("Exchange_Tenant_ID");
        private static ExchangeSettings _settings;
        private static ExchangeServer _server;

        [OneTimeSetUp]
        public void OneTimeSetup()
        {
            if (string.IsNullOrEmpty(_applicationID) ||
                string.IsNullOrEmpty(_password) ||
                string.IsNullOrEmpty(_tenantID))
            {
                throw new ArgumentException("Password, Application ID or Tenant ID is missing. Please check environment variables.");
            }
            _settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };

            _server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldReadOneItemTest()
        {
            var subject = "One Email Test";
            await SendTestEmail(subject);
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 1);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldReadFiveItemsTest()
        {
            var subject = "Five Emails test";
            for (int i = 0; i < 5; i++)
                await SendTestEmail(subject);

            var options = new ExchangeOptions
            {
                MaxEmails = 3,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.That(result.Count, Is.EqualTo(3));
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldGetUnreadMailsTest()
        {
            var subject = "Get Unread Emails Test";
            await SendTestEmail(subject);

            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = true,
                MarkEmailsAsRead = true,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 1);
            result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 0);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldMarkEmailAsReadTest()
        {
            var subject = "Mark Email As Read Test";
            await SendTestEmail(subject);

            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = true,
                MarkEmailsAsRead = true,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 1);
            result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 0);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldGetAttachmentTest()
        {
            var dirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../AttachmentData/");
            var filename = "GetAttachmentTest.txt";
            var fullpath = Path.Combine(dirPath, filename);
            var subject = "Get Attachment Test";
            await SendTestEmailWithAttachment(subject, filename);

            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = false,
                AttachmentSaveDirectory = dirPath
            };

            Directory.CreateDirectory(dirPath);
            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result[0].AttachmentSaveDirs.Count, 1);
            Assert.AreEqual(result[0].AttachmentSaveDirs[0], fullpath);
            Assert.IsTrue(File.Exists(fullpath));
            Directory.Delete(dirPath, true);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldFilterBySenderTest()
        {
            var subject = "Filter By Sender Test";
            await SendTestEmail(subject);

            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSenderFilter = _username
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 1);
            options.EmailSenderFilter = "SomeRandom@foobar.com";
            result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 0);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldFilterBySubjectTest()
        {
            var subject = "Filter By Subject Test";
            await SendTestEmail(subject);

            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 1);
            options.EmailSubjectFilter = "Some Random Subject";
            result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 0);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldGetOnlyEmailsWithAttachmentsTest()
        {
            var dirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../OnlyAttachmentData/");
            var subject = "Get Only Emails With Attachments Test";
            var fullpath = Path.Combine(dirPath, "OnlyAttachment.txt");
            await SendTestEmailWithAttachment(subject, "OnlyAttachment.txt");
            await SendTestEmail(subject);

            // There has been some hicups where the attachment already exists.
            // This step will make sure that it doesn't.
            if (File.Exists(fullpath)) File.Delete(fullpath);

            var options = new ExchangeOptions
            {
                MaxEmails = 2,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = false,
                GetOnlyEmailsWithAttachments = true,
                AttachmentSaveDirectory = dirPath,
                EmailSubjectFilter = subject
            };

            Directory.CreateDirectory(dirPath);
            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.AreEqual(result.Count, 1);
            Directory.Delete(dirPath, true);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldOverwriteAttachmentsTest()
        {
            var dirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../OverwriteData/");
            var subject = "Overwrite Attchment Test";
            await SendTestEmailWithAttachment(subject, "OverwriteAttachment.txt");

            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = false,
                GetOnlyEmailsWithAttachments = true,
                AttachmentSaveDirectory = dirPath,
                OverwriteAttachment = true
            };

            Directory.CreateDirectory(dirPath);
            var result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.IsTrue(File.Exists(result[0].AttachmentSaveDirs[0]));
            Assert.AreEqual(Directory.GetFiles(dirPath).Length, 1);
            result = await ReadEmailTask.ReadEmailFromExchangeServer(_settings, options, new CancellationToken());
            Assert.IsTrue(File.Exists(result[0].AttachmentSaveDirs[0]));
            Assert.AreEqual(Directory.GetFiles(dirPath).Length, 1);
            Directory.Delete(dirPath, true);
            await DeleteMessages(subject);
        }

        #region HelperMethods

        private async Task DeleteMessages(string subject)
        {
            Thread.Sleep(5000); // Give some time for emails to get through before deletion.
            var options = new List<QueryOption>
            {
                new QueryOption("$search", "\"subject:" + subject + "\"")
            };
            var credentials = new UsernamePasswordCredential(_username, _password, _tenantID, _applicationID);
            var graph = new GraphServiceClient(credentials);
            var messages = await graph.Me.Messages.Request(options).GetAsync();
            foreach (var message in messages)
                await graph.Me.Messages[message.Id].Request().DeleteAsync();
        }

        private async Task SendTestEmail(string subject)
        {
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = subject
            };

            await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Thread.Sleep(2000); // Give the email some time to get through.
        }

        private async Task SendTestEmailWithAttachment(string subject, string filename)
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../..", filename);
            File.WriteAllText(filePath, "This is a test attachment file.");
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has a file attachment.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = subject
            };

            var attachment = new Attachment
            {
                AttachmentType = AttachmentType.FileAttachment,
                FilePath = filePath,
                ThrowExceptionIfAttachmentNotFound = false,
                SendIfNoAttachmentsFound = false
            };

            var attachmentArray = new Attachment[] { attachment };

            await EmailTask.SendEmailToExchangeServer(input, attachmentArray, _server, new CancellationToken());
            File.Delete(filePath);
            Thread.Sleep(5000); // Give the email some time to get through.
        }

        #endregion
    }
}
