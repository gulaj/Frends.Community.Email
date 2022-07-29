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
    public class ExchangeEmailTests
    {
        private readonly string _mailbox = "inbox";
        private readonly string _username = "frends_exchange_test_user@frends.com";
        private readonly string _password = Environment.GetEnvironmentVariable("Exchange_User_Password");
        private readonly string _applicationID = Environment.GetEnvironmentVariable("Exchange_Application_ID");
        private readonly string _tenantID = Environment.GetEnvironmentVariable("Exchange_Tenant_ID");

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldReadOneItemTest()
        {
            var subject = "One Email Test";
            await SendTestEmail(subject);
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.That(result.Count, Is.EqualTo(1));
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldReadFiveItemsTest()
        {
            var subject = "Five Emails test";
            for (int i = 0; i < 5; i++)
                await SendTestEmail(subject);
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 5,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.That(result.Count, Is.EqualTo(5));
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldGetUnreadMailsTest()
        {
            var subject = "Get Unread Emails Test";
            await SendTestEmail(subject);
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = true,
                MarkEmailsAsRead = true,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 1);
            result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 0);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldMarkEmailAsReadTest()
        {
            var subject = "Mark Email As Read Test";
            await SendTestEmail(subject);
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = true,
                MarkEmailsAsRead = true,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 1);
            result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 0);
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
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
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
            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result[0].AttachmentSaveDirs.Count == 1);
            Assert.IsTrue(result[0].AttachmentSaveDirs[0] == fullpath);
            Assert.IsTrue(File.Exists(fullpath));
            Directory.Delete(dirPath, true);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldFilterBySenderTest()
        {
            var subject = "Filter By Sender Test";
            await SendTestEmail(subject);
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSenderFilter = _username
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 1);
            options.EmailSenderFilter = "SomeRandom@foobar.com";
            result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 0);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldFilterBySubjectTest()
        {
            var subject = "Filter By Subject Test";
            await SendTestEmail(subject);
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 1);
            options.EmailSubjectFilter = "Some Random Subject";
            result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 0);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldGetOnlyEmailsWithAttachmentsTest()
        {
            var dirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../OnlyAttachmentData/");
            var subject = "Get Only Emails With Attachments Test";
            await SendTestEmailWithAttachment(subject, "OnlyAttachment.txt");
            await SendTestEmail(subject);

            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
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
            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(result.Count == 1);
            Directory.Delete(dirPath, true);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task ReadEmailFromExchangeServer_ShouldOverwriteAttachmentsTest()
        {
            var dirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../OverwriteData/");
            var subject = "Overwrite Attchment Test";
            await SendTestEmailWithAttachment(subject, "OverwriteAttachment.txt");
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password,
                Mailbox = _mailbox
            };
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
            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(File.Exists(result[0].AttachmentSaveDirs[0]));
            Assert.IsTrue(Directory.GetFiles(dirPath).Length == 1);
            result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            Assert.IsTrue(File.Exists(result[0].AttachmentSaveDirs[0]));
            Assert.IsTrue(Directory.GetFiles(dirPath).Length == 1);
            Directory.Delete(dirPath, true);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithPlainTextTest()
        {
            var input = new ExchangeInput
            {   
                To = _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = "Email test - PlainText"
            };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailWithPlainTextToMultipleUsingSemicolonTest()
        {
            var input = new ExchangeInput
            {
                To = _username + ";" + _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = "Email test - MultipleSemiColon"
            };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailWithPlainTextToMultipleUsingColonTest()
        {
            var input = new ExchangeInput
            {
                To = _username + "," + _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = "Email test - MultipleColon"
            };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailWithCCTest()
        {
            var input = new ExchangeInput
            {
                To = _username,
                Cc = _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = "Email test - CC"
            };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailWithBCCTest()
        {
            var input = new ExchangeInput
            {
                To = _username,
                Bcc = _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = "Email test - BCC"
            };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailWithHtmlTest()
        {
            var input = new ExchangeInput
            {
                To = _username,
                Message = "<div><h1>This is a header text.</h1></div>",
                IsMessageHtml = true,
                MessageEncoding = "utf-8",
                Subject = "Email test - HTML"
            };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailSwitchEncodingTest()
        {
            var input = new ExchangeInput
            {
                To = _username,
                Message = "The letter Ä changes to ?.",
                IsMessageHtml = false,
                MessageEncoding = "ascii",
                Subject = "Email test - ASCII"
            };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailWithFileAttachmentTest()
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../attachmentFile.txt");
            File.WriteAllText(filePath, "This is a test attachment file.");
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has a file attachment.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = "Email test - FileAttachment"
            };

            var attachment = new Attachment
            {
                AttachmentType = AttachmentType.FileAttachment,
                FilePath = filePath,
                ThrowExceptionIfAttachmentNotFound = false,
                SendIfNoAttachmentsFound = false
            };

            var attachmentArray = new Attachment[] { attachment };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, attachmentArray, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            File.Delete(filePath);
            await DeleteMessages(input.Subject);
        }

        [Test]
        public async Task SendEmailWithStringAttachmentTest()
        {
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has an attachment written from a string.",
                IsMessageHtml = false,
                MessageEncoding = "utf-8",
                Subject = "Email test - StringAttachment"
            };

            var stringAttachment = new AttachmentFromString
            {
                FileContent = "This is a test attachment from string.",
                FileName = "attachmentFile.txt"
            };

            var attachment = new Attachment
            {
                AttachmentType = AttachmentType.AttachmentFromString,
                StringAttachment = stringAttachment
            };

            var attachmentArray = new Attachment[] { attachment };

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, attachmentArray, server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            await DeleteMessages(input.Subject);
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

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            await EmailTask.SendEmailToExchangeServer(input, null, server, new CancellationToken());
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

            var server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };

            await EmailTask.SendEmailToExchangeServer(input, attachmentArray, server, new CancellationToken());
            File.Delete(filePath);
            Thread.Sleep(5000); // Give the email some time to get through.
        }

        #endregion
    }
}
