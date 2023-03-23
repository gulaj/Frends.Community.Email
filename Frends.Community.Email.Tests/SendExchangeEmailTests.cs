﻿using Azure.Identity;
using Microsoft.Graph;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using File = System.IO.File;
using Path = System.IO.Path;

namespace Frends.Community.Email.Tests
{
    [TestFixture]
    public class SendExchangeEmailTests
    {
        private readonly string _username = "frends_exchange_test_user@frends.com";
        private readonly string _password = Environment.GetEnvironmentVariable("Exchange_User_Password");
        private readonly string _applicationID = Environment.GetEnvironmentVariable("Exchange_Application_ID");
        private readonly string _tenantID = Environment.GetEnvironmentVariable("Exchange_Tenant_ID");
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

            _server = new ExchangeServer
            {
                Username = _username,
                Password = _password,
                AppId = _applicationID,
                TenantId = _tenantID
            };
        }

        [Test]
        public async Task SendEmailWithPlainTextTest()
        {
            var subject = "Email test - PlainText";
            var message = "This is a test message from Frends.Commmunity.Email Unit Tests.";
            var input = new ExchangeInput
            {
                To = _username,
                Message = message,
                IsMessageHtml = false,
                Subject = subject
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.AreEqual(email[0].BodyText, message);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailToMultipleUsingSemicolonTest()
        {
            var subject = "Email test - MultipleSemiColon";
            var input = new ExchangeInput
            {
                To = _username + ";" + _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                Subject = subject
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.AreEqual(_username + ", " + _username, email[0].To);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailToMultipleUsingCommaTest()
        {
            var subject = "Email test - MultipleColon";
            var input = new ExchangeInput
            {
                To = _username + "," + _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                Subject = subject
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.AreEqual(_username + ", " + _username, email[0].To);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithCCTest()
        {
            var subject = "Email test - CC";
            var input = new ExchangeInput
            {
                To = _username,
                Cc = _username,
                Message = "This is a test message from Frends.Commmunity.Email Unit Tests.",
                IsMessageHtml = false,
                Subject = subject
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.AreEqual(_username, email[0].Cc);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithHtmlTest()
        {
            var subject = "Email test - HTML";
            var message = "<div><h1>This is a header text.</h1></div>";

            var input = new ExchangeInput
            {
                To = _username,
                Message = message,
                IsMessageHtml = true,
                Subject = subject
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.IsTrue(email[0].BodyHtml.Contains(message));
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailNordicLettersTest()
        {
            var subject = "Email test - ääööåå";
            var input = new ExchangeInput
            {
                To = _username,
                Message = "Tämä testimaili tuo yöllä ålannista.",
                IsMessageHtml = false,
                Subject = subject
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.AreEqual("Tämä testimaili tuo yöllä ålannista.", email[0].BodyText);
            Assert.AreEqual(subject, email[0].Subject);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithFileAttachmentTest()
        {
            var subject = "Email test - FileAttachment";
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../attachmentFile.txt");
            File.WriteAllText(filePath, "This is a test attachment file.");
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has a file attachment.",
                IsMessageHtml = false,
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

            var result = await EmailTask.SendEmailToExchangeServer(input, attachmentArray, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            File.Delete(filePath);
            Thread.Sleep(2000); // Give the email some time to get through.
            await ReadTestEmailWithAttachment(subject);
            Assert.IsTrue(File.Exists(filePath));
            File.Delete(filePath);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithByteArrayAttachmentTest()
        {
            var subject = "Email test - ByteArrayAttachment";
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has a byte array attachment.",
                IsMessageHtml = false,
                Subject = subject
            };

            var byteArrayAttachment = new AttachmentFromByteArray
            {
                FileBuffer = new byte[16 * 1024],
                FileName = "fileAttachment.txt",
            };

            var attachment = new Attachment
            {
                AttachmentType = AttachmentType.AttachmentFromByteArray,
                ByteArrayAttachment =byteArrayAttachment,
                ThrowExceptionIfAttachmentNotFound = false,
                SendIfNoAttachmentsFound = false
            };

            var attachmentArray = new Attachment[] { attachment };

            var result = await EmailTask.SendEmailToExchangeServer(input, attachmentArray, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            await ReadTestEmailWithAttachment(subject);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithBase64StringAttachmentTest()
        {
            var subject = "Email test - Base64StringAttachment";
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has a base64 string attachment.",
                IsMessageHtml = false,
                Subject = subject
            };

            var base64StringAttachment = new AttachmentFromBase64String
            {
                Content = Convert.ToBase64String(new byte[16 * 1024]),
                FileName = "fileAttachment.txt",
            };

            var attachment = new Attachment
            {
                AttachmentType = AttachmentType.AttachmentFromBase64String,
                Base64StringAttachment = base64StringAttachment,
                ThrowExceptionIfAttachmentNotFound = false,
                SendIfNoAttachmentsFound = false
            };

            var attachmentArray = new Attachment[] { attachment };

            var result = await EmailTask.SendEmailToExchangeServer(input, attachmentArray, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            await ReadTestEmailWithAttachment(subject);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithBigFileAttachmentTest()
        {
            var subject = "Email test - BigFileAttachment";
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../BigAttachmentFile.txt");

            // Write 9MB file.
            var stream = new FileStream(filePath, FileMode.CreateNew);
            stream.Seek(9L * 1024, SeekOrigin.Begin);
            stream.WriteByte(0);
            stream.Close();

            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has a file attachment.",
                IsMessageHtml = false,
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

            var result = await EmailTask.SendEmailToExchangeServer(input, attachmentArray, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            File.Delete(filePath);
            Thread.Sleep(2000); // Give the email some time to get through.
            await ReadTestEmailWithAttachment(subject);
            Assert.IsTrue(File.Exists(filePath));
            File.Delete(filePath);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithStringAttachmentTest()
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../stringAttachmentFile.txt");
            var subject = "Email test - StringAttachment";
            var input = new ExchangeInput
            {
                To = _username,
                Message = "This email has an attachment written from a string.",
                IsMessageHtml = false,
                Subject = subject
            };

            var stringAttachment = new AttachmentFromString
            {
                FileContent = "This is a test attachment from string.",
                FileName = "stringAttachmentFile.txt"
            };

            var attachment = new Attachment
            {
                AttachmentType = AttachmentType.AttachmentFromString,
                StringAttachment = stringAttachment
            };

            var attachmentArray = new Attachment[] { attachment };

            var result = await EmailTask.SendEmailToExchangeServer(input, attachmentArray, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            await ReadTestEmailWithAttachment(subject);
            Assert.IsTrue(File.Exists(filePath));
            File.Delete(filePath);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailWithEmptyAttachmentTest()
        {
            var subject = "Email test - EmptyEmail";
            var message = "This is a test message from Frends.Commmunity.Email Unit Tests.";
            var input = new ExchangeInput
            {
                To = _username,
                Message = message,
                IsMessageHtml = false,
                Subject = subject
            };

            var attachment = new List<Attachment>();

            var result = await EmailTask.SendEmailToExchangeServer(input, attachment.ToArray(), _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.AreEqual(email[0].BodyText, message);
            await DeleteMessages(subject);
        }

        [Test]
        public async Task SendEmailAsAnotherUserTest()
        {
            var subject = "Email test - AnotherUserTest";
            var message = "This is a test message from Frends.Commmunity.Email Unit Tests.";
            var input = new ExchangeInput
            {
                To = _username,
                From = "frends_exchange_test_user_2@frends.com",
                Message = message,
                IsMessageHtml = false,
                Subject = subject
            };

            var result = await EmailTask.SendEmailToExchangeServer(input, null, _server, new CancellationToken());
            Assert.IsTrue(result.EmailSent);
            Thread.Sleep(2000); // Give the email some time to get through.
            var email = await ReadTestEmail(subject);
            Assert.AreEqual("frends_exchange_test_user_2@frends.com", email[0].From);
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

        private async Task<List<EmailMessageResult>> ReadTestEmail(string subject)
        {
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password
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
            return result;
        }

        private async Task<List<EmailMessageResult>> ReadTestEmailWithAttachment(string subject)
        {
            var dirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../..");
            var settings = new ExchangeSettings
            {
                TenantId = _tenantID,
                AppId = _applicationID,
                Username = _username,
                Password = _password
            };

            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = false,
                AttachmentSaveDirectory = dirPath,
                EmailSubjectFilter = subject
            };

            var result = await ReadEmailTask.ReadEmailFromExchangeServer(settings, options, new CancellationToken());
            return result;
        }

        #endregion
    }
}
