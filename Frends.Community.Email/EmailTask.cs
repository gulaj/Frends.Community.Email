using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net;
using System.Threading;
using System;
using MimeKit;
using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;

#pragma warning disable 1591

namespace Frends.Community.Email
{
    public class EmailTask
    {
        /// <summary>
        /// Sends email message with optional attachments. See https://github.com/CommunityHiQ/Frends.Community.Email
        /// </summary>
        /// <returns>
        /// Object { bool EmailSent, string StatusString }
        /// </returns>
        public static Output SendEmail([PropertyTab]Input message, [PropertyTab]Attachment[] attachments, [PropertyTab]Options SMTPSettings, CancellationToken cancellationToken)
        {
            var output = new Output();

            var mail = CreateMimeMessage(message);

            if (attachments != null)
            {
                // Email object is created using BodyBuilder.
                var builder = new BodyBuilder();

                if (message.IsMessageHtml) builder.HtmlBody = string.Format(message.Message);
                else builder.TextBody = string.Format(message.Message);

                foreach (var attachment in attachments)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (attachment.AttachmentType == AttachmentType.FileAttachment)
                    {
                        var allAttachmentFilePaths = GetAttachmentFiles(attachment.FilePath);

                        if (attachment.ThrowExceptionIfAttachmentNotFound && allAttachmentFilePaths.Count == 0) throw new FileNotFoundException(string.Format("The given filepath \"{0}\" had no matching files", attachment.FilePath), attachment.FilePath);

                        if (allAttachmentFilePaths.Count == 0 && !attachment.SendIfNoAttachmentsFound)
                        {
                            output.StatusString = string.Format("No attachments found matching path \"{0}\". No email sent.", attachment.FilePath);
                            output.EmailSent = false;
                            return output;
                        }

                        foreach (var filePath in allAttachmentFilePaths) builder.Attachments.Add(filePath);
                    }

                    if (attachment.AttachmentType == AttachmentType.AttachmentFromString)
                    {
                        // Create attachment only if content is not empty.
                        if (!string.IsNullOrEmpty(attachment.stringAttachment.FileContent))
                        {
                            var path = CreateTemporaryFile(attachment);
                            builder.Attachments.Add(path);
                            CleanUpTempWorkDir(path);
                        }
                    }
                }

                mail.Body = builder.ToMessageBody();
            }
            else
            {
                mail.Body = (message.IsMessageHtml) 
                    ? new TextPart(MimeKit.Text.TextFormat.Html) { Text = message.Message } 
                    : new TextPart(MimeKit.Text.TextFormat.Plain) { Text = message.Message };
            }

            // Initialize new MailKit SmtpClient.
            using (var client = new SmtpClient())
            {
                // Accept all certs?
                if (SMTPSettings.AcceptAllCerts) client.ServerCertificateValidationCallback = (s, x509certificate, x590chain, sslPolicyErrors) => true;
                else client.ServerCertificateValidationCallback = MailService.DefaultServerCertificateValidationCallback;

                if (SMTPSettings.UseSsl) client.Connect(SMTPSettings.SMTPServer, SMTPSettings.Port, SecureSocketOptions.SslOnConnect);
                else client.Connect(SMTPSettings.SMTPServer, SMTPSettings.Port);

                client.AuthenticationMechanisms.Remove("XOAUTH2");

                if (string.IsNullOrEmpty(SMTPSettings.UserName) || string.IsNullOrEmpty(SMTPSettings.Password)) throw new ArgumentException("SMTP credentials were not given for authentication.");

                client.Authenticate(new NetworkCredential(SMTPSettings.UserName, SMTPSettings.Password));

                client.Send(mail);

                client.Disconnect(true);

                client.Dispose();
            }

            output.EmailSent = true;
            output.StatusString = string.Format("Email sent to: {0}", mail.To.ToString());

            return output;
        }

        #region HelperMethods

        /// <summary>
        /// Create MimeMessage.
        /// </summary>
        private static MimeMessage CreateMimeMessage([PropertyTab] Input message)
        {
            // Split recipients, either by comma or semicolon.
            var separators = new[] { ',', ';' };

            var recipients = message.To.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            var ccRecipients = message.Cc.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            var bccRecipients = message.Bcc.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            // Create mail object.
            var mail = new MimeMessage();
            mail.From.Add(new MailboxAddress(message.SenderName, message.From));
            mail.Subject = message.Subject;

            // Add recipients.
            foreach (var recipientAddress in recipients) mail.To.Add(MailboxAddress.Parse(recipientAddress));

            // Add CC recipients.
            foreach (var ccRecipient in ccRecipients) mail.Cc.Add(MailboxAddress.Parse(ccRecipient));

            // Add BCC recipients.
            foreach (var bccRecipient in bccRecipients) mail.Bcc.Add(MailboxAddress.Parse(bccRecipient));

            return mail;
        }

        /// <summary>
        /// Gets all actual file names of attachments matching given file path.
        /// </summary>
        /// <param name="filePath"></param>
        private static ICollection<string> GetAttachmentFiles(string filePath)
        {
            var folder = Path.GetDirectoryName(filePath);
            var fileMask = Path.GetFileName(filePath) != "" ? Path.GetFileName(filePath) : "*";
            var filePaths = Directory.GetFiles(folder, fileMask);
            return filePaths;
        }

        /// <summary>
        /// Create temp file of attachment from string.
        /// </summary>
        /// <param name="attachment"></param>
        private static string CreateTemporaryFile(Attachment attachment)
        {
            var TempWorkDirBase = InitializeTemporaryWorkPath();
            var filePath = Path.Combine(TempWorkDirBase, attachment.stringAttachment.FileName);
            var content = attachment.stringAttachment.FileContent;

            using (StreamWriter sw = File.CreateText(filePath)) sw.Write(content);

            return filePath;
        }

        /// <summary>
        /// Remove the temporary workdir.
        /// </summary>
        /// <param name="tempWorkDir"></param>
        private static void CleanUpTempWorkDir(string tempWorkDir)
        {
            if (!string.IsNullOrEmpty(tempWorkDir) && Directory.Exists(tempWorkDir)) Directory.Delete(tempWorkDir, true);
        }

        /// <summary>
        /// Create temperary directory for temp file.
        /// </summary>
        private static string InitializeTemporaryWorkPath()
        {
            var tempWorkDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempWorkDir);

            return tempWorkDir;
        }

        #endregion
    }
}
