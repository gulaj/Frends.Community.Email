using System.IO;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using Frends.Tasks.Attributes;
using System;
using System.Text;

#pragma warning disable 1591


namespace Frends.Community.Email
{
    public class EmailTask
    {
        /// <summary>
        /// Sends email message with optional attachments. See https://github.com/CommunityHiQ/Frends.Community.Sftp
        /// </summary>
        /// <returns>
        /// Object { bool EmailSent, string StatusString }
        /// </returns>
        public static async Task<Output> SendEmail([CustomDisplay(DisplayOption.Tab)]Input message, [CustomDisplay(DisplayOption.Tab)]Attachment[] attachments, [CustomDisplay(DisplayOption.Tab)]Options SMTPSettings, CancellationToken cancellationToken)
        {
            var output = new Output();

            using (var client = InitializeSmtpClient(SMTPSettings))
            {
                using (var mail = InitializeMailMessage(message))
                {                 
                    if (attachments != null)
                        foreach(var attachment in attachments)
                        {
                            if (attachment.AttachmentType == AttachmentType.FileAttachment)
                            {
                                ICollection<string> allAttachmentFilePaths = GetAttachmentFiles(attachment.FilePath);

                                if (attachment.ThrowExceptionIfAttachmentNotFound && allAttachmentFilePaths.Count == 0)
                                    throw new FileNotFoundException(string.Format("The given filepath \"{0}\" had no matching files", attachment.FilePath), attachment.FilePath);

                                if (allAttachmentFilePaths.Count == 0 && !attachment.SendIfNoAttachmentsFound)
                                {
                                    output.StatusString = string.Format("No attachments found matching path \"{0}\". No email sent.", attachment.FilePath);
                                    output.EmailSent = false;
                                    return output;
                                }

                                foreach (var fp in allAttachmentFilePaths)
                                {
                                    mail.Attachments.Add(new System.Net.Mail.Attachment(fp));
                                }
                            }

                            if (attachment.AttachmentType == AttachmentType.AttachmentFromString)
                            {
                                //Create attachment only if content is not empty
                                if (!string.IsNullOrEmpty(attachment.stringAttachment.FileContent))
                                    mail.Attachments.Add(System.Net.Mail.Attachment.CreateAttachmentFromString(attachment.stringAttachment.FileContent, attachment.stringAttachment.FileName));
                            }
                        
                        }
                    
                    cancellationToken.ThrowIfCancellationRequested();

                    await client.SendMailAsync(mail);                 

                    output.EmailSent = true;
                    output.StatusString = string.Format("Email sent to: {0}", mail.To.ToString());

                    return output;
                }
            }
        }
        
        /// <summary>
        /// Initializes new SmtpClient with given parameters.
        /// </summary>
        private static SmtpClient InitializeSmtpClient(Options settings)
        {
            var smtpClient = new SmtpClient
            {
                Port = settings.Port,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = settings.UseWindowsAuthentication,
                EnableSsl = settings.UseSsl,
                Host = settings.SMTPServer
            };

            if (!settings.UseWindowsAuthentication)
                smtpClient.Credentials = new NetworkCredential(settings.UserName, settings.Password);            

            return smtpClient;
        }

        /// <summary>
        /// Initializes new MailMessage with given parameters. Uses default value 'true' for IsBodyHtml
        /// </summary>
        private static MailMessage InitializeMailMessage(Input input)
        {
            //split recipients, either by comma or semicolon
            var separators = new[] { ',', ';' };

            string[] recipients = input.To.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            //Create mail object
            var mail = new MailMessage()
            {
                From = new MailAddress(input.From, input.SenderName),
                Subject = input.Subject,
                Body = input.Message,
                IsBodyHtml = input.IsMessageHtml
            };
            //Add recipients
            foreach (var recipientAddress in recipients)
            {
                mail.To.Add(recipientAddress);
            }

            //Set message encoding
            Encoding encoding = Encoding.GetEncoding(input.MessageEncoding);

            mail.BodyEncoding = encoding;
            mail.SubjectEncoding = encoding;

            return mail;
        }

        /// <summary>
        /// Gets all actual file names of attachments matching given file path
        /// </summary>
        private static ICollection<string> GetAttachmentFiles(string filePath)
        {            
            string folder = Path.GetDirectoryName(filePath);
            string fileMask = Path.GetFileName(filePath) != "" ? Path.GetFileName(filePath) : "*";

            string[] filePaths = Directory.GetFiles(folder, fileMask);            

            return filePaths;
        }
    }
}
