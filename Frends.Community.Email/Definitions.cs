using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

#pragma warning disable 1591

namespace Frends.Community.Email
{
    public class Input
    {
        /// <summary>
        /// Recipient addresses separated by ',' or ';'
        /// </summary>
        [DefaultValue("\"jane.doe@somedomain.com\"")]
        public string To { get; set; }

        /// <summary>
        /// Cc recipient addresses separated by ',' or ';'
        /// </summary>
        [DefaultValue("\"jane.doe@somedomain.com\"")]
        public string Cc { get; set; }

        /// <summary>
        /// Bcc recipient addresses separated by ',' or ';'
        /// </summary>
        [DefaultValue("\"jane.doe@somedomain.com\"")]
        public string Bcc { get; set; }

        /// <summary>
        /// Sender address.
        /// </summary>
        [DefaultValue("\"john.doe@somedomain.com\"")]
        public string From { get; set; }

        /// <summary>
        /// Name of the sender.
        /// </summary>
        [DefaultValue("\"\"")]
        public string SenderName { get; set; }

        /// <summary>
        /// Email message's subject.
        /// </summary>
        [DefaultValue("\"Hello Jane\"")]
        public string Subject { get; set; }

        /// <summary>
        /// Body of the message.
        /// </summary>
        [DefaultValue("\"You've got mail!\"")]
        public string Message { get; set; }

        /// <summary>
        /// Set this true if the message is HTML.
        /// </summary>
        [DefaultValue("false")]
        public bool IsMessageHtml { get; set; }

        /// <summary>
        /// Encoding of message body and subject. Use following table's name column for other options. https://msdn.microsoft.com/en-us/library/system.text.encoding(v=vs.110).aspx#Anchor_5 
        /// </summary>
        [DefaultValue("\"utf-8\"")]
        public string MessageEncoding { get; set; }

    }

    public class Options
    {
        /// <summary>
        /// SMTP server address.
        /// </summary>
        [DefaultValue("\"smtp.somedomain.com\"")]
        public string SMTPServer { get; set; }

        /// <summary>
        /// SMTP server port.
        /// </summary>
        [DefaultValue("25")]
        public int Port { get; set; }

        /// <summary>
        /// Set this true if SMTP expects to be connected using SSL.
        /// </summary>
        [DefaultValue("false")]
        public bool UseSsl { get; set; }

        /// <summary>
        /// Should the task accept all certificates from IMAP server, including invalid ones?
        /// </summary>
        [DefaultValue(false)]
        public bool AcceptAllCerts { get; set; }

        /// <summary>
        /// Use this username to log in to the SMTP server
        /// </summary>
        [DefaultValue("\"\"")]
        public string UserName { get; set; }

        /// <summary>
        /// Use this password to log in to the SMTP server
        /// </summary>
        [PasswordPropertyText(true)]
        [DefaultValue("\"\"")]
        public string Password { get; set; }
    }

    public class Output
    {
        /// <summary>
        /// Value is true if email was sent.
        /// </summary>
        public bool EmailSent { get; set; }
        /// <summary>
        /// Contains information about the task's result.
        /// </summary>
        public string StatusString { get; set; }
    }

    public class Attachment
    {
        /// <summary>
        /// Chooses if the attachment file is created from a string or copied from disk.
        /// </summary>
        public AttachmentType AttachmentType { get; set; }

        [UIHint(nameof(AttachmentType), "", AttachmentType.AttachmentFromString)]
        public AttachmentFromString stringAttachment { get; set; }

        /// <summary>
        /// Attachment file's path. Uses Directory.GetFiles(string, string) as a pattern matching technique. See https://msdn.microsoft.com/en-us/library/wz42302f(v=vs.110).aspx.
        /// Exception: If the path ends in a directory, all files in that folder are added as attachments.
        /// </summary>
        [DefaultValue("\"\"")]
        [UIHint(nameof(AttachmentType), "", AttachmentType.FileAttachment)]
        public string FilePath { get; set; }

        /// <summary>
        /// If set true and no files match the given path, an exception is thrown.
        /// </summary>
        [UIHint(nameof(AttachmentType), "", AttachmentType.FileAttachment)]
        public bool ThrowExceptionIfAttachmentNotFound { get; set; }

        /// <summary>
        /// If set true and no files match the given path, email will be sent nevertheless.
        /// </summary>
        [UIHint(nameof(AttachmentType), "", AttachmentType.FileAttachment)]
        public bool SendIfNoAttachmentsFound { get; set; }
    }
    public enum AttachmentType
    {
        /// <summary>
        /// Select this if the attachment is a file.
        /// </summary>
        FileAttachment,

        /// <summary>
        /// Select this if the attachment file should be created from a string.
        /// </summary>
        AttachmentFromString
    }

    public class AttachmentFromString
    {
        /// <summary>
        /// Content of the attachment file
        /// </summary>
        [DefaultValue("\"\"")]
        public string FileContent { get; set; }

        /// <summary>
        /// Name of the attachment file
        /// </summary>
        [DefaultValue("\"\"")]
        public string FileName { get; set; }
    }
}
