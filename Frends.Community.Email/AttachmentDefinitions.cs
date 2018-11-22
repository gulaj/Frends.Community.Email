using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Community.Email
{
    /// <summary>
    /// Attachment options
    /// </summary>
    public class FetchAttachmentOptions
    {
        /// <summary>
        /// Maximum number of emails to retrieve
        /// </summary>
        [DefaultValue(10)]
        public int MaxEmails { get; set; }

        /// <summary>
        /// Directory where attachments will be saved to.
        /// </summary>
        [DefaultValue("")]
        [DisplayFormat(DataFormatString = "Text")]
        public string AttachmentSaveDirectory { get; set; }

        /// <summary>
        /// Should the attachment be overwritten, if the save directory
        /// already contains an attachment with the same name? If no,
        /// a GUID will be added to the filename.
        /// </summary>
        public bool OverwriteAttachment { get; set; }

        /// <summary>
        /// Optional. If a sender is given, it will be used to filter emails.
        /// </summary>
        [DefaultValue("")]
        [DisplayFormat(DataFormatString = "Text")]
        public string EmailSenderFilter { get; set; }

        /// <summary>
        /// Optional. If a subject is given, it will be used to filter emails.
        /// </summary>
        [DefaultValue("")]
        [DisplayFormat(DataFormatString = "Text")]
        public string EmailSubjectFilter { get; set; }

        /// <summary>
        /// If true, the task throws an error if no messages matching
        /// search criteria were found.
        /// </summary>
        public bool ThrowErrorIfNoMessagesFound { get; set; }

        /// <summary>
        /// If true, the task gets attachments of only unread emails.
        /// </summary>
        public bool GetOnlyUnreadEmails { get; set; }

        /// <summary>
        /// If true, the task fetches only emails with attachments.
        /// </summary>
        public bool GetOnlyEmailsWithAttachments { get; set; }

        /// <summary>
        /// If true, then marks queried emails as read unless task execution is cancelled during processing.
        /// </summary>
        public bool MarkEmailsAsRead { get; set; }

        /// <summary>
        /// If true, then received emails with their attachments
        /// will be permanently deleted (hard delete) unless task
        /// execution is cancelled during processing. 
        public bool DeleteReadEmails { get; set; }
    }

    /// <summary>
    /// Output result for fetch attachment operation
    /// </summary>
    public class EmailAttachmentResult
    {
        /// <summary>
        /// Email id
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// Email id
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// To-field from email
        /// </summary>
        public string To { get; set; }

        /// <summary>
        /// Cc-field from email
        /// </summary>
        public string Cc { get; set; }

        /// <summary>
        /// From-field from email
        /// </summary>
        public string From { get; set; }

        /// <summary>
        /// Email received date
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// Title of the email
        /// </summary>
        public string Subject { get; set; }
        
        /// <summary>
        /// Body of the email as text
        /// </summary>
        public string BodyText { get; set; }

        /// <summary>
        /// Attachment download path
        /// </summary>
        public List<string> AttachmentSaveDirs { get; set; }
    }
}
