using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Net;
using MimeKit;

namespace Frends.Community.Email
{
    /// <summary>
    /// Fetch attachment operations for Exchange servers
    /// </summary>
    public class FetchExchangeAttachmentsTask
    {
        /// <summary>
        /// For fetching file attachments from emails on an Exchange server
        /// </summary>
        /// <param name="settings">Settings to use</param>
        /// <param name="options">Options to use</param>
        /// <returns>
        /// List of
        /// {
        /// string Status
        /// string Id.
        /// string To.
        /// string Cc.
        /// string From.
        /// DateTime Date.
        /// string Title.
        /// string BodyText.
        /// List of {string AttachmentSavePath} .
        /// }
        /// </returns>
        public static List<EmailAttachmentResult> FetchExchangeAttachments([PropertyTab]ExchangeSettings settings,[PropertyTab]FetchAttachmentOptions options, CancellationToken cToken)
        {
            // Check that save directory is given
            if (options.AttachmentSaveDirectory == "")
                throw new ArgumentException("No save directory given. ",
                    nameof(options.AttachmentSaveDirectory));

            // Check that save directory exists
            if (!Directory.Exists(options.AttachmentSaveDirectory))
                throw new ArgumentException("Could not find or access attachment save directory. ",
                    nameof(options.AttachmentSaveDirectory));

            List<EmailAttachmentResult> result = new List<EmailAttachmentResult> { };

            // Connect
            ExchangeService exchangeService = Services.ConnectToExchangeService(settings);

            // Exchange search view
            ItemView view = new ItemView(options.MaxEmails);

            // For exchange search results
            FindItemsResults<Item> exchangeResults;

            // Create search filter collection. Always search only for emails with attachments.
            var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(ItemSchema.HasAttachments, true));

            // Construct rest of search filter based on options

            if (options.GetOnlyUnreadEmails)
                searchFilter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

            if (options.EmailSenderFilter != "")
                searchFilter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, options.EmailSenderFilter));

            if (options.EmailSubjectFilter != "")
                searchFilter.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, options.EmailSubjectFilter));

            // Query desired items
            exchangeResults = exchangeService.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

            // Get email items
            List<EmailMessage> emails = exchangeResults.Where(msg => msg is EmailMessage).Cast<EmailMessage>().ToList();

            // Check if list is empty and if an error needs to be thrown.
            // If not, return a result with a notification of no found messages.
            if (emails.Count == 0 && options.ThrowErrorIfNoMessagesFound)
            {
                throw new ArgumentException("No messages matching the search filter found. ",
                    nameof(options.ThrowErrorIfNoMessagesFound));
            }
            else if (emails.Count == 0)
            {
                result.Add(new EmailAttachmentResult
                {
                    Status = "No messages matching the search filter found.",
                    Id = "",
                    Date = DateTime.Now,
                    Subject = "",
                    BodyText = "",
                    To = "",
                    From = "",
                    Cc = "",
                    AttachmentSaveDirs = new List<string> { }
                });
                return result;
            }

            // Load properties for each email and process attachments
            foreach (EmailMessage email in emails)
            {
                // Initialize attachment save path list
                List<string> pathList = new List<string> { };

                cToken.ThrowIfCancellationRequested();

                // Define property set
                var propSet = new PropertySet(
                        BasePropertySet.FirstClassProperties,
                        EmailMessageSchema.Body,
                        EmailMessageSchema.Attachments);
                propSet.RequestedBodyType = BodyType.Text;

                // Bind and load email message with desired properties
                var newEmail = EmailMessage.Bind(
                    exchangeService,
                    email.Id,
                    propSet);

                // Save all attachments to given directory
                foreach (var attachment in newEmail.Attachments)
                {
                    FileAttachment file = attachment as FileAttachment;
                    string path = "";
                    if (options.OverwriteAttachment)
                    {
                        path = Path.Combine(options.AttachmentSaveDirectory, file.Name);
                    }
                    else
                    {
                        string uniqueName = Path.GetFileNameWithoutExtension(file.Name)
                            + "_" + Guid.NewGuid().ToString()
                            + Path.GetExtension(file.Name);
                        path = Path.Combine(options.AttachmentSaveDirectory, uniqueName);
                    }
                    file.Load(path);
                    pathList.Add(path);

                    cToken.ThrowIfCancellationRequested();
                }

                // Build result for email message
                result.Add(new EmailAttachmentResult
                {
                    Status = "Ok.",
                    Id = newEmail.Id.UniqueId,
                    Date = newEmail.DateTimeReceived,
                    Subject = newEmail.Subject,
                    BodyText = newEmail.Body.Text,
                    To = string.Join(",", newEmail.ToRecipients.Select(j => j.Address)),
                    From = newEmail.From.Address,
                    Cc = string.Join(",", newEmail.CcRecipients.Select(j => j.Address)),
                    AttachmentSaveDirs = pathList
                });
            }

            // Should delete mails?
            if (options.DeleteReadEmails)
            {
                emails.ForEach(msg => msg.Delete(DeleteMode.HardDelete));
            }

            // Should mark mails as read?
            if (!options.DeleteReadEmails && options.MarkEmailsAsRead)
            {
                foreach (EmailMessage msg in emails)
                {
                    msg.IsRead = true;
                    msg.Update(ConflictResolutionMode.AutoResolve);
                }
            }

            return result;
        }
    }
}
