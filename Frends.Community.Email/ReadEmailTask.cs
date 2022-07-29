﻿using Azure.Identity;
using MailKit;
using MailKit.Net.Imap;
using Microsoft.Graph;
using MimeKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SearchQuery = MailKit.Search.SearchQuery;
using Directory = System.IO.Directory;
using Path = System.IO.Path;
using File = System.IO.File;

namespace Frends.Community.Email
{
    /// <summary>
    /// Read email.
    /// </summary>
    public class ReadEmailTask
    {
        /// <summary>
        /// Read emails from IMAP server
        /// </summary>
        /// <param name="settings">IMAP server settings</param>
        /// <param name="options">Email options</param>
        /// <returns>List{ string Id, string To, string Cc, string From, DateTime Date, string Subject, string BodyText, string BodyHtml }</returns>
        public static List<EmailMessageResult> ReadEmailWithIMAP([PropertyTab]ImapSettings settings, [PropertyTab]ImapOptions options)
        {
            var result = new List<EmailMessageResult>();

            using (var client = new ImapClient())
            {
                // Accept all certs?
                if (settings.AcceptAllCerts)
                    client.ServerCertificateValidationCallback = (s, x509certificate, x590chain, sslPolicyErrors) => true;
                else
                    client.ServerCertificateValidationCallback = MailService.DefaultServerCertificateValidationCallback;

                // Connect to imap server.
                client.Connect(settings.Host, settings.Port, settings.UseSSL);

                // Authenticate with imap server.
                client.Authenticate(settings.UserName, settings.Password);

                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);

                // Get all or only unread emails?
                IList<UniqueId> messageIds = options.GetOnlyUnreadEmails
                    ? inbox.Search(SearchQuery.NotSeen)
                    : inbox.Search(SearchQuery.All);

                // Read as many as there are unread emails or as many as defined in options.MaxEmails.
                for (int i = 0; i < messageIds.Count && i < options.MaxEmails; i++)
                {
                    MimeMessage msg = inbox.GetMessage(messageIds[i]);
                    result.Add(new EmailMessageResult
                    {
                        Id = msg.MessageId,
                        Date = msg.Date.DateTime,
                        Subject = msg.Subject,
                        BodyText = msg.TextBody,
                        BodyHtml = msg.HtmlBody,
                        From = string.Join(",", msg.From.Select(j => j.ToString())),
                        To = string.Join(",", msg.To.Select(j => j.ToString())),
                        Cc = string.Join(",", msg.Cc.Select(j => j.ToString()))
                    });

                    // Should mark emails as read?
                    if (!options.DeleteReadEmails && options.MarkEmailsAsRead)
                        inbox.AddFlags(messageIds[i], MessageFlags.Seen, true);
                }

                // Should delete emails?
                if (options.DeleteReadEmails && messageIds.Any())
                {
                    inbox.AddFlags(messageIds, MessageFlags.Deleted, false);
                    inbox.Expunge();
                }

                client.Disconnect(true);
            }

            return result;
        }

        /// <summary>
        /// Reads emails from an Exchange server.
        /// </summary>
        /// <param name="settings">Settings to use</param>
        /// <param name="options">Options to use</param>
        /// <param name="cancellationToken"></param>
        /// <returns>List{ string Id, string To, string Cc, string From, DateTime Date, string Subject, string BodyText, string BodyHtml, List(string) AttachmentSaveDirs }</returns>
        public static async Task<List<EmailMessageResult>> ReadEmailFromExchangeServer([PropertyTab]ExchangeSettings settings, [PropertyTab]ExchangeOptions options, CancellationToken cancellationToken)
        {
            var queryOptions = new List<QueryOption>();
            var result = new List<EmailMessageResult>();
            var credentials = new UsernamePasswordCredential(settings.Username, settings.Password, settings.TenantId, settings.AppId);
            var graphServiceClient = new GraphServiceClient(credentials);
            var searchQuery = "";
            var queryInUse = false;

            if (string.IsNullOrEmpty(settings.Username) || string.IsNullOrEmpty(settings.Password) || string.IsNullOrEmpty(settings.AppId) || string.IsNullOrEmpty(settings.TenantId))
                throw new ArgumentException("Username, Password, Application ID and Tenant ID cannot be empty. Please check Exchange settings.");

            if (string.IsNullOrEmpty(settings.Mailbox))
                throw new ArgumentException("No mailbox provided. Please provide mailbox where emails will be read.");

            if (!string.IsNullOrWhiteSpace(options.EmailSenderFilter))
            {
                searchQuery = $"from:{options.EmailSenderFilter}";
                queryInUse = true;
            }

            if (!string.IsNullOrWhiteSpace(options.EmailSubjectFilter))
            {
                searchQuery += queryInUse ? $" AND subject:{options.EmailSubjectFilter}" : $"subject:{options.EmailSubjectFilter}";
                queryInUse = true;
            }

            if (options.GetOnlyEmailsWithAttachments)
            {
                searchQuery += queryInUse ? " AND hasAttachments:true" : $"hasAttachments:true";
                queryInUse = true;
            }

            if (queryInUse) queryOptions.Add(new QueryOption("$search", $"\"{searchQuery}\""));
            var messages = await graphServiceClient.Me.MailFolders[settings.Mailbox].Messages.Request(queryOptions).Top(options.MaxEmails).GetAsync(cancellationToken);

            if (messages.Count == 0 && options.ThrowErrorIfNoMessagesFound)
                throw new Exception("No emails were found.");
            foreach (var email in messages)
            {
                var attachmentPaths = new List<string>();
                if (!options.IgnoreAttachments && email.HasAttachments.Value)
                {
                    if (string.IsNullOrEmpty(options.AttachmentSaveDirectory))
                        throw new ArgumentException("No attachment save directory provided.");
                    else if (!Directory.Exists(options.AttachmentSaveDirectory))
                        Directory.CreateDirectory(options.AttachmentSaveDirectory);

                    var attachments = await graphServiceClient.Me.Messages[email.Id].Attachments.Request().GetAsync(cancellationToken);

                    foreach (FileAttachment attachment in attachments)
                    {
                        var path = Path.Combine(options.AttachmentSaveDirectory, attachment.Name);
                        if (File.Exists(path) && options.OverwriteAttachment)
                            File.Delete(path);
                        else if (File.Exists(path))
                            throw new Exception("Attachment file " + attachment.Name + " already exists in the given directory.");
                        File.WriteAllBytes(path, attachment.ContentBytes);
                        attachmentPaths.Add(path);
                    }
                }

                var singleResult = new EmailMessageResult
                {
                    Id = email.Id,
                    To = string.Join(", ", email.ToRecipients),
                    Cc = email.CcRecipients != null ? string.Join(", ", email.CcRecipients) : "",
                    From = email.From.EmailAddress.Address,
                    Date = email.ReceivedDateTime.Value.DateTime,
                    Subject = email.Subject,
                    BodyText = email.Body.ContentType == BodyType.Text ? email.Body.Content : "",
                    BodyHtml = email.Body.ContentType == BodyType.Html ? email.Body.Content : "",
                    AttachmentSaveDirs = attachmentPaths
                };

                if (options.GetOnlyUnreadEmails && !email.IsRead.Value)
                {
                    var added = false;
                    if (options.GetOnlyEmailsWithAttachments && singleResult.AttachmentSaveDirs.Count != 0)
                    {
                        result.Add(singleResult);
                        added = true;
                    }
                    else if (!options.GetOnlyEmailsWithAttachments)
                    {
                        result.Add(singleResult);
                        added = true;
                    }
                    if (added && options.DeleteReadEmails)
                        await graphServiceClient.Me.Messages[email.Id].Request().DeleteAsync(cancellationToken);
                    else if (added && options.MarkEmailsAsRead)
                        await graphServiceClient.Me.Messages[email.Id].Request().Select("IsRead").UpdateAsync(new Message { IsRead = true }, cancellationToken);
                }
                else if (!options.GetOnlyUnreadEmails)
                {
                    var added = false;
                    if (options.GetOnlyEmailsWithAttachments && singleResult.AttachmentSaveDirs.Count != 0)
                    {
                        result.Add(singleResult);
                        added = true;
                    }
                    else if (!options.GetOnlyEmailsWithAttachments)
                    {
                        result.Add(singleResult);
                        added = true;
                    }
                    if (added && options.DeleteReadEmails)
                        await graphServiceClient.Me.Messages[email.Id].Request().DeleteAsync(cancellationToken);
                    else if (added && options.MarkEmailsAsRead)
                        await graphServiceClient.Me.Messages[email.Id].Request().Select("IsRead").UpdateAsync(new Message { IsRead = true }, cancellationToken);
                }
            }

            return result;
        }
    }
}
