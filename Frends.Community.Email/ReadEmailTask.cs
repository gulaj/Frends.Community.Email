using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using MimeKit;

namespace Frends.Community.Email
{
    /// <summary>
    /// Read email operations for Exchange servers
    /// </summary>
    public class ReadEmailTask
    {
        /// <summary>
        /// For reading emails from an server
        /// </summary>
        /// <param name="settings">Settings to use</param>
        /// <param name="options">Options to use</param>
        /// <returns>
        /// List of
        /// {
        /// string Id.
        /// string To.
        /// string Cc.
        /// string From.
        /// DateTime Date.
        /// string Title.
        /// string BodyText.
        /// string BodyHtml.
        /// }
        /// </returns>
        public static List<EmailMessageResult> ReadEmail([PropertyTab]ReadEmailSettings settings, [PropertyTab]ReadEmailOptions options)
        {
            switch (settings.MailProtocol)
            {
                case MailProtocol.Exchange:
                    return ReadEmailFromExchangeServer(settings, options);
                case MailProtocol.IMAP:
                    return ReadEmailWithIMAP(settings.ServerSettings, options);
                default:
                    return new List<EmailMessageResult>();

            }
        }

        /// <summary>
        /// Read emails from IMAP server
        /// </summary>
        /// <param name="settings">IMAP server settings</param>
        /// <param name="options">Email options</param>
        public static List<EmailMessageResult> ReadEmailWithIMAP(ServerSettings settings, ReadEmailOptions options)
        {
            var result = new List<EmailMessageResult>();

            using (var client = new ImapClient())
            {
                // accept all certs?
                if (settings.AcceptAllCerts)
                {
                    client.ServerCertificateValidationCallback = (s, x509certificate, x590chain, sslPolicyErrors) => true;
                }
                else
                {
                    client.ServerCertificateValidationCallback = MailService.DefaultServerCertificateValidationCallback;
                }

                // connect to imap server
                client.Connect(settings.Host, settings.Port, settings.UseSSL);

                // authenticate with imap server
                client.Authenticate(settings.UserName, settings.Password);

                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);

                // get all or only unread emails?
                IList<UniqueId> messageIds = options.GetOnlyUnreadEmails
                    ? inbox.Search(SearchQuery.NotSeen)
                    : inbox.Search(SearchQuery.All);
                
                // read as many as there are unread emails or as many as defined in options.MaxEmails
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

                    // should mark emails as read?
                    if (!options.DeleteReadEmails && options.MarkEmailsAsRead)
                    {
                        inbox.AddFlags(messageIds[i], MessageFlags.Seen, true);
                    }
                }

                // should delete emails?
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
        /// Reads emails from an Exchange server
        /// </summary>
        /// <param name="settings">Settings to use</param>
        /// <param name="options">Options to use</param>
        /// <returns></returns>
        public static List<EmailMessageResult> ReadEmailFromExchangeServer(ReadEmailSettings settings, ReadEmailOptions options)
        {
            // connect
            ExchangeService exchangeService = ConnectToExchangeService(settings.ExchangeSettings);

            // exchange search view
            ItemView view = new ItemView(options.MaxEmails);

            // for exchange search results
            FindItemsResults<Item> exchangeResults;

            if (options.GetOnlyUnreadEmails)
            {
                // query only unread items
                var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                exchangeResults = exchangeService.FindItems(WellKnownFolderName.Inbox, searchFilter, view);
            }
            else
            {
                // query all items
                exchangeResults = exchangeService.FindItems(WellKnownFolderName.Inbox, view);
            }

            // get email items
            List<EmailMessage> emails = exchangeResults.Where(msg => msg is EmailMessage).Cast<EmailMessage>().ToList();

            // load properties for emails
            exchangeService.LoadPropertiesForItems(emails, new PropertySet(
                BasePropertySet.FirstClassProperties,
                ItemSchema.TextBody,
                EmailMessageSchema.Body));

            // map exchange items to task output results
            List<EmailMessageResult> result = emails
                .Select(msg => new EmailMessageResult
                {
                    Id = msg.Id.UniqueId,
                    Date = msg.DateTimeReceived,
                    Subject = msg.Subject,
                    BodyText = msg.TextBody.Text,
                    BodyHtml = msg.Body.Text,
                    To = string.Join(",", msg.ToRecipients.Select(j => j.Address)),
                    From = msg.From.Address,
                    Cc = string.Join(",", msg.CcRecipients.Select(j => j.Address)),
                })
                .ToList();

            // should delete mails?
            if (options.DeleteReadEmails)
            {
                emails.ForEach(msg => msg.Delete(DeleteMode.HardDelete));
            }

            // should mark mails as read?
            if (!options.DeleteReadEmails && options.MarkEmailsAsRead)
            {
                foreach(EmailMessage msg in emails)
                {
                    msg.IsRead = true;
                    msg.Update(ConflictResolutionMode.AutoResolve);
                }
            }
            
            return result;
        }

        /// <summary>
        /// Helper for connecting to Exchange service
        /// </summary>
        /// <param name="settings">Exchange server related settings</param>
        /// <returns></returns>
        public static ExchangeService ConnectToExchangeService(ExchangeSettings settings)
        {
            ExchangeVersion ev;
            switch (settings.ExchangeServerVersion)
            {
                case ExchangeServerVersion.Exchange2007_SP1:
                    ev = ExchangeVersion.Exchange2007_SP1;
                    break;
                case ExchangeServerVersion.Exchange2010:
                    ev = ExchangeVersion.Exchange2010;
                    break;
                case ExchangeServerVersion.Exchange2010_SP1:
                    ev = ExchangeVersion.Exchange2010_SP1;
                    break;
                case ExchangeServerVersion.Exchange2010_SP2:
                    ev = ExchangeVersion.Exchange2010_SP2;
                    break;
                case ExchangeServerVersion.Exchange2013:
                    ev = ExchangeVersion.Exchange2013;
                    break;
                case ExchangeServerVersion.Exchange2013_SP1:
                    ev = ExchangeVersion.Exchange2013_SP1;
                    break;
                default:
                    ev = ExchangeVersion.Exchange2013;
                    break;
            }

            ExchangeService service = new ExchangeService(ev);

            if (string.IsNullOrWhiteSpace(settings.EmailAddress))
            {
                service.UseDefaultCredentials = true;
            }
            else
            {
                service.Credentials = new NetworkCredential(settings.EmailAddress, settings.Password);
            }

            if (settings.UseAutoDiscover)
            {
                service.AutodiscoverUrl(settings.EmailAddress, RedirectionUrlValidationCallback);
            }
            else
            {
                service.Url = new Uri(settings.ServerAddress);
            }

            return service;
        }

        // The following is a basic redirection validation callback method. It 
        // inspects the redirection URL and only allows the Service object to 
        // follow the redirection link if the URL is using HTTPS. 
        //
        // This redirection URL validation callback provides sufficient security
        // for development and testing of your application. However, it may not
        // provide sufficient security for your deployed application. You should
        // always make sure that the URL validation callback method that you use
        // meets the security requirements of your organization.
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }

            return result;
        }
    }
}
