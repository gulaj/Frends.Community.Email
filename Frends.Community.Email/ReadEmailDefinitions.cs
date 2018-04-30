using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Community.Email
{
    /// <summary>
    /// Protocol to use
    /// </summary>
    public enum MailProtocol
    {
        /// <summary>
        /// Pop3 protocol
        /// </summary>
        // Pop3, // NOT IMPLEMENTED

        /// <summary>
        /// IMAP
        /// </summary>
        IMAP, // NOT TESTED

        /// <summary>
        /// Exchange servers
        /// </summary>
        Exchange
    }

    /// <summary>
    /// Connection settings
    /// </summary>
    public class ReadEmailSettings
    {
        /// <summary>
        /// Protocol to use
        /// </summary>
        public MailProtocol MailProtocol { get; set; }

        /// <summary>
        /// Settings for IMAP & POP3 servers
        /// </summary>
        [UIHint(nameof(Email.MailProtocol), "", MailProtocol.IMAP)]
        public ServerSettings ServerSettings { get; set; }

        /// <summary>
        /// Exchange server specific options
        /// </summary>
        [UIHint(nameof(Email.MailProtocol), "", MailProtocol.Exchange)]
        public ExchangeSettings ExchangeSettings { get;set; }
    }

    /// <summary>
    /// Settings for IMAP and POP3 servers
    /// </summary>
    public class ServerSettings
    {
        /// <summary>
        /// Host address
        /// </summary>
        [DefaultValue("imap.frends.com")]
        [DisplayFormat(DataFormatString = "Text")]
        public string Host { get; set; }

        /// <summary>
        /// Host port
        /// </summary>
        [DefaultValue(993)]
        public int Port { get; set; }

        /// <summary>
        /// Use SSL or not
        /// </summary>
        [DefaultValue(true)]
        public bool UseSSL { get; set; }

        /// <summary>
        /// Account name to login with
        /// </summary>
        [DefaultValue("accountName")]
        [DisplayFormat(DataFormatString = "Text")]
        public string UserName { get; set; }

        /// <summary>
        /// Account password
        /// </summary>
        [PasswordPropertyText]
        public string Password { get; set; }
    }

    /// <summary>
    /// Exchange server spesific options
    /// </summary>
    public class ExchangeSettings
    {
        /// <summary>
        /// Which exchange server to target
        /// </summary>
        public ExchangeServerVersion ExchangeServerVersion { get; set; }

        /// <summary>
        /// If true, will try to auto discover server address from user name. In this cae Host and Port values are not used.
        /// </summary>
        public bool UseAutoDiscover { get; set; }

        /// <summary>
        /// Exchange server address
        /// </summary>
        [DefaultValue("exchange.frends.com")]
        [DisplayFormat(DataFormatString = "Text")]
        [UIHint(nameof(UseAutoDiscover), "", false)]
        public string ServerAddress { get; set; }

        /// <summary>
        /// Try to login with agent account?
        /// </summary>
        [DefaultValue(false)]
        [Description("Authorize with agent account")]
        public bool UseAgentAccount { get; set; }

        /// <summary>
        /// Email account to use
        /// </summary>
        [DefaultValue("agent@frends.com")]
        [DisplayFormat(DataFormatString = "Text")]
        public string EmailAddress { get; set; }

        /// <summary>
        /// Account password
        /// </summary>
        [PasswordPropertyText]
        [UIHint(nameof(UseAgentAccount), "", false)]
        public string Password { get; set; }
    }

    /// <summary>
    /// Wich exchange version to target
    /// </summary>
    public enum ExchangeServerVersion
    {
        Exchange2007_SP1,
        Exchange2010,
        Exchange2010_SP1,
        Exchange2010_SP2,
        Exchange2013,
        Exchange2013_SP1
    }

    /// <summary>
    /// Options related to read operation
    /// </summary>
    public class ReadEmailOptions
    {
        /// <summary>
        /// Maximum number of emails to retrieve
        /// </summary>
        [DefaultValue(10)]
        public int MaxEmails { get; set; }

        /// <summary>
        /// Should get only unread emails?
        /// </summary>
        public bool GetOnlyUnreadEmails { get; set; }

        /// <summary>
        /// If true, then marks queried emails as read
        /// </summary>
        public bool MarkEmailsAsRead { get; set; }

        /// <summary>
        /// If true, then received emails will be hard deleted
        /// </summary>
        public bool DeleteReadEmails { get; set; }
    }

    /// <summary>
    /// Output result for read operation
    /// </summary>
    public class EmailMessageResult
    {
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
        public string Title { get; set; }

        /// <summary>
        /// Body of the email
        /// </summary>
        public string Body { get; set; }
    }
}
