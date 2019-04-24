using System.Collections.Generic;
using NUnit.Framework;

namespace Frends.Community.Email.Tests
{
    [TestFixture]
    [Ignore("Tests requires external SMTP server to work")]
    public class ReadEmailTaskExchangeTest
    {
        private readonly string _userName = "";
        private readonly string _password = "";

        [Test]
        public void ReadEmailFromExchangeServer_ShouldReadOneItem()
        {
            var settings = new ExchangeSettings
            {
                ExchangeServerVersion = ExchangeServerVersion.Office365,
                UseAutoDiscover = true,
                EmailAddress = _userName,
                Password = _password
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true
            };

            List<EmailMessageResult> result = ReadEmailTask.ReadEmailFromExchangeServer(settings, options);

            Assert.That(result.Count, Is.EqualTo(1));
        }
    }
}
