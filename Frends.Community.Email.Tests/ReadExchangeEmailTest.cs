using NUnit.Framework;

namespace Frends.Community.Email.Tests
{
    [TestFixture]
    [Ignore("Tests requires external SMTP server to work")]
    public class ReadExchangeEmailTest
    {
        private readonly string _userName = "";
        private readonly string _password = "";
        private readonly string _mailbox = "";

        [Test]
        public void ReadEmailFromExchangeServer_ShouldReadOneItem()
        {
            var settings = new ExchangeSettings
            {
                ExchangeServerVersion = ExchangeServerVersion.Office365,
                UseAutoDiscover = true,
                Username = _userName,
                Password = _password,
                Mailbox = _mailbox
            };
            var options = new ExchangeOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = false,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false,
                IgnoreAttachments = true
            };

            var result = ReadEmailTask.ReadEmailFromExchangeServer(settings, options);

            Assert.That(result.Count, Is.EqualTo(1));
        }
    }
}
