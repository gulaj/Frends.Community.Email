using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;

namespace Frends.Community.Email.Tests
{
    [TestFixture]
    [Ignore("Tests requires external SMTP server to work")]
    class ReadEmailTask_ExchangeTest
    {
        private readonly string _userName = "";
        private readonly string _password = "";

        [Test]
        public void ReadEmailFromExchangeServer_ShouldReadOneItem()
        {
            var settings = new ReadEmailSettings
            {
                ExchangeSettings = new ExchangeSettings
                {
                    ExchangeServerVersion = ExchangeServerVersion.Exchange2010,
                    UseAutoDiscover = true,
                    EmailAddress = _userName,
                    Password = _password
                }
            };
            var options = new ReadEmailOptions
            {
                MaxEmails = 1,
                DeleteReadEmails = true,
                GetOnlyUnreadEmails = false,
                MarkEmailsAsRead = false
            };

            List<EmailMessageResult> result = ReadEmailTask.ReadEmailFromExchangeServer(settings, options);

            Assert.That(result.Count, Is.EqualTo(1));
        }
    }
}
