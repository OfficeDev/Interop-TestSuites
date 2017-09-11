namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation of the SUT Control Adapter interface which is used by test cases in the test suite to send an email to the recipient.
    /// </summary>
    public class MS_OXORULESUTControlAdapter : ManagedAdapterBase, IMS_OXORULESUTControlAdapter
    {
        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
        }

        /// <summary>
        /// A method used to send an email to the specified user account.
        /// </summary>
        /// <param name="senderUserName">The sender's name.</param>
        /// <param name="senderPassword">The sender's password.</param>
        /// <param name="recipientUserName">The recipient's name.</param>
        /// <param name="subject">The email's subject.</param>
        public void SendMailToRecipient(string senderUserName, string senderPassword, string recipientUserName, string subject)
        {
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string ewsUrl = Common.GetConfigurationPropertyValue("EwsUrl", this.Site);

            StringBuilder soapRequestBuilder = new StringBuilder();
            soapRequestBuilder.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            soapRequestBuilder.AppendLine("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            soapRequestBuilder.AppendLine("<soap:Header>");
            soapRequestBuilder.AppendLine("<RequestServerVersion xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\" Version=\"Exchange2016\" />");
            soapRequestBuilder.AppendLine("</soap:Header>");
            soapRequestBuilder.AppendLine("<soap:Body>");
            soapRequestBuilder.AppendLine("<CreateItem MessageDisposition=\"SendOnly\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
            soapRequestBuilder.AppendLine("<Items><Message xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\">");
            soapRequestBuilder.AppendFormat("<Subject>{0}</Subject>", subject);
            soapRequestBuilder.AppendLine("<Body BodyType=\"Text\">The body part is not important, these words are totally useless!</Body>");
            soapRequestBuilder.AppendLine("<Sender><Mailbox>");
            soapRequestBuilder.AppendFormat("<EmailAddress>{0}</EmailAddress>", senderUserName + "@" + domainName);
            soapRequestBuilder.AppendLine("</Mailbox></Sender><ToRecipients><Mailbox>");
            soapRequestBuilder.AppendFormat("<EmailAddress>{0}</EmailAddress>", recipientUserName + "@" + domainName);
            soapRequestBuilder.AppendLine("</Mailbox></ToRecipients><IsRead>false</IsRead></Message></Items>");
            soapRequestBuilder.AppendLine("</CreateItem></soap:Body></soap:Envelope>");

            byte[] requestBytes = System.Text.Encoding.UTF8.GetBytes(soapRequestBuilder.ToString());
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(ewsUrl);
            request.Method = "POST";
            request.ContentType = "text/xml; charset=utf-8";
            request.Headers.Add("SOAPAction", "\"http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem\"");
            request.Credentials = new NetworkCredential(senderUserName, senderPassword, domainName);
            request.ContentLength = requestBytes.Length;
            Stream webRequestStream = request.GetRequestStream();
            webRequestStream.Write(requestBytes, 0, requestBytes.Length);
            webRequestStream.Flush();
            webRequestStream.Dispose();

            HttpWebResponse webResponse = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(webResponse.GetResponseStream(), Encoding.UTF8);
            string responseXml = reader.ReadToEnd();
            reader.Close();
            webResponse.Close();

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(responseXml);
            this.Site.Assert.AreEqual<string>("NoError", doc.DocumentElement.InnerText, "Send a mail to specified user should successfully.");
        }
    }
}