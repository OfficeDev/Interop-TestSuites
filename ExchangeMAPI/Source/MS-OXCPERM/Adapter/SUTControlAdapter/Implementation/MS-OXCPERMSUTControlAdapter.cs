namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Policy;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The Implementation of the SUT Control Adapter interface.
    /// </summary>
    public class MS_OXCPERMSUTControlAdapter : ManagedAdapterBase, IMS_OXCPERMSUTControlAdapter
    {
        /// <summary>
        /// The Availability Web Service object.
        /// </summary>
        private ExchangeServiceBinding availability = new ExchangeServiceBinding();

        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            string url = Common.GetConfigurationPropertyValue("EwsUrl", this.Site);
            this.availability.Url = url;
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                this.AcceptServerCertificate();
            }
        }

        /// <summary>
        /// Gets the free/busy status appointment information for User2 (as specified in ptfconfig) through testUserName's account.
        /// </summary>
        /// <param name="testUserName">The user who gets the free/busy status information.</param>
        /// <param name="password">The testUserName's password.</param>
        /// <returns>
        /// <para>"0": means "FreeBusy", which indicates brief information about the appointments on the calendar;</para>
        /// <para>"1": means "Detailed", which indicates detailed information about the appointments on the calendar;</para>
        /// <para>"2": means the appointment free/busy information can't be viewed or error occurs, which indicates the user has no permission to get information about the appointments on the calendar;</para>
        /// </returns>
        public string GetUserFreeBusyStatus(string testUserName, string password)
        {
            // Use the specified user for the web service client authentication
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.availability.Credentials = new NetworkCredential(testUserName, password, domain);

            GetUserAvailabilityRequestType availabilityRequest = new GetUserAvailabilityRequestType();

            SerializableTimeZone timezone = new SerializableTimeZone()
            {
                Bias = 480,
                StandardTime = new SerializableTimeZoneTime() { Bias = 0, Time = "02:00:00", DayOrder = 5, Month = 10, DayOfWeek = "Sunday" },
                DaylightTime = new SerializableTimeZoneTime() { Bias = -60, Time = "02:00:00", DayOrder = 1, Month = 4, DayOfWeek = "Sunday" }
            };
            availabilityRequest.TimeZone = timezone;

            // Specifies the mailbox to query for availability information.
            string user = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
            EmailAddress emailAddress = new EmailAddress()
            {
                Address = string.IsNullOrEmpty(user) ? string.Empty : user + "@" + domain
            };

            MailboxData mailboxData = new MailboxData()
            {
                Email = emailAddress,
                AttendeeType = MeetingAttendeeType.Organizer,
            };

            availabilityRequest.MailboxDataArray = new MailboxData[] { mailboxData };

            // Identify the time to compare free/busy information.
            FreeBusyViewOptionsType freeBusyViewOptions = new FreeBusyViewOptionsType()
            {
                TimeWindow = new Duration() { StartTime = DateTime.Now, EndTime = DateTime.Now.AddHours(3) },
                RequestedView = FreeBusyViewType.Detailed,
                RequestedViewSpecified = true
            };

            availabilityRequest.FreeBusyViewOptions = freeBusyViewOptions;

            GetUserAvailabilityResponseType availabilityInfo = null;
            try
            {
                availabilityInfo = this.availability.GetUserAvailability(availabilityRequest);
            }
            catch (SoapException exception)
            {
                Site.Assert.Fail("Error occurs when getting free/busy status: {0}", exception.Message);
            }

            string freeBusyStatus = "3";
            FreeBusyResponseType[] freeBusyArray = availabilityInfo.FreeBusyResponseArray;
            if (freeBusyArray != null)
            {
                foreach (FreeBusyResponseType freeBusy in freeBusyArray)
                {
                    ResponseClassType responseClass = freeBusy.ResponseMessage.ResponseClass;
                    if (responseClass == ResponseClassType.Success)
                    {
                        // If the response FreeBusyViewType value is FreeBusy or Detailed, the freeBusyStatus is Detailed.
                        // If all the response FreeBusyViewType values are FreeBusy, the freeBusyStatus is FreeBusy;
                        if (freeBusy.FreeBusyView.FreeBusyViewType == FreeBusyViewType.Detailed)
                        {
                            freeBusyStatus = "1";
                        }
                        else if (freeBusy.FreeBusyView.FreeBusyViewType == FreeBusyViewType.FreeBusy)
                        {
                            if (freeBusyStatus != "1")
                            {
                                freeBusyStatus = "0";
                            }
                        }
                    }
                    else if (responseClass == ResponseClassType.Error)
                    {
                        if (freeBusy.ResponseMessage.ResponseCode == ResponseCodeType.ErrorNoFreeBusyAccess)
                        {
                            return "2";
                        }
                        else
                        {
                            Site.Assert.Fail("Error occurs when getting free/busy status. ErrorCode: {0}; ErrorMessage: {1}.", freeBusy.ResponseMessage.ResponseCode, freeBusy.ResponseMessage.MessageText);
                        }
                    }
                }
            }

            return freeBusyStatus;
        }

        /// <summary>
        /// Verify the remote Secure Sockets Layer (SSL) certificate used for authentication.
        /// In adapter, this method always return true, make client can communicate with server under HTTPS without a certification. 
        /// </summary>
        /// <param name="sender">An object that contains state information for this validation.</param>
        /// <param name="certificate">The certificate used to authenticate the remote party.</param>
        /// <param name="chain">The chain of certificate authorities associated with the remote certificate.</param>
        /// <param name="sslPolicyErrors">One or more errors associated with the remote certificate.</param>
        /// <returns>A Boolean value that determines whether the specified certificate is accepted for authentication.</returns>
        private static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            SslPolicyErrors errors = sslPolicyErrors;

            if ((errors & SslPolicyErrors.RemoteCertificateNameMismatch) == SslPolicyErrors.RemoteCertificateNameMismatch)
            {
                Zone zone = Zone.CreateFromUrl(((HttpWebRequest)sender).RequestUri.ToString());
                if (zone.SecurityZone == SecurityZone.Intranet || zone.SecurityZone == SecurityZone.MyComputer)
                {
                    errors -= SslPolicyErrors.RemoteCertificateNameMismatch;
                }
            }

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == SslPolicyErrors.RemoteCertificateChainErrors)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        // Self-signed certificates have the issuer in the subject field. 
                        if ((certificate.Subject == certificate.Issuer) && (status.Status == X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else if (status.Status != X509ChainStatusFlags.NoError)
                        {
                            // If there are any other errors in the certificate chain, the certificate is invalid, the method returns false.
                            return false;
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are untrusted root errors for self-signed certificates. 
                // These certificates are valid.
                errors -= SslPolicyErrors.RemoteCertificateChainErrors;
            }

            return errors == SslPolicyErrors.None;
        }

        /// <summary>
        /// If the SOAP over HTTPS is used as transport, the adapter uses this function 
        /// to avoid closing base connection.
        /// Local client will accept any valid server certificate after executing this function.
        /// </summary>
        private void AcceptServerCertificate()
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
        }
    }
}