namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Policy;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation of the SUT Control Adapter interface which is used by test cases in the test suite to set or clear Out of Office state by calling Exchange OOF Web Service. 
    /// </summary>
    public class MS_OXWOOFSUTControlAdapter : ManagedAdapterBase, IMS_OXWOOFSUTControlAdapter
    {
        /// <summary>
        /// Set user mailbox Out of Office state.
        /// </summary>
        /// <param name="mailAddress">User's email address.</param>
        /// <param name="password">Password of user mailbox.</param>
        /// <param name="isOOF">If true, set OOF state, else make sure OOF state is not set (clear OOF state).</param>
        /// <returns>If the operation succeed then return true, otherwise return false.</returns>
        public bool SetUserOOFSettings(string mailAddress, string password, bool isOOF)
        {
            using (ExchangeServiceBinding service = new ExchangeServiceBinding())
            {
                service.Url = Common.GetConfigurationPropertyValue(Constants.SetOOFWebServiceURL, this.Site);
                if (service.Url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                {
                    this.AcceptServerCertificate();
                }

                service.Credentials = new System.Net.NetworkCredential(mailAddress, password);

                EmailAddress emailAddress = new EmailAddress
                {
                    Address = mailAddress, Name = string.Empty
                };
                UserOofSettings userSettings = new UserOofSettings
                {
                    // Identify the external audience.
                    ExternalAudience = ExternalAudience.Known
                };

                // Create the OOF reply messages.
                ReplyBody replyBody = new ReplyBody
                {
                    Message = Constants.MessageOfOOFReply
                };

                userSettings.ExternalReply = replyBody;
                userSettings.InternalReply = replyBody;

                // Set OOF state.
                if (isOOF)
                {
                    userSettings.OofState = OofState.Enabled;
                }
                else
                {
                    userSettings.OofState = OofState.Disabled;
                }

                // Create the request.
                SetUserOofSettingsRequest request = new SetUserOofSettingsRequest
                {
                    Mailbox = emailAddress,
                    UserOofSettings = userSettings
                };

                bool success = false;

                try
                {
                    SetUserOofSettingsResponse response = service.SetUserOofSettings(request);
                    if (response.ResponseMessage.ResponseCode == ResponseCodeType.NoError)
                    {
                        success = true;
                    }
                }
                catch (System.Xml.Schema.XmlSchemaValidationException e)
                {
                    // Catch the following critical exceptions, other unexpected exceptions will be emitted to protocol test framework.
                    Site.Log.Add(LogEntryKind.Debug, "An XML schema exception happened. The exception type is {0}.\n The exception message is {1}.", e.GetType().ToString(), e.Message);
                }
                catch (System.Xml.XmlException e)
                {
                    Site.Log.Add(LogEntryKind.Debug, "An XML schema exception happened. The exception type is {0}.\n The exception message is {1}.", e.GetType().ToString(), e.Message);
                }
                catch (System.Reflection.TargetException e)
                {
                    Site.Log.Add(LogEntryKind.Debug, "An operation exception happened when invoke Soap operation. The exception type is {0}.\n The exception message is {1}.", e.GetType().ToString(), e.Message);
                    throw;
                }
                catch (System.IO.IOException e)
                {
                    Site.Log.Add(LogEntryKind.Debug, "An IO exception happened when invoke Soap operation. The exception type is {0}.\n The exception message is {1}.", e.GetType().ToString(), e.Message);
                    throw;
                }

                return success;
            }
        }

        /// <summary>
        /// Verifies the remote Secure Sockets Layer (SSL) certificate used for authentication.
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