namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the ResolveRecipients command.
    /// </summary>
    [TestClass]
    public class S13_ResolveRecipients : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is used to verify whether the ResolveRecipients command is responded by server, if the client specifies the CertificateRetrieval element, server will return the corresponding status.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC01_ResolveRecipients_CertificateRetrieval()
        {
            #region Call ResolveRecipients command and set the CertificateRetrieval value to 1 that specifies server does not retrieve certificates for the recipient.
            Request.ResolveRecipientsOptions requestResolveRecipientsOption = new Request.ResolveRecipientsOptions
            {
                CertificateRetrieval = "1"
            };

            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items =
                        new object[] { requestResolveRecipientsOption, Common.GetConfigurationPropertyValue("User3Name", Site) }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R856");

            // If CertificateRetrieval value is set to 1 , server returns a null Certificates.
            // Verify MS-ASCMD requirement: MS-ASCMD_R856
            Site.CaptureRequirementIfIsNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates,
                856,
                @"[In CertificateRetrieval] Value 1 means do not retrieve certificates for the recipient (default).");
            #endregion

            #region Call ResolveRecipients command and set the CertificateRetrieval value to 2 that specifies server should return the full certificate for each resolved recipient.
            requestResolveRecipientsOption.CertificateRetrieval = "2";
            resolveRecipientsRequest.RequestData.Items = new object[] { requestResolveRecipientsOption, Common.GetConfigurationPropertyValue("User3Name", Site) };
            resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R843");

            // Resolve a list of supplied recipients, server returns a non-null Certificates.
            // Verify MS-ASCMD requirement: MS-ASCMD_R843
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates,
                843,
                @"[In Certificate(ResolveRecipients)] This element [Certificate] is returned by the server only if the client specifies a value of 2 in the CertificateRetrieval element (section 2.2.3.22) in the ResolveRecipients command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R857");

            // Resolve a list of supplied recipients, server returns a non-null Certificates.
            // Verify MS-ASCMD requirement: MS-ASCMD_R857
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates,
                857,
                @"[In CertificateRetrieval] Value 2 means retrieve the full certificate for each resolved recipient.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3825");

            // Resolve a list of supplied recipients, server returns a non-null Type.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3825
            Site.CaptureRequirementIfIsTrue(
                (resolveRecipientsResponse.ResponseData.Response[0].Recipient.Length > (int)0) && (resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].EmailAddress != null) && (resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates != null),
                3825,
                @"[In Response(ResolveRecipients)] If the recipient was resolved, the element also contains the type of recipient, the email address that the recipient resolved to, and, optionally, the S/MIME certificate for the recipient.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3764");

            // Resolve a list of supplied recipients, server returns a non-null Certificates.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3764
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates,
                3764,
                @"[In Recipient] A Certificates element is returned as a child element of the Recipient element if the client requested certificates to be returned in the response.");
            #endregion

            #region Call ResolveRecipients command and set the CertificateRetrieval value to 3 that specifies server should return the mini certificate for each resolved recipient.
            requestResolveRecipientsOption.CertificateRetrieval = "3";
            resolveRecipientsRequest.RequestData.Items = new object[] { requestResolveRecipientsOption, Common.GetConfigurationPropertyValue("User3Name", Site) };
            resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
            Site.Assert.AreEqual<string>("1", resolveRecipientsResponse.ResponseData.Status, "The server should return a status code 1 in the ResolveRecipients command response to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R858");

            // Resolve a list of supplied recipients, server returns a non-null Certificates.
            // Verify MS-ASCMD requirement: MS-ASCMD_R858
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates.MiniCertificate,
                858,
                @"[In CertificateRetrieval] Value 3 means retrieve the mini certificate for each resolved recipient.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3435");

            // Resolve a list of supplied recipients, server returns a non-null Certificates.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3435
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates.MiniCertificate,
                3435,
                @"[In MiniCertificate] This [MiniCertificate] element is returned only if the client specifies a value of 3 in the CertificateRetrieval element (section 2.2.3.22) in the ResolveRecipients command request and the resolved recipient has a valid S/MIME certificate.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4302");

            // Resolve a list of supplied recipients, server returns a non-null Certificates.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4302
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates.Status),
                4302,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 1 [is] One or more certificates were successfully returned.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if EndTime is invalid, server returns status 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC02_ResolveRecipients_EndTimeInvalid()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The EndTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command and set the EndTime element with invalid time format.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2012-04-11T10:00:00.000Z",
                                EndTime = "TimeWithInvalidFormat"
                            }
                        },
                        Common.GetConfigurationPropertyValue("User1Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2264");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2264
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                int.Parse(resolveRecipientsResponse.ResponseData.Status),
                2264,
                @"[In EndTime(ResolveRecipients)] If the client sends an invalid EndTime element value, then the server returns a Status element (section 2.2.3.162.11) value of 5 in the ResolveRecipients command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3762");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3762
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                int.Parse(resolveRecipientsResponse.ResponseData.Status),
                3762,
                @"[In Recipient] The status code returned in the Response element can be used to determine if the recipient was found to be ambiguous.");
        }

        /// <summary>
        /// This test case is used to verify if StartTime is invalid, server returns status 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC03_ResolveRecipients_StartTimeInvalid()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The StartTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command and set the StartTime element with invalid time format.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "TimeWithInvalidFormat",
                                EndTime = "2012-04-11T10:00:00.000Z"
                            }
                        },
                        Common.GetConfigurationPropertyValue("User1Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3986");

            // If server returns ResolveRecipients Status 5. The R3986 should be covered.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3986
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                int.Parse(resolveRecipientsResponse.ResponseData.Status),
                3986,
                @"[In StartTime(ResolveRecipients)] If the client sends an invalid StartTime element value, then the server returns a Status element (section 2.2.3.162.11) value of 5 in the ResolveRecipients command response.");
        }

        /// <summary>
        /// This test case is used to verify when the EndTime is smaller than StartTime, server should return status 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC04_ResolveRecipients_EndTimeSmallThanStartTime()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The StartTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // The client calls ResolveRecipients command with valid value of EndTime and StartTime but the value of EndTime is smaller than StartTime.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2010-03-01T00:20:00.000Z",
                                EndTime = "2010-03-01T00:00:00.000Z"
                            }
                        },
                        Common.GetConfigurationPropertyValue("User1Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2266");

            // If server returns ResolveRecipients Status 5. The R2266 should be covered.
            // Verify MS-ASCMD requirement: MS-ASCMD_R2266
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                int.Parse(resolveRecipientsResponse.ResponseData.Status),
                2266,
                @"[In EndTime(ResolveRecipients)] If the EndTime element value specified in the request is smaller than the StartTime element value plus 30 minutes,[ or the duration spanned by the StartTime element value and the EndTime element value is greater than a specific number<37> of days,] then the server returns a Status element value of 5 in the ResolveRecipients command response.");
        }

        /// <summary>
        /// This test case is used to verify StartTime and EndTime of the spans over maximum value, then server returns a status element value of 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC05_ResolveRecipients_DurationIsTooLarge()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The StartTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command with the duration spanned by the StartTime and the EndTime is greater than a specific number.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2010-03-01T00:00:00.000Z",
                                EndTime = "2010-06-03T00:00:01.000Z"
                            }
                        },
                        Common.GetConfigurationPropertyValue("User1Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Server returns ResolveRecipients Status 5 if the StartTime and the EndTime is greater than a specific number.
            if (Common.IsRequirementEnabled(5175, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5175");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5175
                Site.CaptureRequirementIfAreEqual<int>(
                    5,
                    int.Parse(resolveRecipientsResponse.ResponseData.Status),
                    5175,
                    @"[In Appendix A: Product Behavior] If the duration spanned by the StartTime element value and the EndTime element value is greater than 42 days, then the implementation does return a Status element value of 5 in the ResolveRecipients command response. (<37> Section 2.2.3.58.1: Exchange 2010 and Exchange 2013 use 42 days.)");
            }
        }

        /// <summary>
        /// This test case is used to verify the length of To element can be up to 256.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC06_ResolveRecipients_Over256()
        {
            #region The client calls ResolveRecipients command with To element value length of 256.
            string toElementWith256Chars = CreateFixedLengthString(256);
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            CertificateRetrieval = "2"
                        },
                        toElementWith256Chars
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
            int status256 = int.Parse(resolveRecipientsResponse.ResponseData.Status);
            #endregion

            #region The client calls ResolveRecipients command with To elements value length of 257.
            string toElementWith257Chars = CreateFixedLengthString(257);
            ResolveRecipientsRequest resolveRecipientsRequest257 = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            CertificateRetrieval = "2"
                        },
                        toElementWith257Chars
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse257 = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest257);
            int status257 = int.Parse(resolveRecipientsResponse257.ResponseData.Status);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5873");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5873
            Site.CaptureRequirementIfIsTrue(
                (status256 == 1) && (status257 == 5),
                5873,
                @"[In To] Its [The To element] value is not larger than 256 characters in length.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if the picture does not exist, sever should return 173.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC07_ResolveRecipients_NoPictureExist()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MaxPictures element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MaxPictures element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // The client calls ResolveRecipients command with MaxPictures element setting to 3
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Picture = new Request.ResolveRecipientsOptionsPicture
                            {
                                MaxPictures = 3
                            }
                        },
                        Common.GetConfigurationPropertyValue("User3Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3287");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3287
            Site.CaptureRequirementIfAreEqual<string>(
                "173",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Status,
                3287,
                @"[In MaxPictures(ResolveRecipients)] After the MaxPictures limit is reached, the server returns Status element (section 2.2.3.162) value 173 (NoPicture) if the contact has no photo.");
        }

        /// <summary>
        /// This test case is used to verify one or more Recipient elements are returned to the client in a Response element, when using ambiguous name in the request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC08_ResolveRecipients_AmbiguousUser()
        {
            // The client calls ResolveRecipients command with an ambiguous name in the command request. 
            string displayName = Common.GetConfigurationPropertyValue("AmbiguousSearchName", Site);
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[] { displayName }
                }
            };
            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
            Site.Assert.AreEqual<string>("2", resolveRecipientsResponse.ResponseData.Response[0].Status, "If the recipient was found to be ambiguous, the status should be 2.");

            bool isFullNameReturned = true;
            bool isNullCertificates = true;
            foreach (Response.ResolveRecipientsResponseRecipient recipient in resolveRecipientsResponse.ResponseData.Response[0].Recipient)
            {
                if (!recipient.DisplayName.ToLower(System.Globalization.CultureInfo.InvariantCulture).Contains(displayName.ToLower(System.Globalization.CultureInfo.InvariantCulture)))
                {
                    isFullNameReturned = false;
                }

                if (recipient.Certificates != null)
                {
                    isNullCertificates = false;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4279");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4279
            // If the recipient was found to be ambiguous, the status should be 2 and no certificate nodes returned.
            Site.CaptureRequirementIfIsTrue(
                isNullCertificates,
                4279,
                @"[In Status(ResolveRecipients)] [The meaning of the status value 2 is] No certificate nodes were returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3761");

            // If server returns the ambiguous user full name, the R3761 should be covered.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3761
            Site.CaptureRequirementIfIsTrue(
                isFullNameReturned,
                3761,
                @"[In Recipient] One or more Recipient elements are returned to the client in a Response element by the server if the To element specified in the request was either resolved to a distribution list or found to be ambiguous.");
        }

        /// <summary>
        /// The test case is used to verify response will not include free/busy data and MergedFreeBusy element, when To element specifies an ambiguous name and the Availability element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC09_ResolveRecipients_AmbiguousUserAvailability()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The StartTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call method ResolveRecipients with a To element specifies an ambiguous name to resolve all recipients which their name match the ambiguous name.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2013-08-11T10:00:00.000Z",
                                EndTime = "2013-08-11T11:00:00.000Z"
                            }
                        },
                        Common.GetConfigurationPropertyValue("AmbiguousSearchName", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
            Site.Assert.IsNotNull(resolveRecipientsResponse.ResponseData.Response[0].Recipient, "Server should return recipients data");

            // Checks whether server returns free/busy data and MergedFreeBusy element.
            bool mergedFreeBusyisNull = true;
            foreach (Response.ResolveRecipientsResponseRecipient recipient in resolveRecipientsResponse.ResponseData.Response[0].Recipient)
            {
                mergedFreeBusyisNull = recipient.Availability == null;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4623");

            // If request includes the Availability element and includes a To element for an ambiguous user, server returns a null Availability without free/busy data.
            Site.CaptureRequirementIfIsTrue(
                mergedFreeBusyisNull,
                4623,
                @"[In To] If the ResolveRecipients command request includes the Availability element and includes a To element for an ambiguous user, the response does not include a MergedFreeBusy element (section 2.2.3.97) for that user.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4617");

            // If request includes the Availability element and includes a To element for an ambiguous user, server returns a null Availability without free/busy data.
            Site.CaptureRequirementIfIsTrue(
                mergedFreeBusyisNull,
                4617,
                @"[In To] If the To element specifies an ambiguous name and the Availability element (section 2.2.3.16) is included in the request, the response will not include free/busy data for that user.");
        }

        /// <summary>
        /// This test case is used to verify if the request contains Availability element and a group name, server will return the merger of the data for the individual members of the specified group.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC10_ResolveRecipients_Group_Availability()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Availability element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // The client calls ResolveRecipients command with a group name and Availability element. 
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2010-03-01T00:20:00.000Z",
                                EndTime = "2010-03-01T01:00:00.000Z"
                            }
                        },
                        Common.GetConfigurationPropertyValue("GroupDisplayName", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4625");

            // If request includes the Availability element and includes a To element for a distribution group, server returns a non-null Availability.
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability,
                4625,
                @"[In To] If the ResolveRecipients command request includes the Availability element and the To element specifies a distribution group, then the availability data is returned as a single string that merges the data for the individual members of the distribution group.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R819");

            // If request includes the Availability element and includes a To element for a distribution group, server returns a non-null MergedFreeBusy.
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.MergedFreeBusy,
                819,
                @"[In Availability] When the Availability element is included in a ResolveRecipients request, the server retrieves free/busy information for the users identified in the To elements included in the request, and returns the free/busy information in the MergedFreeBusy element in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3339");

            // If call ResolveRecipients command includes the Availability element and includes a To element for a distribution group successfully, server returns status 1 and a non-null MergedFreeBusy.
            Site.CaptureRequirementIfIsTrue(
                (resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.MergedFreeBusy != null) && (resolveRecipientsResponse.ResponseData.Status == "1"),
                3339,
                @"[In MergedFreeBusy] The MergedFreeBusy element is also included if the Status element value indicates success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3338");

            // If call ResolveRecipients command includes the Availability element and includes a To element for a distribution group successfully, server returns status 1.
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.Status,
                3338,
                @"[In MergedFreeBusy] If the Availability element is included in the response, the response MUST also include the Status element (section 2.2.3.162.11).");
        }

        /// <summary>
        /// This test case is used to verify if the distribution group contains more than 20 members, a Status element value of 161 is returned in the response.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC11_ResolveRecipients_Status161()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Availability element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call method ResolveRecipients to resolve a list of recipients in a special distribution group which contains more than 20 members.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2013-08-11T10:00:00.000Z",
                                EndTime = "2013-08-11T11:00:00.000Z"
                            }
                        },
                        Common.GetConfigurationPropertyValue("LargeGroupDisplayName", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4626");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4626
            Site.CaptureRequirementIfAreEqual<string>(
                "161",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.Status,
                4626,
                @"[In To] If the distribution group contains more than 20 members, a Status element value of 161 is returned in the response indicating that the merged free busy information of such a large distribution group is not useful.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4294");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4294
            Site.CaptureRequirementIfAreEqual<string>(
                "161",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.Status,
                4294,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 161 [is] The distribution group identified by the To element of the ResolveRecipient request included more than 20 recipients.");
        }

        /// <summary>
        /// This test case is used to verify To element(s) that are returned in the response correspond directly to the To element(s) that are specified in the request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC12_ResolveRecipients_To()
        {
            // Call method ResolveRecipients to resolve a specified recipients with To element.
            string recipientName = Common.GetConfigurationPropertyValue("User1Name", Site);
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[] { recipientName }
                }
            };
            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4268");

            // Server return status code 1 in response level to indicate ResolveRecipients command process success.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4268
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                resolveRecipientsResponse.ResponseData.Status,
                4268,
               @"[In Status(ResolveRecipients)] [The meaning of the status value] 1 [is] Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4275");

            // Server return status code 1 under response level to indicate ResolveRecipients command resolve recipients success.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4275
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                resolveRecipientsResponse.ResponseData.Response[0].Status,
                4275,
               @"[In Status(ResolveRecipients)] [The meaning of the status value] 1 [is] The recipient was resolved successfully.");
        }

        /// <summary>
        /// This test case is used to verify ResolveRecipients command, if the To element with valid SMTP addresses in the request message, the MergedFreeBusy elements will be in the response.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC13_ResolveRecipients_MergedFreeBusy()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The StartTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command with To element set to valid SMTP addresses to retrieve the free/busy information of the specified recipient.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2013-08-17T23:00:00.000Z",
                                EndTime = "2013-08-17T23:30:00.000Z"
                            }
                        },
                        Common.GetMailAddress(Common.GetConfigurationPropertyValue("User1Name", Site), Common.GetConfigurationPropertyValue("Domain", Site))
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3321");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3321
            Site.CaptureRequirementIfAreEqual<string>(
                "0",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.MergedFreeBusy,
                3321,
                @"[In MergedFreeBusy] Value 0 means Free.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4291");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4291
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.Status,
                4291,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 1 [is] Free/busy data was successfully retrieved for a given recipient.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4618");

            // Server return Availability element if the To element includes a valid SMTP address in ResolveRecipients request.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4618
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability,
                4618,
                @"[In To] The Availability element is only included when the To element includes a valid SMTP address or name that resolves to a unique individual on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4624");

            // If specified with valid SMTP addresses, To element includes MergedFreeBusy element in the response.
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.MergedFreeBusy,
                4624,
                @"[In To] Only users or distribution lists specified with valid SMTP addresses or a uniquely identifiable string in the request message, To element have MergedFreeBusy elements included in the response.");
        }

        /// <summary>
        /// This test case is used to verify whether the ResolveRecipients command is responded by server, if the client specifies a value of 2 in the CertificateRetrieval element, but there is no certificate on the server, the status is equal to 7.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC14_ResolveRecipients_NoCertificates()
        {
            // Call ResolveRecipients command with the CertificateRetrieval value set to 2 to resolve the recipient which have no certificate.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            CertificateRetrieval = "2"
                        },
                        Common.GetConfigurationPropertyValue("User1Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R343");

            // Verify MS-ASCMD requirement: MS-ASCMD_R343
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                resolveRecipientsResponse.ResponseData.Status,
                343,
                @"[In ResolveRecipients] The ResolveRecipients command is used by clients to resolve a list of supplied recipients, to retrieve their free/busy information, and optionally, to retrieve their S/MIME certificates so that clients can send encrypted S/MIME email messages.<4>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4304");

            // If there is no certificate on the server, the status is equal to 7.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4304
            Site.CaptureRequirementIfAreEqual<string>(
                "7",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Certificates.Status,
                4304,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 7 [is] No certificates were returned.");
        }

        /// <summary>
        /// This test case is used to verify ResolveRecipients command, each digit in the MergedFreeBusy element value string indicates the free/busy status for the user or distribution list for every 30 minute interval.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC15_ResolveRecipients_MergedFreeBusyIntervalTime()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The StartTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Check User2's MergedFreeBusy status before receiving meeting request
            this.SwitchUser(this.User2Information);
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = DateTime.UtcNow.Date.AddDays(1).AddHours(2).ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                                EndTime = DateTime.UtcNow.Date.AddDays(1).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                            }
                        },
                        Common.GetConfigurationPropertyValue("User2Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponseBeforeReceiveMeeting = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
            Site.Assert.AreEqual<string>("1", resolveRecipientsResponseBeforeReceiveMeeting.ResponseData.Status, "If ResolveRecipients command executes successfully, server should return status 1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3319");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3319 
            Site.CaptureRequirementIfAreEqual(
                "0000",
                resolveRecipientsResponseBeforeReceiveMeeting.ResponseData.Response[0].Recipient[0].Availability.MergedFreeBusy,
                3319,
                @"[In MergedFreeBusy] Each digit in the MergedFreeBusy element value string indicates the free/busy status for the user or distribution list for every 30 minute interval.");
            #endregion

            #region User1 calls SendMail command to send one meeting request to user2
            this.SwitchUser(this.User1Information);
            string meetingSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = new Calendar();
            calendar.StartTime = DateTime.UtcNow.Date.AddDays(1).AddHours(2).AddMinutes(40);
            calendar.EndTime = DateTime.UtcNow.Date.AddDays(1).AddHours(3).AddMinutes(10);

            // Send a meeting request email to user2
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0")|| Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                Calendar createdCalendar = this.CreateCalendar(meetingSubject, attendeeEmailAddress, calendar);
                this.SendMeetingRequest(meetingSubject, createdCalendar);
            }
            else
            {
                calendar.OrganizerEmail = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
                calendar.OrganizerName = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
                calendar.UID = Guid.NewGuid().ToString();
                calendar.Attendees = new Response.Attendees();
                calendar.Attendees.Attendee = new Response.AttendeesAttendee[1];
                calendar.Attendees.Attendee[0] = new Response.AttendeesAttendee();
                calendar.Attendees.Attendee[0].Email = attendeeEmailAddress;
                calendar.Attendees.Attendee[0].Name = attendeeEmailAddress;
                this.SendMeetingRequest(meetingSubject, calendar);
            }
            #endregion

            #region Get new added meeting request emails in user2's mailbox
            // Switch to user2's mailbox
            this.SwitchUser(this.User2Information);

            // Sync Inbox folder
            this.GetMailItem(this.User2Information.InboxCollectionId, meetingSubject);

            // Sync Calendar folder
            this.GetMailItem(this.User2Information.CalendarCollectionId, meetingSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingSubject);
            #endregion

            #region Check user2's MergedFreeBusy status before sending meeting response
            ResolveRecipientsResponse resolveRecipientsResponseBeforeMeetingResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R3330, MS-ASCMD_R3332
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3330");

            // Because of value "0" indicates "Free", value "1" indicates "Tentative", then server change the MergedFreeBusy value to "0110" to indicate the recipient is "Busy" during the middle one hour.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3330
            Site.CaptureRequirementIfAreEqual(
                "0110",
                resolveRecipientsResponseBeforeMeetingResponse.ResponseData.Response[0].Recipient[0].Availability.MergedFreeBusy,
                3330,
                @"[In MergedFreeBusy] The MergedFreeBusy element value string is populated from the StartTime element value onwards, therefore the last digit represents between a millisecond and 30 minutes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3332");

            // Because of value "0" indicates "Free", value "1" indicates "Tentative", then server change the MergedFreeBusy value to "0110" to indicate the recipient is "Busy" during the middle one hour.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3332
            Site.CaptureRequirementIfAreEqual(
                "0110",
                resolveRecipientsResponseBeforeMeetingResponse.ResponseData.Response[0].Recipient[0].Availability.MergedFreeBusy,
                3332,
                @"[In MergedFreeBusy] Any appointment that ends inside a second of the interval requested shall impact the digit representing that timeframe.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify ResolveRecipients command, if user does not have a contact photo, the status is equal to 173.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC16_ResolveRecipients_Picture_Status173()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command to resolve the special recipient who does not have any contact photo.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Picture = new Request.ResolveRecipientsOptionsPicture()
                        },
                        Common.GetConfigurationPropertyValue("User3Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4311");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4311
            Site.CaptureRequirementIfAreEqual<string>(
                "173",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Status,
                4311,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 173 [is] The user does not have a contact photo.");
        }

        /// <summary>
        /// This test case is used to verify the ResolveRecipients command request is successful, when setting the parameters of the Picture element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC17_ResolveRecipients_Picture_Success()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command to resolve the special recipient who has a contact photo.
            Request.ResolveRecipientsOptions requestResolveRecipientsOption = new Request.ResolveRecipientsOptions
            {
                Picture = new Request.ResolveRecipientsOptionsPicture
                {
                    MaxSizeSpecified = true,
                    MaxSize = 102400,
                    MaxPicturesSpecified = true,
                    MaxPictures = 3
                }
            };

            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        requestResolveRecipientsOption,
                        Common.GetConfigurationPropertyValue("User1Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4310");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4311
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Status,
                4310,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 1 [is] The contact photo was retrieved successfully.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2133");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2133
            // The Data element value indicates the contact photo size, if the Data element value is not null then MS-ASCMD_R2133 is verified.
            Site.CaptureRequirementIfIsNotNull(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Data,
                2133,
                @"[In Data(ResolveRecipients)] The Data element<27> is an optional child element of the Picture element in ResolveRecipients command responses that contains the binary data of the contact photo.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5350");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5350
            // If the Data element value is less than MaxSize value set in request, then MS-ASCMD_R5350 is verified.
            Site.CaptureRequirementIfIsTrue(
                Convert.ToInt32(resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Data) <= requestResolveRecipientsOption.Picture.MaxSize,
                5350,
                @"[In MaxSize] The MaxSize element specifies the maximum size of an individual contact photo that is returned in the response, in bytes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5351");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5351
            Site.CaptureRequirementIfIsTrue(
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture.Length <= requestResolveRecipientsOption.Picture.MaxPictures,
                5351,
                @"[In MaxSize] The MaxPictures element (section 2.2.3.94) specifies the maximum number of contact photos to return in the server response.");
        }

        /// <summary>
        /// This test case is used to verify ResolveRecipients command, if contact photo exceeded the size limit set by the MaxSize element, the status is equal to 174.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC18_ResolveRecipients_Picture_Status174()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command with MaxSize element to resolve the special recipient who has a contact photo, the photo size is great than the value of MaxSize element.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Picture = new Request.ResolveRecipientsOptionsPicture
                            {
                                MaxSizeSpecified = true,
                                MaxSize = 1
                            }
                        },
                        Common.GetConfigurationPropertyValue("User2Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4312");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4312
            Site.CaptureRequirementIfAreEqual<string>(
                "174",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Status,
                4312,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 174 [is] The contact photo exceeded the size limit set by the MaxSize element (section 2.2.3.95.1).");
        }

        /// <summary>
        /// This test case is used to verify ResolveRecipients command, if number of contact photos returned exceeded the size limit set by the MaxPictures element, the status is equal to 175.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC19_ResolveRecipients_Picture_Status175()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Picture element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Call ResolveRecipients command with MaxPictures element set to "0" to resolve the special recipient who has one contact photo.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest();
            Request.ResolveRecipients requestResolveRecipients = new Request.ResolveRecipients();
            Request.ResolveRecipientsOptions requestResolveRecipientsOption = new Request.ResolveRecipientsOptions
            {
                Picture = new Request.ResolveRecipientsOptionsPicture
                {
                    MaxPicturesSpecified = true,
                    MaxPictures = 0
                }
            };
            requestResolveRecipients.Items = new object[] { requestResolveRecipientsOption, Common.GetConfigurationPropertyValue("User2Name", Site) };
            resolveRecipientsRequest.RequestData = requestResolveRecipients;
            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4313");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4313
            Site.CaptureRequirementIfAreEqual<string>(
                "175",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Status,
                4313,
                @"[In Status(ResolveRecipients)] [The meaning of the status value] 175 [is] The number of contact photos returned exceeded the size limit set by the MaxPictures element (section 2.2.3.94.1).");
        }

        /// <summary>
        /// This test case is used to verify ResolveRecipients command, if the number of To element in the request exceeded the size limit, the status is equal to 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC20_ResolveRecipients_Status5()
        {
            // Create a ResolveRecipients request with 101 recipients.
            object[] items = new object[101];
            for (int i = 0; i < 101; i++)
            {
                items[i] = "User" + i;
            }

            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = items
                }
            };

            // Call method ResolveRecipients to resolve the request with 101 To elements.
            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            if (Common.IsRequirementEnabled(7500, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R7500");

                // Verify MS-ASCMD requirement: MS-ASCMD_R7500
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "1",
                    resolveRecipientsResponse.ResponseData.Status,
                    7500,
                    @"[In Appendix A: Product Behavior] Implementation does not limit the number of elements in command requests and not return the specified error if the limit is exceeded. (<17> Section 2.2.3.173: Exchange 2007 SP1 and Exchange 2010 do not limit the number of To elements in command requests.)");
            }

            if (Common.IsRequirementEnabled(7501, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R7501");

                // Verify MS-ASCMD requirement: MS-ASCMD_R7501
                this.Site.CaptureRequirementIfAreNotEqual<string>(
                    "1",
                    resolveRecipientsResponse.ResponseData.Status,
                    7501,
                    @"[In Appendix A: Product Behavior] Implementation does limit the number of elements in command requests and return the specified error if the limit is exceeded. (<17> Section 2.2.3.173: Update Rollup 6 for Exchange 2010 Service Pack 2 (SP2), Exchange 2013, and Exchange 2016 Preview do limit the number of To elements in command requests.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4270");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4270
                Site.CaptureRequirementIfAreEqual<string>(
                    "5",
                    resolveRecipientsResponse.ResponseData.Status,
                    4270,
                   @"[In Status(ResolveRecipients)] [The meaning of the status value 5 is] Either an invalid parameter was specified or the range exceeded limits.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5656");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5656
                Site.CaptureRequirementIfAreEqual<string>(
                    "5",
                    resolveRecipientsResponse.ResponseData.Status,
                    5656,
                   @"[In Limiting Size of Command Requests] In ResolveRecipients (section 2.2.2.13) command request, when the limit value of To element is bigger than 100 (minimum 1, maximum 2,147,483,647), the error returned by server is Status element (section 2.2.3.162.11) value of 5.");
            }
        }

        /// <summary>
        /// This test case is used to verify ResolveRecipients command, if free/busy data could not be retrieved from the server for a given recipient, the status is equal to 163.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S13_TC21_ResolveRecipients_Status163()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The StartTime element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Get the current value of user1's AccessRights.
            string originalAccessRights = this.GetMailboxFolderPermission(Common.GetConfigurationPropertyValue("SutComputerName", this.Site), this.User3Information);
            Site.Assert.IsFalse(string.IsNullOrEmpty(originalAccessRights), "The AccessRights property should have a valid value.");

            // Set the AccessRights to None.
            this.SetMailboxFolderPermission(Common.GetConfigurationPropertyValue("SutComputerName", this.Site), this.User3Information, "None");
            string currentAccessRights = this.GetMailboxFolderPermission(Common.GetConfigurationPropertyValue("SutComputerName", this.Site), this.User3Information);
            Site.Assert.AreEqual<string>("None", currentAccessRights, "The value of AccessRights should be None.");

            // The client calls ResolveRecipients command with valid value of StartTime and EndTime.
            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest
            {
                RequestData = new Request.ResolveRecipients
                {
                    Items = new object[]
                    {
                        new Request.ResolveRecipientsOptions
                        {
                            Availability = new Request.ResolveRecipientsOptionsAvailability
                            {
                                StartTime = "2010-03-01T00:00:00.000Z",
                                EndTime = "2010-03-01T00:30:00.000Z"
                            }
                        },
                        Common.GetConfigurationPropertyValue("User3Name", Site)
                    }
                }
            };

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);

            if (originalAccessRights != currentAccessRights)
            {
                // Restore the original value of AccessRights property.
                this.SetMailboxFolderPermission(Common.GetConfigurationPropertyValue("SutComputerName", this.Site), this.User3Information, originalAccessRights);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4298");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4298
            Site.CaptureRequirementIfAreEqual<string>(
                "163",
                resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Availability.Status,
                4298,
                @"[In Status(ResolveRecipients)][The meaning of the status value] 163 [is] Free/busy data could not be retrieved from the server for a given recipient.");
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Create a fixed length string.
        /// </summary>
        /// <param name="length">The length of the string</param>
        /// <returns>Return a string</returns>
        private static string CreateFixedLengthString(int length)
        {
            int integer = length / 7;
            int remainder = length % 7;
            string createString = string.Empty;
            for (int i = 0; i < integer; i++)
            {
                createString += "MSASCMD";
            }

            for (int i = 0; i < remainder; i++)
            {
                createString += "a";
            }

            return createString;
        }
        #endregion
    }
}