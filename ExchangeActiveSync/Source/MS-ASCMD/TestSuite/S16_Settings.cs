namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the Settings command.
    /// </summary>
    [TestClass]
    public class S16_Settings : TestSuiteBase
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
        /// This test case is used to verify OOF status using Settings command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC01_Settings_Oof_Success()
        {
            #region Creates Settings request
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    Oof = new Request.SettingsOof
                    {
                        Item = new Request.SettingsOofSet { OofState = Request.OofState.Item2, OofStateSpecified = true }
                    }
                }
            };
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R4392, MS-ASCMD_R4396
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4392");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4392
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                settingsResponse.ResponseData.Status,
                4392,
                @"[In Status(Settings)] [The meaning of the status value] 1 [is] Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4396");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4396
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                settingsResponse.ResponseData.Oof.Status,
                4396,
                @"[In Status(Settings)] [The meaning of the status value] 1 [is] Success.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the empty DevicePassword using Settings command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC02_Settings_DevicePassword_Status2()
        {
            #region Creates invalid Settings request
            // The client calls Settings command to send a bad settings request. Set element under DevicePassword is present, but empty.
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    DevicePassword = new Request.SettingsDevicePassword
                    {
                        Item = new Request.SettingsDevicePasswordSet()
                    }
                }
            };
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R4398, MS-ASCMD_R483, MS-ASCMD_R5845
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4398");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4398
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                settingsResponse.ResponseData.Status,
                4398,
                @"[In Status(Settings)] [The meaning of the status value 2 is]The XML code is formatted incorrectly.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5845");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5845
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                settingsResponse.ResponseData.Status,
                5845,
                @"[In Status(Settings)] [The meaning of the status value] 2 [is] The XML code is formatted incorrectly.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R483");

            // Verify MS-ASCMD requirement: MS-ASCMD_R483
            Site.CaptureRequirementIfIsTrue(
                settingsResponse.ResponseData.Status != "1" && settingsResponse.ResponseData.DeviceInformation == null && settingsResponse.ResponseData.DevicePassword == null && settingsResponse.ResponseData.Oof == null && settingsResponse.ResponseData.RightsManagementInformation == null && settingsResponse.ResponseData.UserInformation == null,
                483,
                @"[In Settings] If the command was not successful, the processing of the request cannot begin, no property responses are returned, and the Status node MUST indicate a protocol error.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if either the StartTime element or the EndTime element is included in the request without the other, server return status 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC03_Settings_OofState()
        {
            #region Creates Settings request with only StartTime element
            // The client calls Settings command with only setting StartTime element.
            SettingsRequest settingsRequestWithStartTimeOnly = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    Oof = new Request.SettingsOof
                    {
                        Item = new Request.SettingsOofSet
                        {
                            OofState = Request.OofState.Item2,
                            OofStateSpecified = true,
                            StartTime = DateTime.Today,
                            StartTimeSpecified = true
                        }
                    }
                }
            };
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseWithStartTimeOnly = this.CMDAdapter.Settings(settingsRequestWithStartTimeOnly);
            #endregion

            #region Verify requirements MS-ASCMD_R2272, MS-ASCMD_R3992
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2272");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2272
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                settingsResponseWithStartTimeOnly.ResponseData.Oof.Status,
                2272,
                @"[In EndTime(Settings)] If [either] the StartTime element [or the EndTime element] is included in the request without the other, a Status element (section 2.2.3.162.14) value of 2 is returned as a child of the Oof element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3992");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3992
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                settingsResponseWithStartTimeOnly.ResponseData.Oof.Status,
                3992,
                @"[In StartTime(Settings)] If [either] the StartTime element [or the EndTime element] is included in the request without the other, a Status element (section 2.2.3.162.14) value of 2 is returned as a child element of the Oof element.");
            #endregion

            #region Creates Settings request with only EndTime element
            // The client calls Settings command with only setting EndTime element.
            SettingsRequest settingsRequestWithEndTimeOnly = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    Oof = new Request.SettingsOof
                    {
                        Item = new Request.SettingsOofSet
                        {
                            OofState = Request.OofState.Item2,
                            OofStateSpecified = true,
                            EndTime = DateTime.Today,
                            EndTimeSpecified = true
                        }
                    }
                }
            };
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseWithEndTimeOnly = this.CMDAdapter.Settings(settingsRequestWithEndTimeOnly);
            #endregion

            #region requirements MS-ASCMD_R5681, MS-ASCMD_R5827
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5681");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5681
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                settingsResponseWithEndTimeOnly.ResponseData.Oof.Status,
                5681,
                @"[In EndTime(Settings)] If [either the StartTime element or] the EndTime element is included in the request without the other, a Status element (section 2.2.3.162.14) value of 2 is returned as a child of the Oof element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5827");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5827
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                settingsResponseWithEndTimeOnly.ResponseData.Oof.Status,
                5827,
                @"[In StartTime(Settings)] If [either the StartTime element or] the EndTime element is included in the request without the other, a Status element (section 2.2.3.162.14) value of 2 is returned as a child element of the Oof element.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify using Settings command to get user information.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC04_Settings_UserInformation()
        {
            #region Creates Settings request to get user information
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    UserInformation = new Request.SettingsUserInformation { Item = string.Empty }
                }
            };
            #endregion

            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);

            #region Verify Requirements MS-ASCMD_R5898, MS-ASCMD_R4387, MS-ASCMD_R5901, MS-ASCMD_R5905, MS-ASCMD_R5915, MS-ASCMD_R733, MS-ASCMD_R3679, MS-ASCMD_R5178, MS-ASCMD_R5844, MS-ASCMD_R5169, MS-ASCMD_R5843
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4387");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4387
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                settingsResponse.ResponseData.UserInformation.Status,
                4387,
                @"[In Status(Settings)] [The meaning of the status value] 1 [is] Success.");

            if (Common.IsRequirementEnabled(5898, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5898");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5898
                Site.CaptureRequirementIfIsFalse(
                    settingsResponse.ResponseDataXML.Contains("<AccountId>"),
                    5898,
                    @"[In Appendix A: Product Behavior] The implementation does not return this element [AccountId] in Settings command responses. (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(5901, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5901");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5901
                Site.CaptureRequirementIfIsFalse(
                    settingsResponse.ResponseDataXML.Contains("<AccountName>"),
                    5901,
                    @"[In Appendix A: Product Behavior] The implementation does not return this element [AccountName] in Settings command responses. (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(5905, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5905");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5905
                Site.CaptureRequirementIfIsFalse(
                    settingsResponse.ResponseDataXML.Contains("<SendDisabled>"),
                    5905,
                    @"[In Appendix A: Product Behavior] The implementation does not return this element [SendDisabled] in Settings command responses. (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(5915, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5915");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5915
                Site.CaptureRequirementIfIsFalse(
                    settingsResponse.ResponseDataXML.Contains("<UserDisplayName>"),
                    5915,
                    @"[In Appendix A: Product Behavior] The implementation does not return this element [UserDisplayName] in Settings command responses. (Exchange 2007 and above follow this behavior.)");
            }

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R733");

                // Verify MS-ASCMD requirement: MS-ASCMD_R733
                // If response contains Accounts element under Get element, then MS-ASCMD_R733 is verified
                Site.CaptureRequirementIfIsNotNull(
                    settingsResponse.ResponseData.UserInformation.Get.Accounts,
                    733,
                    @"[In Accounts] The Accounts element<10> is an optional child element of the Get element in Settings command responses that contains all aggregate accounts that the user subscribes to.");

                Site.Assert.AreEqual<int>(1, settingsResponse.ResponseData.UserInformation.Get.Accounts.Length, "Server should return one account element");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3679");

                // Verify MS-ASCMD requirement: MS-ASCMD_R3679
                // If response contains PrimarySmtpAddress element under Get element, then MS-ASCMD_R3679 is verified
                Site.CaptureRequirementIfIsNotNull(
                    settingsResponse.ResponseData.UserInformation.Get.Accounts[0].EmailAddresses.PrimarySmtpAddress,
                    3679,
                    @"[In PrimarySmtpAddress] The PrimarySmtpAddress element<67> is an optional child element of the EmailAddresses element in Settings command responses that specifies the primary SMTP address for the given account.");
            }
            else
            {
                int getElementStarIndex = settingsResponse.ResponseDataXML.IndexOf("<Get>", StringComparison.OrdinalIgnoreCase);
                int getElementEndIndex = settingsResponse.ResponseDataXML.IndexOf("</Get>", StringComparison.OrdinalIgnoreCase);
                int emailAddressElementIndex = settingsResponse.ResponseDataXML.IndexOf("<EmailAddresses>", StringComparison.OrdinalIgnoreCase);

                if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5178");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R5178
                    Site.CaptureRequirementIfIsTrue(
                        emailAddressElementIndex > getElementStarIndex && emailAddressElementIndex < getElementEndIndex,
                        5178,
                        @"[In Appendix A: Product Behavior] <39> Section 2.2.3.75: The EmailAddresses element is only supported as a child element of the Get element when the MS-ASProtocolVersion header is set to 12.1 [or 14.0].");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5169");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R5169
                    Site.CaptureRequirementIfIsTrue(
                        emailAddressElementIndex > getElementStarIndex && emailAddressElementIndex < getElementEndIndex,
                        5169,
                        @"[In Appendix A: Product Behavior] <33> Section 2.2.3.54: The EmailAddresses element is only supported as a child element of the Get element when the MS-ASProtocolVersion header is set to 12.1 [or 14.0].");
                }

                if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5844");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R5844
                    Site.CaptureRequirementIfIsTrue(
                        emailAddressElementIndex > getElementStarIndex && emailAddressElementIndex < getElementEndIndex,
                        5844,
                        @"[In Appendix A: Product Behavior] <39> Section 2.2.3.75: The EmailAddresses element is only supported as a child element of the Get element when the MS-ASProtocolVersion header is set to [12.1 or] 14.0.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5843");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R5843
                    Site.CaptureRequirementIfIsTrue(
                        emailAddressElementIndex > getElementStarIndex && emailAddressElementIndex < getElementEndIndex,
                        5843,
                        @"[In Appendix A: Product Behavior] <33> Section 2.2.3.54: The EmailAddresses element is only supported as a child element of the Get element when the MS-ASProtocolVersion header is set to [12.1 or] 14.0.");
                }
            }
            #endregion
        }

        /// <summary>
        /// This test case is used to specify the OOF setting to known external users.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC05_Settings_OofSetToExternalKnown()
        {
            #region Creates Settings request to enable OOF setting to known external users
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();
            settingsOofSet.OofState = Request.OofState.Item2;
            settingsOofSet.OofStateSpecified = true;

            // Set OofMessage
            string bodyType = "TEXT";
            string enabled = "1";
            string replyMessage = "Out of office";
            Request.OofMessage setEnableToExternalUser = CreateOofMessage(bodyType, enabled, replyMessage);
            setEnableToExternalUser.AppliesToExternalKnown = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToExternalUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");
            this.IsOofSettingsChanged = true;
            #endregion

            #region Get AppliesToExternalKnown OOF message
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            Response.OofMessage oofMessage = null;
            do
            {
                Thread.Sleep(waitTime);
                SettingsResponse settingsResponse = this.GetOofSettings();
                oofMessage = TestSuiteBase.GetAppliesToExternalKnownOofMessage(settingsResponse);
                counter++;
            }
            while (counter < retryCount &&
                (oofMessage.AppliesToExternalKnown == null ||
                oofMessage.Enabled != enabled ||
                oofMessage.ReplyMessage.Trim() != replyMessage ||
                oofMessage.BodyType != bodyType));

            Site.Assert.AreEqual<string>(enabled, oofMessage.Enabled, "The oof message settings for external known users should be enabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(replyMessage, oofMessage.ReplyMessage.Trim(), "The reply message to external known users should be set successfully. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(bodyType, oofMessage.BodyType, "The body type of the oof settings for external known users should be set successfully. Retry count: {0}", counter);
            #endregion

            #region Verify requirements MS-ASCMD_R783, MS-ASCMD_R785, MS-ASCMD_R786, MS-ASCMD_R787, MS-ASCMD_R2251
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R783");

            // Verify MS-ASCMD requirement: MS-ASCMD_R783
            Site.CaptureRequirementIfIsTrue(
                oofMessage.AppliesToExternalKnown != null && oofMessage.Enabled != null && oofMessage.ReplyMessage != null && oofMessage.BodyType != null,
                783,
                @"[In AppliesToExternalKnown] When the AppliesToExternalKnown element is present, its [AppliesToExternalKnown's] peer elements (that is, the other elements within the OofMessage element) specify the OOF settings with regard to known external users.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R785");

            // Verify MS-ASCMD requirement: MS-ASCMD_R785
            Site.CaptureRequirementIfAreEqual<string>(
                enabled,
                oofMessage.Enabled,
                785,
                @"[In AppliesToExternalKnown] The following are the peer elements of the AppliesToExternalKnown element: Enabled (section 2.2.3.56)-Specifies whether an OOF message is sent to this audience while the sending user is OOF.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R786");

            // Verify MS-ASCMD requirement: MS-ASCMD_R786
            Site.CaptureRequirementIfAreEqual<string>(
                replyMessage,
                oofMessage.ReplyMessage.Trim(),
                786,
                @"[In AppliesToExternalKnown] [The following are the peer elements of the AppliesToExternalKnown element:] ReplyMessage (section 2.2.3.136)-Specifies the OOF reply message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R787");

            // Verify MS-ASCMD requirement: MS-ASCMD_R787
            Site.CaptureRequirementIfIsTrue(
                oofMessage.BodyType.Equals("TEXT", StringComparison.OrdinalIgnoreCase) || oofMessage.BodyType.Equals("HTML", StringComparison.OrdinalIgnoreCase),
                787,
                @"[In AppliesToExternalKnown] [The following are the peer elements of the AppliesToExternalKnown element:] BodyType (section 2.2.3.17)-Specifies the format of the OOF message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2251");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2251
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                oofMessage.Enabled,
                2251,
                @"[In Enabled] The value of the Enabled element is 1 if an OOF message is sent while the sending user is OOF.");
            #endregion
        }

        /// <summary>
        /// This test case is used to specify the OOF setting to unknown external users.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC06_Settings_OofSetToExternalUnknown()
        {
            #region Creates Settings request to enable oof setting to unknown external users
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();
            settingsOofSet.OofState = Request.OofState.Item2;
            settingsOofSet.OofStateSpecified = true;

            // Set OofMessage
            string bodyType = "TEXT";
            string enabled = "1";
            string replyMessage = "Out of office";
            Request.OofMessage setEnableToUnknownExternalUser = CreateOofMessage(bodyType, enabled, replyMessage);
            setEnableToUnknownExternalUser.AppliesToExternalUnknown = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToUnknownExternalUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");
            this.IsOofSettingsChanged = true;
            #endregion

            #region Get AppliesToExternalUnknown OOF message
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            Response.OofMessage oofMessage = null;
            do
            {
                Thread.Sleep(waitTime);
                SettingsResponse settingsResponse = this.GetOofSettings();
                oofMessage = TestSuiteBase.GetAppliesToExternalUnknownOofMessage(settingsResponse);
                counter++;
            }
            while (counter < retryCount &&
                (oofMessage.AppliesToExternalUnknown == null ||
                oofMessage.Enabled != enabled ||
                oofMessage.ReplyMessage.Trim() != replyMessage ||
                oofMessage.BodyType != bodyType));

            Site.Assert.AreEqual<string>(enabled, oofMessage.Enabled, "The oof message settings for unknown external users should be enabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(replyMessage, oofMessage.ReplyMessage.Trim(), "The reply message to unknown external users should be set successfully. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(bodyType, oofMessage.BodyType, "The body type of the oof settings for unknown external users should be set successfully. Retry count: {0}", counter);
            #endregion

            #region Verify requirements MS-ASCMD_R5867, MS-ASCMD_R793, MS-ASCMD_R794, MS-ASCMD_R795, MS-ASCMD_R796
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5867");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5867
            Site.CaptureRequirementIfIsTrue(
                oofMessage.AppliesToExternalUnknown != null && oofMessage.Enabled.Equals("1"),
                5867,
                @"[In AppliesToExternalUnknown] [The AppliesToExternalUnknown element] indicates that the OOF message applies to unknown external users.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R793");

            // Verify MS-ASCMD requirement: MS-ASCMD_R793
            Site.CaptureRequirementIfIsTrue(
                oofMessage.AppliesToExternalUnknown != null && oofMessage.Enabled != null && oofMessage.ReplyMessage != null && oofMessage.BodyType != null,
                793,
                @"[In AppliesToExternalUnknown] When the AppliesToExternalUnknown element is present, its [AppliesToExternalUnknown] peer elements (that is, the other elements within the OofMessage element) specify the OOF settings with regard to unknown external users.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R794");

            // Verify MS-ASCMD requirement: MS-ASCMD_R794
            Site.CaptureRequirementIfAreEqual<string>(
                enabled,
                oofMessage.Enabled,
                794,
                @"[In AppliesToExternalUnknown] The following are the peer elements of the AppliesToExternalUnknown element: Enabled (section 2.2.3.56)-Specifies whether an OOF message is sent to this audience while the sending user is OOF.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R795");

            // Verify MS-ASCMD requirement: MS-ASCMD_R795
            Site.CaptureRequirementIfAreEqual<string>(
                replyMessage,
                oofMessage.ReplyMessage.Trim(),
                795,
                @"[In AppliesToExternalUnknown] [The following are the peer elements of the AppliesToExternalUnknown element:] ReplyMessage (section 2.2.3.136)-Specifies the OOF reply message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R796");

            // Verify MS-ASCMD requirement: MS-ASCMD_R796
            Site.CaptureRequirementIfIsTrue(
                oofMessage.BodyType.Equals("TEXT", StringComparison.OrdinalIgnoreCase) || oofMessage.BodyType.Equals("HTML", StringComparison.OrdinalIgnoreCase),
                796,
                @"[In AppliesToExternalUnknown] [The following are the peer elements of the AppliesToExternalUnknown element:] BodyType (section 2.2.3.17)-Specifies the format of the OOF message.");
            #endregion
        }

        /// <summary>
        /// This test case is used to specify the OOF setting to internal users.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC07_Settings_OofSetToInternal()
        {
            #region Creates Settings request to enable oof setting to internal users
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();
            settingsOofSet.OofState = Request.OofState.Item2;
            settingsOofSet.OofStateSpecified = true;

            // Set OofMessage
            string bodyType = "TEXT";
            string enabled = "1";
            string replyMessage = "Out of office";
            Request.OofMessage setEnableToInternalUser = CreateOofMessage(bodyType, enabled, replyMessage);
            setEnableToInternalUser.AppliesToInternal = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToInternalUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");
            this.IsOofSettingsChanged = true;
            #endregion

            #region Get AppliesToInternal OOF message
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            Response.OofMessage oofMessage = null;
            do
            {
                Thread.Sleep(waitTime);
                SettingsResponse settingsResponse = this.GetOofSettings();
                oofMessage = TestSuiteBase.GetAppliesToInternalOofMessage(settingsResponse);
                counter++;
            }
            while (counter < retryCount &&
                (oofMessage.AppliesToInternal == null ||
                oofMessage.Enabled != enabled ||
                oofMessage.ReplyMessage.Trim() != replyMessage ||
                oofMessage.BodyType != bodyType));

            Site.Assert.AreEqual<string>(enabled, oofMessage.Enabled, "The oof message settings for internal users should be enabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(replyMessage, oofMessage.ReplyMessage.Trim(), "The reply message to internal users should be set successfully. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(bodyType, oofMessage.BodyType, "The body type of the oof settings for internal users should be set successfully. Retry count: {0}", counter);
            #endregion

            #region Verify requirements MS-ASCMD_R5866, MS-ASCMD_R802, MS-ASCMD_R803, MS-ASCMD_R804, MS-ASCMD_R805, MS-ASCMD_R826, MS-ASCMD_R1106
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5866");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5866
            Site.CaptureRequirementIfIsTrue(
                oofMessage.AppliesToInternal != null && oofMessage.Enabled.Equals("1"),
                5866,
                @"[In AppliesToInternal] [The AppliesToInternal element] indicates that the OOF message applies to internal users.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R802");

            // Verify MS-ASCMD requirement: MS-ASCMD_R802
            Site.CaptureRequirementIfIsTrue(
                oofMessage.AppliesToInternal != null && oofMessage.Enabled != null && oofMessage.ReplyMessage != null && oofMessage.BodyType != null,
                802,
                @"[In AppliesToInternal] When the AppliesToInternal element is present, its [AppliesToInternal element] peer elements (that is, the other elements within the OofMessage element) specify the OOF settings with regard to internal users.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R803");

            // Verify MS-ASCMD requirement: MS-ASCMD_R803
            Site.CaptureRequirementIfAreEqual<string>(
                enabled,
                oofMessage.Enabled,
                803,
                @"[In AppliesToInternal] The following are the peer elements of the AppliesToInternal element: Enabled (section 2.2.3.56)-Specifies whether an OOF message is sent to this audience while the sending user is OOF.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R804");

            // Verify MS-ASCMD requirement: MS-ASCMD_R804
            Site.CaptureRequirementIfAreEqual<string>(
                replyMessage,
                oofMessage.ReplyMessage.Trim(),
                804,
                @"[In AppliesToInternal] [The following are the peer elements of the AppliesToInternal element:] ReplyMessage (section 2.2.3.136)-Specifies the OOF message itself.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R805");

            // Verify MS-ASCMD requirement: MS-ASCMD_R805
            Site.CaptureRequirementIfIsTrue(
                oofMessage.BodyType.Equals("TEXT", StringComparison.OrdinalIgnoreCase) || oofMessage.BodyType.Equals("HTML", StringComparison.OrdinalIgnoreCase),
                805,
                @"[In AppliesToInternal] [The following are the peer elements of the AppliesToInternal element:] BodyType (section 2.2.3.17)-Specifies the format of the OOF message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R826");

            // Verify MS-ASCMD requirement: MS-ASCMD_R826
            Site.CaptureRequirementIfIsTrue(
                oofMessage.ReplyMessage != null && oofMessage.BodyType != null,
                826,
                @"[In BodyType] [The BodyType element ] MUST always be present (with a non-NULL value) in a Settings command Oof Get operation response if the ReplyMessage element is returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1106");

            // Verify MS-ASCMD requirement: MS-ASCMD_R1106
            Site.CaptureRequirementIfIsTrue(
                oofMessage.ReplyMessage != null && oofMessage.BodyType != null,
                1106,
                @"[In BodyType] Element BodyType in Settings command Oof response, the number allowed is 1�1 (required) if a ReplyMessage element is present [; otherwise, 0�1 (optional)].");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify that BodyType element is required if the ReplyMessage element is present.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC08_Settings_OofSetBodyTypeIsRequiredWithReplyMessage()
        {
            #region Creates Setting request with ReplyMessage element and with BodyType element
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();
            settingsOofSet.OofState = Request.OofState.Item2;
            settingsOofSet.OofStateSpecified = true;

            // Set OofMessage
            string bodyType = "TEXT";
            string enabled = "1";
            string replyMessage = "Out of office";
            Request.OofMessage setEnableToInternalUser = CreateOofMessage(bodyType, enabled, replyMessage);
            setEnableToInternalUser.AppliesToInternal = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToInternalUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");
            this.IsOofSettingsChanged = true;
            #endregion

            #region Get AppliesToInternal OOF message
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            Response.OofMessage oofMessage = null;
            do
            {
                Thread.Sleep(waitTime);
                SettingsResponse settingsResponse = this.GetOofSettings();
                oofMessage = TestSuiteBase.GetAppliesToInternalOofMessage(settingsResponse);
                counter++;
            }
            while (counter < retryCount &&
                (oofMessage.AppliesToInternal == null ||
                oofMessage.Enabled != enabled ||
                oofMessage.ReplyMessage.Trim() != replyMessage ||
                oofMessage.BodyType != bodyType));

            Site.Assert.AreEqual<string>(enabled, oofMessage.Enabled, "The oof message settings for internal users should be enabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(replyMessage, oofMessage.ReplyMessage.Trim(), "The reply message to internal users should be set successfully. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(bodyType, oofMessage.BodyType, "The body type of the oof settings for internal users should be set successfully. Retry count: {0}", counter);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R816");

            // Verify MS-ASCMD requirement: MS-ASCMD_R816
            Site.CaptureRequirementIfIsTrue(
                settingsResponseAfterSet.ResponseData.Oof.Status.Equals("1") && oofMessage.BodyType != null && oofMessage.ReplyMessage != null,
                816,
                @"[In BodyType] It [BodyType element] is a required child element of the OofMessage element in Settings command requests and responses if a ReplyMessage element (section 2.2.3.136) is present as a child of the OofMessage element.");
        }

        /// <summary>
        /// This test case is used to verify the BodyType element is optional if the ReplyMessage element is not present.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC09_Settings_OofSetBodyTypeOptional()
        {
            #region Creates Setting request without ReplyMessage element and without BodyType element
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();
            settingsOofSet.OofState = Request.OofState.Item2;
            settingsOofSet.OofStateSpecified = true;

            // Set OofMessage
            Request.OofMessage setEnableToInternalUser = CreateOofMessage(null, "0", null);
            setEnableToInternalUser.AppliesToInternal = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToInternalUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");
            this.IsOofSettingsChanged = true;
            #endregion

            #region Get AppliesToInternal OOF message
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            Response.OofMessage oofMessage = null;
            do
            {
                Thread.Sleep(waitTime);
                SettingsResponse settingsResponse = this.GetOofSettings();
                oofMessage = TestSuiteBase.GetAppliesToInternalOofMessage(settingsResponse);
                counter++;
            }
            while (counter < retryCount &&
                (oofMessage.AppliesToInternal == null ||
                oofMessage.ReplyMessage != null ||
                oofMessage.BodyType != null ||
                oofMessage.Enabled != "0"));

            Site.Assert.AreEqual<string>("0", oofMessage.Enabled, "The oof message settings for internal users should be disenabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(null, oofMessage.ReplyMessage, "The reply message to internal users should be set to null. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(null, oofMessage.BodyType, "The body type of the oof settings for internal users should be set to null. Retry count: {0}", counter);
            #endregion

            #region Verify Requirements MS-ASCMD_R5677, MS-ASCMD_R2252
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5677");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5677
            Site.CaptureRequirementIfIsTrue(
                oofMessage.ReplyMessage == null && oofMessage.BodyType == null,
                5677,
                @"[In BodyType] [Element BodyType in Settings command Oof response, the number allowed is 1�1 (required) if a ReplyMessage element is present]; otherwise, [the number allowed is ] 0�1 (optional).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2252");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2252
            Site.CaptureRequirementIfAreEqual<string>(
                "0",
                oofMessage.Enabled,
                2252,
                @"[In Enabled][The value of the Enabled element is 1 if an OOF message is sent while the sending user is OOF]; otherwise, the value is 0 (zero).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify set OOF message for three type audiences which are external known users , external unknown users and internal users.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC10_Settings_OofSetForThreeTypeAudiences()
        {
            #region Creates Setting request for three types of audiences
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();
            settingsOofSet.OofState = Request.OofState.Item2;
            settingsOofSet.OofStateSpecified = true;

            // Set internal enabled
            string bodyType = "TEXT";
            string internalEnabled = "1";
            string internalReplyMessage = "Out of office";
            Request.OofMessage setEnableToInternalUser = CreateOofMessage(bodyType, internalEnabled, internalReplyMessage);
            setEnableToInternalUser.AppliesToInternal = string.Empty;

            // Set external known enabled
            string externalKnownBodyType = "TEXT";
            string externalKnownEnabled = "1";
            string externalKnownReplyMessage = "Call my mobile";
            Request.OofMessage setEnableToExternalKnownlUser = CreateOofMessage(externalKnownBodyType, externalKnownEnabled, externalKnownReplyMessage);
            setEnableToExternalKnownlUser.AppliesToExternalKnown = string.Empty;

            // Set external unknown enabled
            string externalUnknownBodyType = "TEXT";
            string externalUnknownEnabled = "1";
            string externalUnknownReplyMessage = "Call my mobile";
            Request.OofMessage setEnableToExternalUnknownlUser = CreateOofMessage(externalUnknownBodyType, externalUnknownEnabled, externalUnknownReplyMessage);
            setEnableToExternalUnknownlUser.AppliesToExternalUnknown = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToInternalUser, setEnableToExternalKnownlUser, setEnableToExternalUnknownlUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");
            this.IsOofSettingsChanged = true;
            #endregion

            if (Common.IsRequirementEnabled(5194, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5194");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5194
                Site.CaptureRequirementIfAreEqual<string>(
                    "1",
                    settingsResponseAfterSet.ResponseData.Oof.Status,
                    5194,
                    @"[In Appendix A: Product Behavior] <54> Section 2.2.3.113: Exchange 2007, Exchange 2010 and Exchange 2013 require that the reply message for unknown external and known external audiences be the same.");
            }

            #region Get OOF message
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            Response.OofMessage internalOofMessage = null;
            Response.OofMessage externalKnownOofMessage = null;
            Response.OofMessage externalUnknownOofMessage = null;
            do
            {
                Thread.Sleep(waitTime);
                SettingsResponse settingsResponse = this.GetOofSettings();
                internalOofMessage = TestSuiteBase.GetAppliesToInternalOofMessage(settingsResponse);
                externalKnownOofMessage = TestSuiteBase.GetAppliesToExternalKnownOofMessage(settingsResponse);
                externalUnknownOofMessage = TestSuiteBase.GetAppliesToExternalUnknownOofMessage(settingsResponse);

                counter++;
            }
            while (counter < retryCount &&
                (internalOofMessage.AppliesToInternal == null ||
                internalOofMessage.Enabled != internalEnabled ||
                internalOofMessage.ReplyMessage.Trim() != internalReplyMessage ||
                internalOofMessage.BodyType != bodyType ||
                externalKnownOofMessage.AppliesToExternalKnown == null ||
                externalKnownOofMessage.Enabled != externalKnownEnabled ||
                externalKnownOofMessage.ReplyMessage.Trim() != externalKnownReplyMessage ||
                externalKnownOofMessage.BodyType != externalKnownBodyType ||
                externalUnknownOofMessage.AppliesToExternalUnknown == null ||
                externalUnknownOofMessage.Enabled != externalUnknownEnabled ||
                externalUnknownOofMessage.ReplyMessage.Trim() != externalUnknownReplyMessage ||
                externalUnknownOofMessage.BodyType != externalUnknownBodyType));

            Site.Assert.AreEqual<string>(internalEnabled, internalOofMessage.Enabled, "The oof message settings for internal users should be enabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(internalReplyMessage, internalOofMessage.ReplyMessage.Trim(), "The reply message to internal users should be set successfully. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(bodyType, internalOofMessage.BodyType, "The body type of the oof settings for internal users should be set successfully. Retry count: {0}", counter);

            Site.Assert.AreEqual<string>(externalKnownEnabled, externalKnownOofMessage.Enabled, "The oof message settings for known external users should be enabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(externalKnownReplyMessage, externalKnownOofMessage.ReplyMessage.Trim(), "The reply message to known external users should be set successfully. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(externalKnownBodyType, externalKnownOofMessage.BodyType, "The body type of the oof settings for known external users should be set successfully. Retry count: {0}", counter);

            Site.Assert.AreEqual<string>(externalUnknownEnabled, externalUnknownOofMessage.Enabled, "The oof message settings for unknown external users should be enabled. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(externalUnknownReplyMessage, externalUnknownOofMessage.ReplyMessage.Trim(), "The reply message to unknown external users should be set successfully. Retry count: {0}", counter);
            Site.Assert.AreEqual<string>(externalUnknownBodyType, externalUnknownOofMessage.BodyType, "The body type of the oof settings for unknown external users should be set successfully. Retry count: {0}", counter);
            #endregion
        }

        /// <summary>
        /// This test case is used to test if set BodyType to invalid data, server should return status 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC11_Settings_Status5()
        {
            #region Create Oof set request with invalid arguments
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();

            // Create OofMessage with invalid bodyType value.
            string bodyType = "InvalidValue";
            string enabled = "1";
            string replyMessage = "Oof";
            Request.OofMessage setEnableToExternalUser = CreateOofMessage(bodyType, enabled, replyMessage);
            setEnableToExternalUser.AppliesToExternalKnown = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToExternalUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R5702, MS-ASCMD_R5582, MS-ASCMD_R484
            // If user send Oof set request with invalid arguments, server returns status 5 which means invalid arguments, then MS-ASCMD_R5702, MS-ASCMD_R5582 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5702");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5702
            Site.CaptureRequirementIfAreEqual<string>(
                "5",
                settingsResponseAfterSet.ResponseData.Oof.Status,
                5702,
                @"[In Status(Settings)] [The meaning of the status value] 5 [is] Invalid arguments.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5582");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5582
            Site.CaptureRequirementIfAreEqual<string>(
                "5",
                settingsResponseAfterSet.ResponseData.Oof.Status,
                5582,
                @"[In Status(Settings)] [The meaning of the status value] 5 [is] Invalid arguments.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R484");

            // Verify MS-ASCMD requirement: MS-ASCMD_R484
            Site.CaptureRequirementIfAreNotEqual<string>(
                "2",
                settingsResponseAfterSet.ResponseData.Oof.Status,
                484,
                @"[In Settings] Any error other than a protocol error is returned in the Status elements of the individual property responses.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test if set different reply message to external known users and external unknown users, server should respond status 6.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC12_Settings_Status6()
        {
            #region Creates Settings request for two types of audiences
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            Request.SettingsOofSet settingsOofSet = new Request.SettingsOofSet();

            // Set external known enabled
            string externalKnownBodyType = "TEXT";
            string externalKnownEnabled = "1";
            string externalKnownReplyMessage = "Message to external known";
            Request.OofMessage setEnableToExternalKnownlUser = CreateOofMessage(externalKnownBodyType, externalKnownEnabled, externalKnownReplyMessage);
            setEnableToExternalKnownlUser.AppliesToExternalKnown = string.Empty;

            // Set external Unknown enabled
            string externalUnknownBodyType = "TEXT";
            string externalUnknownEnabled = "1";
            string externalUnknownReplyMessage = "Message to unknown external users";
            Request.OofMessage setEnableToExternalUnknownlUser = CreateOofMessage(externalUnknownBodyType, externalUnknownEnabled, externalUnknownReplyMessage);
            setEnableToExternalUnknownlUser.AppliesToExternalUnknown = string.Empty;

            settingsOofSet.OofMessage = new Request.OofMessage[] { setEnableToExternalKnownlUser, setEnableToExternalUnknownlUser };
            settingsRequest.RequestData.Oof.Item = settingsOofSet;
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R5385, MS-ASCMD_R5583
            // If use send Oof set request with conflicting arguments, server returns status 6, then MS-ASCMD_R5385, MS-ASCMD_R5583 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5385");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5385
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                settingsResponseAfterSet.ResponseData.Oof.Status,
                5385,
                @"[In Status(Settings)] [The meaning of the status value] 6 [is] Conflicting arguments.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5583");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5583
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                settingsResponseAfterSet.ResponseData.Oof.Status,
                5583,
                @"[In Status(Settings)] [The meaning of the status value] 6 [is] Conflicting arguments.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test the implication of batching mechanism, server response should match the order of request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC13_Settings_ResponseInOrder()
        {
            #region Creates Settings request to get Oof and user information
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    Oof = new Request.SettingsOof { Item = new Request.SettingsOofGet { BodyType = "TEXT" } },
                    UserInformation = new Request.SettingsUserInformation { Item = string.Empty }
                }
            };
            #endregion

            int oofElementRequestPosition = settingsRequest.GetRequestDataSerializedXML().IndexOf("<Oof>", StringComparison.OrdinalIgnoreCase);
            int userInfomationRequestPosition = settingsRequest.GetRequestDataSerializedXML().IndexOf("<UserInformation>", StringComparison.OrdinalIgnoreCase);

            #region Calls Settings command
            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponse.ResponseData.Oof.Status, "Server should response status 1, if set enabled successful");
            Site.Assert.IsTrue(settingsResponse.ResponseData.Oof != null && settingsResponse.ResponseData.UserInformation != null && settingsResponse.ResponseData.DeviceInformation == null && settingsResponse.ResponseData.DevicePassword == null && settingsResponse.ResponseData.RightsManagementInformation == null, "Server should response only requested data");
            #endregion

            int oofElementResponsePosition = settingsResponse.ResponseDataXML.IndexOf("<Oof>", StringComparison.OrdinalIgnoreCase);
            int userInformationResponsePosition = settingsResponse.ResponseDataXML.IndexOf("<UserInformation>", StringComparison.OrdinalIgnoreCase);

            #region Verify Requirements MS-ASCMD_R469, MS-ASCMD_R5713, MS-ASCMD_R480, MS-ASCMD_R481
            // If server response contains user settings properties in sequence order as required in request, then MS-ASCMD_R469, MS-ASCMD_R5713, MS-ASCMD_R480, MS-ASCMD_R481 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R469");

            // Verify MS-ASCMD requirement: MS-ASCMD_R469
            Site.CaptureRequirementIfIsTrue(
                oofElementRequestPosition < userInfomationRequestPosition && oofElementResponsePosition < userInformationResponsePosition,
                469,
                @"[In Settings] The implication of this batching mechanism is that commands are executed in the order in which they are received.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5713");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5713
            Site.CaptureRequirementIfIsTrue(
                oofElementRequestPosition < userInfomationRequestPosition && oofElementResponsePosition < userInformationResponsePosition,
                5713,
                @"[In Settings] The implication of this batching mechanism is that the ordering of Get and Set responses will match the order of those commands in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R480");

            // Verify MS-ASCMD requirement: MS-ASCMD_R480
            Site.CaptureRequirementIfIsTrue(
                oofElementRequestPosition < userInfomationRequestPosition && oofElementResponsePosition < userInformationResponsePosition,
                480,
                @"[In Settings] The server will return responses in the same order [the order of elements in the Settings request] in which they [responses] were requested.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R481");

            // Verify MS-ASCMD requirement: MS-ASCMD_R481
            Site.CaptureRequirementIfIsTrue(
                oofElementRequestPosition < userInfomationRequestPosition && oofElementResponsePosition < userInformationResponsePosition,
                481,
                @"[In Settings] Each response message contains a Status element (section 2.2.3.162.14) value for the command, which addresses the success or failure of the Settings command, followed by Status values for each of the changes made to the Oof, DeviceInformation, DevicePassword or UserInformation elements.");
            #endregion

            // Set OOF state to disabled
            this.SetOofDisabled();
        }

        /// <summary>
        /// This test case is used to test the RightsManagementInformation element in Settings response and request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC14_Settings_GetRightManagementInformation()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The RightsManagementInformation element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The RightsManagementInformation element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.SwitchUser(this.User3Information);
            #region Creates Settings request to get RightsManagementInformation
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    RightsManagementInformation = new Request.SettingsRightsManagementInformation
                    {
                        Get = string.Empty
                    }
                }
            };
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3863");

            if (settingsResponse.ResponseData.RightsManagementInformation != null)
            {
                // If RightsManagementInformation element is returned, R3863 can be verified.
                Site.CaptureRequirement(
                    3863,
                    @"[In RightsManagementInformation] The RightsManagementInformation element<68> is an optional child element of the Settings element in Settings command requests and responses.");
            }
        }

        /// <summary>
        /// This test case is used to test when the EnableOutboundSMS element is set to 1, the PhoneNumber is required.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC15_Settings_SetDeviceInformationPhoneNumber()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5201, this.Site), "Only on the initial release version of Exchange 2010, the PhoneNumber element is required to have a value, when the EnableOutboundSMS element is set to 1.");

            #region Set EnableOutboundSMS value to "1" and set PhoneNumber value
            Request.DeviceInformation deviceInformation = TestSuiteBase.GenerateDeviceInformation();
            deviceInformation.Set.EnableOutboundSMS = "1";
            deviceInformation.Set.PhoneNumber = "88888888888";

            SettingsRequest settingsRequest = CreateDefaultDeviceInformationRequest();
            settingsRequest.RequestData.DeviceInformation = deviceInformation;

            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5201");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5201
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                settingsResponse.ResponseData.DeviceInformation.Status,
                5201,
                @"[In Appendix A: Product Behavior] <62> Section 2.2.3.124: When the EnableOutboundSMS element is set to 1 and the MS-ASProtocolVersion header is set to 14.0, the PhoneNumber element is required to have a value.");

            #region Record the notification email from SUT.
            string emailSubject = "Send and receive mobile text messages on your computer!";
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            string serverId = string.Empty;

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            SyncResponse syncResponse = this.Sync(syncRequest);
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;

            do
            {
                syncResponse = this.Sync(syncRequest);

                if (!string.IsNullOrEmpty(syncResponse.ResponseDataXML))
                {
                    serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", emailSubject);

                    if (!string.IsNullOrEmpty(serverId))
                    {
                        break;
                    }
                }

                System.Threading.Thread.Sleep(waitTime);
                counter++;
            }
            while (counter < retryCount);

            Site.Assert.IsFalse(string.IsNullOrEmpty(serverId), "The notification message named '{0}' should be found. Retry count: {1}.", emailSubject, counter);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, emailSubject);
            #endregion

            #region Restore the default settings on SUT.
            deviceInformation.Set.EnableOutboundSMS = "0";
            settingsRequest.RequestData.DeviceInformation = deviceInformation;
            settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponse.ResponseData.DeviceInformation.Status, "The default settings should be restored successfully.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if set password length exceeds 256, server should return status 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC16_Settings_DevicePassword_LongPassword()
        {
            #region Creates Settings request
            // The client calls Settings command to send a settings request. Set element DevicePassword.
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    DevicePassword = new Request.SettingsDevicePassword
                    {
                        Item = new Request.SettingsDevicePasswordSet
                        {
                            Password = "password".PadRight(256, 'a')
                        }
                    }
                }
            };
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4400");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4400
            Site.CaptureRequirementIfAreEqual<string>(
                "5",
                settingsResponse.ResponseData.DevicePassword.Status,
                4400,
                @"[In Status(Settings)] [The meaning of the status value 5 is] The specified password is too long.");
        }

        /// <summary>
        /// This test case is used to verify if OofState not set to 2 and both StartTime and EndTime are present, server returns a successful response.
        /// </summary>        
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S16_TC17_Settings_OofStateNotTwo()
        {
            #region Creates Settings request with only StartTime element
            // The client calls Settings command with both StartTime and EndTime elements, but OofState not set as 2.
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings
                {
                    Oof = new Request.SettingsOof
                    {
                        Item = new Request.SettingsOofSet
                        {
                            OofState = Request.OofState.Item1,
                            OofStateSpecified = true,
                            StartTime = DateTime.Today,
                            StartTimeSpecified = true,
                            EndTime = DateTime.Today.AddDays(1),
                            EndTimeSpecified = true
                        }
                    }
                }
            };
            #endregion

            #region Calls Settings command
            SettingsResponse settingsResponse = this.CMDAdapter.Settings(settingsRequest);
            #endregion Calls Settings command

            Site.CaptureRequirementIfAreEqual(
                "1",
                settingsResponse.ResponseData.Oof.Status,
                3535,
                @"[In OofState] If the OofState element value is not set to 2 and the StartTime and EndTime elements are submitted in the request, the client does receive a successful response message.");
        }
        #endregion

        #region Private Method
        /// <summary>
        /// Create one DeviceInformation request with empty Settings
        /// </summary>
        /// <returns>The settings request</returns>
        private static SettingsRequest CreateDefaultDeviceInformationRequest()
        {
            SettingsRequest settingsRequest = new SettingsRequest
            {
                RequestData = new Request.Settings { DeviceInformation = new Request.DeviceInformation() }
            };
            return settingsRequest;
        }

        /// <summary>
        /// Set OOF disabled for internal and external users
        /// </summary>
        private void SetOofDisabled()
        {
            SettingsRequest settingsRequest = CreateDefaultOofRequest();
            settingsRequest.RequestData.Oof.Item = new Request.SettingsOofSet { OofState = Request.OofState.Item0, OofStateSpecified = true };

            SettingsResponse settingsResponseAfterSet = this.CMDAdapter.Settings(settingsRequest);
            Site.Assert.AreEqual<string>("1", settingsResponseAfterSet.ResponseData.Oof.Status, "Server should response status 1, if OOF disabled successful");
        }
        #endregion
    }
}