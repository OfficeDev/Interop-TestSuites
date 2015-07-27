//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is used to test the common status codes.
    /// </summary>
    [TestClass]
    public class S21_CommonStatusCode : TestSuiteBase
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
        /// This test case is used to verify the server will return 166, when AccountId is invalid.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S21_TC01_CommonStatusCode_166()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 166 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 166 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call method SendMail to send e-mail messages with invalid AccountID value.
            string emailSubject = Common.GenerateResourceName(Site, "subject");

            // Send email with invalid AccountID value
            SendMailResponse sendMailResponse = this.SendPlainTextEmail("InvalidAccountID", emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4956");
            Site.Log.Add(LogEntryKind.Debug, "When sending mail with invalid AccountID, server returns status {0}", sendMailResponse.ResponseData.Status);

            // Verify MS-ASCMD requirement: MS-ASCMD_R4956
            Site.CaptureRequirementIfAreEqual<string>(
                "166",
                sendMailResponse.ResponseData.Status,
                4956,
                @"[In Common Status Codes] [The meaning of the status value 166 is] The AccountId (section 2.2.3.3) value is not valid.<100>");

            #region Sync user2 mailbox changes
            // Switch to user2's mailbox
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);

            // Record user name, folder collectionId and item subject that is used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify server will return 173, when the picture does not exist.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S21_TC02_CommonStatusCode_173()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 173 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 173 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call method ResolveRecipients to resolve a list of supplied recipients, to retrieve their free/busy information, or retrieve their S/MIME certificates so that clients can send encrypted S/MIME e-mail messages.
            string displayName = this.User3Information.UserName;

            ResolveRecipientsRequest resolveRecipientsRequest = new ResolveRecipientsRequest();
            Request.ResolveRecipients requestResolveRecipients = new Request.ResolveRecipients();

            Request.ResolveRecipientsOptions requestResolveRecipientsOption = new Request.ResolveRecipientsOptions
            {
                Picture = new Request.ResolveRecipientsOptionsPicture { MaxPictures = 3 }
            };

            requestResolveRecipients.Items = new object[] { requestResolveRecipientsOption, displayName };
            resolveRecipientsRequest.RequestData = requestResolveRecipients;

            ResolveRecipientsResponse resolveRecipientsResponse = this.CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4970");
            Site.Log.Add(LogEntryKind.Debug, "When the contact picture does not exit, server returns status {0}", resolveRecipientsResponse.ResponseData.Response.Recipient[0].Picture[0].Status);

            // Verify MS-ASCMD requirement: MS-ASCMD_R4970
            Site.CaptureRequirementIfAreEqual<string>(
                "173",
                resolveRecipientsResponse.ResponseData.Response.Recipient[0].Picture[0].Status,
                4970,
                @"[In Common Status Codes] [The meaning of the status value 173 is] The user does not have a contact photo.<107>");
        }

        /// <summary>
        /// This test case is used to verify the server will return 165, when the required DeviceInformation element is missing in the Provision request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S21_TC03_CommonStatusCode_165()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 165 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 165 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User calls Provision command without the DeviceInformation element

            ProvisionRequest provisionRequest = TestSuiteBase.GenerateDefaultProvisionRequest();
            provisionRequest.RequestData.DeviceInformation = null;

            ProvisionResponse provisionResponse = this.CMDAdapter.Provision(provisionRequest);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4954");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4954
            Site.CaptureRequirementIfAreEqual<byte>(
                165,
                provisionResponse.ResponseData.Status,
                4954,
                @"[In Common Status Codes] [The meaning of the status value 165 is] The required DeviceInformation element (as specified in [MS-ASPROV] section 2.2.2.52) is missing in the Provision request.<99>");
        }

        /// <summary>
        /// This test case is used to verify the server will return 105, when the request contains a combination of parameters that is invalid.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S21_TC04_CommonStatusCode_105()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The DstFldId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User2 sends mail to User1 and do FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server, and get the ServerId of sent email item and the SyncKey
            SyncResponse syncResponseInbox = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponseInbox, "Subject", subject);
            #endregion

            #region Call method MoveItems with the email item's ServerId to move the email item from Inbox folder to recipient information cache.
            Request.MoveItemsMove moveItemsMove = new Request.MoveItemsMove
            {
                DstFldId = this.User1Information.RecipientInformationCacheCollectionId,
                SrcFldId = this.User1Information.InboxCollectionId,
                SrcMsgId = serverId
            };

            MoveItemsRequest moveItemsRequest = Common.CreateMoveItemsRequest(new Request.MoveItemsMove[] { moveItemsMove });
            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4821");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4821
            Site.CaptureRequirementIfAreEqual<byte>(
                105,
                moveItemsResponse.ResponseData.Response[0].Status,
                4821,
                @"[In Common Status Codes] [The meaning of the status value 105 is] The request contains a combination of parameters that is invalid.");
        }

        /// <summary>
        /// This test case is used to verify the server returns 164, when the BodyPartPreference node has an unsupported Type element value.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S21_TC05_CommonStatusCode_164()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 164 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Status value 164 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User calls Sync command with option element

            // Set an unsupported Type element value in the BodyPartPreference node 
            Request.Options option = new Request.Options
            {
                Items = new object[]
                {
                    new Request.BodyPartPreference()
                    {
                        // As specified in [MS-ASAIRS] section 2.2.2.22.3, only a value of 2 (HTML) SHOULD be used in the Type element of a BodyPartPreference element.
                        // Then '3' is an unsupported Type element value.
                        Type = 3
                    }
                },
                ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.BodyPartPreference }
            };

            SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
            this.Sync(syncRequest);

            syncRequest.RequestData.Collections[0].Options = new Request.Options[] { option };
            syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
            SyncResponse syncResponse = this.Sync(syncRequest);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5412");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5412
            Site.CaptureRequirementIfAreEqual<byte>(
                164,
                syncResponse.ResponseData.Status,
                5412,
                @"[In Common Status Codes] [The meaning of the status value 164 is] The BodyPartPreference node (as specified in [MS-ASAIRS] section 2.2.2.7) has an unsupported Type element (as specified in [MS-ASAIRS] section 2.2.2.22.4) value.<98>");
        }

        /// <summary>
        /// This test case is used to verify the server returns 118, when the message was already sent in a previous request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S21_TC06_CommonStatusCode_118()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User1 calls SendMail command to send email messages to user2.

            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string from = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string to = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string content = Common.GenerateResourceName(Site, "Default Email");
            string mime = Common.CreatePlainTextMime(from, to, null, null, emailSubject, content);
            SendMailRequest sendMailRequest = Common.CreateSendMailRequest(TestSuiteBase.ClientId, false, mime);
            SendMailResponse responseSendMail = this.CMDAdapter.SendMail(sendMailRequest);
            Site.Assert.AreEqual<string>(
                string.Empty,
                responseSendMail.ResponseDataXML,
                "The server should return an empty xml response data to indicate SendMail command success.");
            
            this.SwitchUser(this.User2Information);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            this.SwitchUser(this.User1Information);
            #endregion

            #region User1 calls SendMail command with the same ClientId again.

            // Use the same ClientId to call SendMail command again
            responseSendMail = this.CMDAdapter.SendMail(sendMailRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4848");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4848
            Site.CaptureRequirementIfAreEqual<string>(
                "118",
                responseSendMail.ResponseData.Status,
                4848,
                @"[In Common Status Codes] [The meaning of the status value 118 is] The message was already sent in a previous request.");

            #endregion
        }
        #endregion
    }
}