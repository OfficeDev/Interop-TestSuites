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

    /// <summary>
    /// This scenario is used to test the SendMail command.
    /// </summary>
    [TestClass]
    public class S15_SendMail : TestSuiteBase
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
        /// This test case is used to verify the server returns an empty response, when mail sending successfully.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S15_TC01_SendMail_Success()
        {
            #region User1 calls SendMail command to send email messages to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            SendMailResponse sendMailResponse = this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Verify Requirements MSASCMD_R463, MSASCMD_R4379, MSASCMD_R5091
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R463");

            // If the server returns an empty response after user calls SendMail command, it means the message was sent successfully, then MS-ASCMD_R463, MS-ASCMD_R4379, MS-ASCMD_R5091 are verified.
            // Verify MS-ASCMD requirement: MS-ASCMD_R463
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                sendMailResponse.ResponseDataXML,
                463,
                @"[In Response] If the message was sent successfully, the server returns an empty response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4379");

            // If the server returns an empty response after user calls SendMail command, it means the message was sent successfully.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4379
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                sendMailResponse.ResponseDataXML,
                4379,
                @"[In Status(SendMail)] If the [SendMail] command succeeds, no XML body is returned in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5091");

            // If the server returns an empty response after user calls SendMail command, it means the message was sent successfully.
            // Verify MS-ASCMD requirement: MS-ASCMD_R5091
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                sendMailResponse.ResponseDataXML,
                5091,
                @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 5: ] If the message was sent successfully, the server returns an empty response.");
            #endregion

            #region Sync user2 mailbox changes
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            SyncResponse syncResponse = this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R443");

            // If syncReponse is not null, means user2 received the email from user1, then MS-ASCMD_R443 is verified
            // Verify MS-ASCMD requirement: MS-ASCMD_R443
            Site.CaptureRequirementIfIsNotNull(
                syncResponse,
                443,
                @"[In SendMail] The SendMail command is used by clients to send MIME-formatted email messages to the server.");

            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the server returns 166, when AccountId is invalid.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S15_TC02_SendMail_AccountIdInvalid()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The AccountID element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The AccountID element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            #region Call method SendMail to Send e-mail messages with invalid AccountID value.
            string emailSubject = Common.GenerateResourceName(Site, "subject");

            // Send email with invalid AccountID value
            SendMailResponse responseSendMail = this.SendPlainTextEmail("InvalidValueAccountID", emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Verify Requirements MS-ASCMD_R4380, MS-ASCMD_R5092, MS-ASCMD_R724
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4380");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4380
            // If server returns Status code element, means the SendMail operation meet failure, then MS-ASCMD_R4380 is verified.
            Site.CaptureRequirementIfIsNotNull(
                responseSendMail.ResponseData.Status,
                4380,
                @"[In Status(SendMail)] If the[SendMail] command fails, the Status element contains a code that indicates the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R95");

            // Verify MS-ASCMD requirement: MS-ASDTYPE_R95
            // If server returns Status code element, means the SendMail operation meet failure when the value of the AccountID element does not adhere to the expected format, then MS-ASDTYPE_R95 is verified.
            Site.CaptureRequirementIfIsNotNull(
                responseSendMail.ResponseData.Status,
                "MS-ASDTYPE",
                95,
                @"[In string Data Type] Commands that process such elements[defined as string types in XML schemas] can return an error if the value of the element does not adhere to the expected format.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5092");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5092
            // If user calls SendMail command and server returns Status code element which means SendMail operation meet failure.
            Site.CaptureRequirementIfIsNotNull(
                responseSendMail.ResponseData.Status,
                5092,
                @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 5: ] [If the message was sent successfully, the server returns an empty response.] Otherwise, the server responds with a Status element (section 2.2.3.162.8) that indicates the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R724");

            // Verify MS-ASCMD requirement: MS-ASCMD_R724
            Site.CaptureRequirementIfAreEqual<string>(
                "166",
                responseSendMail.ResponseData.Status,
                724,
                @"[In AccountId(SendMail, SmartForward, SmartReply)] A Status element (section 2.2.3.162) value of 166 is returned if the AccountId element value is not valid.");
            #endregion

            #region Sync user2 mailbox changes
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion
        }

        #endregion
    }
}