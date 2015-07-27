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
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is used to test the Ping command.
    /// </summary>
    [TestClass]
    public class S11_Ping : TestSuiteBase
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
        /// This test case is used to verify when the Ping command request omitted required parameters, the Ping command response has status value 3.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC01_Ping_Status3()
        {
            #region Change a new user to call ActiveSync operation with a new DeviceID.
            this.CMDAdapter.SwitchUser(this.User3Information.UserName, this.User3Information.UserPassword, this.User3Information.UserDomain);

            // Record user that uses new DeviceID
            this.CMDAdapter.ChangeDeviceID("NewDevice");
            this.RecordDeviceInfoChanged();
            #endregion

            #region Calls FolderSync command before sending ping request
            this.FolderSync();
            #endregion

            #region Create a Ping request without Folder element
            PingRequest pingRequest = CreatePingRequest(this.User3Information.NotesCollectionId, Request.PingFolderClass.Notes);
            pingRequest.RequestData.Folders = null;
            PingResponse pingReponse = this.CMDAdapter.Ping(pingRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R4240, MS-ASCMD_R4241
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4240");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4240
            Site.CaptureRequirementIfAreEqual<byte>(
                3,
                pingReponse.ResponseData.Status,
                4240,
                @"[In Status(Ping)] [When the scope is Global], [the cause of the status value 3 is] The Ping command request did not specify all the necessary parameters.");

            #region Create a Ping request with Folder element and HeartbeatInterval element
            PingRequest pingRequestWithRequireElement = CreatePingRequest(this.User3Information.NotesCollectionId, Request.PingFolderClass.Notes);
            PingResponse pingResponseWithRequireElement = this.CMDAdapter.Ping(pingRequestWithRequireElement);
            Site.Assert.IsNotNull(pingResponseWithRequireElement.ResponseData.Status, "If Ping command executes successfully, server will return status code.");
            #endregion

            #region Create a Ping request with Folder element and without HeartbeatInterval element
            PingRequest pingRequestWithoutHeartBeatIntervalElement = CreatePingRequest(this.User3Information.NotesCollectionId, Request.PingFolderClass.Notes);
            pingRequestWithoutHeartBeatIntervalElement.RequestData.HeartbeatInterval = null;
            PingResponse pingReponseWithoutHeartBeatIntervalElement = this.CMDAdapter.Ping(pingRequestWithoutHeartBeatIntervalElement);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4241");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4241
            Site.CaptureRequirementIfIsTrue(
                pingReponse.ResponseData.Status == 3 && pingReponseWithoutHeartBeatIntervalElement.ResponseData.Status != 3,
                4241,
                @"[In Status(Ping)] [When the scope is Global], [the cause of the status value 3 is] The client MUST issue a Ping request that includes both the heartbeat interval and the folder list at least once.");
            #endregion

            #region Delete ActiveSync device in User3's mail account
            this.ClearDevice(this.User3Information.UserName, this.User3Information.UserPassword, this.User3Information.UserDomain);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when the Ping command request has the syntax error, the Ping command response has status value 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC02_Ping_Status4()
        {
            #region Create one initial ping request without HeartbeatInterval element that means to create one poorly formatted WBXML.
            PingRequest pingRequest = new PingRequest { RequestData = new Request.Ping() };
            PingResponse pingResponse = this.CMDAdapter.Ping(pingRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R4244, MS-ASCMD_R4228
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4244");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4244
            Site.CaptureRequirementIfAreEqual<byte>(
                4,
                pingResponse.ResponseData.Status,
                4244,
                @"[In Status(Ping)] [When the scope is Global], [the cause of the status value 4 is] Frequently caused by poorly formatted WBXML.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4228");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4228
            Site.CaptureRequirementIfAreEqual<byte>(
                4,
                pingResponse.ResponseData.Status,
                4228,
                @"[In Status(Ping)] If the command failed, the Status element contains a code that indicates the type of failure.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when HeartbeatInterval is setting to a negative value, the Ping command response has status value 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC03_Ping_Status5()
        {
            #region Create a ping request with negative HeartbeatInterval value.
            PingRequest pingRequest = CreatePingRequest(this.User1Information.InboxCollectionId, Request.PingFolderClass.Email);

            // Send the ping request by setting HearbeatInterval to invalid value
            pingRequest.RequestData.HeartbeatInterval = "-100";
            PingResponse pingResponse = this.CMDAdapter.Ping(pingRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R5837, MS-ASCMD_R5690
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5837");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5837
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                pingResponse.ResponseData.Status,
                5837,
                @"[In Status(Ping)] [When the scope is Global], [the cause of the status value 5 is] The client sent a Ping command request with a heartbeat interval that was [either too long or] too short.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5690");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5690
            Site.CaptureRequirementIfIsTrue(
                pingResponse.ResponseData.HeartbeatInterval != null && pingResponse.ResponseData.Status == 5,
                5690,
                @"[In HeartbeatInterval(Ping)] The HeartbeatInterval element is also returned by the server with a status code of 5 and specifies [either] the minimum [or maximum] interval that is allowed when the client has requested a heartbeat interval that is outside the acceptable range.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify folder hierarchy sync is required before calling the Ping command by client.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC04_Ping_Status7_SyncRequired()
        {
            #region Change a new user to call ActiveSync operation with a new DeviceID.
            this.CMDAdapter.SwitchUser(this.User3Information.UserName, this.User3Information.UserPassword, this.User3Information.UserDomain);

            // Record user that uses new DeviceID
            this.CMDAdapter.ChangeDeviceID("ChangeNewDevice");
            this.RecordDeviceInfoChanged();
            #endregion

            #region Create a Ping request.
            PingRequest pingRequest = CreatePingRequest(this.User1Information.InboxCollectionId, Request.PingFolderClass.Email);
            pingRequest.RequestData.HeartbeatInterval = "60";
            PingResponse pingResponse = this.CMDAdapter.Ping(pingRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4258");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4258
            Site.CaptureRequirementIfAreEqual<byte>(
                7,
                pingResponse.ResponseData.Status,
                4258,
                @"[In Status(Ping)] [When the scope is Global], [the cause of the status value 7 is] a folder hierarchy sync is required.");

            #region Delete ActiveSync device in User3 mail account
            this.ClearDevice(this.User3Information.UserName, this.User3Information.UserPassword, this.User3Information.UserDomain);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when a folder is out of date, the Ping command response have status value 7
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC05_Ping_Status7_Outdate()
        {
            #region Change a new user to call ActiveSync operation with a new DeviceID.
            this.CMDAdapter.SwitchUser(this.User3Information.UserName, this.User3Information.UserPassword, this.User3Information.UserDomain);

            // Record user that uses new DeviceID
            this.RecordDeviceInfoChanged();
            this.CMDAdapter.ChangeDeviceID("FirstDeviceID");
            #endregion

            #region Sync user folder changes
            this.FolderSync();
            #endregion

            #region Switch to new device and create one folder
            this.CMDAdapter.ChangeDeviceID("SecondDeviceID");
            string folderSynKey = this.GetFolderSyncKey();

            // Create one mail type folder 
            string newFolderName = Guid.NewGuid().ToString().Replace("-", string.Empty);
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(folderSynKey, 12, newFolderName, this.User1Information.InboxCollectionId);
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual<byte>(
                1,
                folderCreateResponse.ResponseData.Status,
                "After folder create success, server should return status 1");
            #endregion

            #region Switch back to old device and send one ping request
            this.CMDAdapter.ChangeDeviceID("FirstDeviceID");
            PingRequest pingRequest = CreatePingRequest(this.User1Information.InboxCollectionId, Request.PingFolderClass.Email);
            PingResponse pingResponse = this.CMDAdapter.Ping(pingRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4257");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4257
            Site.CaptureRequirementIfAreEqual<byte>(
                7,
                pingResponse.ResponseData.Status,
                4257,
                @"[In Status(Ping)] [When the scope is Global], [the cause of the status value 7 is] The folder hierarchy is out of date.");

            #region Record new created folder collection ID.
            // Get new created folder collection ID
            FolderSyncRequest folderSyncRequestForNewFolder = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse folderSyncResponseForNewFolder = this.CMDAdapter.FolderSync(folderSyncRequestForNewFolder);

            // Gets the server ID for new folder after change DeviceID.
            string serverId = TestSuiteBase.GetCollectionId(folderSyncResponseForNewFolder, newFolderName);
            TestSuiteBase.RecordCaseRelativeFolders(this.User3Information, serverId);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the response specifies the changed folders when changes occurred in two of the monitored folders.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC06_Ping_ChangesOccurred()
        {
            #region User2 calls CreateFolder command to create two sub folders in user Inbox folder
            this.SwitchUser(this.User2Information);
            int folderCreateCount = 2;
            List<string> folderIDList = new List<string>();
            for (int folderIndex = 0; folderIndex < folderCreateCount; folderIndex++)
            {
                string folderID = this.CreateFolder((byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate", (uint)(folderIndex + 1)), this.User2Information.InboxCollectionId);
                folderIDList.Add(folderID);

                // Sync the new created folder
                this.SyncChanges(folderID);
            }

            // Record created folders
            TestSuiteBase.RecordCaseRelativeFolders(this.User2Information, folderIDList.ToArray());
            #endregion

            #region User1 calls SendMail command to send one mail to user2, then user2 move the email to one of new created folder
            this.SwitchUser(this.User1Information);
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);

            // User2 calls Sync command to get the email, then moves to sub folder
            this.SwitchUser(this.User2Information);
            SyncResponse syncInboxFolderResponse = this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            string emailItemServerID = TestSuiteBase.FindServerId(syncInboxFolderResponse, "Subject", emailSubject);

            // User2 calls MoveItems command to move the email item from Inbox folder to one of created subfolder
            MoveItemsRequest moveItemsRequest = TestSuiteBase.CreateMoveItemsRequest(emailItemServerID, this.User2Information.InboxCollectionId, folderIDList[0]);
            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);
            Site.Assert.AreEqual(3, moveItemsResponse.ResponseData.Response[0].Status, " If MoveItems command executes successfully, server should return status 3");
            #endregion

            #region Resend ping request to monitor the changes in sub folders and record duration
            DateTime startTime = DateTime.Now;
            PingRequest pingRequest = CreateMultiFolderPingRequest(folderIDList, Request.PingFolderClass.Email);
            PingResponse pingResponse = this.CMDAdapter.Ping(pingRequest);
            DateTime endTime = DateTime.Now;
            TimeSpan duration = endTime - startTime;
            #endregion

            #region Verify Requirements MS-ASCMD_R306, MS-ASCMD_R4235, MS-ASCMD_R4236, MS-ASCMD_R5714, MS-ASCMD_R5636
            // Ping command is used to monitor specified folder changes, if changes happened in the folders, server will return status 2 in Ping response, then MS-ASCMD_R306, MS-ASCMD_R4235, MS-ASCMD_R4236, MS-ASCMD_R5636 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R306");

            // Verify MS-ASCMD requirement: MS-ASCMD_R306
            Site.CaptureRequirementIfIsNotNull(
                pingResponse.ResponseData.Status,
                306,
                @"[In Ping] The Ping command is used to request that the server monitor specified folders for changes that would require the client to resynchronize.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4235");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4235
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                pingResponse.ResponseData.Status,
                4235,
                @"[In Status(Ping)] [When the scope is] Global, [the meaning of the status value] 2 [is] Changes occurred in at least one of the monitored folders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5636");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5636
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                pingResponse.ResponseData.Status,
                5636,
                @"[In Monitoring Folders for New Items] The server uses the last airsync:SyncKey element (as specified in section 2.2.3.166.4) value returned to the client when determining what to report in the Ping response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4236");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4236
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                pingResponse.ResponseData.Status,
                4236,
                @"[In Status(Ping)] [When the scope is] Global, [the meaning of the status value 2 is] The response specifies the changed folders.");

            if (Common.IsRequirementEnabled(5714, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5714");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5714
                Site.CaptureRequirementIfIsTrue(
                    pingRequest.RequestData.HeartbeatInterval != null && pingResponse.ResponseData.Status == (byte)2 && duration.Milliseconds < Convert.ToInt32(pingRequest.RequestData.HeartbeatInterval) * 1000,
                    5714,
                    @"[In Appendix A: Product Behavior] The body of the request contains an interval of time that specifies how long the implementation does wait before responding if no new items are added to the specified folders. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Send ping request with empty HeartbeatInterval element, if server returns status equal to 2 then MS-ASCMD_R5836 is verified.
            pingRequest = CreateMultiFolderPingRequest(folderIDList, Request.PingFolderClass.Email);
            pingRequest.RequestData.HeartbeatInterval = null;
            pingResponse = this.CMDAdapter.Ping(pingRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5836");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5836
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                pingResponse.ResponseData.Status,
                5836,
                @"[In Status(Ping)] [When the scope is Global], The server saves the heartbeat interval value (section 2.2.3.79.1), so only the folder list is required on subsequent requests.");
        }

        /// <summary>
        /// This test case is used to verify when heartbeat interval is setting to a too large number, the Ping command response have status value 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC07_Ping_HeartbeatIntervalSetTooLarge()
        {
            #region Create a ping request with too large HeartbeatInterval value.
            PingRequest pingRequest = CreatePingRequest(this.User1Information.InboxCollectionId, Request.PingFolderClass.Email);

            // Create invalid ping request by setting HearbeatInterval to too large number which exceeds 3540;
            pingRequest.RequestData.HeartbeatInterval = "6000000";
            #endregion

            #region Calls Ping command
            PingResponse pingResponse = this.CMDAdapter.Ping(pingRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R4249, MS-ASCMD_R5691
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4249");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4249
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                pingResponse.ResponseData.Status,
                4249,
                @"[In Status(Ping)] [When the scope is Global], [the cause of the status value 5 is] The client sent a Ping command request with a heartbeat interval that was [either] too long [or too short].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5691");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5691
            Site.CaptureRequirementIfIsTrue(
                pingResponse.ResponseData.HeartbeatInterval != null && pingResponse.ResponseData.Status == 5,
                5691,
                @"[In HeartbeatInterval(Ping)] The HeartbeatInterval element is also returned by the server with a status code of 5 and specifies [either the minimum or] maximum interval that is allowed when the client has requested a heartbeat interval that is outside the acceptable range.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify a Ping command without request body.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S11_TC08_Ping_WithoutRequestBody()
        {
            #region Call Ping command with request body
            PingRequest pingRequest = CreatePingRequest(this.User1Information.InboxCollectionId, Request.PingFolderClass.Email);
            pingRequest.RequestData.HeartbeatInterval = "60";
            PingResponse pingResponse = this.CMDAdapter.Ping(pingRequest);

            // According to techinical specification, if there were changes occurred in the Inbox folder, the Ping response would return 2.
            // Reissue the Ping command when the Sync command completes to stay up to date.
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            while (counter < retryCount && 2 == pingResponse.ResponseData.Status)
            {
                System.Threading.Thread.Sleep(waitTime);
                SyncRequest syncRequest = TestSuiteBase.CreateEmptySyncRequest(this.User1Information.InboxCollectionId);
                this.Sync(syncRequest);
                syncRequest.RequestData.Collections[0].SyncKey = this.LastSyncKey;
                this.Sync(syncRequest);

                pingResponse = this.CMDAdapter.Ping(pingRequest);
                counter++;
            }

            Site.Assert.AreEqual<byte>(1, pingResponse.ResponseData.Status, "The Ping command should be successful.");
            #endregion

            #region Create a ping request without request body
            pingRequest = new PingRequest();
            pingRequest.RequestData = null;
            #endregion

            #region Call Ping command
            pingResponse = this.CMDAdapter.Ping(pingRequest);
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R335");

            // Verify MS-ASCMD requirement: MS-ASCMD_R335
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                pingResponse.ResponseData.Status,
                335,
                @"[In Ping] A Ping command can be sent with no body, in which case the cached version is used.");
            #endregion
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// This method is used to create Ping request.
        /// </summary>
        /// <param name="folderID">The value of the FolderId element.</param>
        /// <param name="pingFolderClass">An instance of the PingFolderClass.</param>
        /// <returns>The Ping request.</returns>
        private static PingRequest CreatePingRequest(string folderID, Request.PingFolderClass pingFolderClass)
        {
            Request.PingFolder pingFolder = new Request.PingFolder
            {
                Class = pingFolderClass,
                Id = folderID
            };

            return Common.CreatePingRequest(new Request.PingFolder[] { pingFolder }, "80");
        }

        /// <summary>
        /// This method is designed to create Ping request.
        /// </summary>
        /// <param name="folderIDList">The List value of the FolderId element.</param>
        /// <param name="pingFolderClass">An instance of the PingFolderClass.</param>
        /// <returns>The Ping request.</returns>
        private static PingRequest CreateMultiFolderPingRequest(List<string> folderIDList, Request.PingFolderClass pingFolderClass)
        {
            List<Request.PingFolder> pingRequestList = new List<Request.PingFolder>();
            foreach (string folderID in folderIDList)
            {
                Request.PingFolder pingFolder = new Request.PingFolder { Class = pingFolderClass, Id = folderID };
                pingRequestList.Add(pingFolder);
            }

            return Common.CreatePingRequest(pingRequestList.ToArray(), "80");
        }

        /// <summary>
        /// Initial folder sync
        /// </summary>
        /// <returns>SyncKey returned from server</returns>
        private string GetFolderSyncKey()
        {
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse foldersyncResponse = this.CMDAdapter.FolderSync(folderSyncRequest);
            string folderSynKey = foldersyncResponse.ResponseData.SyncKey;
            return folderSynKey;
        }
        #endregion
    }
}