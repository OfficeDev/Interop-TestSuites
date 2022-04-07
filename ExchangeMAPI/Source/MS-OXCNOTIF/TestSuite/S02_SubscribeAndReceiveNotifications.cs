namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test cases for S02_SubscribeAndReceiveNotifications.
    /// </summary>
    [TestClass]
    public class S02_SubscribeAndReceiveNotifications : TestSuiteBase
    {
        #region Class Initialization and Cleanup

        /// <summary>
        /// Class initialize.
        /// </summary>
        /// <param name="testContext">The session context handle</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This test case is designed to implement using RopQueryRows to create a table view and subscribe to the TableModified event, and then verify the TableRowAdded event which is a type of TableModified event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC01_VerifyTableRowAddEventAndCreateTableViewByTableQueryRows()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);

            // The first content table handle is used to trigger the notification.
            uint contentTableHandle1;

            // The second content table handle is used to get the information from the specified table.
            uint contentTableHandle2 = 0;
            this.GetContentsTable(inboxTableHandle, out contentTableHandle1, false);
            #endregion

            #region Create table view by QueryRows

            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[] { PropertyTags.All[PropertyNames.PidTagFolderId], PropertyTags.All[PropertyNames.PidTagInstanceNum], PropertyTags.All[PropertyNames.PidTagMid] };
            this.SetColumns(contentTableHandle1, tags);

            // Retrieves rows from content table to get the data of 30 rows
            this.QueryRows(contentTableHandle1, 30);
            #endregion

            #region Trigger TableRowAdded event and get notification
            ulong messageId = this.TriggerTableRowAddedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for TableRowAdded event
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                    Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableRowAdded, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableRowAdded.");
                    this.VerifyTableModifyNotificationFlag(notifyResponse);
                    this.VerifyTableRowAddedNotificationElements(notifyResponse);
                    if (notifyResponse.NotificationData.TableRowDataSize != null)
                    {
                        #region Verify the elements TableRowDataSize and TableRowData of the notification response

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R163");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R163
                        this.Site.CaptureRequirementIfAreEqual<ushort?>(
                            (ushort)notifyResponse.NotificationData.TableRowData.Length,
                            notifyResponse.NotificationData.TableRowDataSize,
                            163,
                            @"[In NotificationData Structure] TableRowDataSize: An unsigned 16-bit integer that indicates the length of the table row data.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R166");

                        // Convert TableRowData to string type.
                        string tableRowDataToString = BitConverter.ToString(notifyResponse.NotificationData.TableRowData);
                        Site.Log.Add(LogEntryKind.Debug, "The table row data is {0}.", tableRowDataToString);

                        // Convert TableRowFolderID to string type.
                        Site.Assert.IsNotNull(notifyResponse.NotificationData.TableRowFolderID, "The TableRowFolderID should not null.");
                        string tableRowFolderIDToString = BitConverter.ToString(BitConverter.GetBytes((ulong)notifyResponse.NotificationData.TableRowFolderID));
                        Site.Log.Add(LogEntryKind.Debug, "The value of TableRowFolderID is {0}.", tableRowFolderIDToString);

                        // Convert TableRowInstance to string type.
                        Site.Assert.IsNotNull(notifyResponse.NotificationData.TableRowInstance, "The TableRowInstance should not null.");
                        string tableRowInstanceToString = BitConverter.ToString(BitConverter.GetBytes((uint)notifyResponse.NotificationData.TableRowInstance));
                        Site.Log.Add(LogEntryKind.Debug, "The value of TableRowInstance is {0}.", tableRowInstanceToString);

                        // Convert TableRowMessageID to string type.
                        Site.Assert.IsNotNull(notifyResponse.NotificationData.TableRowMessageID, "The TableRowMessageID should not null.");
                        string tableRowMessageIDToString = BitConverter.ToString(BitConverter.GetBytes((ulong)notifyResponse.NotificationData.TableRowMessageID));
                        Site.Log.Add(LogEntryKind.Debug, "The value of TableRowMessageID is {0}.", tableRowMessageIDToString);

                        // Check whether TableRowData contains TableRowFolderID, TableRowInstance and TableRowMessageID or not.
                        bool isContained = tableRowDataToString.Contains(tableRowFolderIDToString) &&
                            tableRowDataToString.Contains(tableRowInstanceToString) &&
                            tableRowDataToString.Contains(tableRowMessageIDToString);

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R166
                        // The added properties values are TableRowFolderID, TableRowInstance and TableRowMessageID. 
                        // So if TableRowData contains the three values, this requirement can be verified.
                        this.Site.CaptureRequirementIfIsTrue(
                            isContained,
                            166,
                            @"[In NotificationData Structure] TableRowData (variable): The table row data, which contains a list of property values in a PropertyRow structure, as specified in [MS-OXCDATA] section 2.8, for the row that was added or modified in the table. ");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R166001");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R166001
                        // Because TableRowFolderID, TableRowInstance and TableRowMessageID are added by previous RopSetColumns ROP.
                        // So if TableRowData contains the three values, this requirement can be verified.
                        this.Site.CaptureRequirementIfIsTrue(
                            isContained,
                            166001,
                            @"[In NotificationData Structure] The property values to be included are determined by a previous RopSetColumns ROP, as specified in [MS-OXCTABL] section 2.2.2.2. ");
                        #endregion
                    }

                    #region Get the location of the previous row and the row where the modified row is inserted
                    // Open Inbox folder and get content table of the Inbox folder
                    this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
                    uint tableRowInstanceFromTable;
                    uint insertAfterTableRowInstanceFromTable = 0;
                    this.GetContentsTable(inboxTableHandle, out contentTableHandle2, false);

                    // Create table view by QueryRows
                    this.SetColumns(contentTableHandle2, tags);

                    // Retrieves rows from content table to get the data of 30 rows
                    RopQueryRowsResponse queryResponse = this.QueryRows(contentTableHandle2, 30);

                    // Get the location of the current row in the table
                    RopQueryPositionResponse queryPositionResponse = this.QueryPosition(contentTableHandle2);

                    // Get the location of the previous row in the table
                    tableRowInstanceFromTable = BitConverter.ToUInt32(queryResponse.RowData.PropertyRows[(int)queryPositionResponse.Numerator - 2].PropertyValues[1].Value, 0);

                    // Get the location of the row where the modified row is inserted
                    foreach (PropertyRow row in queryResponse.RowData.PropertyRows)
                    {
                        if (BitConverter.ToUInt64(row.PropertyValues[2].Value, 0) == this.LocatingMessageId)
                        {
                            insertAfterTableRowInstanceFromTable = BitConverter.ToUInt32(row.PropertyValues[1].Value, 0);
                        }
                    }
                    #endregion

                    #region Verify the elements of the notification response
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R149");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R149
                    this.Site.CaptureRequirementIfAreEqual<uint?>(
                        tableRowInstanceFromTable,
                        notifyResponse.NotificationData.TableRowInstance,
                        149,
                        @"[In NotificationData Structure] TableRowInstance: An identifier of the instance of the previous row in the table.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R161");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R161
                    this.Site.CaptureRequirementIfAreEqual<uint?>(
                        insertAfterTableRowInstanceFromTable,
                        notifyResponse.NotificationData.InsertAfterTableRowInstance,
                        161,
                        @"[In NotificationData Structure] InsertAfterTableRowInstance: An unsigned 32-bit identifier of the instance of the row where the modified row is inserted");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R161002");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R161002
                    this.Site.CaptureRequirementIfIsTrue(
                        notifyResponse.NotificationData.InsertAfterTableRowInstance!=null && (notifyResponse.NotificationData.NotificationFlags & 0x8000) == 0x8000 && notifyResponse.NotificationData.TableEventType== 0x0003,
                        161002,
                        @"[In NotificationData Structure] This field [InsertAfterTableRowInstance] is available when bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0003.");


                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R145");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R145
                    // If the value of TableRowMessageID is the message which trigger the notification, this requirement can be verified.
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        messageId,
                        notifyResponse.NotificationData.TableRowMessageID,
                        145,
                        @"[In NotificationData Structure] TableRowMessageID: The value of the Message ID structure, as specified in [MS-OXCDATA] section 2.2.1.2, of the item triggering the notification.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R141");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R141
                    // If the value of TableRowFolderID is the folder which trigger the notification, this requirement can be verified.
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        this.InboxFolderId,
                        notifyResponse.NotificationData.TableRowFolderID,
                        141,
                        @"[In NotificationData Structure] TableRowFolderID: The value of the Folder ID structure, as specified in [MS-OXCDATA] section 2.2.1.1, of the item triggering the notification.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R97");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R97
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        contentTableHandle1,
                        notifyResponse.NotificationHandle,
                        97,
                        @"[In RopNotify ROP Response Buffer] [NotificationHandle] The target object can be a table.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R261: The value of the Folder ID and Message ID separately are {0},{1}", notifyResponse.NotificationData.TableRowFolderID, notifyResponse.NotificationData.TableRowMessageID);

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R261
                    bool isVerifiedR261 = notifyResponse.NotificationData.TableRowFolderID != null && notifyResponse.NotificationData.TableRowMessageID != null;

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR261,
                        261,
                        @"[In Creating and Sending TableModified Event Notifications] [When a TableModified event occurs, the server generates a notification using one of the following three methods, listed in descending order of usefulness to the client.] For TableRowAdded event, the server generates an informative notification that specifies the nature of the change (content or hierarchy), the value of the Folder ID structure, as specified in [MS-OXCDATA] section 2.2.1.1, the value of the Message ID structure, as specified in [MS-OXCDATA] section 2.2.1.2, and new table values.");

                    this.VeriyServerGenerateInformativeNotification(isVerifiedR261);

                    Site.Assert.IsNotNull(notifyResponse.NotificationData.TableEventType, "The TableEventType in the RopNotifyResponse should not null.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R129");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R129
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0003,
                        (int)notifyResponse.NotificationData.TableEventType,
                        129,
                        @"[In NotificationData Structure] [TableEventType value] 0x0003: The notification is for TableRowAdded events.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R23");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R23
                    // The client can receive TableRowAdded notification indicates that a new row has been added to the table.
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0003,
                        (int)notifyResponse.NotificationData.TableEventType,
                        23,
                        @"[In TableModified Event Types] TableRowAdded: A new row has been added to the table.");

                    // Only Exchange 2010 and above require a table view, this server version restrict of MS-OXCNTOIF_R249 is the same with MS-OXCNTOIF_R245.
                    if (Common.IsRequirementEnabled(245, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R252");

                        // If the server can return notification response, it can verify RopQueryRows create a table view. so the following requirements can be captured directly. 
                        Site.CaptureRequirement(
                            252,
                            @"[In Creating and Sending TableModified Event Notifications] If a table view is required on the server, the server MUST receive a request from one of the following ROPs, each of which cause a table view to be created on the server: RopQueryRows ([MS-OXCROPS] section 2.2.5.4).");

                        this.VerifyTableViewCreated();
                    }
                    #endregion
                }
            }
            #endregion

            // Only Exchange 2010 and above will stop sending notifications if the RopResetTable ROP is received, until a new table view is created.
            if (Common.IsRequirementEnabled(281, this.Site))
            {
                #region Reset the contents table
                // Reset the first content table handle is used to trigger the notification.
                this.ResetTable(contentTableHandle1);

                // Reset the second content table handle is used to get the information from the specified table.
                this.ResetTable(contentTableHandle2);
                #endregion

                #region Trigger TableRowAdded event after ResetTable and get notification
                this.TriggerTableRowAddedEvent();

                // Trigger table event after ResetTable, server shouldn't send notification response on Exchange server 2010 and above
                rsp = this.CNOTIFAdapter.GetNotification(false);
                bool isServerCreateSubscription = false;
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsFalse(isServerCreateSubscription, "The server can't send notify response when reset table.");
                #endregion

                #region Create table view by QueryRows
                this.SetColumns(contentTableHandle1, tags);

                // Retrieves rows from content table to get the data of 2 rows
                this.QueryRows(contentTableHandle1, 2);
                #endregion

                #region Trigger TableRowAdded event again and get notification
                this.TriggerTableRowModifiedEvent();

                // Create table view by QueryRows after ResetTable, and trigger table event again the server should send notification response.
                rsp = this.CNOTIFAdapter.GetNotification(true);
                Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsTrue(isServerCreateSubscription, "After resetting table, the server should send notify response when a new table view is created by RopQueryRows.");
                #endregion

                #region Verify notification response after reset Table and create table view by QueryRows

                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R281");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R281
                // Since the server return the notification after a table view is created by RopQueryRows, this requirement can be verified directly.
                Site.CaptureRequirement(
                    281,
                    @"[In Appendix A: Product Behavior] The implementation does stop sending notifications if the RopResetTable ROP ([MS-OXCROPS] section 2.2.5.15) is received, until a new table view is created using one of the following ROPs: RopQueryRows. (Exchange 2010 and above follow this behavior.)");
                #endregion
            }
        }

        /// <summary>
        /// This test case is designed to implement using RopFindRow to create a table view and subscribe to the TableModified event, and then verify the TableRowDeleted event which is a type of TableModified event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC02_VerifyTableRowDeleteEventAndCreateTableViewByTableFindRow()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
            uint contentTableHandle;
            this.GetContentsTable(inboxTableHandle, out contentTableHandle, false);
            #endregion

            #region Create table view by FindRow
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[] { PropertyTags.All[PropertyNames.PidTagInstID], PropertyTags.All[PropertyNames.PidTagImportance], PropertyTags.All[PropertyNames.PidTagInstanceNum], PropertyTags.All[PropertyNames.PidTagFolderId], PropertyTags.All[PropertyNames.PidTagLastModificationTime] };
            this.SetColumns(contentTableHandle, tags);

            // Set the value of restriction data
            byte[] restrictionData = new byte[5];
            ushort pidTagMessageClassID = PropertyTags.All[PropertyNames.PidTagLastModificationTime].PropertyId;
            ushort typeOfPidTagMessageClass = PropertyTags.All[PropertyNames.PidTagLastModificationTime].PropertyType;
            restrictionData[0] = (byte)Restrictions.ExistRestriction;
            Array.Copy(BitConverter.GetBytes(typeOfPidTagMessageClass), 0, restrictionData, 1, sizeof(ushort));
            Array.Copy(BitConverter.GetBytes(pidTagMessageClassID), 0, restrictionData, 3, sizeof(ushort));
            this.FindRow(contentTableHandle, restrictionData);
            #endregion

            #region Trigger TableRowDeleted event and get notification
            this.TriggerTableRowDeletedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for TableRowDeleted event
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                    Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableRowDeleted, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableRowDeleted.");
                    this.VerifyTableModifyNotificationFlag(notifyResponse);
                    this.VerifyTableRowDeletedNotificationElements(notifyResponse);

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R262: The value of the Folder ID and Message ID separately are {0},{1}", notifyResponse.NotificationData.TableRowFolderID, notifyResponse.NotificationData.TableRowMessageID);

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R262
                    bool isVerifiedR262 = notifyResponse.NotificationData.TableRowFolderID != null && notifyResponse.NotificationData.TableRowMessageID != null;

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR262,
                        262,
                        @"[In Creating and Sending TableModified Event Notifications] [When a TableModified event occurs, the server generates a notification using one of the following three methods, listed in descending order of usefulness to the client.] For TableRowDeleted event, the server generates an informative notification that specifies the nature of the change (content or hierarchy), the value of the Folder ID structure, as specified in [MS-OXCDATA] section 2.2.1.1, the value of the Message ID structure, as specified in [MS-OXCDATA] section 2.2.1.2, and new table values.");

                    this.VeriyServerGenerateInformativeNotification(isVerifiedR262);

                    Site.Assert.IsNotNull(notifyResponse.NotificationData.TableEventType, "The TableEventType in the RopNotifyResponse should not null.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R130");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R130
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0004,
                        (int)notifyResponse.NotificationData.TableEventType,
                        130,
                        @"[In NotificationData Structure] [TableEventType value] 0x0004: The notification is for TableRowDeleted events.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R24");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R24
                    // The client can receive TableRowDeleted notification indicates that an existing row has been deleted from the table.
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0004,
                        (int)notifyResponse.NotificationData.TableEventType,
                        24,
                        @"[In TableModified Event Types] TableRowDeleted: An existing row has been deleted from the table.");

                    // Only Exchange 2010 and above require a table view, this server version restrict is the same with MS-OXCNTOIF_R245.
                    if (Common.IsRequirementEnabled(245, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R249");

                        // If the server can return notification response, it can verify RopQueryRows create a table view. so the following requirements can be captured directly.   
                        this.Site.CaptureRequirement(
                            249,
                            @"[In Creating and Sending TableModified Event Notifications] If a table view is required on the server, the server MUST receive a request from one of the following ROPs, each of which cause a table view to be created on the server: RopFindRow ([MS-OXCROPS] section 2.2.5.13).");

                        this.VerifyTableViewCreated();
                    }
                }
            }
            #endregion

            // Only Exchange 2010 and above will stop sending notifications if the RopResetTable ROP is received, until a new table view is created.
            if (Common.IsRequirementEnabled(275, this.Site))
            {
                #region Reset the contents table
                this.ResetTable(contentTableHandle);
                #endregion

                #region Trigger TableRowAdded event after ResetTable and get notification
                this.TriggerTableRowAddedEvent();
                rsp = this.CNOTIFAdapter.GetNotification(false);
                bool isServerCreateSubscription = false;
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsFalse(isServerCreateSubscription, "The server can't send notify response when reset table.");
                #endregion

                #region Create table view by FindRow
                this.SetColumns(contentTableHandle, tags);
                this.FindRow(contentTableHandle, restrictionData);
                #endregion

                #region Trigger TableRowAdded event again and get notification
                this.TriggerTableRowAddedEvent();

                // Create table view by FindRow after ResetTable, and trigger table event again the server should send notification response.
                rsp = this.CNOTIFAdapter.GetNotification(true);
                Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }
                #endregion

                #region Verify notification response after reset Table and create table view by FindRow
                // Since the server return the notification after a table view is created by RopFindRow, this requirement can be verified.
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R275: The implementation does return notifications after get RopFindRow", isServerCreateSubscription ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R275
                bool isR275Satisfied = isServerCreateSubscription;

                Site.CaptureRequirementIfIsTrue(
                    isR275Satisfied,
                    275,
                    @"[In Appendix A: Product Behavior] Implementation does stop sending notifications if the RopResetTable ROP ([MS-OXCROPS] section 2.2.5.15) is received, until a new table view is created using one of the following ROPs: RopFindRow. (Exchange 2010 and above follow this behavior.)");
                #endregion
            }
        }

        /// <summary>
        ///  This test case is designed to implement using RopQueryColumnsAll to create a table view and subscribe to the TableModified event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC03_VerifyCreateTableViewByTableQueryColumnsAll()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxHandle;
            this.OpenFolder(this.InboxFolderId, out inboxHandle);
            uint contentTableHandle;
            this.GetContentsTable(inboxHandle, out contentTableHandle, false);
            #endregion

            #region Create table view by QueryColumnsAll
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[] { PropertyTags.All[PropertyNames.PidTagFolderId] };
            this.SetColumns(contentTableHandle, tags);
            this.QueryColumnsAll(contentTableHandle);
            #endregion

            #region Trigger TableRowAdded event and get notification
            this.TriggerTableRowAddedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify table view is created by QueryColumnsAll
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                    Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableRowAdded, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableRowAdded.");
                    this.VerifyTableModifyNotificationFlag(notifyResponse);

                    // Only Exchange 2010 and above require a table view, this server version restrict is the same with MS-OXCNTOIF_R245.
                    if (Common.IsRequirementEnabled(245, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R250");

                        // If the server can return notification response, it can verify RopQueryRows create a table view. so the following requirements can be captured directly. 
                        this.Site.CaptureRequirement(
                            250,
                            @"[In Creating and Sending TableModified Event Notifications] If a table view is required on the server, the server MUST receive a request from one of the following ROPs, each of which cause a table view to be created on the server: RopQueryColumnsAll ([MS-OXCROPS] section 2.2.5.12).");

                        this.VerifyTableViewCreated();
                    }
                }
            }
            #endregion

            // Only Exchange 2010 and above will stop sending notifications if the RopResetTable ROP is received, until a new table view is created. 
            if (Common.IsRequirementEnabled(277, this.Site))
            {
                #region Reset the contents table
                this.ResetTable(contentTableHandle);
                #endregion

                #region Trigger TableRowAdded event after ResetTable and get notification
                this.TriggerTableRowAddedEvent();

                // Trigger table event after ResetTable, server shouldn't send notification response on Exchange server 2010 and above
                rsp = this.CNOTIFAdapter.GetNotification(false);
                bool isServerCreateSubscription = false;
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsFalse(isServerCreateSubscription, "The server can't send notify response when reset table.");
                #endregion

                #region Create table view by QueryColumnsAll
                this.SetColumns(contentTableHandle, tags);
                this.QueryColumnsAll(contentTableHandle);
                #endregion

                #region Trigger TableRowAdded event again and get notification
                this.TriggerTableRowAddedEvent();

                // Create table view by QueryColumnsAll after ResetTable, and trigger table event again the server should send notification response.
                rsp = this.CNOTIFAdapter.GetNotification(true);
                Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }
                #endregion

                #region Verify notification response after reset Table and create table view by QueryColumnsAll
                // Since the server return the notification after a table view is created by RopQueryColumnsAll, this requirement can be verified.
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R277: The implementation does return notifications after get RopQueryColumnsAll", isServerCreateSubscription ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R277
                bool isR277Satisfied = isServerCreateSubscription;

                Site.CaptureRequirementIfIsTrue(
                    isR277Satisfied,
                    277,
                    @"[In Appendix A: Product Behavior] Implementation does stop sending notifications if the RopResetTable ROP ([MS-OXCROPS] section 2.2.5.15) is received, until a new table view is created using one of the following ROPs: RopQueryColumnsAll. (Exchange 2010 and above follow this behavior.)");
                #endregion
            }
        }

        /// <summary>
        /// This test case is designed to implement using RopSeekRow to create a table view and subscribe to the TableModified event, and then verify the TableRowModified event which is a type of TableModified event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC04_VerifyTableRowModifyEventAndCreateTableViewByTableSeekRow()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);

            // The first content table handle is used to trigger the notification.
            uint contentTableHandle1;

            // The second content table handle is used to get the information from the specified table.
            uint contentTableHandle2 = 0;
            this.GetContentsTable(inboxTableHandle, out contentTableHandle1, false);
            #endregion

            #region Create table view by SeekRow
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[]
            { 
                PropertyTags.All[PropertyNames.PidTagInstID], 
                PropertyTags.All[PropertyNames.PidTagImportance],
                PropertyTags.All[PropertyNames.PidTagInstanceNum],
                PropertyTags.All[PropertyNames.PidTagFolderId],
                PropertyTags.All[PropertyNames.PidTagMessageClass]           
            };

            this.SetColumns(contentTableHandle1, tags);

            // Seek the table from the position 1
            this.SeekRow(contentTableHandle1, 1);
            #endregion

            #region Trigger tableRowModified event and get notification
            this.TriggerTableRowModifiedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for tableRowModified event
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                    Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableRowModified, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableRowModified.");
                    this.VerifyTableModifyNotificationFlag(notifyResponse);
                    this.VerifyTableRowModifiedNotificationElements(notifyResponse);

                    PropertyTag[] tags1 = new PropertyTag[] { PropertyTags.All[PropertyNames.PidTagFolderId], PropertyTags.All[PropertyNames.PidTagInstanceNum], PropertyTags.All[PropertyNames.PidTagMid] };

                    // Open Inbox folder and get content table of the Inbox folder
                    this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
                    this.GetContentsTable(inboxTableHandle, out contentTableHandle2, false);

                    // Create table view by QueryRows
                    this.SetColumns(contentTableHandle2, tags1);

                    // Retrieves rows from content table to get the data of 30 rows
                    RopQueryRowsResponse queryResponse = this.QueryRows(contentTableHandle2, 30);

                    ulong insertAfterTableRowIDFromTable = 0, insertAfterTableRowFolderIDFromTable = 0;

                    // Get the old value of the Message ID structure of the item triggering the notification.
                    for (int i = 0; i < queryResponse.RowData.PropertyRows.Count; i++)
                    {
                        if (BitConverter.ToUInt64(queryResponse.RowData.PropertyRows[i].PropertyValues[2].Value, 0) == this.TriggerMessageId)
                        {
                            // If the message that triggers the notification is in the first row, the insert row should be 0. 
                            // Otherwise, the inserted row is the previous row the message which trigger the notification.
                            if (i == 0)
                            {
                                insertAfterTableRowIDFromTable = 0;
                                insertAfterTableRowFolderIDFromTable = 0;
                            }
                            else
                            {
                                insertAfterTableRowIDFromTable = BitConverter.ToUInt64(queryResponse.RowData.PropertyRows[i - 1].PropertyValues[2].Value, 0);
                                insertAfterTableRowFolderIDFromTable = BitConverter.ToUInt64(queryResponse.RowData.PropertyRows[i - 1].PropertyValues[0].Value, 0);
                            }
                        }
                    }

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R157");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R157
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        insertAfterTableRowIDFromTable,
                        notifyResponse.NotificationData.InsertAfterTableRowID,
                        157,
                        @"[In NotificationData Structure] InsertAfterTableRowID: The old value of the Message ID structure of the item triggering the notification.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R153");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R153
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        insertAfterTableRowFolderIDFromTable,
                        notifyResponse.NotificationData.InsertAfterTableRowFolderID,
                        153,
                        @"[In NotificationData Structure] InsertAfterTableRowFolderID: The old value of the Folder ID structure of the item triggering the notification.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R161003");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R161003
                    this.Site.CaptureRequirementIfIsTrue(
                        notifyResponse.NotificationData.InsertAfterTableRowInstance != null && (notifyResponse.NotificationData.NotificationFlags & 0x8000) == 0x8000 && notifyResponse.NotificationData.TableEventType == 0x0005,
                        161003,
                        @"[In NotificationData Structure] This field [InsertAfterTableRowInstance] is available when bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0005.");


                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R263: The value of the Folder ID and Message ID separately are {0},{1}", notifyResponse.NotificationData.TableRowFolderID, notifyResponse.NotificationData.TableRowMessageID);

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R263
                    bool isVerifiedR263 = notifyResponse.NotificationData.TableRowFolderID != null && notifyResponse.NotificationData.TableRowMessageID != null;

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR263,
                        263,
                        @"[In Creating and Sending TableModified Event Notifications] [When a TableModified event occurs, the server generates a notification using one of the following three methods, listed in descending order of usefulness to the client.] For TableRowModified event, the server generates an informative notification that specifies the nature of the change (content or hierarchy), the value of the Folder ID structure, as specified in [MS-OXCDATA] section 2.2.1.1, the value of the Message ID structure, as specified in [MS-OXCDATA] section 2.2.1.2, and new table values.");

                    this.VeriyServerGenerateInformativeNotification(isVerifiedR263);

                    Site.Assert.IsNotNull(notifyResponse.NotificationData.TableEventType, "The TableEventType in the RopNotifyResponse should not null.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R131");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R131
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0005,
                        (int)notifyResponse.NotificationData.TableEventType,
                        131,
                        @"[In NotificationData Structure] [TableEventType value] 0x0005: The notification is for TableRowModified events.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R25");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R25
                    // The client can receive TableRowModified notification indicates that an existing row has been modified in the table.
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0005,
                        (int)notifyResponse.NotificationData.TableEventType,
                        25,
                        @"[In TableModified Event Types] TableRowModified: An existing row has been modified in the table.");

                    // Only Exchange 2010 and above require a table view, this server version restrict is the same with MS-OXCNTOIF_R245.
                    if (Common.IsRequirementEnabled(245, this.Site))
                    {
                        // If the server can return notification response, it can verify RopQueryRows create a table view. so the following requirements can be captured directly. 
                        Site.CaptureRequirement(
                            253,
                            @"[In Creating and Sending TableModified Event Notifications] If a table view is required on the server, the server MUST receive a request from one of the following ROPs, each of which cause a table view to be created on the server: RopSeekRow ([MS-OXCROPS] section 2.2.5.8)");

                        this.VerifyTableViewCreated();
                    }
                }
            }
            #endregion

            // Only Exchange 2010 and above will stop sending notifications if the RopResetTable ROP is received, until a new table view is created. 
            if (Common.IsRequirementEnabled(283, this.Site))
            {
                #region Reset the contents table
                // Reset the first content table handle is used to trigger the notification.
                this.ResetTable(contentTableHandle1);

                // Reset the second content table handle is used to get the information from the specified table.
                this.ResetTable(contentTableHandle2);
                #endregion

                #region Trigger TableRowAdded event after ResetTable and get notification
                this.TriggerTableRowAddedEvent();

                // Trigger table event after ResetTable, server shouldn't send notification response on Exchange server 2010 and above
                rsp = this.CNOTIFAdapter.GetNotification(false);
                bool isServerCreateSubscription = false;
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsFalse(isServerCreateSubscription, "The server can't send notify response when resetting table.");
                #endregion

                #region Create table view by SeekRow
                this.SetColumns(contentTableHandle1, tags);

                // Seek the table from the position 1
                this.SeekRow(contentTableHandle1, 1);
                #endregion

                #region Trigger TableRowAdded event again and get notification
                this.TriggerTableRowAddedEvent();

                // Create table view by SeekRow after ResetTable, and trigger table event again the server should send notification response.
                rsp = this.CNOTIFAdapter.GetNotification(true);
                Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsTrue(isServerCreateSubscription, "After resetting table, the server should send notify response when a new table view is created by RopSeekRow.");
                #endregion

                #region Verify notification response after reset Table and create table view by SeekRow

                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R283");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R283
                // Since the server return the notification after a table view is created by RopSeekRow, this requirement can be verified directly.
                Site.CaptureRequirement(
                    283,
                    @"[In Appendix A: Product Behavior] The implementation does stop sending notifications if the RopResetTable ROP ([MS-OXCROPS] section 2.2.5.15) is received, until a new table view is created using one of the following ROPs: RopSeekRow. (Exchange 2010 and above follow this behavior.)");
                #endregion
            }
        }

        /// <summary>
        /// This test case is designed to implement using RopSeekRowBookmark to create a table view and subscribe to the TableModified event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC05_VerifyCreateTableViewByTableSeekRowBookmark()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
            uint contentTableHandle;
            this.GetContentsTable(inboxTableHandle, out contentTableHandle, false);
            #endregion

            #region Create table view by SeekRowBookmark
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[]
            { 
                PropertyTags.All[PropertyNames.PidTagInstID], 
                PropertyTags.All[PropertyNames.PidTagImportance],
                PropertyTags.All[PropertyNames.PidTagInstanceNum],
                PropertyTags.All[PropertyNames.PidTagFolderId],
                PropertyTags.All[PropertyNames.PidTagMessageClass]
            };

            // Create table view by SeekRowBookmark
            this.SetColumns(contentTableHandle, tags);
            RopCreateBookmarkResponse bookmarkRsp = this.CreateBookmark(contentTableHandle);

            // Move the cursor to the position 2
            this.SeekRowBookmark(contentTableHandle, bookmarkRsp.Bookmark, 2);
            #endregion

            #region Trigger TableRowAdded event and get notification
            this.TriggerTableRowAddedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify table view is created by SeekRowBookmark
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                    Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableRowAdded, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableRowAdded.");
                    this.VerifyTableModifyNotificationFlag(notifyResponse);

                    // Only Exchange 2010 and above require a table view, this server version restrict is the same with MS-OXCNTOIF_R245.
                    if (Common.IsRequirementEnabled(245, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R255");

                        // If the server can return notification response, it can verify RopQueryRows create a table view. so the following requirements can be captured directly. 
                        Site.CaptureRequirement(
                            255,
                            @"[In Creating and Sending TableModified Event Notifications] If a table view is required on the server, the server MUST receive a request from one of the following ROPs, each of which cause a table view to be created on the server: RopSeekRowBookmark ([MS-OXCROPS] section 2.2.5.9)");

                        this.VerifyTableViewCreated();
                    }
                }
            }
            #endregion

            // Only Exchange 2010 and above will stop sending notifications if the RopResetTable ROP is received, until a new table view is created. 
            if (Common.IsRequirementEnabled(286, this.Site))
            {
                #region Reset the contents table
                this.ResetTable(contentTableHandle);
                #endregion

                #region Trigger ObjectDeleted event and get notification
                this.TriggerObjectDeletedEvent();

                // Trigger table event after ResetTable, server shouldn't send notification response on Exchange server 2010 and above
                rsp = this.CNOTIFAdapter.GetNotification(false);
                bool isServerCreateSubscription = false;
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsFalse(isServerCreateSubscription, "The server shouldn't send notify response when reset table");
                #endregion

                #region Create table view by SeekRowBookmark
                this.SetColumns(contentTableHandle, tags);
                this.CreateBookmark(contentTableHandle);

                // Move the cursor to the position 2
                this.SeekRowBookmark(contentTableHandle, bookmarkRsp.Bookmark, 2);
                #endregion

                #region Trigger tableRowModified event and get notification
                this.TriggerTableRowModifiedEvent();

                // Create table view by SeekRowBookmark after ResetTable, and trigger table event again the server should send notification response.
                rsp = this.CNOTIFAdapter.GetNotification(true);
                Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");

                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }
                #endregion

                #region Verify notification response after reset Table and create table view by RopSeekRowBookmark
                // Since the server return the notification after a table view is created by RopSeekRowBookmark, this requirement can be verified.
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R286: The implementation does return notifications after get RopSeekRowBookmark", isServerCreateSubscription ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R286
                bool isR286Satisfied = isServerCreateSubscription;

                Site.CaptureRequirementIfIsTrue(
                    isR286Satisfied,
                    286,
                    @"[In Appendix A: Product Behavior] The Implementation does stop sending notifications if the RopResetTable ROP ([MS-OXCROPS] section 2.2.5.15) is received, until a new table view is created using one of the following ROPs: RopSeekRowBookmark. (Exchange 2010 and above follow this behavior.)");
                #endregion
            }
        }

        /// <summary>
        /// This test case is designed to implement using RopQueryPosition to create a table view and subscribe to the TableModified event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC06_VerifyCreateTableViewByTableQueryPosition()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
            uint contentTableHandle;
            this.GetContentsTable(inboxTableHandle, out contentTableHandle, false);
            #endregion

            #region Create table view by QueryPosition
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[]
            {
                PropertyTags.All[PropertyNames.PidTagInstID], 
                PropertyTags.All[PropertyNames.PidTagImportance],
                PropertyTags.All[PropertyNames.PidTagInstanceNum],
                PropertyTags.All[PropertyNames.PidTagFolderId],
                PropertyTags.All[PropertyNames.PidTagMessageClass]
            };
            this.SetColumns(contentTableHandle, tags);
            this.QueryPosition(contentTableHandle);
            #endregion

            #region Trigger TableRowAdded event and get notification
            this.TriggerTableRowAddedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify table view is created by RopQueryPosition
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                    Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableRowAdded, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableRowAdded.");
                    this.VerifyTableModifyNotificationFlag(notifyResponse);

                    // Only Exchange 2010 and above require a table view, this server version restrict is the same with MS-OXCNTOIF_R245.
                    if (Common.IsRequirementEnabled(245, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R251");

                        // If the server can return notification response, it can verify RopQueryRows create a table view. so the following requirements can be captured directly. 
                        Site.CaptureRequirement(
                            251,
                            @"[In Creating and Sending TableModified Event Notifications] If a table view is required on the server, the server MUST receive a request from one of the following ROPs, each of which cause a table view to be created on the server: RopQueryPosition ([MS-OXCROPS] section 2.2.5.7).");

                        this.VerifyTableViewCreated();
                    }
                }
            }
            #endregion

            // Only Exchange 2010 and above will stop sending notifications if the RopResetTable ROP is received, until a new table view is created. 
            if (Common.IsRequirementEnabled(279, this.Site))
            {
                #region Reset the contents table
                this.ResetTable(contentTableHandle);
                #endregion

                #region Trigger TableRowAdded event and get notification
                this.TriggerTableRowAddedEvent();

                // Trigger table event after ResetTable, server shouldn't send notification response on Exchange server 2010 and above
                rsp = this.CNOTIFAdapter.GetNotification(false);
                bool isServerCreateSubscription = false;
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }

                Site.Assert.IsFalse(isServerCreateSubscription, "The server can't send notify response when reset table.");
                #endregion

                #region Create table view by QueryPosition
                this.SetColumns(contentTableHandle, tags);
                this.QueryPosition(contentTableHandle);
                #endregion

                #region Trigger TableRowAdded event and get notification
                this.TriggerTableRowAddedEvent();

                // Create table view by QueryPosition after ResetTable, and trigger table event again the server should send notification response.
                rsp = this.CNOTIFAdapter.GetNotification(true);
                Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        isServerCreateSubscription = true;
                    }
                }
                #endregion

                #region Verify notification response after reset Table and create table view by RopQueryPosition
                // Since the server return the notification after a table view is created by RopQueryPosition, this requirement can be verified.
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R279: The implementation does return notifications after get RopQueryPosition", isServerCreateSubscription ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R279
                bool isR279Satisfied = isServerCreateSubscription;

                Site.CaptureRequirementIfIsTrue(
                    isR279Satisfied,
                    279,
                    @"[In Appendix A: Product Behavior] Implementation does stop sending notifications if the RopResetTable ROP ([MS-OXCROPS] section 2.2.5.15) is received, until a new table view is created using one of the following ROPs: RopQueryPosition. (Exchange 2010 and above follow this behavior.)");
                #endregion
            }
        }

        /// <summary>
        /// This test case is designed to verify that Exchange 2007 does not require that a table view is created in order to send table notifications. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC07_VerifyTableViewNotRequired()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(244, this.Site), "This case runs only under Exchange 2007, since Exchange 2010 and Exchange 2013 does require that a table view is created in order to send table notifications");

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);

            uint contentTableHandle;
            this.GetContentsTable(inboxTableHandle, out contentTableHandle, false);
            #endregion

            #region Trigger TableRowDeleted event and get notification
            this.TriggerTableRowDeletedEvent();

            bool isGetNotification = false;
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    isGetNotification = true;
                }
            }
            #endregion

            #region Verify notification response when no table view is created

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R244: the implementation does {0} get the notification without table view.", isGetNotification ? string.Empty : "not");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R244 
            // When there is no table view is created on the server, client can also get notifications from server, this requirement can be verified.
            this.Site.CaptureRequirementIfIsTrue(
                isGetNotification,
                244,
                @"[In Appendix A: Product Behavior] [If the client has subscribed to TableModified event notifications, by using the RopRegisterNotification ROP] Implementation does not require that a table view is created in order to send tale notifications. (<12> Section 3.1.4.3: Exchange 2007 does not require that a table view is created in order to send table notifications.)");

            if (Common.IsRequirementEnabled(372, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R372: the implementation does {0} get the notification without table view.", isGetNotification ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R372
                // Server does create the subscription automatically without table view is created.                    
                this.Site.CaptureRequirementIfIsTrue(
                    isGetNotification,
                    372,
                    @"[In Appendix A: Product Behavior] Implementation does created the subscription automatically, when the client creates a Table object on the implementation. (In Exchange 2007, the subscription is created automatically when the client creates a Table object on the server.)");
            }
            #endregion

            #region Reset the contents table
            this.ResetTable(contentTableHandle);
            #endregion

            #region Trigger TableRowAdded event and get notification
            this.TriggerTableRowAddedEvent();
            rsp = this.CNOTIFAdapter.GetNotification(true);
            bool isServerCreateSubscription = false;
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    isServerCreateSubscription = true;
                }
            }
            #endregion

            #region Verify whether the implementation stops sending notifications when the RopResetTable ROP is received.
            if (Common.IsRequirementEnabled(294, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R294: the implementation does {0} send the notification when the RopResetTable ROP is received", isServerCreateSubscription ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R294
                bool isR294Satisfied = isServerCreateSubscription;

                this.Site.CaptureRequirementIfIsTrue(
                    isR294Satisfied,
                    294,
                    @"[In Appendix A: Product Behavior] Implementation does not stop sending notification if the RopResetTable ROP is received. (<13> Section 3.1.4.3: Exchange 2007 does not stop sending notifications if the RopResetTable ROP ([MS-OXCROPS] section 2.2.5.15) is received.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the NewMail event and the elements related to the event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC08_VerifyNewMailEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe NewMail event on server
            uint notificationHandle;

            // Creates a subscription for notifications in the entire database, so the WantWholeStore should be true (0x01), flolderId and messageId should be not present.
            this.CNOTIFAdapter.RegisterNotificationWithParameter(NotificationType.NewMail, 1, 0, 0, out notificationHandle);
            #endregion

            #region Trigger NewMail event and get notification
            this.TriggerNewMailEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for NewMail event
            bool isNewMail = false;

            // Check whether the value of NotificationHandle is same as the handle created by the method RegisterNotification.
            bool isNotificationSubcription = false;
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.NewMail, "The notification type should be NewMail.");
                    isNewMail = true;

                    #region Get messageFlags from table
                    uint folderHandle, messageHandle;
                    this.OpenFolder(this.InboxFolderId, out folderHandle);

                    Site.Assert.IsNotNull(notifyResponse.NotificationData.MessageId, "The MessageId in the RopNotifyResponse should not null.");
                    this.OpenMessage(folderHandle, this.InboxFolderId, (ulong)notifyResponse.NotificationData.MessageId, out messageHandle);
                    RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.GetPropertiesSpecific(messageHandle, new PropertyTag[] { PropertyTags.All[PropertyNames.PidTagMessageFlags] });
                    uint messageFlagsFromTable = BitConverter.ToUInt32(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, 0);
                    #endregion

                    #region Verify element MessageFlags of the notification response
                    if (Common.IsRequirementEnabled(214001, this.Site))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R214001");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R214001
                        this.Site.CaptureRequirementIfAreEqual<uint?>(
                            messageFlagsFromTable,
                            notifyResponse.NotificationData.MessageFlags,
                            214001,
                            @"[In Appendix A: Product Behavior] MessageFlags does specify the message flags of new mail that has been received.(Exchange 2007, Exchange 2010, Exchange 2016 and above follow this behavior.)");
                    }

                    if(Common.IsRequirementEnabled(214002,this.Site))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R214002");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R214002
                        this.Site.CaptureRequirementIfAreEqual<uint?>(
                            0,
                            notifyResponse.NotificationData.MessageFlags,
                            214002,
                            @"[In Appendix A: Product Behavior] Implementation does return zero for MessageFlags. <10> Section 2.2.1.4.1.2:  In Exchange 2013 the value of MessageFlags is zero. (Exchange 2013 follows this behavior.)");
                    }
                    #endregion

                    #region Get messageId from table
                    uint contentTableHandle;
                    RopGetContentsTableResponse contentResponse = this.GetContentsTable(folderHandle, out contentTableHandle, false);

                    // The properties need to be set
                    PropertyTag[] tags = new PropertyTag[] { PropertyTags.All[PropertyNames.PidTagMid] };
                    this.SetColumns(contentTableHandle, tags);
                    RopQueryRowsResponse queryResponse = this.QueryRows(contentTableHandle, (ushort)contentResponse.RowCount);
                    byte[] messageID = queryResponse.RowData.PropertyRows[0].PropertyValues[0].Value;
                    ulong messageIDFromTable = BitConverter.ToUInt64(messageID, 0);
                    #endregion

                    #region Verify element MessageId of the notification response
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R174");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R174
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        messageIDFromTable,
                        notifyResponse.NotificationData.MessageId,
                        174,
                        @"[In NotificationData Structure] MessageId: The Message ID structure, as specified in [MS-OXCDATA] section 2.2.1.2, of the item triggering the event.");
                    #endregion

                    // Check whether the value of NotificationHandle is same as the handle created by the method RegisterNotification.
                    if (notifyResponse.NotificationHandle == notificationHandle)
                    {
                        isNotificationSubcription = true;
                    }

                    this.VerifyNewMailNotificationElements(notifyResponse);
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R96");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R96
            // If the value of NotificationHandle is same as the handle created by the method RegisterNotification, this requirement can be verified.
            this.Site.CaptureRequirementIfIsTrue(
                isNotificationSubcription,
                96,
                @"[In RopNotify ROP Response Buffer] [NotificationHandle] The target object can be a notification subscription.");

            Site.Assert.IsTrue(isNewMail, "Received the Notification for the new mail.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R103");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R103
            this.Site.CaptureRequirement(
                103,
                @"[In NotificationData Structure] [NotificationType value] 0x0002: The notification is for a NewMail event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R11");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R11
            // Client can receive NewMail notification indicates that a new email message has been received by the server.
            this.Site.CaptureRequirement(
                11,
                @"[In Server Event Types] NewMail: A new email message has been received by the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R41");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R41
            bool isVerifiedR41 = isNewMail;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR41,
                41,
                @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0002: The server sends notifications to the client when NewMail events occur within the scope of interest.");

            this.VerifyServerCreateSessionContaintAndSaveInformation(isNewMail);
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the ObjectCopied event and the elements related to the event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC09_VerifyObjectCopiedEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe ObjectCopied event on server
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCopied);
            #endregion

            #region Trigger ObjectCopied event and get notification
            this.TriggerObjectCopiedEvent();

            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for ObjectCopied event
            bool isObjectCopied = false;
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;

                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.ObjectCopied, "The notification type should be ObjectCopied.");
                    isObjectCopied = true;

                    #region Verify elements OldFolderId and OldMessageId of the notification response
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R182");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R182
                    // The InboxFolderId is the folder of the message originally exists. So if the value of OldFolderId is equal to InboxFolderId, this requirement can be verified.
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        this.InboxFolderId,
                        notifyResponse.NotificationData.OldFolderId,
                        182,
                        @"[In NotificationData Structure] OldFolderId: The old Folder ID structure of the item triggering the event.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R186");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R186
                    // The newMessageId1 is the message originally is. So if the value of OldMessageId is equal to newMessageId1, this requirement can be verified.
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        this.TriggerMessageId,
                        notifyResponse.NotificationData.OldMessageId,
                        186,
                        @"[In NotificationData Structure] OldMessageId: The old Message ID structure of the item triggering the event.");
                    #endregion

                    this.VerifyObjectCopiedNotificationElements(notifyResponse);
                }
            }

            Site.Assert.IsTrue(isObjectCopied, "The notification for an ObjectCopied event should be received.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R108");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R108
            this.Site.CaptureRequirement(
                108,
                @"[In NotificationData Structure] [NotificationType value] 0x0040: The notification is for an ObjectCopied event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R12");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R12
            // Client can receive a ObjectCopied notification indicates that an existing object has been copied on the server.
            this.Site.CaptureRequirement(
                12,
                @"[In Server Event Types] ObjectCopied: An existing object has been copied on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R46");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R46            
            this.Site.CaptureRequirement(
                46,
                @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0040: The server sends notifications to the client when ObjectCopied events occur within the scope of interest.");

            this.VeriyServerGenerateInformativeNotification(isObjectCopied);
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the ObjectCreated event and the elements related to the event. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC10_VerifyObjectCreatedEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe ObjectCreated event on server
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCreated);
            #endregion

            #region Trigger ObjectCreated event and get notification
            ulong createdFodlerId = this.TriggerObjectCreatedEvent();

            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for ObjectCreated event
            bool isObjectCreated = false;
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.ObjectCreated, "The notification type should be ObjectCreated.");
                    isObjectCreated = true;
                    this.VerifyObjectCreatedNotificationElements(notifyResponse);

                    #region Verify elements FolderId, ParentFolderId and Tags of the notification response
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R170");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R170
                    // The folder of createdFodlerId is the folder which trigger the notification.
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        createdFodlerId,
                        notifyResponse.NotificationData.FolderId,
                        170,
                        @"[In NotificationData Structure] FolderId: The Folder ID structure of the item triggering the event.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R178");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R178
                    // The folder of InboxFolderId is the parent of the folder which trigger the notification.
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        this.InboxFolderId,
                        notifyResponse.NotificationData.ParentFolderId,
                        178,
                        @"[In NotificationData Structure] ParentFolderId: The Folder ID structure of the parent folder of the item triggering the event.");

                    if (notifyResponse.NotificationData.TagCount == 0)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R20302");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R20302
                        this.Site.CaptureRequirementIfIsNull(
                            notifyResponse.NotificationData.Tags,
                            20302,
                            @"[In NotificationData Structure] This field [Tags] is not available if the TagCount field is available and the value of the TagCount field is 0x0000.");
                    }

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R161004");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R161004
                    this.Site.CaptureRequirementIfIsTrue(
                        (notifyResponse.NotificationData.NotificationFlags & 0x8000)==0x0000 && notifyResponse.NotificationData.InsertAfterTableRowInstance==null,
                        161004,
                        @"[In NotificationData Structure] This field [InsertAfterTableRowInstance] is not available when bit 0x8000 is not set in the NotificationFlags field.");

                    #endregion
                }
            }

            Site.Assert.IsTrue(isObjectCreated, "The Notification response for ObjectCreated event should be received.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R104");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R104
            this.Site.CaptureRequirement(
                104,
                @"[In NotificationData Structure] [NotificationType value] 0x0004: The notification is for an ObjectCreated event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R13");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R13
            // Client can receive ObjectCreated notification indicates that a new object has been created on the server.
            this.Site.CaptureRequirement(
                13,
                @"[In Server Event Types] ObjectCreated: A new object has been created on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R42");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R42
            this.Site.CaptureRequirement(
                42,
                @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0004: The server sends notifications to the client when ObjectCreated events occur within the scope of interest.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R52");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R52
            // Set WantWholeStore to 1 when register the notification. So if the created folder can trigger successfully, which means a random folder can be triggered, this requirement can be verified.
            this.Site.CaptureRequirement(
                52,
                @"[In RopRegisterNotification ROP Request Buffer] WantWholeStore: A value of TRUE (0x01) if the scope for notifications is the entire mailbox.");
            this.VeriyServerGenerateInformativeNotification(isObjectCreated);
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the ObjectDeleted event and the elements related to the event. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC11_VerifyObjectDeletedEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(53, this.Site), "This case runs only under Exchange 2007 and Exchange 2010, since Exchange 2013 does not support the value of WantWholeStore is 0x00 when the notification type is ObjectDeleted.");

            #region Subscribe ObjectDeleted event on server
            uint notificationHandle;
            this.CNOTIFAdapter.RegisterNotificationWithParameter(NotificationType.ObjectDeleted, 0, this.NewFolderId, this.TriggerMessageId, out notificationHandle);
            #endregion

            #region Trigger ObjectDeleted event and get notification
            this.TriggerObjectDeletedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            #endregion

            #region Verify notification response for ObjectDeleted event
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R105");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R105
                    this.Site.CaptureRequirementIfAreEqual<NotificationType>(
                        NotificationType.ObjectDeleted,
                        (NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF),
                        105,
                        @"[In NotificationData Structure] [NotificationType value] 0x0008: The notification is for an ObjectDeleted event.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R337");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R337
                    // Client send a request without any ROP request. So if server can successfully return ROP request, this requirement can be covered.
                    this.Site.CaptureRequirementIfAreNotEqual<int>(
                        0,
                        rsp.Count,
                        337,
                        @"[In Receiving an EcDoRpcExt] The server does not require that the EcDoRpcExt2 method call include a ROP request.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14
                    // Client can receive a ObjectDeleted notification indicates that an existing object has been deleted from the server.
                    this.Site.CaptureRequirement(
                        14,
                        @"[In Server Event Types] ObjectDeleted: An existing object has been deleted from the server.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R43");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R43
                    this.Site.CaptureRequirement(
                        43,
                        @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0008: The server sends notifications to the client when ObjectDeleted events occur within the scope of interest.");

                    if (Common.IsRequirementEnabled(53, this.Site))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R53");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R53
                        // Set WantWholeStore to 0, and  when register the notification and delete the folder which trigger the notification. So if the deleted folder can trigger successfully, which means a random folder can be triggered, this requirement can be covered.
                        this.Site.CaptureRequirement(
                            53,
                            @"[In RopRegisterNotification ROP Request Buffer] otherwise [if the scope for notifications is not the entire mailbox], [the value of WantWholeStore is] FALSE (0x00).");
                    }

                    this.VeriyServerGenerateInformativeNotification(true);
                }
            }

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the ObjectModified event and the elements related to the event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC12_VerifyObjectModifiedEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe ObjectModified event on server
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectModified);
            #endregion

            #region Trigger ObjectModified event and get notification
            this.TriggerObjectModifiedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for ObjectModified event
            bool isObjectModified = false;
            bool gotRopNotifyResponse = false;
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    gotRopNotifyResponse = true;
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.ObjectModified, "The notification type should be ObjectModified.");
                    isObjectModified = true;

                    #region Verify elements TagCount and Tags of the notification response
                    if (notifyResponse.NotificationData.TagCount != null)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R194");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R194
                        bool isVerifiedR194;

                        if (notifyResponse.NotificationData.TagCount == 0)
                        {
                            isVerifiedR194 = notifyResponse.NotificationData.Tags == null;
                        }
                        else
                        {
                            isVerifiedR194 = notifyResponse.NotificationData.Tags.Length == notifyResponse.NotificationData.TagCount;
                        }

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR194,
                            194,
                            @"[In NotificationData Structure] TagCount: An unsigned 16-bit integer that specifies the number of property tags in the Tags field.");

                        if (notifyResponse.NotificationData.TagCount != 0x0000 && notifyResponse.NotificationData.TagCount != 0xFFFF)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R20301");

                            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R20301
                            this.Site.CaptureRequirementIfIsNotNull(
                                notifyResponse.NotificationData.Tags,
                                20301,
                                @"[In NotificationData Structure] This field [Tags] is available if the TagCount field is available and the value of the TagCount field is  not 0x0000 or 0xFFFF .");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R202");

                            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R202
                            this.Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(uint[]),
                                notifyResponse.NotificationData.Tags.GetType(),
                                202,
                                @"[In NotificationData Structure] Tags (variable): An array of unsigned 32-bit integers that identifies the IDs of properties that have changed.");

                            if (Common.IsRequirementEnabled(198, this.Site))
                            {
                                Site.Assert.IsNotNull(notifyResponse.NotificationData.TagCount, "The TagCount should not null.");

                                // Add the debug information
                                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R198");

                                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R198
                                this.Site.CaptureRequirementIfAreEqual<int>(
                                    notifyResponse.NotificationData.Tags.Length,
                                    (int)notifyResponse.NotificationData.TagCount,
                                    198,
                                    @"[In Appendix A: Product Behavior] [If the value of the NotificationType field in the NotificationFlags field is 0x0010] Implementation does not set the value of the TagCount field to 0x0000. (<9> Section 2.2.1.4.1.2:  Exchange 2007, and Exchange 2010 do not set the value of the TagCount field to 0x0000; they set the value of the field to the number of property tags in the Tags field.)");
                            }
                        }
                    }

                    if (Common.IsRequirementEnabled(199, this.Site))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R199");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R199
                        this.Site.CaptureRequirementIfAreEqual<ushort?>(
                            0,
                            notifyResponse.NotificationData.TagCount,
                            199,
                            @"[In Appendix A: Product Behavior] [If the value of the NotificationType field in the NotificationFlags field is 0x0010] Implementation will set the value of this field [TagCount] to 0x0000. (Exchange 2013 and above follow this behavior)");
                    }
                    #endregion

                    this.VerifyObjectModifiedNotificationElements(notifyResponse);
                }
            }

            if (Common.IsRequirementEnabled(510, this.Site) && Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http")
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R510");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R510
                // Get the RopNotifyResponse via MAPI transport, MS-OXCNOTIF_R510 can be verified.
                this.Site.CaptureRequirementIfIsTrue(
                    gotRopNotifyResponse,
                    510,
                    @"[In Appendix A: Product Behavior] Implementation does support the Execute request type. (<8> Section 2.2.1.4.1.1:  The Execute request type was introduced in Exchange 2013 SP1.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R106");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R106
            bool isVerifiedR106 = isObjectModified;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR106,
                106,
                @"[In NotificationData Structure] [NotificationType value] 0x0010: The notification is for an ObjectModified event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15
            // Client can receive a ObjectModified notification indicates that an existing object has been modified on the server.
            bool isVerifiedR15 = isObjectModified;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR15,
                15,
                @"[In Server Event Types] ObjectModified: An existing object has been modified on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R44");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R44
            bool isVerifiedR44 = isObjectModified;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR44,
                44,
                @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0010: The server sends notifications to the client when ObjectModified events occur within the scope of interest.");

            this.VerifyServerCreateSessionContaintAndSaveInformation(isObjectModified);
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the ObjectMoved event with moving folder and the elements related to the event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC13_VerifyObjectMovedFolderEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe ObjectMoved event on server
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectMoved);
            #endregion

            #region Trigger ObjectModified event and get notification
            this.TriggerObjectMovedFolderEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for ObjectMoved event
            bool isObjectMoved = false;
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.ObjectMoved, "The notification type should be ObjectMoved.");
                    isObjectMoved = true;

                    #region Verify element OldParentFolderId of notification response
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R190");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R190
                    // The folder of InboxFolderId is the older parent folder which trigger the notification. So if the value of OldParentFolderId equals to InboxFolderId, this requirement can be verified.
                    this.Site.CaptureRequirementIfAreEqual<ulong?>(
                        this.InboxFolderId,
                        notifyResponse.NotificationData.OldParentFolderId,
                        190,
                        @"[In NotificationData Structure] OldParentFolderId: The old parent Folder ID structure of the item triggering the event.");
                    #endregion

                    this.VerifyObjectMovedNotificationElements(notifyResponse);
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R107");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R107
            bool isVerifiedR107 = isObjectMoved;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR107,
                107,
                @"[In NotificationData Structure] [NotificationType value] 0x0020: The notification is for an ObjectMoved event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R16");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R16
            // Client can receive a ObjectMoved notification indicates that an existing object has been moved to another location on the server.
            bool isVerifiedR16 = isObjectMoved;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR16,
                16,
                @"[In Server Event Types] ObjectMoved: An existing object has been moved to another location on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R45");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R45
            bool isVerifiedR45 = isObjectMoved;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR45,
                45,
                @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0020: The server sends notifications to the client when ObjectMoved events occur within the scope of interest.");

            this.VerifyServerCreateSessionContaintAndSaveInformation(isObjectMoved);
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the SearchCompleted event and the elements related to the event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC14_VerifySearchCompletedEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe SearchCompleted event on server
            this.CNOTIFAdapter.RegisterNotification(NotificationType.SearchCompleted);
            #endregion

            #region Trigger SearchCompleted event and get notification
            this.TriggerSearchCompletedEvent();

            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response for SearchCompleted event
            bool isSearchCompleted = false;
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;

                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.SearchCompleted, "The notification type should be SearchCompleted.");
                    isSearchCompleted = true;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R109");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R109
            bool isVerifiedR109 = isSearchCompleted;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR109,
                109,
                @"[In NotificationData Structure] [NotificationType value] 0x0080: The notification is for a SearchCompleted event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17
            // Client can receive SearchCompleted notification indicates that a search operation has been completed on the server.
            bool isVerifiedR17 = isSearchCompleted;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR17,
                17,
                @"[In Server Event Types] SearchCompleted: A search operation has been completed on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R47");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R47
            bool isVerifiedR47 = isSearchCompleted;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR47,
                47,
                @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0080: The server sends notifications to the client when SearchCompleted events occur within the scope of interest.");

            this.VerifyServerCreateSessionContaintAndSaveInformation(isSearchCompleted);
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the server allows multiple Notification subscription objects to be created and associated with the same session context.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC15_VerifyMultipleServerEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Trigger SearchCompleted event
            this.TriggerSearchCompletedEvent();
            #endregion

            #region Register and trigger ObjectCopied, ObjectCreated and ObjectDeleted events
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCopied);
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCreated);
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectDeleted);

            this.TriggerNewMailEvent();
            this.TriggerObjectCopiedEvent();
            this.TriggerObjectCreatedEvent();
            this.TriggerObjectDeletedEvent();
            #endregion

            #region Wait the above events finished and get notification

            bool isObjectCopied = false;
            bool isObjectCreated = false;
            bool isObjectDeleted = false;

            IList<IDeserializable> rsp;
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            List<List<uint>> responseSOHs;
            bool isGetNotification = false;
            do
            {
                if (!isGetNotification)
                {
                    // Try to get the notifications until the server generates the notifications.
                    rsp = this.CNOTIFAdapter.GetNotification(true);
                    Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
                    isGetNotification = true;
                }
                else
                {
                    // Once the server generates the notification, the notifications can be gotten directly.
                    rsp = this.CNOTIFAdapter.Process(
                        null,
                        this.CNOTIFAdapter.LogonHandle,
                        out responseSOHs);
                }

                retryCount--;

                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                        switch ((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF))
                        {
                            case NotificationType.ObjectCopied:
                                isObjectCopied = true;
                                break;
                            case NotificationType.ObjectCreated:
                                isObjectCreated = true;
                                break;
                            case NotificationType.ObjectDeleted:
                                isObjectDeleted = true;
                                if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) != (ushort)FlagsBit.M && (notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.S) != (ushort)FlagsBit.S)
                                {
                                    // Add the debug information
                                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17902");

                                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17902
                                    this.Site.CaptureRequirementIfIsNotNull(
                                        notifyResponse.NotificationData.ParentFolderId,
                                        17902,
                                        @"[In NotificationData Structure] This field [ParentFolderId] is available if the value of the NotificationType field is 0x0008, and it [RopNotify ROP] is sent for a message in a folder (both bit 0x4000 and bit 0x8000 are not set in the NotificationFlags field).");
                                }

                                break;
                        }
                    }
                }
            }
            while (rsp.Count > 0 && retryCount > 0);
            Site.Assert.IsTrue(
                rsp.Count == 0,
                "The left notifications aren't all received in {0} retry times. Try to configure RetryCount property in configure file.",
                Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            #endregion

            #region Verify multiple notification subscription objects to be created

            bool isSubscriptionCreated = isObjectCreated || isObjectDeleted || isObjectCopied;

            this.VerifyServerCreateSessionContaintAndSaveInformation(isSubscriptionCreated);

            if (Common.IsRequirementEnabled(292, this.Site))
            {
                // The implementation does allow multiple Notification Subscription objects to be created and associated with the same session context. Here uses ObjectCopied, ObjectCreated and ObjectDeleted to test.
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R292");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R292
                bool isR292Satisfied = isObjectCopied && isObjectCreated && isObjectDeleted;

                Site.Assert.IsTrue(isObjectCopied, "Received the notification is for ObjectCopied event.");
                Site.Assert.IsTrue(isObjectCreated, "Received the notification is for ObjectCreated event.");
                Site.Assert.IsTrue(isObjectDeleted, "Received the notification is for ObjectDeleted event.");

                Site.CaptureRequirementIfIsTrue(
                    isR292Satisfied,
                    292,
                    @"[In Appendix A: Product Behavior] The implementation does allow multiple Notification Subscription objects to be created and associated with the same session context. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that the server does not create a subscription to table notifications for a table with a NoNotifications flag.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC16_VerifyNoNotificationFlagForTableEventType()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Open Inbox folder and get content table that were created with a NoNotifications flag
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
            uint contentTableHandle;

            // Disable table notification.
            RopGetContentsTableResponse contentTableRsp = this.GetContentsTable(inboxTableHandle, out contentTableHandle, true);
            #endregion

            #region Create tableView by QueryRows
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[] 
            { 
                PropertyTags.All[PropertyNames.PidTagInstID], 
                PropertyTags.All[PropertyNames.PidTagImportance],
                PropertyTags.All[PropertyNames.PidTagInstanceNum],
                PropertyTags.All[PropertyNames.PidTagFolderId],
                PropertyTags.All[PropertyNames.PidTagMessageClass]
            };

            // Create tableView by QueryRows.
            this.SetColumns(contentTableHandle, tags);
            this.QueryRows(contentTableHandle, (ushort)contentTableRsp.RowCount);
            #endregion

            #region Trigger TableRowDeleted event and get notification
            this.TriggerTableRowDeletedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(false);
            Site.Assert.AreEqual<int>(0, rsp.Count, "The response should contain notification message");
            #endregion

            #region Verify notification response for the tables that were created with a NoNotifications flag
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R258");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R258
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x0000,
                (int)rsp.Count,
                258,
                @"[In Creating and Sending TableModified Event Notifications] The server MUST NOT create a subscription to table notifications for the tables that were created with a NoNotifications flag.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the TableChanged event and the elements related to the event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC17_VerifyTableChangedEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(128, this.Site), "This case runs only under Exchange 2007, since Exchange 2010 and Exchange 2013 implementation does not support TableChanged events.");

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxHandle;
            this.OpenFolder(this.InboxFolderId, out inboxHandle);
            uint contentTableHandle;
            this.GetContentsTable(inboxHandle, out contentTableHandle, false);
            #endregion

            #region Create table view by QueryColumnsAll
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[] { PropertyTags.All[PropertyNames.PidTagFolderId] };
            this.SetColumns(contentTableHandle, tags);

            // Retrieves rows from content table to get the data of row 1.
            this.QueryRows(contentTableHandle, 1);
            this.QueryColumnsAll(contentTableHandle);
            #endregion

            #region Trigger TableChanged event and get notification
            this.TriggerTableChangedEvent();
            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            #endregion

            #region Verify notification response for TableChanged event
            if (Common.IsRequirementEnabled(128, this.Site))
            {
                foreach (IDeserializable response in rsp)
                {
                    Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                    if (response is RopNotifyResponse)
                    {
                        RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                        Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                        Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableChanged, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableChanged.");
                        this.VerifyTableModifyNotificationFlag(notifyResponse);
                        this.VerifyTableChangedNotificationElements(notifyResponse);

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R264");

                        Site.Assert.IsNull(notifyResponse.NotificationData.TableRowFolderID, "The TableRowFolderID of the basic notification should be null.");
                        Site.Assert.IsNull(notifyResponse.NotificationData.TableRowMessageID, "The TableRowMessageID of the basic notification should be null.");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R264
                        this.Site.CaptureRequirement(
                            264,
                            @"[In Creating and Sending TableModified Event Notifications] [When a TableModified event occurs, the server generates a notification using one of the following three methods, listed in descending order of usefulness to the client.] For TableChanged event, the server generates a basic notification that does not include specifics about the change made.");

                        this.VerifyTriggerBasicNotification(true);

                        Site.Assert.IsNotNull(notifyResponse.NotificationData.TableEventType, "The TableEventType in the RopNotifyResponse should not null.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R128");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R128
                        this.Site.CaptureRequirementIfAreEqual<int>(
                            0x0001,
                            (int)notifyResponse.NotificationData.TableEventType,
                            128,
                            @"[In NotificationData Structure] [TableEventType value] 0x0001: The notification is for TableChanged events.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R22");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R22
                        // Client can receive TableChanged notification indicates that a table has been changed on the server.
                        this.Site.CaptureRequirementIfAreEqual<int>(
                            0x0001,
                            (int)notifyResponse.NotificationData.TableEventType,
                            22,
                            @"[In TableModified Event Types] TableChanged: A table has been changed.");
                    }
                }
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify ObjectModified event with the U bit and T bit are set.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC18_VerifyUbitAndTbit()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe ObjectModified event on server
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectModified);
            #endregion

            #region Trigger NewMail event and get notification
            // Trigger NewMailEvent to get mail modified information to get U bit and T bit notification flag.
            this.TriggerNewMailEvent();
            DateTime beginTime = DateTime.Now;
            DateTime endTime;
            int sleepTime = int.Parse(Common.GetConfigurationPropertyValue("SleepTime", this.Site));

            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() != "mapi_http")
            {
                // Do a loop to keep the connection and wait for the new mail being created on server.
                do
                {
                    this.CNOTIFAdapter.EcDoAsyncConnectEx();
                    System.Threading.Thread.Sleep(sleepTime);
                    endTime = DateTime.Now;
                }
                while ((endTime - beginTime) < TimeSpan.FromMilliseconds(sleepTime * 50));
            }

            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify notification response elements
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.ObjectModified, "The notification type should be ObjectModified.");
                    #region Verify elements TotalMessageCount and UnreadMessageCount of the notification response
                    if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.T) == (ushort)FlagsBit.T)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R20701");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R20701
                        this.Site.CaptureRequirementIfIsNotNull(
                            notifyResponse.NotificationData.TotalMessageCount,
                            20701,
                            @"[In NotificationData Structure] This field [TotalMessageCount] is available if bit 0x1000 is set in the NotificationFlags field.");
                    }

                    if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.U) == (ushort)FlagsBit.U)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R21101");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R21101
                        this.Site.CaptureRequirementIfIsNotNull(
                            notifyResponse.NotificationData.UnreadMessageCount,
                            21101,
                            @"[In NotificationData Structure] This field [UnreadMessageCount] is available if bit 0x2000 is set in the NotificationFlags field.");
                    }

                    if (notifyResponse.NotificationData.TotalMessageCount != null)
                    {
                        uint inboxTableHandle;
                        this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
                        uint contentTableHandle;
                        RopGetContentsTableResponse getContentsTableResponse = this.GetContentsTable(inboxTableHandle, out contentTableHandle, false);

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R206");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R206
                        this.Site.CaptureRequirementIfAreEqual<uint?>(
                            getContentsTableResponse.RowCount,
                            notifyResponse.NotificationData.TotalMessageCount,
                            206,
                            @"[In NotificationData Structure] TotalMessageCount: An unsigned 32-bit integer that specifies the total number of items in the folder triggering this event.");
                    }

                    if (notifyResponse.NotificationData.UnreadMessageCount != null)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R210");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R210
                        // In TriggerNewMailEvent() method, 1 mail is created and unread. So if the value of UnreadMessageCount is 1, this requirement can be verified.
                        this.Site.CaptureRequirementIfAreEqual<uint?>(
                            1,
                            notifyResponse.NotificationData.UnreadMessageCount,
                            210,
                            @"[In NotificationData Structure] UnreadMessageCount: An unsigned 32-bit integer that specifies the number of unread items in a folder triggering this event.");
                    }
                    #endregion
                }
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the TableRestrictionChanged event, which is a type of TableModified event.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC19_VerifyTableRestrictionChangedEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(13201, this.Site), "This case runs only under Exchange 2007, since Exchange 2010 and Exchange 2013 implementation does not support TableRestrictionChanged events.");

            #region Open Inbox folder and get content table of the Inbox folder
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
            uint contentTableHandle;
            this.GetContentsTable(inboxTableHandle, out contentTableHandle, false);
            #endregion

            #region Create table view by QueryPosition
            // The properties need to be set
            PropertyTag[] tags = new PropertyTag[] 
                { 
                    PropertyTags.All[PropertyNames.PidTagInstID], 
                    PropertyTags.All[PropertyNames.PidTagImportance],
                    PropertyTags.All[PropertyNames.PidTagInstanceNum],
                    PropertyTags.All[PropertyNames.PidTagFolderId],
                    PropertyTags.All[PropertyNames.PidTagMessageClass]
                };

            // Create table view by QueryPosition.
            this.SetColumns(contentTableHandle, tags);
            this.QueryPosition(contentTableHandle);
            #endregion

            #region Trigger TableRestrictionChanged event and get notification
            this.RestrictTable(contentTableHandle, new byte[] { });

            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            #endregion

            #region Verify notification response for TableRestrictionChanged event
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>(NotificationType.TableModified, notifyResponse.NotificationData.NotificationType, "The notification type for the RopNotify response should be TableModified.");
                    Site.Assert.AreEqual<EventTypeOfTable>(EventTypeOfTable.TableRestrictionChanged, notifyResponse.NotificationData.TableEvent, "The table event type for the RopNotify response should be TableRestrictionChanged.");
                    this.VerifyTableModifyNotificationFlag(notifyResponse);

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R265");

                    this.Site.Assert.IsNull(
                        notifyResponse.NotificationData.TableRowFolderID,
                        "The server generates a basic notification should not have TableRowFolderID.");

                    this.Site.Assert.IsNull(
                        notifyResponse.NotificationData.TableRowMessageID,
                        "The server generates a basic notification should not have TableRowMessageID.");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R265
                    // The server generates a basic notification and MS-OXCNOTIF_R265 can be verified directly.
                    this.Site.CaptureRequirement(
                        265,
                        @"[In Creating and Sending TableModified Event Notifications] [When a TableModified event occurs, the server generates a notification using one of the following three methods, listed in descending order of usefulness to the client.] For TableRestrictionChanged event, the server generates a basic notification that does not include specifics about the change made.");

                    this.Site.Assert.IsNotNull(notifyResponse.NotificationData.TableEventType, "TableEventType in the RopNotify response should not null.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R13201");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R13201
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        0x0007,
                        (int)notifyResponse.NotificationData.TableEventType,
                        13201,
                        @"[In NotificationData Structure] [TableEventType value] Implementation does support TableRestrictionChanged events (0x0007). (Exchange 2007 follows this behavior.)");
                }
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the ObjectMoved event, which is a type of TableModified event. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S02_TC20_VerifyObjectMovedMessageEvent()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe ObjectMoved event on server
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectMoved);
            #endregion

            #region Trigger ObjectMessageMove event and get notification
            this.TriggerObjectMessageMoveEvent();

            IList<IDeserializable> rsp = this.CNOTIFAdapter.GetNotification(true);
            Site.Assert.IsTrue(rsp.Count > 0, "The response should contain notification message.");
            #endregion

            #region Verify element OldMessageId of the notification response
            foreach (IDeserializable response in rsp)
            {
                Site.Assert.IsTrue(response.GetType() == typeof(RopNotifyResponse) || response.GetType() == typeof(RopPendingResponse), "The ROP response type should be RopNotifyResponse or RopPendingResponse.");
                if (response is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response;
                    Site.Assert.AreEqual<NotificationType>((NotificationType)(notifyResponse.NotificationData.NotificationFlags & 0x0FFF), NotificationType.ObjectMoved, "The notification type should be ObjectMoved.");
                    if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R18701");

                        // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R18701
                        this.Site.CaptureRequirementIfIsNotNull(
                            notifyResponse.NotificationData.OldMessageId,
                            18701,
                            @"[In NotificationData Structure] This field [OldMessageId] is available if the value of the NotificationType in the NotificationFlags field is 0x0020 and bit 0x8000 is set in the NotificationFlags field.");
                    }
                }
            }
            #endregion
        }
        #endregion

        #region Private Methods

        /// <summary>
        /// The method is used to verify when RegisterNotification the server will create a  new Notification Subscription object and associate it with the session context, and save the information.
        /// </summary>
        /// <param name="isSubscriptionCreated">Whether the server return the right response</param>
        private void VerifyServerCreateSessionContaintAndSaveInformation(bool isSubscriptionCreated)
        {
            if (Common.IsRequirementEnabled(288, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R288: The RopPending ROP response is {0} returned", isSubscriptionCreated ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R288 
                this.Site.CaptureRequirementIfIsTrue(
                    isSubscriptionCreated,
                    288,
                    @"[In Appendix A: Product Behavior] When a RopRegisterNotification ROP (section 2.2.1.2.1) message is received by the server, the implementation does create a new Notification Subscription object and associate it with the session context. (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(290, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R290: The implementation does {0} save the information provided in the RopRegisterNotification ROP request", isSubscriptionCreated ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R290
                this.Site.CaptureRequirementIfIsTrue(
                    isSubscriptionCreated,
                    290,
                    @"[In Appendix A: Product Behavior] [When a RopRegisterNotification ROP message is received by the server.] The implementation does save the information provided in the RopRegisterNotification ROP request fields for future use. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to verify the after the table view is created the notification response can be get on server 2010.
        /// </summary>
        private void VerifyTableViewCreated()
        {
            if (Common.IsRequirementEnabled(245, this.Site))
            {
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R245");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R245
                // Since the server return the notification, it verified the table view is created.
                // so this requirement will be verified directly.
                Site.CaptureRequirement(
                    245,
                    @"[In Appendix A: Product Behavior] If the client has subscribed to TableModified event notifications, by using the RopRegisterNotification ROP (section 2.2.1.2.1), the implementation does require that a table view is created in order to send the TableModified event notifications, as specified in section 2.2.1.1.1. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(257, this.Site))
            {
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNTOIF_R257");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R257
                // Since the server return the notification, so this requirement will be verified directly.
                Site.CaptureRequirement(
                    257,
                    @"[In Appendix A: Product Behavior] The implementation does then create a subscription to TableModified event notifications automatically for every table created on the server. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to verify the TableModify of NotificationType. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyTableModifyNotificationFlag(RopNotifyResponse notifyResponse)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R110");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R110
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x0100,
                notifyResponse.NotificationData.NotificationFlags & 0x0fff,
                110,
                @"[In NotificationData Structure] [NotificationType value] 0x0100: The notification is for a TableModified event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R18");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R18
            // Client can receive TableModified notification indicates that a table has been modified on the server.
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x0100,
                notifyResponse.NotificationData.NotificationFlags & 0x0fff,
                18,
                @"[In Server Event Types] TableModified: A table has been modified on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R40001");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R40001
            // If R110 and R18 been verified, R40001 will be verified.
            this.Site.CaptureRequirement(
                40001,
                @"[In RopRegisterNotification ROP Request Buffer] [NotificationTypes value] 0x0100: The server sends notifications to the client when TableModified events occur within the scope of interest.");
        }

        /// <summary>
        /// This method is used to verify the basic notification response.
        /// </summary>
        /// <param name="isBasicNotification">Whether the response is Basic notification</param>
        private void VerifyTriggerBasicNotification(bool isBasicNotification)
        {
            if (Common.IsRequirementEnabled(271, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R271: The implementation does {0} generate a basic notification", isBasicNotification ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R271
                // The basic notification that does not include specifics about the change made. So the returned value of the Folder ID and Message ID is null.
                this.Site.CaptureRequirementIfIsTrue(
                    isBasicNotification,
                    271,
                    @"[In Appendix A: Product Behavior] Implementation does only generate a basic notification when it is not feasible to generate an informative notification. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to verify the Server generate informative notification whenever possible.
        /// </summary>
        /// <param name="isInformativeNotification">Whether the notification is informative</param>
        private void VeriyServerGenerateInformativeNotification(bool isInformativeNotification)
        {
            if (Common.IsRequirementEnabled(269, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R269: The implementation does {0} generate a informative notification", isInformativeNotification ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R269
                // The informative notification that specifies the nature of the change,the value of the Folder ID structure, the value of the Message ID structure, and new table values.
                this.Site.CaptureRequirementIfIsTrue(
                    isInformativeNotification,
                    269,
                    @"[In Appendix A: Product Behavior] However, the implementation does generate informative notifications whenever possible. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for TableChanged event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyTableChangedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0001 of TableEventType means this notification is for TableChanged event, 
            // this method is called after notification for TableChanged to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14204");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14204
            // TableEventType with value 0x0001 means the notification is for TableChanged event, so if the TableRowFolderID is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.TableRowFolderID,
                14204,
                @"[In NotificationData Structure] This field [TableRowFolderID] is not available if the TableEventType field is 0x0001.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15403");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15403
            // TableEventType with value 0x0001 means the notification is for TableChanged event, so if the InsertAfterTableRowFolderID is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.InsertAfterTableRowFolderID,
                15403,
                @"[In NotificationData Structure] This field [InsertAfterTableRowFolderID] is not available if the TableEventType field is 0x0001.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15803");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15803
            // TableEventType with value 0x0001 means the notification is for TableChanged event, so if the InsertAfterTableRowID is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.InsertAfterTableRowID,
                15803,
                @"[In NotificationData Structure] This field [InsertAfterTableRowID] is not available if the TableEventType field is 0x0001.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R16403");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R16403
            // TableEventType with value 0x0001 means the notification is for TableChanged event, so if the TableRowDataSize is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.TableRowDataSize,
                16403,
                @"[In NotificationData Structure] This field [TableRowDataSize] is not available if the TableEventType field is 0x0001.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R16703");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R16703
            // TableEventType with value 0x0001 means the notification is for TableChanged event, so if the TableRowData is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.TableRowData,
                16703,
                @"[In NotificationData Structure] This field [TableRowData] is not available  if the TableEventType field is 0x0001.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14604");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14604
            // TableEventType with value 0x0001 means the notification is for TableChanged event, so if the TableRowMessageID is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.TableRowMessageID,
                14604,
                @"[In NotificationData Structure] This field [TableRowMessageID] is not available if the TableEventType field is 0x0001.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15004");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15004
            // TableEventType with value 0x0001 means the notification is for TableChanged event, so if the TableRowInstance is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.TableRowInstance,
                15004,
                @"[In NotificationData Structure] This field [TableRowInstance] is not available if the TableEventType field is 0x0001.");
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for TableRowAdded event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyTableRowAddedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0003 of TableEventType means this notification is for TableRowAdded event, 
            // this method is called after notification for TableRowAdded to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R12401");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R12401
            // The value 0x0100 of NotificationType means that notification is for a TableModified event, cause TableRowAdded is one of the TableModified events.
            // So if the TableEventType in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableEventType,
                12401,
                @"[In NotificationData Structure] This field [TableEventType] is available if the NotificationType value in the NotificationFlags field is 0x0100.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17101");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17101
            // The value 0x0100 of NotificationType means that notification is for a TableModified event, cause TableRowAdded is one of the TableModified events.
            // So if the FolderId in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.FolderId,
                17101,
                @"[In NotificationData Structure] This field [FolderId] is not available if the NotificationType value in the NotificationFlags field is 0x0100.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17502");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17502
            // The value 0x0100 of NotificationType means that notification is for a TableModified event, cause TableRowAdded is one of the TableModified events.
            // So if the MessageId in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.MessageId,
                17502,
                @"[In NotificationData Structure] This field [MessageId] is not available if the NotificationType value in the NotificationFlags field is 0x0100.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14201");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14201
            // TableEventType with value 0x0003 means the notification is for TableRowAdded event, so if the TableRowFolderID is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableRowFolderID,
                14201,
                @"[In NotificationData Structure] This field [TableRowFolderID] is available if the TableEventType field is available and is 0x0003.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15401");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15401
            // TableEventType with value 0x0003 means the notification is for TableRowAdded event, so if the InsertAfterTableRowFolderID is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.InsertAfterTableRowFolderID,
                15401,
                @"[In NotificationData Structure] This field [InsertAfterTableRowFolderID] is available if the TableEventType field is available and is 0x0003.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R16401");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R16401
            // TableEventType with value 0x0003 means the notification is for TableRowAdded event, so if the TableRowDataSize is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableRowDataSize,
                16401,
                @"[In NotificationData Structure] This field [TableRowDataSize] is available if the TableEventType field is available and is 0x0003.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R16701");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R16701
            // TableEventType with value 0x0003 means the notification is for TableRowAdded event, so if the TableRowData is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableRowData,
                16701,
                @"[In NotificationData Structure] This field [TableRowData] is available if the TableEventType field is available and is 0x0003.");

            // FlagsBit.M is 0x8000. If FlagsBit.M & NotificationFlags is itself, indicate that M bit is set. 
            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14601");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14601
                // TableEventType with value 0x0003 means the notification is for TableRowAdded event, so if the TableRowMessageID is not null, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.TableRowMessageID,
                    14601,
                    @"[In NotificationData Structure] This field [TableRowMessageID] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0003.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15001");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15001
                // TableEventType with value 0x0003 means the notification is for TableRowAdded event, so if the TableRowInstance is not null, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.TableRowInstance,
                    15001,
                    @"[In NotificationData Structure] This field [TableRowInstance] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0003.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15801");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15801
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.InsertAfterTableRowID,
                    15801,
                    @"[In NotificationData Structure] This field [InsertAfterTableRowID] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0003.");
            }
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for TableRowDeleted event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyTableRowDeletedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0004 of TableEventType means this notification is for TableRowDeleted event, 
            // this method is called after notification for TableRowDeleted to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14202");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14202
            // TableEventType with value 0x0004 means the notification is for TableRowDeleted event, so if the TableRowFolderID is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableRowFolderID,
                14202,
                @"[In NotificationData Structure] This field [TableRowFolderID] is available if the TableEventType field is available and is 0x0004.");

            // FlagsBit.M is 0x8000. If FlagsBit.M & NotificationFlags is itself, indicate that M bit is set. 
            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14602");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14602
                // TableEventType with value 0x0004 means the notification is for TableRowDeleted event, so if the TableRowMessageID is not null, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.TableRowMessageID,
                    14602,
                    @"[In NotificationData Structure] This field [TableRowMessageID] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0004.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15002");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15002
                // TableEventType with value 0x0004 means the notification is for TableRowDeleted event, so if the TableRowInstance is not null, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.TableRowInstance,
                    15002,
                    @"[In NotificationData Structure] This field [TableRowInstance] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0004.");
            }
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for TableRowModified event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyTableRowModifiedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0005 of TableEventType means this notification is for TableRowModified event, 
            // this method is called after notification for TableRowModified to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14203");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14203
            // TableEventType with value 0x0005 means the notification is for TableRowModified event, so if the TableRowFolderID is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableRowFolderID,
                14203,
                @"[In NotificationData Structure] This field [TableRowFolderID] is available if the TableEventType field is available and is 0x0005.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15402");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15402
            // TableEventType with value 0x0005 means the notification is for TableRowModified event, so if the InsertAfterTableRowFolderID is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.InsertAfterTableRowFolderID,
                15402,
                @"[In NotificationData Structure] This field [InsertAfterTableRowFolderID] is available, if the TableEventType field is available and is 0x0005.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R16402");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R16402
            // TableEventType with value 0x0005 means the notification is for TableRowModified event, so if the TableRowDataSize is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableRowDataSize,
                16402,
                @"[In NotificationData Structure] This field [TableRowDataSize] is available only if the TableEventType field is available and is 0x0005.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R16702");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R16702
            // TableEventType with value 0x0005 means the notification is for TableRowModified event, so if the TableRowData is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TableRowData,
                16702,
                @"[In NotificationData Structure] This field [TableRowData] is available if the TableEventType field is available and is 0x0005.");

            // FlagsBit.M is 0x8000. If FlagsBit.M & NotificationFlags is itself, indicate that M bit is set. 
            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R14603");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R14603
                // TableEventType with value 0x0005 means the notification is for TableRowModified event, so if the TableRowMessageID is not null, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.TableRowMessageID,
                    14603,
                    @"[In NotificationData Structure] This field [TableRowMessageID] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0005.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15003");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15003
                // TableEventType with value 0x0005 means the notification is for TableRowModified event, so if the TableRowInstance is not null, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.TableRowInstance,
                    15003,
                    @"[In NotificationData Structure] This field [TableRowInstance] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0005.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R15802");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R15802
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.InsertAfterTableRowID,
                    15802,
                    @"[In NotificationData Structure] This field [InsertAfterTableRowID] is available if bit 0x8000 is set in the NotificationFlags field and if the TableEventType field is available and is 0x0005.");
            }
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for NewMail event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyNewMailNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0002 of NotificationType means this notification is for NewMail event, 
            // this method is called after notification for NewMail to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R12402");

            bool isR12402Satisfied = false;

            if (notifyResponse.NotificationData.TableEventType == null)
            {
                isR12402Satisfied = true;
            }
            else
            {
                if (notifyResponse.NotificationData.TableEventType == (ushort)EventTypeOfTable.NONE)
                {
                    isR12402Satisfied = true;
                }
            }

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R12402
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the TableEventType in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsTrue(
                isR12402Satisfied,
                12402,
                @"[In NotificationData Structure] This field [TableEventType] is not available if the NotificationType value in the NotificationFlags field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17103");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17103
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the FolderId in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.FolderId,
                17103,
                @"[In NotificationData Structure] This field [FolderId] is available if the NotificationType value in the NotificationFlags field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17905");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17905
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the ParentFolderId in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.ParentFolderId,
                17905,
                @"[In NotificationData Structure] This field [ParentFolderId] is not available if the value of the NotificationType field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R18303");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R18303
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the OldFolderId in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.OldFolderId,
                18303,
                @"[In NotificationData Structure] This field [OldFolderId] is not available if the NotificationType value in the NotificationFlags field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R18703");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R18703
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the OldMessageId in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.OldMessageId,
                18703,
                @"[In NotificationData Structure] This field [OldMessageId] is not available if the value of the NotificationType in the NotificationFlags field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R19103");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R19103
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the OldParentFolderId in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.OldParentFolderId,
                19103,
                @"[In NotificationData Structure] This field [OldParentFolderId] is not available if the value of the NotificationType in the NotificationFlags field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R19503");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R19503
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the TagCount in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.TagCount,
                19503,
                @"[In NotificationData Structure] This field [TagCount] is not available if the value of the NotificationType in the NotificationFlags field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R21501");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R21501
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the MessageFlags in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.MessageFlags,
                21501,
                @"[In NotificationData Structure] This field [MessageFlags] is available if the value of the NotificationType in the NotificationFlags field is 0x0002. For details, see [MS-OXCMSG] section 2.2.1.6.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R22101");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R22101
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the UnicodeFlag in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.UnicodeFlag,
                22101,
                @"[In NotificationData Structure] This field [UnicodeFlag] is available if the value of the NotificationType field in the NotificationFlags field is 0x0002.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R22601");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R22601
            // The value 0x0002 of NotificationType means that notification is for a NewMail event.
            // So if the MessageClass in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.MessageClass,
                22601,
                @"[In NotificationData Structure] This field [MessageClass] is available if the value of the NotificationType in the NotificationFlags field is 0x0002.");

            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17501");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17501
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.MessageId,
                    17501,
                    @"[In NotificationData Structure] This field [MessageId] is available if the NotificationType value in the NotificationFlags field is 0x0002, and bit 0x8000 is set in the NotificationFlags field.");
            }
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for ObjectCreated event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyObjectCreatedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0004 of NotificationType means this notification is for ObjectCreated event, 
            // this method is called after notification for ObjectCreated to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R19501");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R19501
            // The value 0x0004 of NotificationType means that notification is for a ObjectCreated event.
            // So if the TagCount in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TagCount,
                19501,
                @"[In NotificationData Structure] This field [TagCount] is available if the value of the NotificationType in the NotificationFlags field is 0x0004.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R21502");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R21502
            // The value 0x0004 of NotificationType means that notification is for a ObjectCreated event.
            // So if the MessageFlags in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.MessageFlags,
                21502,
                @"[In NotificationData Structure] This field [MessageFlags] is not available if the value of the NotificationType in the NotificationFlags field is 0x0004. For details, see [MS-OXCMSG] section 2.2.1.6.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R22102");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R22102
            // The value 0x0004 of NotificationType means that notification is for a ObjectCreated event.
            // So if the UnicodeFlag in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.UnicodeFlag,
                22102,
                @"[In NotificationData Structure] This field [UnicodeFlag] is not available if the value of the NotificationType field in the NotificationFlags field is 0x0004.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R22602");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R22602
            // The value 0x0004 of NotificationType means that notification is for a ObjectCreated event.
            // So if the MessageClass in the response is null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNull(
                notifyResponse.NotificationData.MessageClass,
                22602,
                @"[In NotificationData Structure] This field [MessageClass] is not available if the value of the NotificationType in the NotificationFlags field is 0x0004.");

            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) != (ushort)FlagsBit.M && (notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.S) != (ushort)FlagsBit.S)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R17901");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R17901
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.ParentFolderId,
                    17901,
                    @"[In NotificationData Structure] This field [ParentFolderId] is available if the value of the NotificationType field is 0x0004, and it [RopNotify ROP] is sent for a message in a folder (both bit 0x4000 and bit 0x8000 are not set in the NotificationFlags field).");
            }
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for ObjectModified event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyObjectModifiedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0010 of NotificationType means this notification is for ObjectModified event, 
            // this method is called after notification for ObjectModified to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R19502");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R19502
            // The value 0x0010 of NotificationType means that notification is for a ObjectModified event.
            // So if the TagCount in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.TagCount,
                19502,
                @"[In NotificationData Structure] This field [TagCount] is available if the value of the NotificationType in the NotificationFlags field is 0x0010.");
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for ObjectMoved event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyObjectMovedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0020 of NotificationType means this notification is for ObjectMoved event, 
            // this method is called after notification for ObjectMoved to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R18301");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R18301
            // The value 0x0020 of NotificationType means that notification is for a ObjectMoved event.
            // So if the OldFolderId in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.OldFolderId,
                18301,
                @"[In NotificationData Structure] This field [OldFolderId] is available if the NotificationType value in the NotificationFlags field is 0x0020.");

            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) != (ushort)FlagsBit.M)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R19101");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R19101
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.OldParentFolderId,
                    19101,
                    @"[In NotificationData Structure] This field [OldParentFolderId] is available if the value of the NotificationType in the NotificationFlags field is 0x0020 and bit 0x8000 is not set in the NotificationFlags field.");
            }
        }

        /// <summary>
        /// This method is used to verify RopNotify response elements for ObjectCopied event. 
        /// </summary>
        /// <param name="notifyResponse">The notification response</param>
        private void VerifyObjectCopiedNotificationElements(RopNotifyResponse notifyResponse)
        {
            // The value 0x0040 of NotificationType means this notification is for ObjectCopied event, 
            // this method is called after notification for ObjectCopied to verify response elements.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R18302");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R18302
            // The value 0x0040 of NotificationType means that notification is for a ObjectCopied event.
            // So if the OldFolderId in the response is not null, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                notifyResponse.NotificationData.OldFolderId,
                18302,
                @"[In NotificationData Structure] This field [OldFolderId] is available if the NotificationType value in the NotificationFlags field is 0x0040.");

            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R18702");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R18702
                this.Site.CaptureRequirementIfIsNotNull(
                    notifyResponse.NotificationData.OldMessageId,
                    18702,
                    @"[In NotificationData Structure] This field [OldMessageId] is available if the value of the NotificationType in the NotificationFlags field is 0x0040 and bit 0x8000 is set in the NotificationFlags field.");
            }

            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.T) != (ushort)FlagsBit.T)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R20702");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R20702
                this.Site.CaptureRequirementIfIsNull(
                    notifyResponse.NotificationData.TotalMessageCount,
                    20702,
                    @"[In NotificationData Structure] This field [TotalMessageCount] is not available if bit 0x1000 is not set in the NotificationFlags field.");
            }

            if ((notifyResponse.NotificationData.NotificationFlags & (ushort)FlagsBit.U) != (ushort)FlagsBit.U)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R21102");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R21102
                this.Site.CaptureRequirementIfIsNull(
                    notifyResponse.NotificationData.UnreadMessageCount,
                    21102,
                    @"[In NotificationData Structure] This field [UnreadMessageCount] is not available if bit 0x2000 is not set in the NotificationFlags field.");
            }
        }
        #endregion
    }
}