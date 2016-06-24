namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test Rops related to Attachment 
    /// </summary>
    [TestClass]
    public class S08_RopAttachment : TestSuiteBase
    {
        #region Const definitions for test
        /// <summary>
        /// Constant string for the test data of PidTagAttachDataBinary.
        /// </summary>
        private const string TestDataOfPidTagAttachDataBinary = "add the content for file";

        /// <summary>
        /// Constant string for the test data of PidTagAttachLongPathname.
        /// </summary>
        private const string TestDataOfPidTagAttachLongPathname = "pidTagAttachLongPathname";

        /// <summary>
        /// Constant string for the test data of PidTagAttachRendering.
        /// </summary>
        private const string TestDataOfPidTagAttachRendering = "PidTagAttachRendering";

        /// <summary>
        /// Constant string for the test data of PidTagAttachTransportName.
        /// </summary>
        private const string TestDataOfPidTagAttachTransportName = "transportName";

        /// <summary>
        /// Constant string for the test data of PidTagAttachMimeTag.
        /// </summary>
        private const string TestDataOfPidTagAttachMimeTag = "image/jpeg";

        /// <summary>
        /// Constant string for the test data of PidTagAttachContentId.
        /// </summary>
        private const string TestDataOfPidTagAttachContentId = "image001.jpg@01CC1FB3.2053ED80";

        /// <summary>
        /// Constant string for the test data of PidTagAttachContentLocation.
        /// </summary>
        private const string TestDataOfPidTagAttachContentLocation = "/asp.net/images/people/";

        /// <summary>
        /// Constant string for the test data of PidTagAttachContentBase.
        /// </summary>
        private const string TestDataOfPidTagAttachContentBase = "http://i3.asp.net/asp.net/images/people/";

        /// <summary>
        /// Constant value for the test data of PidTagAttachMethod.
        /// </summary>
        private const string TestDataOfPidTagAttachMethod = "0x00000001";

        /// <summary>
        /// Constant string for the test data of PidTagAttachLongFileName.
        /// </summary>
        private const string TestDataOfPidTagAttachLongFileName = "image001.jpg";

        /// <summary>
        /// Constant string for the test data of PidTagAttachExtension.
        /// </summary>
        private const string TestDataOfPidTagAttachExtension = ".jpg";

        /// <summary>
        /// Constant string for the test data of PidTagAttachPayloadClass. 
        /// </summary>
        private const string TestDataOfPidTagAttachPayloadClass = "XPayloadClass";

        /// <summary>
        /// Constant string for the test data of PidTagAttachPayloadProviderGuidString. 
        /// </summary>
        private const string TestDataOfPidTagAttachPayloadProviderGuidString = "6F9619FF-8B86-D011-B42D-00C04FC964FF";

        /// <summary>
        /// Constant string for the test data of PidTagAttachTag.
        /// </summary>
        private byte[] testDataOfPidTagAttachTag = new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0A, 0x01 };

        /// <summary>
        /// Constant string for the test data of PidTagAttachEncoding.
        /// </summary>
        private byte[] testDataOfPidTagAttachEncoding = new byte[] { 0X2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0B, 0x01 };
        #endregion

        /// <summary>
        ///  Initializes the test class before running the test cases in the class.
        /// </summary>
        /// <param name="testContext">Test context which used to store information that is provided to unit tests.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        /// <summary>
        /// This case is used to test the successful response of RopCreateAttachment, RopOpenAttachment and RopSaveChangesAttachment.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC01_RopCreateOpenSaveAttachmentSuccessfully()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an attachment and expect a successful response.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSaveChangesAttachment to save the newly created attachment.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopOpenAttachment with InputHandleIndex set to 0x01 which doesn't refer to a message object and expect a failure response
            RopOpenAttachmentRequest openAttachmentRequest = new RopOpenAttachmentRequest()
            {
                RopId = (byte)RopId.RopOpenAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = 0x01, // Set InputHandleIndex to 0x01 which doesn't refer to a Message object.
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                OpenAttachmentFlags = 0x01,
                AttachmentID = attachmentId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openAttachmentRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenAttachmentResponse openAttachmentResponseFirst = (RopOpenAttachmentResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, openAttachmentResponseFirst.ReturnValue, "The InputHandleIndex is wrong.");
            #endregion

            #region Call RopOpenAttachment to open the attachment as read/write and expect a successful response.
            RopOpenAttachmentResponse openAttachmentResponseSecond;
            uint openedAttachmentHandle = this.OpenAttachment(openedMessageHandle, out openAttachmentResponseSecond, attachmentId, OpenAttachmentFlags.ReadWrite);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openAttachmentResponseSecond.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R446. The return value of RopOpenAttachmentRequest before saving the attachment is {0}, the return value of RopOpenAttachmentRequest after saving the attachment is {1}.", openAttachmentResponseFirst.ReturnValue, openAttachmentResponseSecond.ReturnValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R446
            bool isVerifiedR446 = openAttachmentResponseFirst.ReturnValue != TestSuiteBase.Success && openAttachmentResponseSecond.ReturnValue == TestSuiteBase.Success;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR446,
                446,
                @"[In Receiving a RopCreateAttachment ROP Request] When processing the RopCreateAttachment ROP ([MS-OXCROPS] section 2.2.6.13), the server does not commit the new Attachment object until it receives a call to the RopSaveChangesAttachment ROP ([MS-OXCROPS] section 2.2.6.15).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R853");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R853
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                openAttachmentResponseSecond.ReturnValue,
                853,
                @"[In RopOpenAttachment ROP] The RopOpenAttachment ROP ([MS-OXCROPS] section 2.2.6.12) opens an Attachment object stored on the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R861");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R861
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                openAttachmentResponseSecond.ReturnValue,
                861,
                @"[In RopCreateAttachment ROP] The RopCreateAttachment ROP ([MS-OXCROPS] section 2.2.6.13) creates a new Attachment object on the Message object.");
            #endregion

            #region Call RopSetProperties to set the displayname of the attachment and expect a successful response.
            List<PropertyObj> lstPts = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagDisplayName, Common.GetBytesFromUnicodeString("Test"))
            };
            RopSetPropertiesResponse setPropertiesResponse;
            this.SetPropertiesForMessage(openedAttachmentHandle, lstPts, out setPropertiesResponse);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R918");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R918
            // The calling RopSetPropertiesRequest succeeds indicates that the opened attachment can be modified, so MS-OXCMSG_R918 can be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                setPropertiesResponse.ReturnValue,
                918,
                @"[In RopOpenAttachment ROP Request Buffer] [OpenAttachmentFlags] [ReadWrite (0x01)] Attachment will be opened for both reading and writing.");
            #endregion

            #region Call RopSaveChangesAttachment and set SaveFlags to 0x0C and expect a successful response.
            this.SaveAttachment(openedAttachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopOpenAttachment to open the attachment as BestAccess and expect a successful response.
            RopOpenAttachmentResponse openAttachmentResponseForBestAccess;
            openedAttachmentHandle = this.OpenAttachment(openedMessageHandle, out openAttachmentResponseForBestAccess, attachmentId, OpenAttachmentFlags.BestAccess);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openAttachmentResponseSecond.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSetProperties to set the displayname of the attachment and expect a successful response.
           lstPts = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagDisplayName, Common.GetBytesFromUnicodeString("TestForBestAccess"))
            };
         
            this.SetPropertiesForMessage(openedAttachmentHandle, lstPts, out setPropertiesResponse);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R919");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R919
            // Because the user is onwer of the mailbox, so the user has write permissions for the attachment.
            // The calling RopSetPropertiesRequest succeeds indicates that the opened attachment can be modified, so MS-OXCMSG_R919 can be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                setPropertiesResponse.ReturnValue,
                919,
                @"[In RopOpenAttachment ROP Request Buffer] [OpenAttachmentFlags] [BestAccess (0x03)] Attachment will be opened for read/write if the user has write permissions for the attachment; opened for read-only if not.");
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(openedMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test the successful response of RopGetValidAttachments and RopGetAttachmentTable.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC02_RopGetAttachmentSuccessfully()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create the attachment1 of the message
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId1;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId1);
            #endregion

            #region Call SetPropertiesSpecific to set PidTagDisplayName property.
            List<PropertyObj> propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagDisplayName, Common.GetBytesFromUnicodeString("First attachment"))
            };

            this.SetPropertiesForMessage(attachmentHandle, propertyList);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment1.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopRelease to release attachment1.
            this.ReleaseRop(attachmentHandle);
            #endregion

            #region Call RopCreateAttachment to create the attachment2 of the message
            uint attachmentId2;
            attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId2);
            #endregion

            #region Call SetPropertiesSpecific to set PidTagDisplayName property.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagDisplayName, Common.GetBytesFromUnicodeString("Second attachment"))
            };

            this.SetPropertiesForMessage(attachmentHandle, propertyList);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment2.
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopRelease to release attachment2.
            this.ReleaseRop(attachmentHandle);

            bool isVerifiedR1836 = !(attachmentId1 == attachmentId2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1836");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1836
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1836,
                1836,
                @"[In PidTagAttachNumber Property] The value of property PidTagAttachNumber is different among the two Attachment objects in a message.");
            #endregion

            #region Call RopGetValidAttachments and expect a successful response
            if (Common.IsRequirementEnabled(1715, this.Site))
            {
                RopGetValidAttachmentsRequest getValidAttachmentsRequest = new RopGetValidAttachmentsRequest()
                {
                    RopId = (byte)RopId.RopGetValidAttachments,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex
                };
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getValidAttachmentsRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                RopGetValidAttachmentsResponse getValidAttachmentsResponse = (RopGetValidAttachmentsResponse)this.response;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1715");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1715
                // Calling RopGetValidAttachments succeeds indicates that implementation supports the RopGetValidAttachments ROP.
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    getValidAttachmentsResponse.ReturnValue,
                    1715,
                    @"[In Appendix A: Product Behavior] Implementation does support the RopGetValidAttachments ROP. (Exchange 2007 follows this behavior.)");
            }
            #endregion

            #region Call RopGetAttachmentTable with TableFlags set to 0x00 and expect a successful response.
            RopGetAttachmentTableRequest getAttachmentTableRequest = new RopGetAttachmentTableRequest()
            {
                RopId = (byte)RopId.RopGetAttachmentTable,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                TableFlags = 0x00 // Open the table 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getAttachmentTableRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetAttachmentTableResponse getAttachmentTableResponse = (RopGetAttachmentTableResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getAttachmentTableResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            uint attachmentTableHandle = this.ResponseSOHs[0][getAttachmentTableResponse.OutputHandleIndex];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R890");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R890
            // Calling RopGetAttachmentTable succeeds and the handle of the attachment table is not null indicate that RopGetAttachmentTable ROP retrieves a handle to a Table object.
            this.Site.CaptureRequirementIfIsNotNull(
                attachmentTableHandle,
                890,
                @"[In RopGetAttachmentTable ROP] The RopGetAttachmentTable ROP ([MS-OXCROPS] section 2.2.6.17) retrieves a handle to a Table object that represents the attachments stored on the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R905");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R905
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                getAttachmentTableResponse.ReturnValue,
                905,
                @"[In RopGetAttachmentTable ROP Request Buffer] [TableFlags] [Standard (0x00)] Open the table.");
            #endregion

            #region Call RopGetAttachmentTable with TableFlags set to 0x40 and expect a successful response
            getAttachmentTableRequest.TableFlags = 0x40; // Open the table 
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getAttachmentTableRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getAttachmentTableResponse = (RopGetAttachmentTableResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getAttachmentTableResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            attachmentTableHandle = this.ResponseSOHs[0][getAttachmentTableResponse.OutputHandleIndex];
            #endregion

            #region Call RopSetColumns to set the display name of the attachment table and expect a successful response
            PropertyTag[] propertyTags = new PropertyTag[2];
            propertyTags[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagDisplayName];
            propertyTags[1] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachNumber];

            this.SetColumnsSuccess(propertyTags, attachmentTableHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R906");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R906
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                getAttachmentTableResponse.ReturnValue,
                906,
                @"[In RopGetAttachmentTable ROP Request Buffer] [TableFlags] [Unicode (0x40)] Open the table.");
            #endregion

            #region Call RopQueryColumnsAll to get all the string type properties in the attachment table and expect all the string type properties are returned in Unicode format
            RopQueryColumnsAllRequest queryColumnsAllRequest = new RopQueryColumnsAllRequest()
            {
                RopId = (byte)RopId.RopQueryColumnsAll,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(queryColumnsAllRequest, attachmentTableHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopQueryColumnsAllResponse queryColumnsAllResponse = (RopQueryColumnsAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, queryColumnsAllResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopQueryRows to query all rows.
            RopQueryRowsResponse queryRowsResponse = this.QueryRowsSuccess(attachmentTableHandle);

            bool isVerifiedR907 = false;

            for (int i = 0; i < queryRowsResponse.RowCount; i++)
            {
                byte[] rowData = queryRowsResponse.RowData.PropertyRows[i].PropertyValues[0].Value;
                if (Common.IsUtf16LEString(rowData))
                {
                    isVerifiedR907 = true;
                }
                else
                {
                    isVerifiedR907 = false;
                    break;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R907");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R907
            // Property type ptypString indicates that the string data is in Unicode format, and then MS-OXCMSG_R907 can be verified.     
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR907,
                907,
                @"[In RopGetAttachmentTable ROP Request Buffer] [TableFlags] [Unicode (0x40)] Also requests that the columns containing string data be returned in Unicode format.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R312");

            bool isExitsFirstAttach = false;
            bool isExitsSecondAttach = false;

            Site.Assert.AreEqual<int>(2, queryRowsResponse.RowCount, "The message should include 2 attachments.");

            for (int i = 0; i < queryRowsResponse.RowCount; i++)
            {
                int attachmentID = BitConverter.ToInt32(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[1].Value, 0);

                if (attachmentID == attachmentId1)
                {
                    isExitsFirstAttach = true;
                }

                if (attachmentID == attachmentId2)
                {
                    isExitsSecondAttach = true;
                }
            }

            bool isVerifiedR312 = isExitsFirstAttach && isExitsSecondAttach;

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R312
            // Because according to above step, the message include 2 attachments.
            // When a client calls the RopGetAttachmentTable ROP, the server returns a table of properties for each Attachment object associated with the Message object.
            // If the table from server only includes 2 rows and the PidTagAttachNumber property columns include all attachment ids of attachment associated with Message object.
            // R312 will be verified.
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR312,
                312,
                @"[In Sending a RopGetAttachmentTable ROP Request] When a client calls the RopGetAttachmentTable ROP ([MS-OXCROPS] section 2.2.6.17), the server returns a table of properties for each Attachment object associated with the Message object, as specified in [MS-OXCTABL].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R490");
            bool isVerifiedR490 = isVerifiedR312;

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R490
            // Because the PidTagAttachNumber can be obtained from the table return from server when call RopGetAttachmentTable ROP.
            // So R490 will be verified.
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR490,
                490,
                @"[In Receiving a RopGetAttachmentTable ROP Request] The Table object returned by the RopGetAttachmentTable ROP ([MS-OXCROPS] section 2.2.6.17) allows access to the properties of Attachment objects.");
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(openedMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test Rop properties in RopCreateAttachement initialization.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC03_PropertiesInCreateAttachmentInitialization()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create the attachment of the message
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopGetPropertiesSpecific to get specific properties for the created attachment.
            PropertyTag[] tagArray = this.CreateAttachmentPropertyTagsForInitial();
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Verify requirements related with RopCreateAttachment initialization
            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagAccessLevel = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAccessLevel);
            PropertyObj pidTagCreationTime = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagCreationTime);
            PropertyObj pidTagLastModificationTime = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);
            PropertyObj pidTagAttachSize = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachSize);
            PropertyObj pidTagAttachNumber = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachNumber);
            PropertyObj pidTagRenderingPosition = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRenderingPosition);

            DateTime tagCreationTime = Convert.ToDateTime(pidTagCreationTime.Value);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2035");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2035
            // tagCreationTime is not null indicates that PidTagCreationTime is initialized when calling RopCreateAttachment ROP.
            this.Site.CaptureRequirementIfIsNotNull(
                tagCreationTime,
                2035,
                @"[In Receiving a RopCreateAttachment ROP Request] PidTagCreationTime (section 2.2.2.3) [will be initialized when calling RopCreateAttachment ROP].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R448");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R448
            this.Site.CaptureRequirementIfAreEqual<uint>(
                attachmentId,
                Convert.ToUInt32(pidTagAttachNumber.Value),
                448,
                @"[In Receiving a RopCreateAttachment ROP Request] [The initial data of PidTagAttachNumber is] Varies, depending on the number of existing attachments on the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R451");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R451
            this.Site.CaptureRequirementIfAreEqual<int>(
                unchecked((int)0xffffffff),
                Convert.ToInt32(pidTagRenderingPosition.Value),
                451,
                @"[In Receiving a RopCreateAttachment ROP Request] [The initial data of PidTagRenderingPosition is] 0xFFFFFFFF.");

            if (Common.IsRequirementEnabled(1919, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1919");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1919
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000040,
                    Convert.ToInt32(pidTagAttachSize.Value),
                    1919,
                    @"[In Appendix A: Product Behavior] Implementation does set the PidTagAttachSize property (section 2.2.2.5) to 0x00000040. (Exchange 2007 and 2010 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1649, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1649");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1649
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000000,
                    Convert.ToInt32(pidTagAttachSize.Value),
                    1649,
                    @"[In Appendix A: Product Behavior] Implementation does set the PidTagAttachSize property (section 2.2.2.5) to 0x00000000. (Exchange 2013 and above follow this behavior.)");
            }

            int pidTagAccessLevelInitialValue = Convert.ToInt32(pidTagAccessLevel.Value);
            if (Common.IsRequirementEnabled(1920, this.Site))
            {
                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCMSG_R1920,the actual initial Data of PidTagAccessLevel is {0}",
                    pidTagAccessLevelInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1920
                Site.CaptureRequirementIfAreEqual<int>(
                    0x00000001,
                    pidTagAccessLevelInitialValue,
                    1920,
                    @"[In Appendix A: Product Behavior] Implementation does set the PidTagAccessLevel property ([MS-OXCPRPT] section 2.2.1.2) to 0x00000001. (Exchange 2007 and 2010 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1921, this.Site))
            {
                bool isVerifyR1921 = Convert.ToDateTime(pidTagLastModificationTime.Value) == tagCreationTime;

                // Above If condition has verified the initial data of PidTagLastModificationTime.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1921,
                    1921,
                    @"[In Appendix A: Product Behavior] Implementation does set PidTagLastModificationTime the same as the PidTagCreationTime property. (Exchange 2007 and 2010 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1651, this.Site))
            {
                DateTime lastModificationTime = Convert.ToDateTime(pidTagLastModificationTime.Value);
                TimeSpan dateOffsetModificationTime = lastModificationTime - tagCreationTime;

                bool isVerifyR1651 = dateOffsetModificationTime.Milliseconds < 0.1 && dateOffsetModificationTime.Milliseconds > -0.1;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1651, the time period between PidTagLastModificationTime and PidTagCreationTime  is {0}.", dateOffsetModificationTime.Milliseconds);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1651
                this.Site.CaptureRequirementIfIsTrue(
                    isVerifyR1651,
                    1651,
                    @"[In Appendix A: Product Behavior] Implementation does set the PidTagLastModificationTime property (section 2.2.2.2) to a value that is within 100 nanoseconds of the value of the PidTagCreationTime property (section 2.2.2.3). (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1650, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1650,the actual initial Data of PidTagAccessLevel is {0}", pidTagAccessLevelInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1650
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000000,
                    pidTagAccessLevelInitialValue,
                    1650,
                    @"[In Appendix A: Product Behavior] Implementation does set the PidTagAccessLevel property ([MS-OXCPRPT] section 2.2.1.2) to 0x00000000. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            #region Call RopRelease to release the created message and the created attachment.
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This method is used to test Attachment Rop properties.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC04_AttachmentObjectProperties()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of the newly created attachment.
            RopGetPropertiesAllRequest getPropertiesAllRequest = new RopGetPropertiesAllRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesAll,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,

                // Set PropertySizeLimit, which specifies the maximum size allowed for a property value returned, as specified in [MS-OXCROPS].
                PropertySizeLimit = 0xFFFF,
                WantUnicode = (ushort)0
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.AttachmentProperties);
            #endregion

            #region Set specific properties which are not initialized during creating the attachment
            List<PropertyObj> specificProperties = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagAttachDataBinary, Common.AddInt16LengthBeforeBinaryArray(Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachDataBinary))),
                new PropertyObj(PropertyNames.PidTagAttachLongPathname, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachLongPathname)),
                new PropertyObj(PropertyNames.PidTagAttachTag, Common.AddInt16LengthBeforeBinaryArray(this.testDataOfPidTagAttachTag)),
                new PropertyObj(PropertyNames.PidTagAttachRendering, Common.AddInt16LengthBeforeBinaryArray(Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachRendering))),
                new PropertyObj(PropertyNames.PidTagAttachFlags, BitConverter.GetBytes(0x00000004)),
                new PropertyObj(PropertyNames.PidTagAttachTransportName, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachTransportName)),
                new PropertyObj(PropertyNames.PidTagAttachEncoding, Common.AddInt16LengthBeforeBinaryArray(this.testDataOfPidTagAttachEncoding)),
                new PropertyObj(PropertyNames.PidTagAttachmentLinkId, BitConverter.GetBytes(0x00000000)),
                new PropertyObj(PropertyNames.PidTagAttachmentFlags, BitConverter.GetBytes(0x00000000)),
                new PropertyObj(PropertyNames.PidTagAttachMethod, BitConverter.GetBytes(0x00000001)),
                new PropertyObj(PropertyNames.PidTagAttachmentHidden, BitConverter.GetBytes(true)),
                new PropertyObj(PropertyNames.PidTagAttachMimeTag, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachMimeTag)),
                new PropertyObj(PropertyNames.PidTagAttachContentId, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachContentId)),
                new PropertyObj(PropertyNames.PidTagAttachContentLocation, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachContentLocation)),
                new PropertyObj(PropertyNames.PidTagAttachContentBase, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachContentBase))
            };

            this.SetPropertiesForMessage(attachmentHandle, specificProperties);
            #endregion

            #region Send a RopGetPropertiesSpecific request to get all properties of specific message object.
            PropertyTag[] tagArray = this.CreateAttachmentPropertyTagsForCapture();
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertySizeLimit = 0xFFFF,
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.AttachmentProperties);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj pidTagAttachDataBinary = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachDataBinary);
            PropertyObj pidTagAttachMethod = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachMethod);

            byte[] testDataOfPidTagAttachDataBinary = Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachDataBinary);
            byte[] valueOfPidTagAttachDataBinary = new byte[testDataOfPidTagAttachDataBinary.Length];
            Buffer.BlockCopy((byte[])pidTagAttachDataBinary.Value, 2, valueOfPidTagAttachDataBinary, 0, testDataOfPidTagAttachDataBinary.Length);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R236");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R236
            bool isVerfiedR236 = Common.CompareByteArray(testDataOfPidTagAttachDataBinary, valueOfPidTagAttachDataBinary);
            this.Site.CaptureRequirementIfIsTrue(
                isVerfiedR236,
                236,
                @"[In PidTagAttachDataBinary Property] The PidTagAttachDataBinary property ([MS-OXPROPS] section 2.580) contains the contents of the file to be attached.");

            if (Convert.ToInt32(pidTagAttachMethod.Value) == 0x00000001)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R590");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R590
                this.Site.CaptureRequirementIfIsNotNull(
                    pidTagAttachDataBinary.Value,
                    590,
                    @"[In PidTagAttachMethod Property] [afByValue (0x00000001)] The PidTagAttachDataBinary property (section 2.2.2.7) contains the attachment data.");
            }

            if(Common.IsRequirementEnabled(3008,this.Site))
            {
                PropertyObj pidTagObjectType = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagObjectType);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3008");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3008
                this.Site.CaptureRequirementIfIsNotNull(
                    pidTagObjectType.Value,
                    3008,
                    @"[In Appendix A: Product Behavior] Implementation does support the PidTagObjectType property. (Exchange 2007 follows this behavior.)");
            }
            #endregion

            #region Call RopRelease to release the created message and the created attachment
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test the successful response of RopDeleteAttachment.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC05_RopDeleteAttachmentSuccessfully()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagHasAttachments and pidTagMessageFlags properties of the opened message.
            PropertyTag[] tagArray = this.MessageProperties();
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertySizeLimit = 0xFFFF,
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            List<PropertyObj> ptsFirst = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagHasAttachmentsFirst = PropertyHelper.GetPropertyByName(ptsFirst, PropertyNames.PidTagHasAttachments);
            PropertyObj pidTagMessageFlagsFirst = PropertyHelper.GetPropertyByName(ptsFirst, PropertyNames.PidTagMessageFlags);
            bool isHasAttachmentFirst = (Convert.ToInt32(pidTagMessageFlagsFirst.Value) & 0x00000010) == 0x00000010;
            Site.Assert.IsFalse(isHasAttachmentFirst, "The message should not contains any attachment.");
            #endregion

            #region Call RopCreateAttachment to create an attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSaveChangesAttachment to save the newly created attachment.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenAttachment to open the attachment as read/write.
            RopOpenAttachmentResponse openAttachmentResponse;
            this.OpenAttachment(openedMessageHandle, out openAttachmentResponse, attachmentId, OpenAttachmentFlags.ReadWrite);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagHasAttachments and pidTagMessageFlags properties of the message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            List<PropertyObj> ptsSecond = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagHasAttachmentsSecond = PropertyHelper.GetPropertyByName(ptsSecond, PropertyNames.PidTagHasAttachments);
            PropertyObj pidTagMessageFlagsSecond = PropertyHelper.GetPropertyByName(ptsSecond, PropertyNames.PidTagMessageFlags);
            bool isHasAttachmentSecond = (Convert.ToInt32(pidTagMessageFlagsSecond.Value) & 0x00000010) == 0x00000010;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R16");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R16
            this.Site.CaptureRequirementIfAreEqual<bool>(
                isHasAttachmentSecond,
                Convert.ToBoolean(pidTagHasAttachmentsSecond.Value),
                16,
                @"[In PidTagHasAttachments Property] The server computes this property [PidTagHasAttachments] from the mfHasAttach flag of the PidTagMessageFlags property ([MS-OXPROPS] section 2.780).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R498. The value of PidTagHasAttachments before creating the attachment is {0}, the value of PidTagHasAttachments after creating the attachment is {1}.", pidTagHasAttachmentsFirst.Value, pidTagHasAttachmentsSecond.Value);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R498
            // MS-OXCMSG_R498 can be verified if the value of PidTagHasAttachments before creating the attachment is false and the value of PidTagHasAttachments after creating the attachment is true.
            bool isVerifiedR498 = !Convert.ToBoolean(pidTagHasAttachmentsFirst.Value) && Convert.ToBoolean(pidTagHasAttachmentsSecond.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR498,
                498,
                @"[In PidTagHasAttachments Property] The PidTagHasAttachments property ([MS-OXPROPS] section 2.706) indicates whether the Message object contains at least one attachment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R875. The value of PidTagHasAttachments before creating the attachment is {0}, the value of PidTagHasAttachments after creating the attachment is {1}.", pidTagHasAttachmentsFirst.Value, pidTagHasAttachmentsSecond.Value);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R875
            // The value of PidTagHasAttachments before creating the attachment is false and the value of PidTagHasAttachments after creating the attachment is true indicate that RopSaveChangesAttachment commits the changes made to the attachment object.
            bool isVerifiedR875 = !Convert.ToBoolean(pidTagHasAttachmentsFirst.Value) && Convert.ToBoolean(pidTagHasAttachmentsSecond.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR875,
                875,
                @"[In RopSaveChangesAttachment ROP] The RopSaveChangesAttachment ROP ([MS-OXCROPS] section 2.2.6.15) commits the changes made to the Attachment object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R522. The value of isHasAttachmentSecond is {0}.", isHasAttachmentSecond);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R522
            // One attachment is created in the previous step, so MS-OXCMSG_R522 can be verified if the value of isHasAttachmentSecond is true.
            this.Site.CaptureRequirementIfIsTrue(
                isHasAttachmentSecond,
                522,
                @"[In PidTagMessageFlags Property] [mfHasAttach (0x00000010)] The message has at least one attachment.");
            #endregion

            #region Call RopDeleteAttachment to delete the attachment and expect a successful response.
            RopDeleteAttachmentRequest deleteAttachmentRequest = new RopDeleteAttachmentRequest()
            {
                RopId = (byte)RopId.RopDeleteAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the input Server Object is stored.
                AttachmentID = attachmentId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteAttachmentRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopDeleteAttachmentResponse deleteAttachmentResponse = (RopDeleteAttachmentResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenAttachment to open the deleted attachment and expect a failure response.
            this.OpenAttachment(openedMessageHandle, out openAttachmentResponse, attachmentId, OpenAttachmentFlags.ReadWrite);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R868");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R868
            // The deleted attachment can't be opened indicates that RopDeleteAttachment deletes an existing Attachment object from the Message object.
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.Success,
                openAttachmentResponse.ReturnValue,
                868,
                @"[In RopDeleteAttachment ROP] The RopDeleteAttachment ROP ([MS-OXCROPS] section 2.2.6.14) deletes an existing Attachment object from the Message object.");
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagHasAttachments property of the opened message and expect a successful response with its value is false
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            List<PropertyObj> ptsThird = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagHasAttachmentsThird = PropertyHelper.GetPropertyByName(ptsThird, PropertyNames.PidTagHasAttachments);
            PropertyObj pidTagMessageFlagsThird = PropertyHelper.GetPropertyByName(ptsThird, PropertyNames.PidTagMessageFlags);
            bool isHasAttachmentThird = (Convert.ToInt32(pidTagMessageFlagsThird.Value) & 0x00000010) == 0x00000010;
            Site.Assert.IsFalse(isHasAttachmentThird, "The message should not contains any attachment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R469. The value of PidTagHasAttachments before calling RopDeleteAttachment is {0}, the value of PidTagHasAttachments after calling RopDeleteAttachment is {1}", pidTagHasAttachmentsSecond.Value, pidTagHasAttachmentsThird.Value);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R469
            // The value of PidTagHasAttachments before calling RopDeleteAttachment is true and after calling RopDeleteAttachment is false indicate that server recalculates the PidTagHasAttachments property while processing the RopDeleteAttachment ROP.
            bool isVerifiedR469 = !Convert.ToBoolean(pidTagHasAttachmentsThird.Value) && Convert.ToBoolean(pidTagHasAttachmentsSecond.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR469,
                469,
                @"[In Receiving a RopDeleteAttachment ROP Request] The server recalculates the PidTagHasAttachments property (section 2.2.1.2) while processing the RopDeleteAttachment ROP ([MS-OXCROPS] section 2.2.6.14).");
            #endregion

            #region Call RopRelease to release the created message and attachment.
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test MIME properties.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC06_MIMEProperties()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Set MIME properties which are not initialized during creating the attachment
            List<PropertyObj> specificProperties = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagAttachLongFilename, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachLongFileName)),
                new PropertyObj(PropertyNames.PidTagAttachExtension, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachExtension)),
                new PropertyObj(PropertyNames.PidTagAttachMimeTag, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachMimeTag)),
                new PropertyObj(PropertyNames.PidTagAttachContentId, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachContentId)),
                new PropertyObj(PropertyNames.PidTagAttachContentLocation, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachContentLocation)),
                new PropertyObj(PropertyNames.PidTagAttachContentBase, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachContentBase)),
                new PropertyObj(PropertyNames.PidTagAttachPayloadClass, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachPayloadClass)),
                new PropertyObj(PropertyNames.PidTagAttachPayloadProviderGuidString, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachPayloadProviderGuidString))
            };
            
            this.SetPropertiesForMessage(attachmentHandle, specificProperties);
            #endregion

            #region Send a RopGetPropertiesSpecific request to get all properties of specific message object.
            PropertyTag[] tagArray = this.CreateMIMEAttachmentPropertyTagsForCapture();
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertySizeLimit = 0xFFFF,
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.AttachmentProperties);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj pidTagAttachExtension = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachExtension);
            PropertyObj pidTagAttachLongFilename = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachLongFilename);
            PropertyObj pidTagAttachContentBase = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachContentBase);
            PropertyObj pidTagAttachContentLocation = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachContentLocation);
            PropertyObj pidTagAttachContentId = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachContentId);
            PropertyObj pidTagAttachPayloadClass = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachPayloadClass);
            PropertyObj pidTagAttachPayloadProviderGuidString = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachPayloadProviderGuidString);
            PropertyObj pidTagAttachMimeTag = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAttachMimeTag);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R963");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R963
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachMimeTag,
                Convert.ToString(pidTagAttachMimeTag.Value),
                963,
                @"[In MIME Properties] PidTagAttachMimeTag: The Content-Type header (2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R964");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R964
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachContentId,
                Convert.ToString(pidTagAttachContentId.Value),
                964,
                @"[In MIME properties] PidTagAttachContentId: A content identifier unique to this Message object that matches a corresponding ""cid:"" URI scheme reference in the HTML body of the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1284");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1284
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachPayloadClass,
                Convert.ToString(pidTagAttachPayloadClass.Value),
                1284,
                @"[In MIME Properties] PidTagAttachPayloadClass: The class name of an object that can display the contents of the message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1286");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1286
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachPayloadProviderGuidString,
                Convert.ToString(pidTagAttachPayloadProviderGuidString.Value),
                1286,
                @"[In MIME Properties] PidTagAttachPayloadProviderGuidString: The GUID of the software application that can display the contents of the message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R965, the value of PidTagBodyContentLocation is {0}.", pidTagAttachContentLocation.Value.ToString());

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R965
            bool isVerifiedR965 = Uri.IsWellFormedUriString(pidTagAttachContentLocation.Value.ToString(), UriKind.RelativeOrAbsolute);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR965,
                965,
                @"[In MIME properties] PidTagAttachContentLocation: A relative or full URI that matches a corresponding reference in the HTML body of the Message object.");

            if (Uri.IsWellFormedUriString(pidTagAttachContentLocation.Value.ToString(), UriKind.Relative))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R967, the value of PidTagAttachContentBase is {0}.", pidTagAttachContentBase.Value.ToString());

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R967
                this.Site.CaptureRequirementIfIsNotNull(
                    pidTagAttachContentBase.Value,
                    967,
                    @"[In MIME properties] [PidTagAttachContentBase] MUST be set if the PidTagAttachContentLocation property contains a relative URI.");
            }

            int length = Convert.ToString(pidTagAttachLongFilename.Value).Length;
            string fileName = Convert.ToString(pidTagAttachLongFilename.Value);
            char[] longFileName = Convert.ToString(pidTagAttachLongFilename.Value).ToCharArray();

            int temp = length;
            while (temp > 0)
            {
                temp--;
                if (Convert.ToInt32(longFileName[temp]) == Convert.ToInt32('.'))
                {
                    break;
                }
            }

            string fileNameExtension = fileName.Substring(temp, length - temp);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R604");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R604
            this.Site.CaptureRequirementIfAreEqual<string>(
                fileNameExtension,
                Convert.ToString(pidTagAttachExtension.Value),
                604,
                @"[In PidTagAttachExtension Property] The PidTagAttachExtension property ([MS-OXPROPS] section 2.583) contains a file name extension that indicates the document type of an attachment.");
            #endregion

            #region Call RopRelease to release the created message and the created attachment
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test transaction between RopDeleteAttachment and RopSaveChangesMessage.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC07_Transaction()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSaveChangesAttachment to save the newly created attachment.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle, 0x04);
            #endregion

            #region Call RopOpenMessage to open the message and get the message handle 1.
            uint openedMessageHandle1 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message and get the message handle 2.
            uint openedMessageHandle2 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopDeleteAttachment to delete the attachment and expect a successful response.
            RopDeleteAttachmentRequest deleteAttachmentRequest = new RopDeleteAttachmentRequest()
            {
                RopId = (byte)RopId.RopDeleteAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                AttachmentID = attachmentId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteAttachmentRequest, openedMessageHandle1, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopDeleteAttachmentResponse deleteAttachmentResponse = (RopDeleteAttachmentResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Use the message handle 2 to call RopGetPropertiesSpecific to get the PidTagHasAttachments property.

            // PidTagHasAttachments
            List<PropertyTag> tagArray = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagHasAttachments]
            };

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedMessageHandle2, tagArray);

            List<PropertyObj> ptsFirst = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagHasAttachmentsFirst = PropertyHelper.GetPropertyByName(ptsFirst, PropertyNames.PidTagHasAttachments);
            #endregion

            #region Use the message handle 1 to call RopSaveChangesMessage.
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle1, 0x04);
            #endregion

            #region Call RopOpenMessage to open the message and get the message handle 3.
            uint openedMessageHandle3 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Use the message handle 3 to call RopGetPropertiesSpecific to get the PidTagHasAttachments property.
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedMessageHandle3, tagArray);

            List<PropertyObj> ptsSecond = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagHasAttachmentsSecond = PropertyHelper.GetPropertyByName(ptsSecond, PropertyNames.PidTagHasAttachments);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1523. The value of PidTagHasAttachments before committing the pending changes is {0}, The value of PidTagHasAttachments after committing the pending changes is {1}.", pidTagHasAttachmentsFirst.Value, pidTagHasAttachmentsSecond.Value);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1523
            // The value of PidTagHasAttachments before committing the pending changes is true and after committing is false indicates that the attachment is not permanently removed from the message until the client calls the RopSaveChangesMessage.
            bool isVerifiedR1523 = Convert.ToBoolean(pidTagHasAttachmentsFirst.Value) && !Convert.ToBoolean(pidTagHasAttachmentsSecond.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1523,
                1523,
                @"[In Receiving a RopDeleteAttachment ROP Request] The attachment is not permanently removed from the message until the client calls the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1523. The value of PidTagHasAttachments before committing the pending changes is {0}, The value of PidTagHasAttachments after committing the pending changes is {1}.", pidTagHasAttachmentsFirst.Value, pidTagHasAttachmentsSecond.Value);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1062
            // The value of PidTagHasAttachments before committing the pending changes is true and after committing is false indicates that the changes are not committed to the database until the RopSaveChangesMessage is called.
            bool isVerifiedR1062 = Convert.ToBoolean(pidTagHasAttachmentsFirst.Value) && !Convert.ToBoolean(pidTagHasAttachmentsSecond.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1062,
                1062,
                @"[In Receiving a RopSaveChangesAttachment ROP Request] Although the server commits any pending changes to the Attachment object in the context of its containing Message object, the changes MUST NOT be committed to the database until the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) has been executed on the handle of the Message object.");
            #endregion

            #region Call RopRelease to release the created message and attachment.
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This method is used to test Attachment Rop properties.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC08_FlagsOfAttachmentRelatedProperties()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachMethod of created Attachment
            List<PropertyTag> tagArray = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachMethod]
            };

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(attachmentHandle, tagArray);
            List<PropertyObj> pts = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagAttachMethod = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachMethod);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R588");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R588
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000000,
                Convert.ToInt32(pidTagAttachMethod.Value),
                588,
                @"[In PidTagAttachMethod Property] [afNone (0x00000000)] The attachment has just been created.");
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopOpenAttachment to open the attachment as read/write
            RopOpenAttachmentResponse openAttachmentResponse;
            uint openedAttachmentHandle = this.OpenAttachment(openedMessageHandle, out openAttachmentResponse, attachmentId, OpenAttachmentFlags.ReadWrite);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSetProperties to set specific properties which are not initialized in creating attachment and expect a successful response.
            List<PropertyObj> ptsSetProperties = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagAttachLongPathname, Common.GetBytesFromUnicodeString(TestDataOfPidTagAttachLongPathname))
            };

            this.SetPropertiesForMessage(openedAttachmentHandle, ptsSetProperties);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment.
            this.SaveAttachment(openedAttachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopOpenAttachment to open the attachment as read/write
            openedAttachmentHandle = this.OpenAttachment(openedMessageHandle, out openAttachmentResponse, attachmentId, OpenAttachmentFlags.ReadWrite);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSetProperties to set PidTagAttachMethod to afByReference.
            int size;
            TaggedPropertyValue[] taggedPropertyValueArray = this.CreateMessageTaggedPropertyValueArrays(out size, PidTagAttachMethodFlags.afByReference);
            this.SetFlagsOfPidTagAttachMethod(taggedPropertyValueArray, openedAttachmentHandle, size);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachLongPathname of created Attachment
            tagArray.Add(PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachLongPathname]);

            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArray);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagAttachLongPathname = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachLongPathname);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R436");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R436
            // The value of PidTagAttachLongPathname got by calling RopGetPropertiesSpecific is same with the value set by calling RopSetProperties indicates that the handle returned by the RopOpenAttachment can be used in RopGetPropertiesSpecific ROP.
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachLongPathname,
                Convert.ToString(pidTagAttachLongPathname.Value),
                436,
                @"[In Receiving a RopOpenAttachment ROP Request] The handle returned by the RopOpenAttachment ROP ([MS-OXCROPS] section 2.2.6.12) is used in subsequent ROPs, such as the RopGetPropertiesSpecific ROP ([MS-OXCROPS] section 2.2.8.3).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R592");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R592
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachLongPathname,
                Convert.ToString(pidTagAttachLongPathname.Value),
                592,
                @"[In PidTagAttachMethod Property] [afByReference (0x00000002)] The PidTagAttachLongPathname property (section 2.2.2.13) contains a fully qualified path identifying the attachment To recipients with access to a common file server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R302");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R302
            // The value of PidTagAttachLongPathname got by calling RopGetPropertiesSpecific is same with the value set by calling RopSetProperties indicates that a client adds the contents of a file to an Attachment object by sending a RopSetProperties ROP request
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachLongPathname,
                Convert.ToString(pidTagAttachLongPathname.Value),
                302,
                @"[In Setting Attachment Object Content] A client adds the contents of a file to an Attachment object by sending a RopSetProperties ROP request ([MS-OXCROPS] section 2.2.8.6) as specified in [MS-OXCPRPT] section 2.2.5.");
            #endregion

            #region Call RopSetProperties to set PidTagAttachMethod to afByReferenceOnly
            taggedPropertyValueArray = this.CreateMessageTaggedPropertyValueArrays(out size, PidTagAttachMethodFlags.afByReferenceOnly);
            this.SetFlagsOfPidTagAttachMethod(taggedPropertyValueArray, openedAttachmentHandle, size);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachMethod of created Attachment.
            tagArray.Add(PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachMethod]);

            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArray);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            pidTagAttachMethod = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachMethod);
            pidTagAttachLongPathname = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachLongPathname);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R594");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R594
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagAttachLongPathname,
                Convert.ToString(pidTagAttachLongPathname.Value),
                594,
                @"[In PidTagAttachMethod Property] [afByReferenceOnly (0x00000004)] The PidTagAttachLongPathname property contains a fully qualified path identifying the attachment.");
            #endregion

            #region Call RopSetProperties to set PidTagAttachFlags to attInvisibleInHtml (0x00000001)
            List<PropertyObj> pidTagAttachFlags = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagAttachFlags, BitConverter.GetBytes(0x00000001))
            };
            RopSetPropertiesResponse setPropertiesResponse;
            this.SetPropertiesForMessage(openedAttachmentHandle, pidTagAttachFlags, out setPropertiesResponse);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachFlags of created Attachment
            tagArray.Add(PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachFlags]);

            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArray);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj valueOfPidTagAttachFlags = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2031");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2031
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                Convert.ToInt32(valueOfPidTagAttachFlags.Value),
                2031,
                @"[In PidTagAttachFlags Property] [The value of attInvisibleInHtml is] 0x00000001.");
            #endregion

            #region Call RopSetProperties to set PidTagAttachFlags to attInvisibleInRtf (0x00000002)
            pidTagAttachFlags.Add(new PropertyObj(PropertyNames.PidTagAttachFlags, BitConverter.GetBytes(0x00000002)));
            this.SetPropertiesForMessage(openedAttachmentHandle, pidTagAttachFlags, out setPropertiesResponse);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachFlags of created Attachment
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArray);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            valueOfPidTagAttachFlags = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2032");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2032
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000002,
                Convert.ToInt32(valueOfPidTagAttachFlags.Value),
                2032,
                @"[In PidTagAttachFlags Property] [The value of attInvisibleInRtf is] 0x00000002.");
            #endregion

            #region Call RopSetProperties to set PidTagAttachFlags to attRenderedInBody (0x00000004)
            pidTagAttachFlags.Add(new PropertyObj(PropertyNames.PidTagAttachFlags, BitConverter.GetBytes(0x00000004)));
            this.SetPropertiesForMessage(openedAttachmentHandle, pidTagAttachFlags, out setPropertiesResponse);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachFlags of created Attachment
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArray);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            valueOfPidTagAttachFlags = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2033");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2033
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000004,
                Convert.ToInt32(valueOfPidTagAttachFlags.Value),
                2033,
                @"[In PidTagAttachFlags Property] [The value of attRenderedInBody is] 0x00000004.");
            #endregion

            #region Call RopSetProperties to set PidTagAttachTag to TNEF (0x2A,86,48,86,F7,14,03,0A,01)
            List<PropertyObj> pidTagAttachTagSetValue = new List<PropertyObj>();
            byte[] pidTagAttachTagText = new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0A, 0x01 };
            pidTagAttachTagSetValue.Add(new PropertyObj(PropertyNames.PidTagAttachTag, PropertyHelper.GetBinaryFromGeneral(pidTagAttachTagText)));
            this.SetPropertiesForMessage(openedAttachmentHandle, pidTagAttachTagSetValue, out setPropertiesResponse);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachTag of created Attachment
            List<PropertyTag> tagArrayOfPidTagAttachTag = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachTag]
            };

            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArrayOfPidTagAttachTag);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArrayOfPidTagAttachTag.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            PropertyObj pidTagAttachTag = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachTag);

            byte[] valueOfPidTagAttachTag = new byte[9];
            Buffer.BlockCopy((byte[])pidTagAttachTag.Value, 2, valueOfPidTagAttachTag, 0, 9);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R953");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R953
            bool isVerifiedR953 = Common.CompareByteArray(pidTagAttachTagText, valueOfPidTagAttachTag);
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR953,
                953,
                @"[In PidTagAttachTag Property] The data of TNEF is {0x2A,86,48,86,F7,14,03,0A,01}.");
            #endregion

            #region Call RopSetProperties to set PidTagAttachTag to afStorage (0x2A,86,48,86,F7,14,03,0A,03,02,01)
            pidTagAttachTagText = new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0A, 0x03, 0x02, 0x01 };
            pidTagAttachTagSetValue.Add(new PropertyObj(PropertyNames.PidTagAttachTag, PropertyHelper.GetBinaryFromGeneral(pidTagAttachTagText)));
            this.SetPropertiesForMessage(openedAttachmentHandle, pidTagAttachTagSetValue, out setPropertiesResponse);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachTag of created Attachment
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArrayOfPidTagAttachTag);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArrayOfPidTagAttachTag.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            pidTagAttachTag = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachTag);

            valueOfPidTagAttachTag = new byte[11];
            Buffer.BlockCopy((byte[])pidTagAttachTag.Value, 2, valueOfPidTagAttachTag, 0, 11);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R954");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R954
            bool isVerifiedR954 = Common.CompareByteArray(pidTagAttachTagText, valueOfPidTagAttachTag);
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR954,
                954,
                @"[In PidTagAttachTag Property] The data of afStorage is {0x2A,86,48,86,F7,14,03,0A,03,02,01}.");
            #endregion

            #region Call RopSetProperties to set PidTagAttachTag to MIME (0x2A,86,48,86,F7,14,03,0A,04)
            pidTagAttachTagText = new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x14, 0x03, 0x0A, 0x04 };
            pidTagAttachTagSetValue.Add(new PropertyObj(PropertyNames.PidTagAttachTag, PropertyHelper.GetBinaryFromGeneral(pidTagAttachTagText)));
            this.SetPropertiesForMessage(openedAttachmentHandle, pidTagAttachTagSetValue, out setPropertiesResponse);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachTag of created Attachment
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(openedAttachmentHandle, tagArrayOfPidTagAttachTag);
            pts = PropertyHelper.GetPropertyObjFromBuffer(tagArrayOfPidTagAttachTag.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test case requirement
            pidTagAttachTag = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachTag);

            valueOfPidTagAttachTag = new byte[9];
            Buffer.BlockCopy((byte[])pidTagAttachTag.Value, 2, valueOfPidTagAttachTag, 0, 9);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R955");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R955
            bool isVerifiedR955 = Common.CompareByteArray(pidTagAttachTagText, valueOfPidTagAttachTag);
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR955,
                955,
                @"[In PidTagAttachTag Property] The data of MIME is {0x2A,86,48,86,F7,14,03,0A,04}.");
            #endregion

            #region Call RopRelease to release the created message and the created attachment
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test the attachment read-only property PidTagAttachSize.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC09_ReadOnlyProperties()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopGetPropertiesSpecific to get the read-only property PidTagAttachSize.
            List<PropertyTag> tagArray = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachSize]
            };

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(attachmentHandle, tagArray);
            List<PropertyObj> propertySpecsFir = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response to get Property Value to verify test  case requirement
            PropertyObj pidTagAttachSizeFir = PropertyHelper.GetPropertyByName(propertySpecsFir, PropertyNames.PidTagAttachSize);
            Site.Assert.IsNotNull(pidTagAttachSizeFir.Value, "The PidTagAttachSize property should not null.");
            #endregion

            #region Call RopSetProperties to set the read-only property PidTagAttachSize.
            List<PropertyObj> pts = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagAttachSize, BitConverter.GetBytes(0x00000001))
            };

            this.SetPropertiesForMessage(attachmentHandle, pts);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment.
            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest = new RopSaveChangesAttachmentRequest()
            {
                RopId = (byte)RopId.RopSaveChangesAttachment,
                LogonId = CommonLogonId,
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response. 
                InputHandleIndex = CommonInputHandleIndex,
                SaveFlags = (byte)SaveFlags.ForceSave
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesAttachmentRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            saveChangesAttachmentResponse = (RopSaveChangesAttachmentResponse)this.response;
            
            if (Common.IsRequirementEnabled(1922, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1922");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1922
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesAttachmentResponse.ReturnValue,
                    1922,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only property PidTagAttachSize. (Exchange 2007 follows this behavior.)");
            }
            else
            {
                Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesAttachmentResponse.ReturnValue, "If pending changes include changes to read-only property PidTagAttachSize, call RopSaveChangesAttachment should failure.");
            }
            #endregion

            #region Call RopGetPropertiesSpecific to get the read-only property PidTagAttachSize again.
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(attachmentHandle, tagArray);
            PropertyObj pidTagAttachSizeAgain = PropertyHelper.GetPropertyByName(propertySpecsFir, PropertyNames.PidTagAttachSize);
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R231");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R231
            this.Site.CaptureRequirementIfAreEqual<int>(
                Convert.ToInt32(pidTagAttachSizeFir.Value),
                Convert.ToInt32(pidTagAttachSizeAgain.Value),
                231,
                @"[In PidTagAttachSize Property] This property [PidTagAttachSize] is read-only for the client.");
            #endregion

            #region Call RopRelease to release the created message and attachment.
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test error codes related to attachment ROPs.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S08_TC10_ErrorCodeRelatedWithAttachmentROPs()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment with InputHandleIndex set to 0x01 and expect a failure response.
            RopCreateAttachmentRequest createAttachmentRequest = new RopCreateAttachmentRequest()
            {
                RopId = (byte)RopId.RopCreateAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = 0x01, // Set InputHandleIndex to 0x01 which doesn't refer to a Message object.
                OutputHandleIndex = CommonOutputHandleIndex // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createAttachmentRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopCreateAttachmentResponse createAttachmentResponse = (RopCreateAttachmentResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1712");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1712
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                createAttachmentResponse.ReturnValue,
                1712,
                @"[In Receiving a RopCreateAttachment ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopCreateAttachment] was called does not refer to a Message object.");
            #endregion

            #region Call RopCreateAttachment to create an attachment and expect a successful response.
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSaveChangesAttachment and set the attachmentID to the message ID and expect a failure response.
            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest = new RopSaveChangesAttachmentRequest()
            {
                RopId = (byte)RopId.RopSaveChangesAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                SaveFlags = (byte)SaveFlags.ForceSave
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesAttachmentRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse = (RopSaveChangesAttachmentResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R467");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R467
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                saveChangesAttachmentResponse.ReturnValue,
                467,
                @"[In Receiving a RopSaveChangesAttachment ROP Request] [ecNotSupported (0x80040102)] The value of the InputHandleIndex field on which this ROP [RopSaveChangesAttachment] was called does not refer to an Attachment object.");
            #endregion

            #region Call RopSaveChangesAttachment and set SaveFlags to 0x03 which doesn't specified in the Open Specification and expect a failure response.
            saveChangesAttachmentRequest.SaveFlags = 0x03; // Set SaveFlags to 0x03 which doesn't specified in the Open Specification
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesAttachmentRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            saveChangesAttachmentResponse = (RopSaveChangesAttachmentResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R466");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R466
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                saveChangesAttachmentResponse.ReturnValue,
                466,
                @"[In Receiving a RopSaveChangesAttachment ROP Request] [ecNotSupported (0x80040102)] The value of the SaveFlags field is not a supported combination as specified in section 2.2.3.3.1.");
            #endregion

            #region Call RopSaveChangesAttachment to save the newly created attachment.
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopOpenAttachment with InputHandleIndex set to 0x01 which doesn't refer to a message object and expect a failure response
            RopOpenAttachmentRequest openAttachmentRequest = new RopOpenAttachmentRequest()
            {
                RopId = (byte)RopId.RopOpenAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = 0x01, // Set InputHandleIndex to 0x01 which doesn't refer to a Message object.
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                OpenAttachmentFlags = 0x01,
                AttachmentID = attachmentId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openAttachmentRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenAttachmentResponse openAttachmentResponseFirst = (RopOpenAttachmentResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, openAttachmentResponseFirst.ReturnValue, "The InputHandleIndex is wrong.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1620");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1620
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                openAttachmentResponseFirst.ReturnValue,
                1620,
                @"[In Receiving a RopOpenAttachment ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopOpenAttachment] was called does not refer to a Message object.");
            #endregion

            #region Call RopRelease to release attachment.
            this.ReleaseRop(attachmentHandle);
            #endregion

            #region Call RopGetAttachmentTable with InputHandleIndex set to 0x01 and expect a failure response
            RopGetAttachmentTableRequest getAttachmentTableRequest = new RopGetAttachmentTableRequest()
            {
                RopId = (byte)RopId.RopGetAttachmentTable,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = 0x01, // Set InputHandleIndex to 0x01 which doesn't refer to a Message object
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                TableFlags = 0x00 // Open the table 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getAttachmentTableRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetAttachmentTableResponse getAttachmentTableResponse = (RopGetAttachmentTableResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1526");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1526
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                getAttachmentTableResponse.ReturnValue,
                1526,
                @"[In Receiving a RopGetAttachmentTable ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopGetAttachmentTable] was called does not refer to a Message object.");
            #endregion

            #region Call RopDeleteAttachment and set InputHandleIndex to 0x01 and expect a failure response.
            RopDeleteAttachmentRequest deleteAttachmentRequest = new RopDeleteAttachmentRequest()
            {
                RopId = (byte)RopId.RopDeleteAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // Set InputHandleIndex to 0x00 which doesn't refer to a Message object.
                AttachmentID = attachmentId
            };

            // Set InputHandleIndex to 0x00 which doesn't refer to a Message object.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteAttachmentRequest, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopDeleteAttachmentResponse deleteAttachmentResponse = (RopDeleteAttachmentResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1524");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1524
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                deleteAttachmentResponse.ReturnValue,
                1524,
                @"[In Receiving a RopDeleteAttachment ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopDeleteAttachment] was called does not refer to a Message object.");
            #endregion

            #region Call RopDeleteAttachment and set AttachmentID to a nonexisting one and expect a failure response.
            deleteAttachmentRequest.InputHandleIndex = 0x00;
            deleteAttachmentRequest.AttachmentID = attachmentId + 5;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteAttachmentRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            deleteAttachmentResponse = (RopDeleteAttachmentResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1063");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1063
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                deleteAttachmentResponse.ReturnValue,
                1063,
                @"[In Receiving a RopDeleteAttachment ROP Request] [ecNotFound (0x8004010F)] The value of the AttachmentID field does not correspond to an attachment on the Message object.");
            #endregion

            #region Call RopDeleteAttachment to delete the attachment and expect a successful response.
            deleteAttachmentRequest.AttachmentID = attachmentId;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteAttachmentRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            deleteAttachmentResponse = (RopDeleteAttachmentResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenAttachment to open the deleted attachment and expect a failure response.
            RopOpenAttachmentResponse openAttachmentResponse;
            this.OpenAttachment(openedMessageHandle, out openAttachmentResponse, attachmentId, OpenAttachmentFlags.ReadWrite);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R440");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R440
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                openAttachmentResponse.ReturnValue,
                440,
                @"[In Receiving a RopOpenAttachment ROP Request] [ecNotFound (0x8004010F)] The value of the AttachmentID field does not correspond to an attachment on the Message object.");
            #endregion

            #region Call RopRelease to release the created message and attachment.
            this.ReleaseRop(attachmentHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        #region Private methods
        /// <summary>
        /// The method is used to set flag of PidTagAttachMethod.
        /// </summary>
        /// <param name="taggedPropertyValueArray">TaggedPropertyValue array </param>
        /// <param name="attachmentHandle">The attachment handle.</param>
        /// <param name="size">The size of TaggedPropertyValue array.</param>
        private void SetFlagsOfPidTagAttachMethod(TaggedPropertyValue[] taggedPropertyValueArray, uint attachmentHandle, int size)
        {
            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertyValueSize = (ushort)(size + 2),
                PropertyValueCount = (ushort)taggedPropertyValueArray.Length,
                PropertyValues = taggedPropertyValueArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setPropertiesRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetPropertiesResponse setPropertiesResponse = (RopSetPropertiesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setPropertiesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
        }

        /// <summary>
        /// Message properties PidTagHasAttachments and PidTagMessageFlags.
        /// </summary>
        /// <returns>Property tags array</returns>
        private PropertyTag[] MessageProperties()
        {
            PropertyTag[] tags = new PropertyTag[2];
            tags[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagHasAttachments];
            tags[1] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags];

            return tags;
        }

        /// <summary>
        /// Open the attachment
        /// </summary>
        /// <param name="objectHandle">A Server object handle.</param>
        /// <param name="openAttachmentResponse">The RopOpenAttachmentResponse value.</param>
        /// <param name="attachmentId">The ID of an attachment to be opened.</param>
        /// <param name="openFlags">The OpenModeFlags value.</param>
        /// <returns>A Server object handle of the opened attachment.</returns>
        private uint OpenAttachment(uint objectHandle, out RopOpenAttachmentResponse openAttachmentResponse, uint attachmentId, OpenAttachmentFlags openFlags)
        {
            RopOpenAttachmentRequest openAttachmentRequest = new RopOpenAttachmentRequest()
            {
                RopId = (byte)RopId.RopOpenAttachment,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                OpenAttachmentFlags = (byte)openFlags,
                AttachmentID = attachmentId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openAttachmentRequest, objectHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openAttachmentResponse = (RopOpenAttachmentResponse)this.response;

            return this.ResponseSOHs[0][openAttachmentResponse.OutputHandleIndex];
        }

        /// <summary>
        /// Create property tags for create attachment initialization
        /// </summary>
        /// <returns>PropertyTag array</returns>
        private PropertyTag[] CreateAttachmentPropertyTagsForInitial()
        {
            PropertyTag[] tags = new PropertyTag[6];
            tags[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachNumber];
            tags[1] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachSize];
            tags[2] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel];
            tags[3] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagRenderingPosition];
            tags[4] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagCreationTime];
            tags[5] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModificationTime];

            return tags;
        }

        /// <summary>
        /// Create property tags for attachment for capture code 
        /// </summary>
        /// <returns>PropertyTag array</returns>
        private PropertyTag[] CreateAttachmentPropertyTagsForCapture()
        {
            PropertyTag[] tags = new PropertyTag[29];
            tags[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModificationTime];
            tags[1] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagCreationTime];
            tags[2] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagDisplayName];
            tags[3] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachSize];
            tags[4] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachNumber];
            tags[5] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachDataBinary];
            tags[6] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachDataObject];
            tags[7] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachMethod];
            tags[8] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachLongFilename];
            tags[9] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachFilename];
            tags[10] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachExtension];
            tags[11] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachLongPathname];
            tags[12] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachPathname];
            tags[13] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachTag];
            tags[14] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagRenderingPosition];
            tags[15] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachRendering];
            tags[16] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachFlags];
            tags[17] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachTransportName];
            tags[18] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachEncoding];
            tags[19] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachAdditionalInformation];
            tags[20] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachmentLinkId];
            tags[21] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachmentFlags];
            tags[22] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachmentHidden];
            tags[23] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachMimeTag];
            tags[24] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachContentId];
            tags[25] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachContentLocation];
            tags[26] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachContentBase];
            tags[27] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagTextAttachmentCharset];
            tags[28] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagObjectType];
            return tags;
        }

        /// <summary>
        /// Create MIME property tags for attachment for capture code 
        /// </summary>
        /// <returns>PropertyTag array</returns>
        private PropertyTag[] CreateMIMEAttachmentPropertyTagsForCapture()
        {
            PropertyTag[] tags = new PropertyTag[8];
            tags[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachMimeTag];
            tags[1] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachContentId];
            tags[2] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachContentLocation];
            tags[3] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachContentBase];
            tags[4] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachPayloadClass];
            tags[5] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachLongFilename];
            tags[6] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachExtension];
            tags[7] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachPayloadProviderGuidString];

            return tags;
        }
        #endregion
    }
}