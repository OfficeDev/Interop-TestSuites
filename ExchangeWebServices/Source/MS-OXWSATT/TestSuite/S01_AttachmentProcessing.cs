//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSATT
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the following operations with parameters.
    /// <list type="bullet">
    ///     <item>CreateAttachment</item>
    ///     <item>GetAttachment</item>
    ///     <item>DeleteAttachment</item>
    /// </list>
    /// </summary>
    [TestClass]
    public class S01_AttachmentProcessing : TestSuiteBase
    {
        #region Class initialize and clean up

        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="testContext">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases
        /// <summary>
        /// This test case is designed to verify processing a file attachment.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC01_ProcessFileAttachment()
        {
            #region Step 1 Create a file attachment on an item.
     
            // Create a file attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.FileAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R532");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R532
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createAttachmentInfoResponse.ResponseClass,
                532,
                @"[In Message Processing Events and Sequencing Rules][The CreateAttachment operation] Creates an item and attaches it to the specified item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R53201");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R53201
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createAttachmentInfoResponse.ResponseClass,
                53201,
                @"[In Message Processing Events and Sequencing Rules][The CreateAttachment operation] Creates file attachment and attaches it to the specified item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R457");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R457
            // If the returned attachment's RootItemId is same as the id of root store item, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                this.ItemId.Id,
                createAttachmentInfoResponse.Attachments[0].AttachmentId.RootItemId.ToString(),
                457,
                @"[In m:CreateAttachmentType Complex Type][The ParentId element] Identifies the parent item in the server store that contains the attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R43");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R43
            // Root item id is set in request , so if the returned root item id equals to the value in request , this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                this.ItemId.Id,
                createAttachmentInfoResponse.Attachments[0].AttachmentId.RootItemId.ToString(),
                43,
                @"[In t:AttachmentIdType Complex Type][The RootItemId attribute] represents the unique identifier of the root store item to which the attachment is attached.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R45");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R45
            // If the returned attachment's RootItemChangeKey exists, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                createAttachmentInfoResponse.Attachments[0].AttachmentId.RootItemChangeKey,
                45,
                @"[In t:AttachmentIdType Complex Type][The RootItemChangeKey attribute] represents the change key of the root store item to which the attachment is attached.");

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentType createdAttachment = createAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentId = createdAttachment.AttachmentId;

            #region Step 2 Get the file attachment created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Text, false, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoResponse = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R38");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R38
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getAttachmentInfoResponse.Attachments[0],
                typeof(FileAttachmentType),
                38,
                @"[In t:ArrayOfAttachmentsType Complex Type][The FileAttachment element] specifies a file that is attached to another item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R59");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R59
            // If the returned attachment name is same as the name of attachment created in step 1, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "attach.jpg",
                getAttachmentInfoResponse.Attachments[0].Name,
                59,
                @"[In t:AttachmentType Complex Type][The Name element] specifies the descriptive name of the attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R60");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R60
            // If the returned attachment's MIME type is same as the MIME type of attachment created in step 1, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "image/jpeg",
                getAttachmentInfoResponse.Attachments[0].ContentType,
                60,
                @"[In t:AttachmentType Complex Type][The ContentType element] specifies the MIME type of the attachment content.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R64");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R64
            // If the returned attachment's ContentLocation is same as the ContentLocation of attachment created in step 1, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                @"http://www.contoso.com/xyz.abc",
                getAttachmentInfoResponse.Attachments[0].ContentLocation,
                64,
                @"[In t:AttachmentType Complex Type] The ContentLocation element can be used to associate an attachment with a URL that defines its location on the Web.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #region Step 3 Delete the file attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);
            DeleteAttachmentResponseMessageType deleteAttachmentResponseMessage = deleteAttachmentResponse.ResponseMessages.Items[0] as DeleteAttachmentResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R443");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R443
            // If the deleted attachment's RootItemId is same as the id of store item on server, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                this.ItemId.Id,
                deleteAttachmentResponseMessage.RootItemId.RootItemId,
                443,
                @"[In t:RootItemIdType Complex Type][The RootItemId attribute] Identifies the root item of an attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R446");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R446
            // If the deleted attachment's RootItemChangeKey exists, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                deleteAttachmentResponseMessage.RootItemId.RootItemChangeKey,
                446,
                @"[In t:RootItemIdType Complex Type][The RootItemChangeKey attribute] Identifies the new change key of the root item of an attachment.");

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);
        }

        /// <summary>
        /// This test case is designed to verify processing an item attachment which contains a MessageType item as the child item.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC02_ProcessMessageTypeItemAttachment()
        {
            #region Step 1 Create an item attachment, which contains a MessageType item as the child item, on an item.

            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.MessageAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentType createdAttachment = createAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentId = createdAttachment.AttachmentId;

            #endregion

            #region Step 2 Get the item attachment created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Text, true, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoResponse = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 1, this.Site);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R55001");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R55001
            // If the MIMEContent of returned attachment is not null, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.MimeContent,
                55001,
                @"[In t:AttachmentResponseShapeType Complex Type] If the IncludeMimeContent element is set to true in the AttachmentResponseShapeType complex type, the MIME content will be returned for attachment of the item class: IPM.Note. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R202");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R202
            // If the attachment created in step 1 is successfully returned, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                202,
                @"[In DeleteAttachment Operation] Before an attachment can be deleted, the item MUST be retrieved from the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R416");
            bool isR416Verified = getAttachmentInfoResponse.ResponseClass == ResponseClassType.Success && getAttachmentInfoResponse.Attachments[0].AttachmentId.Id != null;

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R416
            Site.CaptureRequirementIfIsTrue(
                isR416Verified,
                416,
                @"[In t:AttachmentType Complex Type][The AttachmentId element] specifies the attachment identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R63");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R63
            // If the ContentLocation of returned attachment is same as the one of attachment created in step 1, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                @"http://www.contoso.com/xyz.abc",
                getAttachmentInfoResponse.Attachments[0].ContentLocation,
                63,
                @"[In t:AttachmentType Complex Type][The ContentLocation element] specifies the URI that corresponds to the location of the content of the attachment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R36");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R36
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getAttachmentInfoResponse.Attachments[0],
                typeof(ItemAttachmentType),
                36,
                @"[In t:ArrayOfAttachmentsType Complex Type][The ItemAttachment element] specifies an item that is attached to another item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R349");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R349
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                349,
                @"[In t:ItemAttachmentType Complex Type][The type of Message element is] t:MessageType ([MS-OXWSMSG] section 2.2.4.1)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R83");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R83
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                83,
                @"[In t:ItemAttachmentType Complex Type][The Message element] Represents a server e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R475");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R475
            // if all of additional properties are returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.DateTimeCreated,
                475,
                @"[In m:GetAttachmentType Complex Type][The AttachmentShape element] Contains additional properties to return for the attachments.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R521");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R521
            // if one of additional properties is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.DateTimeCreated,
                521,
                @"[In Complex Types] [Complex type name]AttachmentResponseShapeType Specifies additional properties for the GetAttachment operation to return.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R549");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R549
            // if one of additional properties is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.DateTimeCreated,
                549,
                @"[In t:AttachmentResponseShapeType Complex Type][The AdditionalProperties element] Contains additional properties to return in a response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R478");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R478
            this.Site.CaptureRequirementIfAreEqual<string>(
                BodyTypeResponseType.Text.ToString(),
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.Body.BodyType1.ToString(),
                478,
                @"[In t:AttachmentResponseShapeType Complex Type][The BodyType element] Represents the format of the body text in a response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1234");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1234
            // The element "t:Path" is contained in additional property if additional properties are returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.DateTimeCreated,
                "MS-OXWSCDATA",
                1234,
                @"[In t:NonEmptyArrayOfPathsToElementType Complex Type] The element ""t:Path"" with type ""t:Path"" specifies a property to be returned in a response.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #region Step 3 Delete the item attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);
            DeleteAttachmentResponseMessageType deleteAttachmentResponseMessage = deleteAttachmentResponse.ResponseMessages.Items[0] as DeleteAttachmentResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R466");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R466
            // if the RootItemId is same as the id of the store item, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                this.ItemId.Id,
                deleteAttachmentResponseMessage.RootItemId.RootItemId,
                466,
                @"[In m:DeleteAttachmentResponseMessageType Complex Type][The RootItemId element] Specifies the parent item of a deleted attachment.");

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);
        }

        /// <summary>
        /// This test case is designed to verify processing an item attachment which contains a CalendarItem item as the child item.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC03_ProcessCalendarItemTypeItemAttachment()
        {
            #region Step 1 Create an item attachment, which contains a CalendarItemType item as the child item, on an item.

            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.CalendarAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentType createdAttachment = createAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentId = createdAttachment.AttachmentId;

            #endregion

            #region Step 2 Get the item attachment created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, true, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoResponse = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 1, this.Site);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R55003");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R55003
            // If the MIMEContent of returned attachment is not null, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.MimeContent,
                55003,
                @"[In t:AttachmentResponseShapeType Complex Type] If the IncludeMimeContent element is set to true in the AttachmentResponseShapeType complex type, the MIME content will be returned for attachment of the item class: IPM.Appointment. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R311");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R311
            // If the MIMEContent of returned attachment is not null, this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.MimeContent,
                311,
                @"[In t:AttachmentResponseShapeType Complex Type][in IncludeMimeContent] A text value of ""true"" indicates that the attachment contains MIME content.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R476");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R476
            // If the returned attachment name is same as the name of attachment created in step 1, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                createAttachmentInfoResponse.Attachments[0].AttachmentId.Id,
                getAttachmentInfoResponse.Attachments[0].AttachmentId.Id,
                476,
                @"[In m:GetAttachmentType Complex Type][The AttachmentIds element] Contains the identifiers of the attachments to return in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R350");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R350
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                350,
                @"[In t:ItemAttachmentType Complex Type][The type of CalendarItem element is] t:CalendarItemType ([MS-OXWSMTGS] section 2.2.4.4)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R85");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R85
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                85,
                @"[In t:ItemAttachmentType Complex Type][The CalendarItem element] Represents a calendar item.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #region Step 3 Delete the item attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify processing an item attachment which contains a ContactItemType item as the child item.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC04_ProcessContactItemTypeItemAttachment()
        {
            #region Step 1 Create an item attachment, which contains a ContactItemType item as the child item, on an item.

            // A bool value indicate whether the R55003 needs to be captured.
            bool isR550Implemented = Common.IsRequirementEnabled(550, Site);
            
            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.ContactAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentType createdAttachment = createAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentId = createdAttachment.AttachmentId;

            #endregion

            #region Step 2 Get the item attachment created in step 1 by the GetAttachment operation.

            // Get attachment include Mime body.
            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, true, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoResponse = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the length.
            this.Site.Assert.AreEqual<int>(
                 1,
                 getAttachmentResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getAttachmentResponse.ResponseMessages.Items.GetLength(0));

            if (isR550Implemented)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R550");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R550
                // If the MIMEContent of returned attachment is not null, this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.MimeContent,
                    550,
                    @"[In Appendix C: Product Behavior] Implementation does return MIME content for attachments of IPM.Contact, when the IncludeMimeContent element is set to true in the AttachmentResponseShapeType complex type.  <1> (Exchange Server 2013 and above follow this behavior.)");
            }

            // Get attachment not include Mime body.
            GetAttachmentResponseType getAttachmentWithoutMimeResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, false, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoWithoutMimeResponse = getAttachmentWithoutMimeResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentWithoutMimeResponse, 1, this.Site);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R351");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R351
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoWithoutMimeResponse.ResponseClass,
                351,
                @"[In t:ItemAttachmentType Complex Type][The type of Contact element is] t:ContactItemType ([MS-OXWSCONT] section 2.2.4.2)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R554");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R554
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoWithoutMimeResponse.ResponseClass,
                554,
                @"[In t:ItemAttachmentType Complex Type][The Contact element] Represents a contact item.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentWithoutMimeResponse);

            #region Step 3 Delete the item attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify processing a MIMEContent-excluded item attachment which contains a TaskType item as the child item.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC05_ProcessTaskTypeItemAttachmentWithoutMIMEContent()
        {
            #region Step 1 Create an item attachment, which contains a TaskType item as the child item, on an item.

            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.TaskAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentType createdAttachment = createAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentId = createdAttachment.AttachmentId;

            #endregion

            #region Step 2 Get the item attachment created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, false, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoResponse = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 1, this.Site);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R356");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R356
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                356,
                @"[In t:ItemAttachmentType Complex Type][The type of Task element is] t:TaskType ([MS-OXWSTASK] section 2.2.4.3)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R97");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R97
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                97,
                @"[In t:ItemAttachmentType Complex Type][The Task element] Represents a task in the server store.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #region Step 3 Delete the item attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify processing an item attachment which contains a PostItemType item as the child item.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC06_ProcessPostItemTypeItemAttachment()
        {
            #region Step 1 Create an item attachment, which contains a PostItemType item as the child item, on an item.

            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.PostAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentIdType createdAttachmentId = createAttachmentInfoResponse.Attachments[0].AttachmentId;

            #endregion

            #region Step 2 Get the item attachment created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, true, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoResponse = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 1, this.Site);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R55002");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R55002
            // If the MIMEContent of returned attachment is not null, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.MimeContent,
                55002,
                @"[In t:AttachmentResponseShapeType Complex Type] If the IncludeMimeContent element is set to true in the AttachmentResponseShapeType complex type, the MIME content will be returned for attachment of the item class: IPM.Post. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R357");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R357
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                357,
                @"[In t:ItemAttachmentType Complex Type][The type of PostItem element is] t:PostItemType ([MS-OXWSPOST] section 2.2.4.1)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R99");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R99
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                99,
                @"[In t:ItemAttachmentType Complex Type][The PostItem element] Represents a post item in the server store.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #region Step 3 Delete the item attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify processing multiple attachments.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC07_ProcessMultipleAttachments()
        {
            #region Configure SOAP header

            this.ConfigureSOAPHeader();

            #endregion

            #region Step 1 Create 2 attachments on an item.

            // Create a file attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.FileAttachment, AttachmentTypeValue.ItemAttachment);

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 2, this.Site);

            List<AttachmentIdType> attachmentIds = new List<AttachmentIdType>();
            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);
            foreach (AttachmentInfoResponseMessageType createAttachmentInfoResponse in createAttachmentResponse.ResponseMessages.Items)
            {
                // Gets the ID of the created attachment.
                foreach (AttachmentType attachment in createAttachmentInfoResponse.Attachments)
                {
                    attachmentIds.Add(attachment.AttachmentId);
                }
            }

            #endregion

            #region Step 2 Get the attachments created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, false, attachmentIds.ToArray());
            AttachmentInfoResponseMessageType getAttachmentInfo1 = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;
            AttachmentInfoResponseMessageType getAttachmentInfo2 = getAttachmentResponse.ResponseMessages.Items[1] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 2, this.Site);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R141");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R141
            // Since the second new-created attachment (item attachment) is gotten successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfo2.ResponseClass,
                141,
                @"[In CreateAttachment Operation] An item attachment does not exist as a store item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R458");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R458
            // If the two attachments are gotten successfully, this requirement can be captured.
            // Since the success of getting item attachment is verified by R141, so following capture code will only verify the success of getting file attachment. 
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfo1.ResponseClass,
                458,
                @"[In m:CreateAttachmentType Complex Type][The Attachments element] Contains the items or files that are attached to an item in the server store.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R6201");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R6201
            // If the returned attachment's ContentId is same as the ContentId of attachment created in step 1, this requirement can be captured.
            Site.CaptureRequirementIfAreNotEqual<string>(
                getAttachmentInfo1.Attachments[0].ContentId,
                getAttachmentInfo2.Attachments[0].ContentId,
                6201,
                @"[In t:AttachmentType Complex Type][The ContentId element] If N (N=2) attachments are not the same, the object identifier for each attachment is different.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #region Step 3 Delete the attachments created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(attachmentIds.ToArray());

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 2, this.Site);

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);

            #endregion

            #region Step 4 Get the attachments created in step 1 again by the GetAttachment operation to see if they have been deleted successfully.

            GetAttachmentResponseType getAttachmentAfterDeleteResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, false, attachmentIds.ToArray());
            AttachmentInfoResponseMessageType getAttachmentAfterDeleteInfo1 = getAttachmentAfterDeleteResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;
            AttachmentInfoResponseMessageType getAttachmentAfterDeleteInfo2 = getAttachmentAfterDeleteResponse.ResponseMessages.Items[1] as AttachmentInfoResponseMessageType;

            // Check the length.
            Site.Assert.AreEqual<int>(2, getAttachmentAfterDeleteResponse.ResponseMessages.Items.GetLength(0), "Expected Item Count: {0}, Actual Item Count: {1}", 2, getAttachmentAfterDeleteResponse.ResponseMessages.Items.GetLength(0));
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R547");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R547
            // Since the second new-created attachment (item attachment) is gotten unsuccessfully, this requirement can be captured.
            Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentAfterDeleteInfo2.ResponseClass,
                547,
                @"[In DeleteAttachment Operation] An item attachment does not exist as a store item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R467");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R467
            // If the two attachments are deleted successfully, this requirement can be captured.
            // Since the success of deleting item attachment is verified by R547, so following capture code will only verify the success of deleting file attachment. 
            Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentAfterDeleteInfo1.ResponseClass,
                467,
                @"[In m:DeleteAttachmentType Complex Type][The AttachmentIds element] Contains the items or files that are attached to an item in the server store to be deleted.");
        }

        /// <summary>
        /// This test case is designed to verify processing an item attachment which contains an ItemType item as the child item.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC08_ProcessItemTypeAttachment()
        {
            #region Step 1 Create an item attachment, which contains an ItemType item as the child item, on an item.

            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.ItemAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentType createdAttachment = createAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentId = createdAttachment.AttachmentId;

            #endregion

            #region Step 2 Get the item attachment created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, false, createdAttachmentId);
            AttachmentInfoResponseMessageType getAttachmentInfoResponse = getAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 1, this.Site);
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R3111");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R3111
            // when the MIMEContent of returned attachment is null, this requirement can be captured.
            this.Site.CaptureRequirementIfIsNull(
                ((ItemAttachmentType)getAttachmentInfoResponse.Attachments[0]).Item.MimeContent,
                3111,
                @"[In t:AttachmentResponseShapeType Complex Type][in IncludeMimeContent] A text value of ""false"" indicates that the attachment doesn't contain MIME content.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R527");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R527.
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                527,
                @"[In t:ItemAttachmentType Complex Type][The type of Item element is] t:ItemType ([MS-OXWSCORE] section 2.2.4.8).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R81");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R81
            // When the created attachment is returned successfully, this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getAttachmentInfoResponse.ResponseClass,
                81,
                @"[In t:ItemAttachmentType Complex Type][The Item element] Represents a generic item in the server store.");

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #region Step 3 Delete the item attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify processing an attachment to an existing attachment.
        /// </summary>
        [TestCategory("MSOXWSATT"), TestMethod()]
        public void MSOXWSATT_S01_TC09_AttachmentToAttachment()
        {
            #region Step 1 Create an item attachment, which contains an ItemType item as the child item, on an item.

            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentResponse = this.CallCreateAttachmentOperation(this.ItemId.Id, AttachmentTypeValue.ItemAttachment);
            AttachmentInfoResponseMessageType createAttachmentInfoResponse = createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentResponse, 1, this.Site);

            this.VerifyCreateAttachmentSuccessfulResponse(createAttachmentResponse);

            // Gets the ID of the created attachment.
            AttachmentType createdAttachment = createAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentId = createdAttachment.AttachmentId;

            #endregion

            #region Step 2 Create an attachment to the attachment created in step 1.

            // Create an item attachment by CreateAttachment operation.
            CreateAttachmentResponseType createAttachmentAttachmentResponse = this.CallCreateAttachmentOperation(createdAttachmentId.Id, AttachmentTypeValue.ItemAttachment);
            AttachmentInfoResponseMessageType createAttachmentAttachmentInfoResponse = createAttachmentAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(createAttachmentAttachmentResponse, 1, this.Site);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R14201");

            // When the parent attachment and its attachment both created successfully, this requirement can be captured.
            bool isR14201Verified = createAttachmentInfoResponse.ResponseClass == ResponseClassType.Success && createAttachmentAttachmentInfoResponse.ResponseClass == ResponseClassType.Success;

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R14201
            Site.CaptureRequirementIfIsTrue(
                isR14201Verified,
                14201,
                @"[In CreateAttachment Operation] It [item attachment] exists as an attachment to an item or another attachment.");

            // Gets the ID of the created attachment.
            AttachmentType createdAttachmentAttachment = createAttachmentAttachmentInfoResponse.Attachments[0];
            AttachmentIdType createdAttachmentAttachmentId = createdAttachmentAttachment.AttachmentId;

            #region Step 3 Get the parent attachment created in step 1 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, false, createdAttachmentId);

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentResponse, 1, this.Site);

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentResponse);

            #endregion

            #region Step 4 Get the attachment's attachment created in step 2 by the GetAttachment operation.

            GetAttachmentResponseType getAttachmentAttachmentResponse = this.CallGetAttachmentOperation(BodyTypeResponseType.Best, false, createdAttachmentAttachmentId);

            // Check the response.
            Common.CheckOperationSuccess(getAttachmentAttachmentResponse, 1, this.Site);

            this.VerifyGetAttachmentSuccessfulResponse(getAttachmentAttachmentResponse);

            #endregion

            #region Step 5 Delete the attachment's attachment created in step 2 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentAttachmentId);
            DeleteAttachmentResponseMessageType deleteAttachmentAttachmentResponseMessage = deleteAttachmentAttachmentResponse.ResponseMessages.Items[0] as DeleteAttachmentResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentAttachmentResponse, 1, this.Site);

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentAttachmentResponse);

            #endregion

            #region Step 6 Delete the parent attachment created in step 1 by the DeleteAttachment operation.

            DeleteAttachmentResponseType deleteAttachmentResponse = this.CallDeleteAttachmentOperation(createdAttachmentId);
            DeleteAttachmentResponseMessageType deleteAttachmentResponseMessage = deleteAttachmentResponse.ResponseMessages.Items[0] as DeleteAttachmentResponseMessageType;

            // Check the response.
            Common.CheckOperationSuccess(deleteAttachmentResponse, 1, this.Site);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R20001");

            // When the parent attachment and its attachment both deleted successfully, this requirement can be captured.
            bool isR20001Verified = deleteAttachmentResponseMessage.ResponseClass == ResponseClassType.Success && deleteAttachmentAttachmentResponseMessage.ResponseClass == ResponseClassType.Success;

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R20001
            Site.CaptureRequirementIfIsTrue(
                isR20001Verified,
                20001,
                @"[In DeleteAttachment Operation] It [the item attachment] exists as an attachment to an item or another attachment.");

            this.VerifyDeleteAttachmentSuccessfulResponse(deleteAttachmentResponse);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Verify the CreateAttachment operation response.
        /// </summary>
        /// <param name="createAttachmentResponse">A CreateAttachmentResponseType instance.</param>
        private void VerifyCreateAttachmentSuccessfulResponse(CreateAttachmentResponseType createAttachmentResponse)
        {
            foreach (AttachmentInfoResponseMessageType createAttachmentInfoResponse in createAttachmentResponse.ResponseMessages.Items)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R144");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R144
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                            ResponseClassType.Success,
                            createAttachmentInfoResponse.ResponseClass,
                            144,
                            @"[In CreateAttachment Operation] A successful CreateAttachment operation request returns a CreateAttachmentResponse element with the ResponseClass attribute of the CreateAttachmentResponseMessage element set to ""Success"".");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R145");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R145
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                            ResponseCodeType.NoError,
                            createAttachmentInfoResponse.ResponseCode,
                            145,
                            @"[In CreateAttachment Operation][A successful CreateAttachment operation request returns a CreateAttachmentResponse element] The ResponseCode element of the CreateAttachmentResponse element is set to ""NoError"".");
            }
        }

        /// <summary>
        /// Verify the DeleteAttachment operation response.
        /// </summary>
        /// <param name="deleteAttachmentResponse">A DeleteAttachmentResponseType instance.</param>
        private void VerifyDeleteAttachmentSuccessfulResponse(DeleteAttachmentResponseType deleteAttachmentResponse)
        {
            foreach (DeleteAttachmentResponseMessageType deleteAttachmentResponseMessage in deleteAttachmentResponse.ResponseMessages.Items)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R203");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R203
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    deleteAttachmentResponseMessage.ResponseClass,
                    203,
                    @"[In DeleteAttachment Operation] A successful DeleteAttachment operation request returns a DeleteAttachmentResponse element with the ResponseClass attribute of the DeleteAttachmentResponseMessage element set to ""Success"".");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R204");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R204
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    deleteAttachmentResponseMessage.ResponseCode,
                    204,
                    @"[In DeleteAttachment Operation][A successful DeleteAttachment operation request returns a DeleteAttachmentResponse element] The ResponseCode element of the DeleteAttachmentResponse element is set to ""NoError"".");
            }
        }

        /// <summary>
        /// Verify the GetAttachment operation response.
        /// </summary>
        /// <param name="getAttachmentResponse">A GetAttachmentResponseType instance.</param>
        private void VerifyGetAttachmentSuccessfulResponse(GetAttachmentResponseType getAttachmentResponse)
        {
            foreach (AttachmentInfoResponseMessageType getAttachmentInfoResponse in getAttachmentResponse.ResponseMessages.Items)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R259");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R259
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                            ResponseClassType.Success,
                            getAttachmentInfoResponse.ResponseClass,
                            259,
                            @"[In GetAttachment Operation] A successful GetAttachment operation request returns a GetAttachmentResponse element with the ResponseClass attribute of the GetAttachmentResponseMessage element set to ""Success"".");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R260");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R260
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                            ResponseCodeType.NoError,
                            getAttachmentInfoResponse.ResponseCode,
                            260,
                            @"[In GetAttachment Operation] [A successful GetAttachment operation request returns a GetAttachmentResponse element ] The ResponseCode element of the GetAttachmentResponse element is set to ""NoError"".");
            }
        }
        #endregion
    }
}