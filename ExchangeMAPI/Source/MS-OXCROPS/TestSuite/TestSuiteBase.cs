namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A Protocol Test Framework test class.
    /// </summary>
    [TestClass]
    [SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1401:FieldsMustBePrivate", Justification = "Disable warning SA1401 because it should not be treated like a property.")]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables and Consts

        /// <summary>
        /// Definition for ByteCount,the MaximumByteCount field specifies the maximum number of bytes to read if it is set.
        /// </summary>
        public const ushort ByteCount = 0xBABE;

        /// <summary>
        /// Indicate False for fields of Boolean type
        /// </summary>
        protected const int Zero = 0;

        /// <summary>
        /// Definition for AddressType which specifies the address type for which options are to be returned.
        /// </summary>
        protected const string AddressType = "X400\0";

        /// <summary>
        /// Definition for BookmarkSize specifies the size of the Bookmark field.
        /// </summary>
        protected const ushort BookmarkSize = 0x0;

        /// <summary>
        /// Definition for code page that will be used for string values associated with the message.
        /// </summary>
        protected const ushort CodePageId = 0x0FFF;

        /// <summary>
        /// Definition for ClientEntryId which specifies the ID of the downloaded message on the client.
        /// </summary>
        protected const byte ClientEntryId = 0x0;

        /// <summary>
        /// Definition for ClientEntryIdSize which specifies the size of the ClientEntryId field.
        /// </summary>
        protected const ushort ClientEntryIdSize = 0x0001;

        /// <summary>
        /// Definition for DataForProperties which is used to construct properties.
        /// </summary>
        protected const string DataForProperties = "ABCDEFGABCDEFGABCDEFGABCDEFGABCDEFGABCDEFGABCDEFGABCDEFG";

        /// <summary>
        /// Definition for default value of folder handle.
        /// </summary>
        protected const int DefaultFolderHandle = 0xFFFF;

        /// <summary>
        /// Definition for Denominator which represents the denominator of the fraction identifying the table position to seek to
        /// </summary>
        protected const uint Denominator = 0x00000004;

        /// <summary>
        /// Definition for ExcludedTagCount which specifies how many tags are present in ExcludedTags.
        /// </summary>
        protected const ushort ExcludedTagCount = 0x0;

        /// <summary>
        /// Definition for Level(NonZero),NonZero means excluding all descendant sub-objects from being copied.
        /// </summary>
        protected const byte LevelOfNonZero = 0xFF;

        /// <summary>
        /// Definition for Level(Zero), which specifies the level at which the copy is occurring.
        /// </summary>
        protected const byte LevelOfZero = 0x00;

        /// <summary>
        /// Definition for LoopCounter.
        /// </summary>
        protected const int LoopCounter = 0x4F000;

        /// <summary>
        /// Definition for MaximumByteCount,This value specifies the maximum number of bytes to read if the value
        /// of the ByteCount field is equal to 0xBABE.
        /// </summary>
        protected const uint MaximumByteCount = 0x0;

        /// <summary>
        /// Definition for MaximumByteCount, this value is larger than 0x80000000.
        /// </summary>
        protected const uint ExceedMaxCount = 0x90000000;

        /// <summary>
        /// Definition for MaximumBufferSize which specifies the maximum size limit when the server determines the buffer size.
        /// </summary>
        protected const ushort MaximumBufferSize = 0xFFFF;

        /// <summary>
        /// Definition for MaxRowCount which specifies the maximum number of expanded rows to return data for.
        /// </summary>
        protected const ushort MaxRowCount = 0x0002;

        /// <summary>
        /// Definition for MessageStatusMask which specifies which bits in the MessageStatusFlags field are to be changed.
        /// </summary>
        protected const uint MessageStatusMask = 0x00001000;

        /// <summary>
        /// Definition for Numerator which represents the numerator of the fraction identifying the table position to seek to.
        /// </summary>
        protected const uint Numerator = 0x00000000;

        /// <summary>
        /// Definition for PropertyId.
        /// </summary>
        protected const ushort PropertyId = 0x00FF;

        /// <summary>
        /// Definition for PropertyIdCount.
        /// </summary>
        protected const ushort PropertyIdCount = 1;

        /// <summary>
        /// Definition for PropertySizeLimit which specifies the maximum size allowed for a property value returned.
        /// </summary>
        protected const ushort PropertySizeLimit = 0xFFFF;

        /// <summary>
        /// Definition for RecipientColumnCount which specifies the number of columns from the RecipientColumns field.
        /// that are included in RecipientProperties.
        /// </summary>
        protected const ushort RecipientColumnCount = 0x000C;

        /// <summary>
        /// Definition for RestrictionDataSize which specifies the length of the RestrictionData field.
        /// </summary>
        protected const ushort RestrictionDataSize1 = 0x0005;

        /// <summary>
        /// Definition for RestrictionDataSize which specifies the length of the RestrictionData field.
        /// </summary>
        protected const ushort RestrictionDataSize2 = 0x00;

        /// <summary>
        /// Definition for ReturnValue of ROP failure response.
        /// </summary>
        protected const int ReturnValueForRopFail = 0x000004B9;

        /// <summary>
        /// Definition for RowCount.
        /// </summary>
        protected const ushort RowCount = 0x1000;

        /// <summary>
        /// Definition for RowId which specifies the ID of the recipient.
        /// </summary>
        protected const uint RowId = 0x00000000;

        /// <summary>
        /// Definition for RowInstanceNumber which specifies the instance number of the row that is to be preserved as the cursor.
        /// </summary>
        protected const uint RowInstanceNumber = 0x0000;

        /// <summary>
        /// Definition for ServerEntryId, which specifies the ID of the message on the server.
        /// </summary>
        protected const byte ServerEntryId = 0x0;

        /// <summary>
        /// Definition for ServerEntryIdSize which specifies the size of the ServerEntryId field.
        /// </summary>
        protected const ushort ServerEntryIdSize = 0x0001;

        /// <summary>
        /// Definition for StreamSize which specifies the size of the stream.
        /// </summary>
        protected const ulong StreamSize = 0x00000000000000FF;

        /// <summary>
        /// Definition for the value field of TaggedPropertyValue for PidTagMemberId.
        /// </summary>
        protected const ulong TaggedPropertyValueForPidTagMemberId = 0xFFFFFFFFFFFFFFFF;

        /// <summary>
        /// Definition for the value field of TaggedPropertyValue for PidTagMemberRights.
        /// </summary>
        protected const int TaggedPropertyValueForPidTagMemberRights = 0x00000800;

        /// <summary>
        /// This property is used to set ByteCount of RopReadStream request.
        /// </summary>
        protected const ushort ByteCountForRopReadStream = 0x04D2;

        /// <summary>
        /// This property is used to set ByteCount of RopCopyToStream request.
        /// </summary>
        protected const ulong ByteCountForRopCopyToStream = 0x000000000000000F;

        /// <summary>
        /// This property is used to set the DataOffset field of PerUserInformation.
        /// </summary>
        protected const uint DataOffset = 0x0000;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the destination Server object is stored.
        /// </summary>
        protected const byte DestHandleIndex = 0x1;

        /// <summary>
        /// This property is used to set the DisplayName and Comment fields for non-search folder.
        /// </summary>
        protected const string DisplayNameAndCommentForNonSearchFolder = "DisplayNameAndComment1";

        /// <summary>
        /// This property is used to set the DisplayName and Comment fields for search folder.
        /// </summary>
        protected const string DisplayNameAndCommentForSearchFolder = "DisplayNameAndComment2";

        /// <summary>
        /// This property is used to set GlobalCounter field of LongTermId.
        /// </summary>
        protected const int GlobalCounter = 255;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the input Server object is stored.
        /// </summary>
        protected const byte InputHandleIndex0 = 0x0;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the input Server object is stored.
        /// </summary>
        protected const byte InputHandleIndex1 = 0x1;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the input Server object is stored.
        /// </summary>
        protected const byte InputHandleIndex2 = 0x2;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the input Server object is stored.
        /// </summary>
        protected const byte InputHandleIndex3 = 0x3;

        /// <summary>
        /// This is an invalid CopyFlags and will be used to invoke the failure response.
        /// </summary>
        protected const byte InvalidCopyFlags = 0xF;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        protected const byte LogonId = 0x0;

        /// <summary>
        /// Another logonId
        /// </summary>
        protected const byte LogonId1 = 0x1;

        /// <summary>
        /// This property is used to set MaxDataSize field of PerUserInformation.
        /// </summary>
        protected const ushort MaxDataSize = 0x001E;

        /// <summary>
        /// This property is used to be as maxValue parameter of Random method.
        /// </summary>
        protected const int MaxValueForRandom = 326777;

        /// <summary>
        /// This property is used to set MessageClass of RopGetReceiveFolder and RopSetReceiveFolder.
        /// </summary>
        protected const string MessageClassForReceiveFolder = "Test.Class";

        /// <summary>
        /// Configure the count of Created messages in RopProgress test case for get success response.
        /// </summary>
        protected const int MessagesCount = 2000;

        /// <summary>
        /// This property is used to set MessageClass of RopTransportNewMail.
        /// </summary>
        protected const string MessageClassForRopTransportNewMail = "IPM.Note";

        /// <summary>
        /// This value represents True for fields of Boolean type.
        /// </summary>
        protected const byte NonZero = 0xFF;

        /// <summary>
        /// This property specifies the seek offset, it should be less than stream size.
        /// </summary>
        protected const ulong Offset = 5;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the output Server object will be stored.
        /// </summary>
        protected const byte OutputHandleIndex0 = 0x0;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the output Server object will be stored.
        /// </summary>
        protected const byte OutputHandleIndex1 = 0x1;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the output Server object will be stored.
        /// </summary>
        protected const byte OutputHandleIndex2 = 0x2;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the output Server object will be stored.
        /// </summary>
        protected const byte OutputHandleIndex3 = 0x3;

        /// <summary>
        /// This property specifies the byte location in the stream where the region begins.
        /// </summary>
        protected const ulong RegionOffset = 0;

        /// <summary>
        /// This property specifies the size of the region.
        /// </summary>
        protected const ulong RegionSize = 0;

        /// <summary>
        /// This field is reversed and MUST be set to 0.
        /// </summary>
        protected const byte Reserved = 0x0;

        /// <summary>
        /// This value specifies the location in the Server object handle table that is referenced in the response.
        /// </summary>
        protected const byte ResponseHandleIndex1 = 0x1;

        /// <summary>
        /// This value specifies the location in the Server object handle table that is referenced in the response.
        /// </summary>
        protected const byte ResponseHandleIndex2 = 0x3;

        /// <summary>
        /// This property is used to construct a stream data.
        /// </summary>
        protected const string SampleStreamData = "SampleStreamData";

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the source Server object is stored.
        /// </summary>
        protected const byte SourceHandleIndex0 = 0x0;

        /// <summary>
        /// This value specifies the location in the Server object handle table where the handle for the source Server object is stored.
        /// </summary>
        protected const byte SourceHandleIndex1 = 0x1;

        /// <summary>
        /// This property is used to be as the user defined PropertyId.
        /// </summary>
        protected const ushort UserDefinedPropertyId = 0x04D2;

        /// <summary>
        /// This property is used to cause the RopCollapseRow and RopExpandRows failure response.
        /// </summary>
        protected const ulong WrongCategoryId = 0;

        /// <summary>
        /// This property is used to cause a RopGetOwningServers failure response.
        /// </summary>
        protected const ulong WrongFolderId = 1;

        /// <summary>
        /// This property is used to cause RopLongTermIdFromId failure response.
        /// </summary>
        protected const ulong WrongObjectId = 0;

        /// <summary>
        /// This property is used to cause Logon failure response.
        /// </summary>
        protected const string WrongUserDN = "Use Wrong UserDN to cause Logon Failure";

        /// <summary>
        /// Definition for ecNullObject.
        /// </summary>
        protected const int ReturnValueForecNullObject = 0x000004B9;

        /// <summary>
        /// Definition for ecAccessDenied.
        /// </summary>
        protected const uint ReturnValueForecAccessDenied = 0x80070005;

        /// <summary>
        /// The value specifies the response buffer out of range.
        /// </summary>
        protected const uint BufferOutOfRange = 0x10;

        /// <summary>
        /// Success response return value.
        /// </summary>
        protected const uint SuccessReturnValue = 0;

        /// <summary>
        /// Adapter object.
        /// </summary>
        protected IMS_OXCROPSAdapter cropsAdapter;

        /// <summary>
        /// Server object handle in request.
        /// </summary>
        protected uint inputObjHandle;

        /// <summary>
        /// The ROP response payload.
        /// </summary>
        protected byte[] rawData;

        /// <summary>
        /// This property is used to Logon a specific mailbox.
        /// </summary>
        protected string userDN = string.Empty;

        /// <summary>
        /// ROP response objects.
        /// </summary>
        protected IDeserializable response;

        /// <summary>
        /// Server objects handles in response.
        /// </summary>
        protected List<List<uint>> responseSOHs;

        /// <summary>
        /// The dictionary holds the PropertyTags
        /// </summary>
        protected Dictionary<PropertyNames, PropertyTag> propertyDictionary = PropertyDictionary();

        /// <summary>
        /// The value to indicate whether to execute the cleanup method
        /// </summary>
        private bool needDoCleanup;

        #endregion

        #region Common protected Type(s) and Function(s)

        /// <summary>
        /// This method is used to log on to a private mailbox or public folder and will be called before other ROPs
        /// </summary>
        /// <param name="logonType">The logon type.</param>
        /// <param name="userDN">The user Essdn to be used to logon a specific mailbox</param>
        /// <param name="objHandle">The Object handle.</param>
        /// <returns>The logon response.</returns>
        protected RopLogonResponse Logon(LogonType logonType, string userDN, out uint objHandle)
        {
            RopLogonRequest logonRequest;

            logonRequest.RopId = (byte)RopId.RopLogon;
            logonRequest.LogonId = LogonId;

            // Set OutputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            logonRequest.OutputHandleIndex = OutputHandleIndex0;
            logonRequest.StoreState = (uint)StoreState.None;

            if (LogonType.PublicFolder == logonType)
            {
                // Set other parameters for logon type of PublicFolder.
                logonRequest.LogonFlags = (byte)LogonFlags.PublicFolder;
                logonRequest.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping + (uint)OpenFlags.Public;

                // Set EssdnSize to 0, which specifies the size of the Essdn field.
                logonRequest.EssdnSize = 0;

                // Initialize the Essdn to null.
                logonRequest.Essdn = null;
            }
            else
            {
                // Set other parameters for logon type of Mailbox (private mailbox).
                logonRequest.LogonFlags = (byte)LogonFlags.Private;
                logonRequest.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;

                // Set EssdnSize to the byte count of user DN, which specifies the size of the Essdn field.
                logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN);

                // Set Essdn to the content of user DN, which specifies it will log on to the mail box of user represented by the user DN.
                logonRequest.Essdn = Encoding.ASCII.GetBytes(userDN);
            }

            // Send the RopLogon request and get the response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                logonRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;

            Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    logonResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            objHandle = this.responseSOHs[0][logonResponse.OutputHandleIndex];
            return logonResponse;
        }

        /// <summary>
        /// Hard delete messages and subfolders under the specified folder.
        /// </summary>
        /// <param name="userName">The user name to be used to logon a specific mailbox</param>
        /// <param name="passWord">The user password to be used to logon a specific mailbox</param>
        /// <param name="userEssdn">The user Essdn to be used to logon a specific mailbox</param>
        /// <param name="folderIndex">The id of folder in which subfolders and messages should be deleted</param>
        /// <param name="needWaitMail">Indicates whether need waiting the mail to be received. True means need waiting, otherwise not.</param>
        protected void HardDeleteMessagesAndSubfolders(string userName, string passWord, string userEssdn, int folderIndex, bool needWaitMail = false)
        {
            // Log on to a private mailbox.
            bool ret = this.cropsAdapter.RpcDisconnect();
            this.Site.Assert.IsTrue(ret, "Rpc disconnect should be success.");
            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                userEssdn,
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                userName,
                passWord);
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, userEssdn, out this.inputObjHandle);

            // Step 1: Open a folder.
            #region Open a folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            openFolderRequest.InputHandleIndex = InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = OutputHandleIndex1;

            // Set FolderId to the 4th folder of the logonResponse, which specifies the folder to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[folderIndex];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Send the RopOpenFolder request and verify the success response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)this.response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send RopGetContentsTable request.
            #region Send RopGetContentsTable request

            RopGetContentsTableRequest getContentsTableRequest;
            RopGetContentsTableResponse getContentsTableResponse;

            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getContentsTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            getContentsTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            // Send a RopGetContentsTable request to the server and verify the success response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                getContentsTableRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getContentsTableResponse = (RopGetContentsTableResponse)this.response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getContentsTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            // The row content before send RopHardDeleteMessagesAndSubfolders Request.
            uint rowCount = getContentsTableResponse.RowCount;
            int waitTime;
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            if (needWaitMail && rowCount == 0)
            {
                do
                {
                    waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                    System.Threading.Thread.Sleep(waitTime);

                    // Send a RopGetContentsTable request to the server and verify the success response.
                    this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                        getContentsTableRequest,
                        openedFolderHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse);
                    getContentsTableResponse = (RopGetContentsTableResponse)this.response;

                    Site.Assert.AreEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        getContentsTableResponse.ReturnValue,
                        "If ROP succeeds, the ReturnValue of its response is 0(success)");

                    rowCount = getContentsTableResponse.RowCount;
                    retryCount--;
                }
                while (rowCount == 0 && retryCount > 0);
            }

            #endregion

            // The time to wait RopHardDeleteMessagesAndSubfolders succeed.
            waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTimeForCleanUp", this.Site));
            retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCountForCleanUp", this.Site));
            do
            {
                // Step 3: Hard Delete Messages and Subfolders
                #region Hard Delete Messages and Subfolders
                RopHardDeleteMessagesAndSubfoldersRequest hardDeleteRequest;
                RopHardDeleteMessagesAndSubfoldersResponse hardDeleteResponse;
                hardDeleteRequest.RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders;
                hardDeleteRequest.LogonId = LogonId;
                hardDeleteRequest.InputHandleIndex = InputHandleIndex0;
                hardDeleteRequest.WantAsynchronous = 0x00; // Synchronously
                hardDeleteRequest.WantDeleteAssociated = 0xFF; // TRUE: delete all messages and subfolders
                // Send a RopHardDeleteMessagesAndSubfolders request to the server and verify the success response.
                this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                    hardDeleteRequest,
                    openedFolderHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                hardDeleteResponse = (RopHardDeleteMessagesAndSubfoldersResponse)this.response;

                // The return value is 1125 (ecNoDelSubmitMsg) which means that deleting a message that has been submitted is not permitted.
                bool returnValue = (hardDeleteResponse.ReturnValue == 0) || (hardDeleteResponse.ReturnValue == 1125);

                Site.Assert.AreEqual<bool>(
                    true,
                    returnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0(success). The returned value is {0}.",
                    hardDeleteResponse.ReturnValue);
                #endregion

                // Send a RopGetContentsTable request to the server and verify the success response.
                this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                    getContentsTableRequest,
                    openedFolderHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                getContentsTableResponse = (RopGetContentsTableResponse)this.response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getContentsTableResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0(success)");

                rowCount = getContentsTableResponse.RowCount;

                if (rowCount != 0)
                {
                    retryCount--;
                    System.Threading.Thread.Sleep(waitTime);
                }
                else
                {
                    break;
                }
            }
            while (rowCount != 0 && retryCount > 0);
        }

        /// <summary>
        /// Commit Stream
        /// </summary>
        /// <param name="streamHandle">The opened stream handle</param>
        protected void CommitStream(uint streamHandle)
        {
            RopCommitStreamRequest commitStreamRequest;
            RopCommitStreamResponse commitStreamResponse;

            commitStreamRequest.RopId = (byte)RopId.RopCommitStream;
            commitStreamRequest.LogonId = LogonId;
            commitStreamRequest.InputHandleIndex = InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopCommitStream request in CommitStream method.");

            // Send the RopCommitStream request and verify success response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                commitStreamRequest,
                streamHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            commitStreamResponse = (RopCommitStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                commitStreamResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
        }

        /// <summary>
        /// Save Message
        /// </summary>
        /// <param name="messageHandle">The message handle</param>
        protected void SaveMessage(uint messageHandle)
        {
            RopSaveChangesMessageRequest saveChangesMessageRequest;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = LogonId;
            saveChangesMessageRequest.InputHandleIndex = InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopSaveChangesMessage request in SaveMessage method.");

            // Send the RopSaveChangesMessage request and verify success response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
        }

        /// <summary>
        /// Create Sample ContentsTable PropertyTags
        /// </summary>
        /// <returns>Return PropertyTag array</returns>
        protected PropertyTag[] CreateSampleContentsTablePropertyTags()
        {
            // The following sample tags is from [MS-OXCTABL].
            PropertyTag[] propertyTags = new PropertyTag[6];

            // PidTagFolderId
            propertyTags[0] = this.propertyDictionary[PropertyNames.PidTagFolderId];

            // PidTagMid
            propertyTags[1] = this.propertyDictionary[PropertyNames.PidTagMid];

            // PidTagInstID
            propertyTags[2] = this.propertyDictionary[PropertyNames.PidTagInstID];

            // PidTagInstanceNum
            propertyTags[3] = this.propertyDictionary[PropertyNames.PidTagInstanceNum];

            // PidTagSubject
            propertyTags[4] = this.propertyDictionary[PropertyNames.PidTagSubject];

            // PidTagMessageDeliveryTime
            propertyTags[5] = this.propertyDictionary[PropertyNames.PidTagMessageDeliveryTime];

            return propertyTags;
        }

        /// <summary>
        /// This method creates Sample RecipientColumns and Sample RecipientRows.
        /// </summary>
        /// <param name="recipientColumns">Sample RecipientColumns</param>
        /// <param name="recipientRows">Sample RecipientRows</param>
        protected void CreateSampleRecipientColumnsAndRecipientRows(out PropertyTag[] recipientColumns, out ModifyRecipientRow[] recipientRows)
        {
            // Step 1: Create Sample RecipientColumns.
            #region recipientColumns

            // The following sample data is from MS-OXCMSG 4.7.1.
            PropertyTag[] sampleRecipientColumns = new PropertyTag[12];
            sampleRecipientColumns[0] = this.propertyDictionary[PropertyNames.PidTagObjectType];
            sampleRecipientColumns[1] = this.propertyDictionary[PropertyNames.PidTagDisplayType];
            sampleRecipientColumns[2] = this.propertyDictionary[PropertyNames.PidTagAddressBookDisplayNamePrintable];
            sampleRecipientColumns[3] = this.propertyDictionary[PropertyNames.PidTagSmtpAddress];
            sampleRecipientColumns[4] = this.propertyDictionary[PropertyNames.PidTagSendInternetEncoding];
            sampleRecipientColumns[5] = this.propertyDictionary[PropertyNames.PidTagDisplayTypeEx];
            sampleRecipientColumns[6] = this.propertyDictionary[PropertyNames.PidTagRecipientDisplayName];
            sampleRecipientColumns[7] = this.propertyDictionary[PropertyNames.PidTagRecipientFlags];
            sampleRecipientColumns[8] = this.propertyDictionary[PropertyNames.PidTagRecipientTrackStatus];
            sampleRecipientColumns[9] = this.propertyDictionary[PropertyNames.PidTagRecipientResourceState];
            sampleRecipientColumns[10] = this.propertyDictionary[PropertyNames.PidTagRecipientOrder];
            sampleRecipientColumns[11] = this.propertyDictionary[PropertyNames.PidTagRecipientEntryId];
            recipientColumns = sampleRecipientColumns;

            #endregion

            // Step 2: Configure a StandardPropertyRow: propertyRow.
            #region Configure a StandardPropertyRow: propertyRow, data is from Page 62 of MS-OXCMSG

            PropertyValue[] propertyValueArray = new PropertyValue[12];
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValueArray[i] = new PropertyValue();
            }

            // PidTagObjectType
            propertyValueArray[0].Value = BitConverter.GetBytes(0x00000006);

            // PidTagDisplayType
            propertyValueArray[1].Value = BitConverter.GetBytes(0x00000000);

            // PidTa7BitDisplayName
            propertyValueArray[2].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("EmailAlias", this.Site) + "\0");

            // PidTagSmtpAddress
            propertyValueArray[3].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("EmailAlias", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site) + "\0");

            // PidTagSendInternetEncoding
            propertyValueArray[4].Value = BitConverter.GetBytes(0x00000000);

            // PidTagDisplayTypeEx
            propertyValueArray[5].Value = BitConverter.GetBytes(0x40000000);

            // PidTagRecipientDisplayName
            propertyValueArray[6].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("EmailAlias", this.Site) + "\0");

            // PidTagRecipientFlags
            propertyValueArray[7].Value = BitConverter.GetBytes(0x00000001);

            // PidTagRecipientTrackStatus
            propertyValueArray[8].Value = BitConverter.GetBytes(0x00000000);

            // PidTagRecipientResourceState
            propertyValueArray[9].Value = BitConverter.GetBytes(0x00000000);

            // PidTagRecipientOrder
            propertyValueArray[10].Value = BitConverter.GetBytes(0x00000000);

            // The following sample data (0x007c and the subsequent 124(0x7c) binary)
            // is copied from Page 62 of MS-OXCMSG
            byte[] sampleData = 
            { 
                0x7c, 0x00, 0x00, 0x00, 0x00, 0x00, 0xdc, 0xa7, 0x40, 0xc8,
                0xc0, 0x42, 0x10, 0x1a, 0xb4, 0xb9, 0x08, 0x00, 0x2b, 0x2f,
                0xe1, 0x82, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x2f, 0x6f, 0x3d, 0x46, 0x69, 0x72, 0x73, 0x74, 0x20, 0x4f,
                0x72, 0x67, 0x61, 0x6e, 0x69, 0x7a, 0x61, 0x74, 0x69, 0x6f,
                0x6e, 0x2f, 0x6f, 0x75, 0x3d, 0x45, 0x78, 0x63, 0x68, 0x61,
                0x6e, 0x67, 0x65, 0x20, 0x41, 0x64, 0x6d, 0x69, 0x6e, 0x69,
                0x73, 0x74, 0x72, 0x61, 0x74, 0x69, 0x76, 0x65, 0x20, 0x47,
                0x72, 0x6f, 0x75, 0x70, 0x20, 0x28, 0x46, 0x59, 0x44, 0x49,
                0x42, 0x4f, 0x48, 0x46, 0x32, 0x33, 0x53, 0x50, 0x44, 0x4c,
                0x54, 0x29, 0x2f, 0x63, 0x6e, 0x3d, 0x52, 0x65, 0x63, 0x69,
                0x70, 0x69, 0x65, 0x6e, 0x74, 0x73, 0x2f, 0x63, 0x6e, 0x3d,
                0x75, 0x73, 0x65, 0x72, 0x32, 0x00 
            };

            // PidTagRecipientEntryId
            propertyValueArray[11].Value = sampleData;

            List<PropertyValue> propertyValues = new List<PropertyValue>();
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValues.Add(propertyValueArray[i]);
            }

            PropertyRow propertyRow = new PropertyRow
            {
                Flag = (byte)PropertyRowFlag.FlaggedPropertyRow,
                PropertyValues = propertyValues
            };

            // For propertyRow.Flag
            int tempLengthForTest = 1;
            foreach (PropertyValue pv in propertyRow.PropertyValues)
            {
                tempLengthForTest = tempLengthForTest + pv.Value.Length;
            }

            #endregion

            // Step 3: Create Sample RecipientRows.
            #region recipientRows

            RecipientRow recipientRow = new RecipientRow
            {
                // 0101 1001 0000 0110 S,D, Type=SMTP,I,U,E
                RecipientFlags = (ushort)(
                    RecipientFlags.SMTP
                    | RecipientFlags.S
                    | RecipientFlags.D
                    | RecipientFlags.I
                    | RecipientFlags.U
                    | RecipientFlags.E),

                // Set DisplayName, which specifies the Email Address of the recipient, as specified in [MS-OXCDATA].
                // This field is present because D is Set.
                DisplayName = Encoding.Unicode.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0"),

                // Set EmailAddress, which specifies the Email Address of the recipient,
                // as specified in [MS-OXCDATA].
                EmailAddress =
                    Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("EmailAlias", this.Site) + "@" +
                                              Common.GetConfigurationPropertyValue("Domain", this.Site) + "\0"),

                // Set SimpleDisplayName, which specifies the Email Address of the recipient,
                // as specified in [MS-OXCDATA].
                SimpleDisplayName =
                    Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("EmailAlias", this.Site) + "\0"),

                // Set RecipientColumnCount, which specifies the number of columns from the RecipientColumns field
                // that are included in RecipientProperties, as specified in [MS-OXCDATA].
                RecipientColumnCount = RecipientColumnCount,
                RecipientProperties = propertyRow
            };
            ModifyRecipientRow modifyRecipientRow = new ModifyRecipientRow
            {
                RowId = RowId, // Set RowId, which value specifies the ID of the recipient.
                RecipientType = (byte)RecipientType.PrimaryRecipient,
                RecipientRowSize = (ushort)recipientRow.Size(), // Set RecipientRowSize, which specifies the size of the RecipientRow field.
                RecptRow = recipientRow.Serialize()
            };
            ModifyRecipientRow[] sampleModifyRecipientRows = new ModifyRecipientRow[1];
            sampleModifyRecipientRows[0] = modifyRecipientRow;
            recipientRows = sampleModifyRecipientRows;

            #endregion
        }

        /// <summary>
        /// Get ContentsTable Handle
        /// </summary>
        /// <param name="targetFolderHandle">The target folder object handle</param>
        /// <returns>Return the ContentsTable handle</returns>
        protected uint GetContentsTableHandle(uint targetFolderHandle)
        {
            RopGetContentsTableRequest getContentsTableRequest;
            RopGetContentsTableResponse getContentsTableResponse;

            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;

            getContentsTableRequest.LogonId = LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            getContentsTableRequest.InputHandleIndex = InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            getContentsTableRequest.OutputHandleIndex = OutputHandleIndex1;

            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetContentsTable request in GetContentsTableHandle method.");

            // Send the RopGetContentsTable request and verify the success response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                getContentsTableRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getContentsTableResponse = (RopGetContentsTableResponse)this.response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getContentsTableResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint contentsTableHandle = this.responseSOHs[0][getContentsTableResponse.OutputHandleIndex];
            return contentsTableHandle;
        }

        /// <summary>
        /// Get Created Message Handle
        /// </summary>
        /// <param name="folderId">The folder Id be used to create message</param>
        /// <param name="logonHandle">The RopLogon object handle</param>
        /// <returns>Return created Message Handle</returns>
        protected uint GetCreatedMessageHandle(ulong folderId, uint logonHandle)
        {
            // Create a message object first.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = LogonId;
            createMessageRequest.InputHandleIndex = InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = CodePageId;

            createMessageRequest.FolderId = folderId;
            createMessageRequest.AssociatedFlag = Convert.ToByte(Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopCreateMessage request in GetCreatedMessageHandle method.");

            // Send the RopCreateMessage to create message.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                logonHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)this.response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success).");

            // Get handle of message object, it will be used as input handle in next ROP
            uint messageHandle = this.responseSOHs[0][createMessageResponse.OutputHandleIndex];
            return messageHandle;
        }

        /// <summary>
        /// Get Folder Object Handle.
        /// </summary>
        /// <param name="logonResponse">The response of RopLogon</param>
        /// <returns>Return the folder object handle</returns>
        protected uint GetFolderObjectHandle(ref RopLogonResponse logonResponse)
        {
            // Step 1: Open a folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = LogonId;
            openFolderRequest.InputHandleIndex = InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request in GetFolderObjectHandle method.");

            // Send the RopOpenFolder request to open folder.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)this.response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success0");

            uint openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create subfolder.
            #region Create subfolder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = LogonId;
            createFolderRequest.InputHandleIndex = InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(Zero);
            createFolderRequest.OpenExisting = NonZero;
            createFolderRequest.Reserved = Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request in GetFolderObjectHandle method.");

            // Send the RopCreateFolder request and verify the success response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success0");

            #endregion

            // Step 3: Get and return the handle of created folder.
            uint folderHandle = this.responseSOHs[0][createFolderResponse.OutputHandleIndex];
            return folderHandle;
        }

        /// <summary>
        /// Find a folder ID in the specified folder and with specified display name.
        /// </summary>
        /// <param name="openedFolderHandle">Handle of the parent folder.</param>
        /// <param name="folderName">The folder display name.</param>
        /// <returns>The folder ID.</returns>
        protected ulong GetSubfolderIDByName(uint openedFolderHandle, string folderName)
        {
            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest();
            RopGetHierarchyTableResponse getHierarchyTableResponse = new RopGetHierarchyTableResponse();
            getHierarchyTableRequest.RopId = (byte)RopId.RopGetHierarchyTable;
            getHierarchyTableRequest.LogonId = TestSuiteBase.LogonId;
            getHierarchyTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            getHierarchyTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.Depth;
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                getHierarchyTableRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.response;
            uint tableHandle = this.responseSOHs[0][getHierarchyTableResponse.OutputHandleIndex];

            PropertyTag[] properties = new PropertyTag[]
            {
                new PropertyTag()
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagDisplayName].PropertyId,
                    PropertyType = (ushort)PropertyType.PtypString
                },
                new PropertyTag()
                {
                    PropertyId = (ushort)this.propertyDictionary[PropertyNames.PidTagFolderId].PropertyId,
                    PropertyType = (ushort)PropertyType.PtypInteger64
                }
            };
            List<PropertyRow> propertyRows = this.GetTableRowValue(tableHandle, (ushort)getHierarchyTableResponse.RowCount, properties);

            ulong folderId = 0;
            foreach (PropertyRow propertyRow in propertyRows)
            {
                byte[] displayNameInBytes = propertyRow.PropertyValues[0].Value;
                string displayName = Encoding.Unicode.GetString(displayNameInBytes, 0, displayNameInBytes.Length);
                if (displayName.ToLower() == folderName.ToLower())
                {
                    byte[] folderIdInBytes = propertyRow.PropertyValues[1].Value;
                    folderId = (ulong)BitConverter.ToInt64(folderIdInBytes, 0);
                    break;
                }
            }

            return folderId;
        }

        /// <summary>
        /// Create Vast Messages In InBox
        /// </summary>
        /// <param name="logonResponse">the logon response be used to create message</param>
        /// <param name="tableHandle">The tableHanlder of the new folder</param>
        /// <param name="count">The count of created messages</param>
        /// <param name="createMessageRequest">The ROP CreateMessageRequest</param>
        /// <param name="saveChangesMessageRequest">The ROP SaveChangesMessageRequest</param>
        /// <param name="releaseRequest">The ROP ReleaseRequest</param>
        protected void CreateVastMessages(ref RopLogonResponse logonResponse, out uint tableHandle, int count, RopCreateMessageRequest createMessageRequest, RopSaveChangesMessageRequest saveChangesMessageRequest, RopReleaseRequest releaseRequest)
        {
            RopCreateMessageResponse createMessageResponse;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            #region Preparing the table: CreateFolder

            // Open a folder first
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopOpenFolder request.");

            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            // Create a new subfolder in the opened folder
            // The new subfolder will be used as target folder
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;
            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];
            tableHandle = this.GetContentsTableHandle(targetFolderHandle);
            ulong folderId = createFolderResponse.FolderId;

            #endregion

            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int maxRetryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int retryCount;
            bool needRetry = false;
            uint returnValue = 0;
            int loopcount = 50;
            List<ISerializable> multipleRopRequests = new List<ISerializable>();
            List<IDeserializable> multipleRopResponses = new List<IDeserializable>();
            List<uint> multipleInputObjects = new List<uint>();
            List<uint> multipleOutputObjects = new List<uint>();

            for (int i = 0; i < count; i++)
            {
                // If the RPC report error code reported by the following three ROP methods is 1726 (The remote procedure call failed), 
                // re-do the common steps of this case.
                if (returnValue == 1726)
                {
                    // Step 1: Create a message.
                    #region Create message
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "If RPC error code is 1726, re-connect to server.");
                    this.cropsAdapter.RpcConnect(
                        Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                        ConnectionType.PrivateMailboxServer,
                        Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                        Common.GetConfigurationPropertyValue("Domain", this.Site),
                        Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                        Common.GetConfigurationPropertyValue("PassWord", this.Site));
                    logonResponse = this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);
                    #endregion

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopOpenFolder request.");

                    this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                        openFolderRequest,
                        this.inputObjHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse);
                    openFolderResponse = (RopOpenFolderResponse)this.response;

                    Site.Assert.AreEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        openFolderResponse.ReturnValue,
                        "if ROP succeeds, the ReturnValue of its response is 0(success)");
                    openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopCreateFolder request.");

                    this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                        createFolderRequest,
                        openedFolderHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse);
                    createFolderResponse = (RopCreateFolderResponse)this.response;
                    Site.Assert.AreEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        createFolderResponse.ReturnValue,
                        "if ROP succeeds, the ReturnValue of its response is 0(success)");
                    targetFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];
                    tableHandle = this.GetContentsTableHandle(targetFolderHandle);
                    folderId = createFolderResponse.FolderId;
                }

                // Step 1: Create a message.
                #region Create message

                createMessageRequest.FolderId = logonResponse.FolderIds[4];

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 1:Begin to send the RopCreateMessage request in CreateVastMessages method.");

                retryCount = maxRetryCount;
                do
                {
                    multipleRopRequests.Clear();
                    multipleRopResponses.Clear();
                    multipleOutputObjects.Clear();
                    multipleInputObjects.Clear();

                    // Send multiple RopCreateMessage to create messages.
                    for (int t = 0; t < loopcount; t++)
                    {
                        createMessageRequest.FolderId = folderId;
                        createMessageRequest.InputHandleIndex = 0;
                        createMessageRequest.OutputHandleIndex = Convert.ToByte(t + 1);
                        multipleRopRequests.Add(createMessageRequest);
                        if (t == 0)
                        {
                            multipleInputObjects.Add(this.inputObjHandle);
                        }
                        else
                        {
                            multipleInputObjects.Add(TestSuiteBase.DefaultFolderHandle);
                        }
                    }

                    // An additional default handle is needed since the output handle index is from 1 to t+1 in the above loop, and the last one is not added in the above loop.
                    multipleInputObjects.Add(TestSuiteBase.DefaultFolderHandle);

                    this.responseSOHs = this.cropsAdapter.ProcessMutipleRops(
                      multipleRopRequests,
                      multipleInputObjects,
                      ref multipleRopResponses,
                      ref this.rawData,
                      RopResponseType.SuccessResponse);

                    for (int t = 0; t < multipleRopResponses.Count; t++)
                    {
                        if (multipleRopResponses[t] is RopBackoffResponse)
                        {
                            needRetry = true;
                            break;
                        }
                    }

                    System.Threading.Thread.Sleep(waitTime);
                    retryCount--;
                }
                while (needRetry && retryCount >= 0);

                Site.Assert.IsTrue(retryCount >= 0, "The case {0} failed since server is busy, reduce your server load and try again.", this.TestContext.TestName);

                // If the error code is 1726, continue this loop.
                if (returnValue == 1726)
                {
                    continue;
                }

                for (int t = 0; t < multipleRopResponses.Count; t++)
                {
                    if (multipleRopResponses[t] is RopCreateMessageResponse)
                    {
                        createMessageResponse = (RopCreateMessageResponse)multipleRopResponses[t];
                        Site.Assert.AreEqual<uint>(
                            TestSuiteBase.SuccessReturnValue,
                            createMessageResponse.ReturnValue,
                            "if ROP succeeds, the ReturnValue of its response is 0(success)");
                        uint targetMessageHandle = this.responseSOHs[0][createMessageResponse.OutputHandleIndex];
                        multipleOutputObjects.Add(targetMessageHandle);
                    }
                }
                #endregion

                // Step 2: Save the created message.
                #region Save message

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopSaveChangesMessage request in CreateVastMessages method.");

                needRetry = false;
                retryCount = maxRetryCount;

                do
                {
                    multipleRopRequests.Clear();
                    multipleRopResponses.Clear();

                    // Send multiple RopSaveChangesMessage requests to save the created messages.
                    for (int t = 0; t < loopcount; t++)
                    {
                        saveChangesMessageRequest.InputHandleIndex = Convert.ToByte(t);
                        multipleRopRequests.Add(saveChangesMessageRequest);
                    }

                    this.responseSOHs = this.cropsAdapter.ProcessMutipleRops(
                     multipleRopRequests,
                     multipleOutputObjects,
                     ref multipleRopResponses,
                     ref this.rawData,
                     RopResponseType.SuccessResponse);

                    for (int t = 0; t < multipleRopResponses.Count; t++)
                    {
                        if (multipleRopResponses[t] is RopSaveChangesMessageResponse)
                        {
                            saveChangesMessageResponse = (RopSaveChangesMessageResponse)multipleRopResponses[t];
                            if (saveChangesMessageResponse.ReturnValue == 0x80040401)
                            {
                                needRetry = true;
                                break;
                            }
                        }
                        else if (this.response is RopBackoffResponse)
                        {
                            needRetry = false;
                            break;
                        }
                    }

                    System.Threading.Thread.Sleep(waitTime);
                    retryCount--;
                }
                while (this.response is RopBackoffResponse && retryCount >= 0);

                Site.Assert.IsTrue(retryCount >= 0, "The case {0} failed since server is busy, reduce your server load and try again.", this.TestContext.TestName);

                // If the error code is 1726, continue this loop.
                if (returnValue == 1726)
                {
                    continue;
                }

                for (int t = 0; t < multipleRopResponses.Count; t++)
                {
                    if (multipleRopResponses[t] is RopSaveChangesMessageResponse)
                    {
                        saveChangesMessageResponse = (RopSaveChangesMessageResponse)multipleRopResponses[t];
                        Site.Assert.AreEqual<uint>(
                            TestSuiteBase.SuccessReturnValue,
                            saveChangesMessageResponse.ReturnValue,
                        "if ROP succeeds, the ReturnValue of its response is 0(success)");
                    }
                }

                #endregion

                // Step 3: Send the RopRelease request to release all resources associated with a Server object.
                #region Release all resources

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopRelease request in CreateVastMessages method.");
                needRetry = false;
                retryCount = maxRetryCount;

                do
                {
                    multipleRopRequests.Clear();
                    multipleRopResponses.Clear();
                    for (int t = 0; t < loopcount; t++)
                    {
                        releaseRequest.InputHandleIndex = Convert.ToByte(t);
                        multipleRopRequests.Add(releaseRequest);
                    }

                    this.responseSOHs = this.cropsAdapter.ProcessMutipleRops(
                     multipleRopRequests,
                     multipleOutputObjects,
                     ref multipleRopResponses,
                     ref this.rawData,
                     RopResponseType.SuccessResponse);

                    if (multipleRopResponses.Count != 0)
                    {
                        for (int t = 0; t < multipleRopResponses.Count; t++)
                        {
                            if (multipleRopResponses[t] is RopBackoffResponse)
                            {
                                needRetry = true;
                                break;
                            }
                        }
                    }

                    System.Threading.Thread.Sleep(waitTime);
                    retryCount--;
                }
                while (needRetry && retryCount >= 0);

                Site.Assert.IsTrue(retryCount >= 0, "The case {0} failed since server is busy, reduce your server load and try again.", this.TestContext.TestName);

                // If the error code is 1726, continue this loop.
                if (returnValue == 1726)
                {
                    continue;
                }
                #endregion
            }

            // If the error code 1726 occurs on the last time of the above "for" loop, re-do the common steps.
            if (returnValue == 1726)
            {
                #region The common steps
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "If RPC error code is 1726, re-connect to server.");
                this.cropsAdapter.RpcConnect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    ConnectionType.PrivateMailboxServer,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                    Common.GetConfigurationPropertyValue("PassWord", this.Site));
                logonResponse = this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);
                #endregion
            }
        }

        /// <summary>
        /// Create message each one loop In InBox
        /// </summary>
        /// <param name="logonResponse">the logon response be used to create message</param>
        /// <param name="tableHandle">The tableHanlder of the new folder</param>
        /// <param name="count">The count of created messages</param>
        /// <param name="createMessageRequest">The ROP CreateMessageRequest</param>
        /// <param name="saveChangesMessageRequest">The ROP SaveChangesMessageRequest</param>
        /// <param name="releaseRequest">The ROP ReleaseRequest</param>
        protected void CreateSingleProcessEachLoop(ref RopLogonResponse logonResponse, out uint tableHandle, int count, RopCreateMessageRequest createMessageRequest, RopSaveChangesMessageRequest saveChangesMessageRequest, RopReleaseRequest releaseRequest)
        {
            RopCreateMessageResponse createMessageResponse = new RopCreateMessageResponse();
            RopSaveChangesMessageResponse saveChangesMessageResponse = new RopSaveChangesMessageResponse();
            #region Preparing the table: CreateFolder

            // Open a folder first
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopOpenFolder request.");

            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            // Create a new subfolder in the opened folder
            // The new subfolder will be used as target folder
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;
            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];
            tableHandle = this.GetContentsTableHandle(targetFolderHandle);
            ulong folderId = createFolderResponse.FolderId;

            #endregion

            #region Preparing the table: RopCreateAndSaveMessages

            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int maxRetryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int retryCount;
            uint returnValue = 0;
            for (int i = 1; i < count; i++)
            {
                // If the RPC report error code reported by the following three ROP methods is 1726 (The remote procedure call failed), 
                // re-do the common steps of this case.
                if (returnValue == 1726)
                {
                    #region The common steps
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "If RPC error code is 1726, re-connect to server.");
                    this.cropsAdapter.RpcConnect(
                        Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                        ConnectionType.PrivateMailboxServer,
                        Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                        Common.GetConfigurationPropertyValue("Domain", this.Site),
                        Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                        Common.GetConfigurationPropertyValue("PassWord", this.Site));
                    logonResponse = this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopOpenFolder request.");

                    this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                        openFolderRequest,
                        this.inputObjHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse);
                    openFolderResponse = (RopOpenFolderResponse)this.response;

                    Site.Assert.AreEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        openFolderResponse.ReturnValue,
                        "if ROP succeeds, the ReturnValue of its response is 0(success)");
                    openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopCreateFolder request.");

                    this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                        createFolderRequest,
                        openedFolderHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse);
                    createFolderResponse = (RopCreateFolderResponse)this.response;
                    Site.Assert.AreEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        createFolderResponse.ReturnValue,
                        "if ROP succeeds, the ReturnValue of its response is 0(success)");
                    targetFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];
                    tableHandle = this.GetContentsTableHandle(targetFolderHandle);
                    folderId = createFolderResponse.FolderId;
                    #endregion
                }

                #region Create message
                createMessageRequest.FolderId = folderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopCreateMessage request:loop counter i={0}", i);
                retryCount = maxRetryCount;

                do
                {
                    // Send the RopCreateMessage to create message.
                    this.responseSOHs = this.cropsAdapter.ProcessSingleRopWithReturnValue(
                        createMessageRequest,
                        this.inputObjHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse,
                        out returnValue);
                    System.Threading.Thread.Sleep(waitTime);
                    retryCount--;
                }
                while (this.response is RopBackoffResponse && retryCount >= 0);

                Site.Assert.IsTrue(retryCount >= 0, "The case {0} failed since server is busy and always returns RopBackoff in the response, reduce your server load and try again.", this.TestContext.TestName);

                // If the error code is 1726, continue this loop.
                if (returnValue == 1726)
                {
                    continue;
                }

                createMessageResponse = (RopCreateMessageResponse)this.response;
                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    createMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
                uint targetMessageHandle = this.responseSOHs[0][createMessageResponse.OutputHandleIndex];
                #endregion

                #region Save message
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopSaveChangesMessage request:loop counter i={0}", i);

                retryCount = maxRetryCount;

                // Do the loop when response is RopBackoffResponse or saveChangesMessageResponse is 0x80040401 (ecTimeout).
                do
                {
                    // Send the RopSaveChangesMessage request to save the created message.
                    this.responseSOHs = this.cropsAdapter.ProcessSingleRopWithReturnValue(
                        saveChangesMessageRequest,
                        targetMessageHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse,
                        out returnValue);

                    if (this.response is RopSaveChangesMessageResponse)
                    {
                        saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
                    }

                    System.Threading.Thread.Sleep(waitTime);
                    retryCount--;
                }
                while ((this.response is RopBackoffResponse || saveChangesMessageResponse.ReturnValue == 0x80040401) && retryCount >= 0);

                Site.Assert.IsTrue(retryCount >= 0, "The case {0} failed since server is busy and always returns RopBackoff in the response, reduce your server load and try again.", this.TestContext.TestName);

                // If the error code is 1726, continue this loop.
                if (returnValue == 1726)
                {
                    continue;
                }

                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    saveChangesMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
                #endregion

                #region Release all resources
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopRelease request.");

                retryCount = maxRetryCount;

                do
                {
                    this.responseSOHs = this.cropsAdapter.ProcessSingleRopWithReturnValue(
                        releaseRequest,
                        targetMessageHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse,
                        out returnValue);
                    System.Threading.Thread.Sleep(waitTime);
                    retryCount--;
                }
                while (this.response is RopBackoffResponse && retryCount >= 0);

                Site.Assert.IsTrue(retryCount >= 0, "The case {0} failed since server is busy and always returns RopBackoff in the response, reduce your server load and try again.", this.TestContext.TestName);

                // If the error code is 1726, continue this loop.
                if (returnValue == 1726)
                {
                    continue;
                }
                #endregion
            }

            // If the error code 1726 occurs on the last time of the above "for" loop, re-do the common steps.
            if (returnValue == 1726)
            {
                #region The common steps
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "If RPC error code is 1726, re-connect to server.");
                this.cropsAdapter.RpcConnect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    ConnectionType.PrivateMailboxServer,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                    Common.GetConfigurationPropertyValue("PassWord", this.Site));
                logonResponse = this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopOpenFolder request.");

                this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                    openFolderRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                openFolderResponse = (RopOpenFolderResponse)this.response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    openFolderResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");
                openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2:Begin to send the RopCreateFolder request.");

                this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                    createFolderRequest,
                    openedFolderHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                createFolderResponse = (RopCreateFolderResponse)this.response;
                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    createFolderResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");
                targetFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];
                tableHandle = this.GetContentsTableHandle(targetFolderHandle);
                folderId = createFolderResponse.FolderId;
                #endregion
            }

            #endregion
        }

        /// <summary>
        /// Prepare the ROP request for RopCreateMessage, RopSaveChangesMessage and RopRelease.
        /// </summary>
        /// <param name="logonResponse">The response of RopLogon.</param>
        /// <param name="createMessageRequest">The request of RopCreateMessage.</param>
        /// <param name="saveChangesMessageRequest">The request of RopSaveChangesMessage.</param>
        /// <param name="releaseRequest">The request of RopRelease.</param>
        protected void PrepareRops(RopLogonResponse logonResponse, ref RopCreateMessageRequest createMessageRequest, ref RopSaveChangesMessageRequest saveChangesMessageRequest, ref RopReleaseRequest releaseRequest)
        {
            #region prepare rops for createmessage, savemessage and release
            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = TestSuiteBase.LogonId;
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            createMessageRequest.FolderId = logonResponse.FolderIds[4];
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Save message 
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;
            releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            #endregion        
        }
        #endregion

        #region Test Initialization and Cleanup

        /// <summary>
        /// Test initialize.
        /// </summary>
        protected override void TestInitialize()
        {
            // Put here code that shall be run before every test case execution.
            // Get the singleton instance of the Adapter
            this.cropsAdapter = Site.GetAdapter<IMS_OXCROPSAdapter>();
            this.userDN = Common.GetConfigurationPropertyValue("UserEssdn", this.Site) + "\0";
            this.needDoCleanup = true;
        }

        /// <summary>
        /// Check the preconditions of this test suite.
        /// </summary>
        protected void CheckTransportIsSupported()
        {
            if ((Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http") && (!Common.IsRequirementEnabled(5000, this.Site)))
            {
                this.needDoCleanup = false;

                Site.Assert.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
            else if ((Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "ncacn_ip_tcp") && (!Common.IsRequirementEnabled(5001, this.Site)))
            {
                this.needDoCleanup = false;

                Site.Assert.Inconclusive("Exchange Server 2013 and above Implementation does not support this given protocol sequence [ncacn_ip_tcp].");
            }
        }

        /// <summary>
        /// Test cleanup.
        /// </summary>
        protected override void TestCleanup()
        {
            if (!this.needDoCleanup)
            {
                return;
            }

            string userName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
            string passWord = Common.GetConfigurationPropertyValue("PassWord", this.Site);

            // Hard Delete Messages and Subfolders of Inbox
            switch (this.TestContext.TestName)
            {
                case "MSOXCROPS_S01_TC01_TestLogonFailed":
                case "MSOXCROPS_S01_TC02_TestLogonRedirect":
                    break;
                default:
                    // Delete subfolders and messages under Inbox folder
                    this.HardDeleteMessagesAndSubfolders(userName, passWord, this.userDN, 4);

                    // Delete subfolders and messages under the 5th folder
                    this.HardDeleteMessagesAndSubfolders(userName, passWord, this.userDN, 5);

                    if (this.TestContext.TestName == "MSOXCROPS_S05_TC01_TestRopSubmitMessage" || this.TestContext.TestName == "MSOXCROPS_S05_TC07_TestRopTransportSend")
                    {
                        userName = Common.GetConfigurationPropertyValue("EmailAlias", this.Site);
                        passWord = Common.GetConfigurationPropertyValue("EmailAliasPassword", this.Site);
                        this.userDN = Common.GetConfigurationPropertyValue("EmailAliasEssdn", this.Site) + "\0";

                        // Delete subfolders and messages under Inbox folder of the receiver's mailbox
                        this.HardDeleteMessagesAndSubfolders(userName, passWord, this.userDN, 4, true);
                    }

                    break;
            }

            // Objects initialization.
            this.rawData = null;
            this.inputObjHandle = 0;
            this.response = null;
            this.responseSOHs = null;

            // Call the Reset method the re-initialize the Site.
            this.cropsAdapter.RpcDisconnect();
        }

        #endregion

        #region Property initialization

        /// <summary>
        /// Initialize StateProperty from property id and property type id.
        /// </summary>
        /// <param name="propertyId">Property id</param>
        /// <param name="propertyTypeId">Property type id</param>
        /// <returns>The StateProperty</returns>
        protected uint GetStatePropertyByIds(ushort propertyId, ushort propertyTypeId)
        {
            // propertyId and propertyTypId should be joined up. So propertyId should be move left 16 positions for propertyTypeID.
            return (uint)((propertyId << 16) | propertyTypeId);
        }

        /// <summary>
        /// Add all property to dictionary for lookup.
        /// </summary>
        /// <returns>Return the dictionary holds the PropertyTags </returns>
        private static Dictionary<PropertyNames, PropertyTag> PropertyDictionary()
        {
            Dictionary<PropertyNames, PropertyTag> dic = new Dictionary<PropertyNames, PropertyTag>();

            InitDictionary(PropertyNames.PidTagHasAttachments, 0x0E1B, PropertyType.PtypBoolean, ref dic);
            InitDictionary(PropertyNames.PidTagMessageClass, 0x001A, PropertyType.PtypString8, ref dic);
            InitDictionary(PropertyNames.PidTagMessageCodepage, 0x3FFD, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagMessageLocaleId, 0x3FF1, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagMessageSize, 0x0E08, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagMessageStatus, 0x0E17, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagSubjectPrefix, 0x003D, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagNormalizedSubject, 0x0E1D, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagImportance, 0x0017, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagPriority, 0x0026, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagSensitivity, 0x0036, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidLidSmartNoAttach, 0x00008514, PropertyType.PtypBoolean, ref dic);
            InitDictionary(PropertyNames.PidLidPrivate, 0x00008506, PropertyType.PtypBoolean, ref dic);
            InitDictionary(PropertyNames.PidLidSideEffects, 0x00008510, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidNameKeywords, 0x0000, PropertyType.PtypMultipleString, ref dic);
            InitDictionary(PropertyNames.PidLidCommonStart, 0x00008516, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidLidCommonEnd, 0x00008517, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidTagAutoForwardComment, 0x0004, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagBody, 0x1000, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagNativeBody, 0x1016, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagBodyHtml, 0x1013, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagRtfCompressed, 0x1009, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagRtfInSync, 0x0E1F, PropertyType.PtypBoolean, ref dic);
            InitDictionary(PropertyNames.PidTagInternetCodepage, 0x3FDE, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidLidContactLinkEntry, 0x00008585, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidLidContacts, 0x0000853A, PropertyType.PtypMultipleString, ref dic);
            InitDictionary(PropertyNames.PidLidContactLinkName, 0x00008586, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidLidContactLinkSearchKey, 0x00008584, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagArchiveTag, 0x3018, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagPolicyTag, 0x3019, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagRetentionPeriod, 0x301A, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagStartDateEtc, 0x301B, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagRetentionDate, 0x301C, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidTagRetentionFlags, 0x301D, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagArchivePeriod, 0x301E, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagArchiveDate, 0x301F, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidTagLastModificationTime, 0x3008, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidTagCreationTime, 0x3007, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidTagDisplayName, 0x3001, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachSize, 0x0E20, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAttachNumber, 0x0E21, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAttachPathname, 0x3708, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachDataBinary, 0x3701, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagAttachMethod, 0x3705, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAttachFilename, 0x3704, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachLongFilename, 0x3707, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachExtension, 0x3703, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachLongPathname, 0x370D, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachTag, 0x370A, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagRenderingPosition, 0X370B, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAttachRendering, 0x3709, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagAttachFlags, 0x3714, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAttachTransportName, 0x370C, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachEncoding, 0x3702, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagAttachAdditionalInformation, 0x370F, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagAttachmentLinkId, 0x7FFA, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAttachmentFlags, 0x7FFD, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAttachmentHidden, 0x7FFE, PropertyType.PtypBoolean, ref dic);
            InitDictionary(PropertyNames.PidTagAttachMimeTag, 0x370E, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachContentId, 0x3712, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachContentLocation, 0x3713, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachContentBase, 0x3711, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagAttachDataObject, 0x3701, PropertyType.PtypComObject, ref dic);
            InitDictionary(PropertyNames.PidTagMessageFlags, 0x0E07, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAccessLevel, 0x0FF7, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagDisplayBcc, 0x0E02, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagDisplayCc, 0x0E03, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagDisplayTo, 0x0E04, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagSecurityDescriptor, 0x0E27, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagUrlCompNameSet, 0x0E62, PropertyType.PtypBoolean, ref dic);
            InitDictionary(PropertyNames.PidTagTrustSender, 0x0E79, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagUrlCompName, 0x10F3, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagSearchKey, 0x300B, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagAccess, 0x0FF4, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagCreatorName, 0x3FF8, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagCreatorEntryId, 0x3FF9, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagLastModifierName, 0x3FFA, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagLastModifierEntryId, 0x3FFB, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagHasNamedProperties, 0x664A, PropertyType.PtypBoolean, ref dic);
            InitDictionary(PropertyNames.PidTagLocaleId, 0x66A1, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagChangeKey, 0x65E2, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagObjectType, 0x0FFE, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRecordKey, 0x0FF9, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagLocalCommitTime, 0x6709, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidTagAutoForwarded, 0x0005, PropertyType.PtypBoolean, ref dic);

            InitDictionary(PropertyNames.PidTagDisplayType, 0x3900, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagAddressBookDisplayNamePrintable, 0x39ff, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagSmtpAddress, 0x39fe, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagSendInternetEncoding, 0x3a71, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagDisplayTypeEx, 0x3905, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRecipientDisplayName, 0x5ff6, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagRecipientFlags, 0x5ffd, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRecipientTrackStatus, 0x5fff, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRecipientResourceState, 0x5fde, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRecipientOrder, 0x5fdf, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRecipientEntryId, 0x5ff7, PropertyType.PtypBinary, ref dic);

            InitDictionary(PropertyNames.PidTagFolderId, 0x6748, PropertyType.PtypInteger64, ref dic);
            InitDictionary(PropertyNames.PidTagMid, 0x674A, PropertyType.PtypInteger64, ref dic);
            InitDictionary(PropertyNames.PidTagInstID, 0x674D, PropertyType.PtypInteger64, ref dic);
            InitDictionary(PropertyNames.PidTagInstanceNum, 0x674E, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagSubject, 0x0037, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagMessageDeliveryTime, 0x0E06, PropertyType.PtypTime, ref dic);
            InitDictionary(PropertyNames.PidTagRowType, 0x0FF5, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagContentCount, 0x3602, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagOfflineAddressBookName, 0x6800, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagOfflineAddressBookSequence, 0x6801, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagOfflineAddressBookContainerGuid, 0x6802, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagOfflineAddressBookMessageClass, 0x6803, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagOfflineAddressBookDistinguishedName, 0x6804, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagSortLocaleId, 0x6705, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagEntryId, 0x0FFF, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagMemberId, 0x6671, PropertyType.PtypInteger64, ref dic);
            InitDictionary(PropertyNames.PidTagMemberName, 0x6672, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagMemberRights, 0x6673, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRuleSequence, 0x6676, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagRuleCondition, 0x6679, PropertyType.PtypRestriction, ref dic);
            InitDictionary(PropertyNames.PidTagRuleActions, 0x6680, PropertyType.PtypRuleAction, ref dic);
            InitDictionary(PropertyNames.PidTagRuleProvider, 0x6681, PropertyType.PtypString, ref dic);
            InitDictionary(PropertyNames.PidTagCnsetSeen, 0x6796, PropertyType.PtypBinary, ref dic);

            InitDictionary(PropertyNames.PidTagSourceKey, 0x65E0, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagPredecessorChangeList, 0x65E3, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagParentSourceKey, 0x65E1, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.PidTagFolderType, 0x3601, PropertyType.PtypInteger32, ref dic);
            InitDictionary(PropertyNames.PidTagTemplateData, 0x0001, PropertyType.PtypBinary, ref dic);
            InitDictionary(PropertyNames.UserSpecified, 0x1234, PropertyType.PtypString, ref dic);

            return dic;
        }

        /// <summary>
        /// Initialize the properties into the dictionary. 
        /// </summary>
        /// <param name="propertyName">The name of property</param>
        /// <param name="propertyId">Property id</param>
        /// <param name="propertyType">Property type id</param>
        /// <param name="propertyTagDictionary">The dictionary to look up PropertyTag by PropertyName</param>
        private static void InitDictionary(
            PropertyNames propertyName,
            ushort propertyId,
            PropertyType propertyType,
            ref Dictionary<PropertyNames,
            PropertyTag> propertyTagDictionary)
        {
            if (propertyTagDictionary == null)
            {
                propertyTagDictionary = new Dictionary<PropertyNames, PropertyTag>();
            }

            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = propertyId, PropertyType = (ushort)propertyType
            };
            propertyTagDictionary.Add(propertyName, propertyTag);
        }

        /// <summary>
        /// Get the properties' value from the rows of the table.
        /// </summary>
        /// <param name="tableHandle">The table handle.</param>
        /// <param name="rowCount">The amount of the rows.</param>
        /// <param name="properties">The properties need to show.</param>
        /// <returns>The property rows in the specified table object.</returns>
        private List<PropertyRow> GetTableRowValue(uint tableHandle, ushort rowCount, PropertyTag[] properties)
        {
            #region The client calls RopSetColumns operation to set the property information to show.

            RopSetColumnsRequest setColumnsRequest = new RopSetColumnsRequest
            {
                RopId = (byte)RopId.RopSetColumns,
                LogonId = TestSuiteBase.LogonId,
                InputHandleIndex = TestSuiteBase.InputHandleIndex0,
                PropertyTagCount = (ushort)properties.Length,
                PropertyTags = properties,
                SetColumnsFlags = (byte)AsynchronousFlags.None
            };
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                    setColumnsRequest,
                    tableHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)this.response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setColumnsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #region The client calls RopQueryRows operation to query the folder which have the special properties.

            RopQueryRowsRequest queryRowsRequest = new RopQueryRowsRequest
            {
                RopId = (byte)RopId.RopQueryRows,
                LogonId = TestSuiteBase.LogonId,
                InputHandleIndex = TestSuiteBase.InputHandleIndex0,
                RowCount = (ushort)rowCount,
                QueryRowsFlags = (byte)QueryRowsFlags.Advance,
                ForwardRead = 0x01
            };
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                    queryRowsRequest,
                    tableHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryRowsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            return queryRowsResponse.RowData.PropertyRows;
        }

        #endregion
    }
}