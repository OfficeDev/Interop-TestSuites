namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    #region Response types
    /// <summary>
    /// Response types.
    /// </summary>
    public enum RopResponseType
    {
        /// <summary>
        /// Success response.
        /// </summary>
        SuccessResponse,

        /// <summary>
        /// Failure response.
        /// </summary>
        FailureResponse,

        /// <summary>
        /// Success response
        /// </summary>
        Response,

        /// <summary>
        /// Null destination failure response.
        /// </summary>
        NullDestinationFailureResponse,

        /// <summary>
        /// Redirect response.
        /// </summary>
        RedirectResponse
    }
    #endregion

    #region HandleIndex
    /// <summary>
    /// Index of Handle in request buffer or response buffer
    /// </summary>
    public enum HandleIndex : byte
    {
        /// <summary>
        /// The first index is 0
        /// </summary>
        FirstIndex = 0,

        /// <summary>
        /// The second index is 1
        /// </summary>
        SecondIndex = 1
    }
    #endregion

    /// <summary>
    /// MS_OXCPRPTAdapter class
    /// </summary>
    public partial class MS_OXCPRPTAdapter : ManagedAdapterBase, IMS_OXCPRPTAdapter
    {
        #region Public fields for consts definitions

        /// <summary>
        /// Definition for several points referring 0xBABE.
        /// </summary>
        public const uint BufferSize = 0xBABE;

        /// <summary>
        /// Flags of RPC_HEADER_EXT,which indicates it is compressed.
        /// </summary>
        public const ushort CompressedForFlagsOfHeader = 0x0001;

        /// <summary>
        /// Definition for default value of Output handle. 
        /// </summary>
        public const uint DefaultOutputHandle = 0xFFFFFFFF;

        /// <summary>
        /// FolderId (8 bytes): 64-bit identifier. This field MUST be set to 0x0000000000000000
        /// </summary>
        public const ulong FolderIdForRopSynchronizationImportHierarchyChange = 0x0000000000000000;

        /// <summary>
        /// Server object handle value 0xFFFFFFFF is used to initialize unused entries of a Server object handle table.
        /// </summary>
        public const uint HandleValueForUnusedEntries = 0xFFFFFFFF;

        /// <summary>
        /// Flags of RPC_HEADER_EXT, which indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT.
        /// </summary>
        public const ushort LastForFlagsOfHeader = 0x0004;

        /// <summary>
        /// The maximum value of pcbOut 
        /// </summary>
        public const uint MaxPcbOut = 0x40000;

        /// <summary>
        /// Definition for MaxrgbOut,which specifies the maximum size of the rgbOut buffer to place in Response
        /// </summary>
        public const int MaxRgbOut = 0x10008;

        /// <summary>
        ///  MessageId (8 bytes): This field MUST be set to 0x0000000000000000.
        /// </summary>
        public const ulong MessageIdForRops = 0x0000000000000000;

        /// <summary>
        /// Definition for the Null-Terminating string.
        /// </summary>
        public const char NullTerminateCharacter = '\0';

        /// <summary>
        /// Definition for PayloadLen which indicates the length of the field that represents the length of payload.
        /// </summary>
        public const int PayloadLen = 0x2;

        /// <summary>
        /// The maximum size of the rgbAuxOut.
        /// </summary>
        public const uint PcbAuxOut = 0x1008;

        /// <summary>
        /// This flags indicates client requests server to not compress or XOR payload of rgbOut and rgbAuxOut.
        /// </summary>
        public const uint PulFlags = 0x00000003;

        /// <summary>
        /// Unsigned 64-bit integer. This value specifies the number of bytes read from the source object or written to the destination object.
        /// </summary>
        public const ulong ReadOrWrittenByteCountForRopCopyToStream = 0x0000000000000000;

        /// <summary>
        /// Definition the one reserved byte: 0x00
        /// </summary>
        public const byte ReservedOneByte = 0x00;

        /// <summary>
        /// Definition the two reserved bytes: 0x0000
        /// </summary>
        public const ushort ReservedTwoBytes = 0x0000;

        /// <summary>
        /// Definition for ReturnValue of PulFlags,which MUST be set to 0x00000000.
        /// </summary>
        public const int ReturnValueForPulFlags = 0x00000000;

        /// <summary>
        /// Definition for the perhaps return value of RopFastTransferSourceGetBufferResponse: 0x00000480
        /// </summary>
        public const uint ReturnValueForRopFastTransferSourceGetBufferResponse = 0x00000480;

        /// <summary>
        /// Definition for the perhaps return value of RopMoveFolderResponse and RopMoveCopyMessagesResponse: 0x00000503
        /// </summary>
        public const uint ReturnValueForRopMoveFolderResponseAndMoveCopyMessage = 0x00000503;

        /// <summary>
        /// Definition for ReturnValue of RopQueryNamedProperties: 0x00040380.
        /// </summary>
        public const uint ReturnValueForRopQueryNamedProperties = 0x00040380;

        /// <summary>
        /// Definition for the perhaps return value of Redirect response: 0x00000478.
        /// </summary>
        public const uint ReturnValueForRopRedirectResponse = 0x00000478;

        /// <summary>
        /// Definition for ReturnValue of ret, which MUST be set to 0x0.
        /// </summary>
        public const int ReturnValueForRet = 0x0;

        /// <summary>
        /// Definition for ReturnValue of success response: 0x00000000.
        /// </summary>
        public const uint SuccessReturnValue = 0x00000000;

        /// <summary>
        /// Definition for RopSize which specifies the size of both this field and the RopsList field.
        /// </summary>
        public const ushort RopSize = 0x2;

        /// <summary>
        /// Definition for Version of RpcHeaderExt,this value MUST be set to 0x00.
        /// </summary>
        public const ushort VersionOfRpcHeaderExt = 0x00;

        /// <summary>
        /// Flags of RPC_HEADER_EXT,which indicates it is obfuscated.
        /// </summary>
        public const ushort XorMagicForFlagsOfHeader = 0x0002;

        /// <summary>
        /// GUID of the Commonly Used Property Set PS_MAPI, the value must not be changed in Exchange products.
        /// </summary>
        public const string PSMAPI = "00020328-0000-0000-C000-000000000046";

        /// <summary>
        /// Value of PropertySizeLimit, used in RopGetPropertiesSpecific and RopGetPropertiesAll, the default value is 300.
        /// </summary>
        public const ushort PropertySizeLimit = 300;

        /// <summary>
        /// Value of property to be written in RopWriteStream, can be any value of string.
        /// </summary>
        public const string WriteData = "1111";

        /// <summary>
        /// Sub folder name used in RopCopyTo and RopCopyProperties, value is arbitrary.
        /// </summary>
        public const string SubFolder = "MS-OXCPRPT_SubFolder";

        /// <summary>
        /// Value of the TestRootFolderName, value is arbitrary.
        /// </summary>
        public const string RootFolderName = "MS-OXCPRPT_TestRootFolder";

        /// <summary>
        /// Name of Folder to be created in InitHandle method for further using, values are arbitrary.
        /// </summary>
        public const string FolderDisplayName = "MS-OXCPRPT_TestFolder";

        /// <summary>
        /// Value of common property used to test common object properties specified in section 2.2.1, which is a special value in the uint range.
        /// </summary>
        public const string CommonValue = "1234";

        /// <summary>
        /// Value of the new property id used to create in RopOpenStream, must set to a special ushort value, which is bigger than 0x8000(32768).
        /// </summary>
        public const ushort NewCreatePropertyID = 36880;

        /// <summary>
        /// The Lid field in PropertyName in which Kind flag is set to 0x00 in GetPropertyName method, which is special value.
        /// </summary>
        public const uint PropertyName00Lid = 1234;

        /// <summary>
        /// The GUID field in PropertyName in which Kind flag is set to 0x00 in GetPropertyName method, which is unique special value in the GUID format.
        /// </summary>
        public const string PropertyName00Guid = "00062002-0000-0000-c000-000000000046";

        /// <summary>
        /// The Name field in PropertyName in which Kind flag is set to 0x01 in GetPropertyName method, must be a string.
        /// </summary>
        public const string PropertyName01Name = "pname1";

        /// <summary>
        /// The GUID field in PropertyName in which Kind flag is set to 0x01 in GetPropertyName method, which is unique special value in the GUID format.
        /// </summary>
        public const string PropertyName01Guid = "00062003-0000-0000-c000-000000000047";

        /// <summary>
        /// The GUID field in PropertyName in which Kind flag is set to 0x02 in GetPropertyName method, which is unique special value in the GUID format.
        /// </summary>
        public const string PropertyNameFFGuid = "00062004-0000-0000-c000-000000000048";

        /// <summary>
        /// The Name field in PropertyName in which Create flag isn't set, and PropertyName doesn't exist in RopGetPropertyIdsFromNamesMethod request, must be a string not equal to the value of PropertyName01Name.
        /// </summary>
        public const string PropertyNameNotMapped = "pnameNotMapped";

        /// <summary>
        /// The Name field in PropertyName in which Create flag isn't set, and PropertyName doesn't exist in RopGetPropertyIdsFromNamesMethod request, which is unique special value in the GUID format.
        /// </summary>
        public const string PropertyNameNotMappedGuid = "00062004-0000-0000-c000-000000000049";
        #endregion

        #region Private Fields

        /// <summary>
        /// Definition for invalid handle. 
        /// </summary>
        private const uint InvalidHandle = 0xFFFFFFFF;

        /// <summary>
        /// Indicates if common configuration has been imported
        /// </summary>
        private static bool commonConfigImported = false;

        /// <summary>
        /// Folder handle array.
        /// </summary>
        private static uint[] cprptFolderHandle = new uint[3];

        /// <summary>
        /// Folder id array.
        /// </summary>
        private static ulong[] cprptFolderId = new ulong[3];

        /// <summary>
        /// Message handle array.
        /// </summary>
        private static uint[] cprptMessageHandle = new uint[3];

        /// <summary>
        /// Message id array.
        /// </summary>
        private static ulong[] cprptMessageId = new ulong[3];

        /// <summary>
        /// Attachment handle array.
        /// </summary>
        private static uint[] cprptAttachmentHandle = new uint[3];

        /// <summary>
        /// Attachment id array.
        /// </summary>
        private static uint[] cprptAttachmentId = new uint[3];

        /// <summary>
        /// Id of root folder.
        /// </summary>
        private static ulong rootFolderId;

        /// <summary>
        /// The folders to be deleted.
        /// </summary>
        private static Dictionary<uint, ulong> deletedFolders = new Dictionary<uint, ulong>();

        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// The OxcropsClient instance (used for the second session).
        /// </summary>
        private OxcropsClient oxcropsClientSession2;

        /// <summary>
        /// Variable stored logon handle.
        /// </summary>
        private uint cprptLogonHandle;

        /// <summary>
        /// The handle which is operating.
        /// </summary>
        private uint cprptCurrentHandle;

        /// <summary>
        /// The first handle used in ROPs.
        /// </summary>
        private uint cprptFirstHandle;

        /// <summary>
        /// The second handle used in ROPs.
        /// </summary>
        private uint cprptSecondHandle;

        /// <summary>
        /// The first object handle used in ROPs.
        /// </summary>
        private uint cprptFirstObjectHandle;

        /// <summary>
        /// The second object handle used in ROPs.
        /// </summary>
        private uint cprptSecondObjectHandle;

        /// <summary>
        /// The current object to be operated
        /// </summary>
        private ObjectToOperate cprptCurrentObj;

        /// <summary>
        /// Handle of inbox folder.
        /// </summary>
        private uint cprptInputBoxHandle;

        /// <summary>
        /// Folder id of inbox.
        /// </summary>
        private ulong cprptInputBoxFolderId;

        /// <summary>
        /// Handle of the root folder of test folder.
        /// </summary>
        private uint cprptTestRootFolderHandle;

        /// <summary>
        /// Folder name of root folder.
        /// </summary>
        private string cprptTestFolderName;

        /// <summary>
        /// The object type which is now operating.
        /// </summary>
        private ServerObjectType cprptCurrentType;

        /// <summary>
        /// All folder Ids returned by [RopLogon] operation.
        /// </summary>
        private ulong[] folderIds;

        /// <summary>
        /// The Logon handle.
        /// </summary>
        private uint logonHandle;

        /// <summary>
        /// The handle of public folder which is created by SUT setup script.
        /// </summary>
        private uint existingPublicFolderHandle;

        /// <summary>
        /// Record the logon type
        /// </summary>
        private bool isPublicFolderLogonType = false;
        
        /// <summary>
        /// Used for two sessions.
        /// </summary>
        private Dictionary<string, TaggedPropertyValue> tagPropertyValuesToVerification = new Dictionary<string, TaggedPropertyValue>();

        /// <summary>
        /// Indicate whether needs to do the cleanup work
        /// </summary>
        private bool needDoCleanup = true;

        #region HandleIndex
        /// <summary>
        /// Index of Handle in request buffer or response buffer
        /// </summary>
        public enum HandleIndex : byte
        {
            /// <summary>
            /// Index is 0
            /// </summary>
            FirstIndex = 0,

            /// <summary>
            /// Index is 1
            /// </summary>
            SecondIndex = 1
        }
        #endregion

        #region CodePageId
        /// <summary>
        /// 2 bytes specifying the code page that the non-Unicode representation of the strings on this Message object are to be encoded;
        /// </summary>
        public enum CodePageId : ushort
        {
            /// <summary>
            /// Value of "0x0FFF" means that the code page of the Logon object is used.
            /// </summary>
            SameAsLogonObject = 0x0FFF
        }
        #endregion

        #endregion

        #region IAdapter members
        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXCPRPT";
            if (!commonConfigImported)
            {
                Common.MergeConfiguration(testSite);

                commonConfigImported = true;
            }

            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
            this.oxcropsClientSession2 = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
        }

        /// <summary>
        /// Reset the adapter.
        /// Adapter will disconnect with server via calling RPC method. 
        /// At the same time, the ITestSite context will be reseated.
        /// </summary>
        public override void Reset()
        {
            this.CleanUpMethod();

            // rpcAdapter would disconnect the RPC link to server.
            bool isDisconnected = this.RpcDisconnect();
            this.Site.Assert.IsTrue(isDisconnected, "RPC disconnect should be executed successfully.");

            // The ITestSite context which was initialized in Initialize(ITestSite testSite) would be reseated.
            base.Reset();
        }

        #endregion

        #region Interface methods
        /// <summary>
        /// This method is used to initialize the test environment for private mailbox.
        /// </summary>
        public void InitializeMailBox()
        {
            this.InitializeCommonProperties();
            this.InitHandle(this.Site);
            this.cprptTestFolderName = RootFolderName;
            bool isConnected = this.RpcConnect(LogonType.Mailbox);
            this.Site.Assert.IsTrue(isConnected, "Rpc connect method should be executed successfully.");
            this.isPublicFolderLogonType = false;
            this.GetHandles();
        }

        /// <summary>
        /// This method is used to initialize the test environment for public folders.
        /// </summary>
        public void InitializePublicFolder()
        {
            this.InitializeCommonProperties();
            this.InitHandle(this.Site, LogonType.PublicFolder);
            this.isPublicFolderLogonType = true;
        }

        /// <summary>
        /// This method is used to get object for different object types.
        /// </summary>
        /// <param name="objType">Specifies the object type.</param>
        /// <param name="objToOperate">Specifies which objects to operate.</param>
        public void GetObject(ServerObjectType objType, ObjectToOperate objToOperate)
        {
            this.cprptCurrentObj = objToOperate;

            if (objType == ServerObjectType.Logon)
            {
                this.cprptFirstHandle = this.cprptFirstObjectHandle = this.cprptCurrentHandle = this.cprptLogonHandle;
            }

            if (objType == ServerObjectType.Folder)
            {
                switch (objToOperate)
                {
                    case ObjectToOperate.FirstObject:
                        this.cprptFirstHandle = this.cprptFirstObjectHandle = this.cprptCurrentHandle = cprptFolderHandle[0];
                        break;
                    case ObjectToOperate.SecondObject:
                        this.cprptSecondHandle = this.cprptSecondObjectHandle = cprptFolderHandle[1];
                        break;
                    case ObjectToOperate.ThirdObject:
                        this.cprptCurrentHandle = cprptFolderHandle[2];
                        break;
                    case ObjectToOperate.FourthObject:
                        this.cprptCurrentHandle = cprptFolderHandle[3];
                        break;
                    case ObjectToOperate.FifthObject:
                        this.cprptCurrentHandle = cprptFolderHandle[1];
                        break;
                    default:
                        Site.Assert.Fail("Invalid ObjectToOperate enum value {0}.", objToOperate);
                        break;
                }
            }

            if (objType == ServerObjectType.Message)
            {
                switch (objToOperate)
                {
                    case ObjectToOperate.FirstObject:
                        this.cprptFirstHandle = this.cprptFirstObjectHandle = this.cprptCurrentHandle = cprptMessageHandle[0];
                        break;
                    case ObjectToOperate.SecondObject:
                        this.cprptSecondHandle = this.cprptSecondObjectHandle = cprptMessageHandle[1];
                        break;
                    case ObjectToOperate.ThirdObject:
                        this.cprptCurrentHandle = cprptMessageHandle[2];
                        break;
                    case ObjectToOperate.FourthObject:
                        this.cprptCurrentHandle = cprptMessageHandle[3];
                        break;
                    default:
                        Site.Assert.Fail("Invalid ObjectToOperate enum value {0}.", objToOperate);
                        break;
                }
            }

            if (objType == ServerObjectType.Attachment)
            {
                switch (objToOperate)
                {
                    case ObjectToOperate.FirstObject:
                        this.cprptFirstHandle = this.cprptFirstObjectHandle = this.cprptCurrentHandle = cprptAttachmentHandle[0];
                        break;
                    case ObjectToOperate.SecondObject:
                        this.cprptSecondHandle = this.cprptSecondObjectHandle = cprptAttachmentHandle[1];
                        break;
                    case ObjectToOperate.ThirdObject:
                        this.cprptCurrentHandle = cprptAttachmentHandle[2];
                        break;
                    case ObjectToOperate.FourthObject:
                        this.cprptCurrentHandle = cprptAttachmentHandle[3];
                        break;
                    default:
                        Site.Assert.Fail("Invalid ObjectToOperate enum value {0}.", objToOperate);
                        break;
                }
            }

            #region Set property

            TaggedPropertyValue[] taggedPropertyValueArray = new TaggedPropertyValue[1];
            RopSetPropertiesResponse ropSetPropertiesResponse;
            if (objType == ServerObjectType.Attachment)
            {
                // Add this property to attachment1.
                this.cprptCurrentType = ServerObjectType.Attachment;
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
                ropSetPropertiesResponse = this.RopSetProperties(cprptAttachmentHandle[0], taggedPropertyValueArray, true);
                Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "The properties should be set for the first object successfully.");
                RopSaveChangesAttachmentResponse ropSaveAttachmentResponse = this.RopSaveChangesAttachment(cprptAttachmentHandle[0], true);
                Site.Assert.AreEqual<uint>(0, ropSaveAttachmentResponse.ReturnValue, "The changes on Attachment should be saved successfully.");
                this.RopSaveChangesMessage(cprptMessageHandle[0], true);

                // Add this property to attachment2.
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.SecondObject);
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
                this.RopSetProperties(cprptAttachmentHandle[1], taggedPropertyValueArray, true);
                this.RopSaveChangesAttachment(cprptAttachmentHandle[1], true);
                this.RopSaveChangesMessage(cprptMessageHandle[0], true);
            }

            if (objType == ServerObjectType.Message)
            {
                // Add this property to message1.
                this.cprptCurrentType = ServerObjectType.Message;
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
                this.RopSetProperties(cprptMessageHandle[0], taggedPropertyValueArray, true);
                this.RopSaveChangesMessage(cprptMessageHandle[0], true);

                // Add this property to message2.
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.SecondObject);
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
                this.RopSetProperties(cprptMessageHandle[1], taggedPropertyValueArray, true);
                this.RopSaveChangesMessage(cprptMessageHandle[1], true);
            }

            if (objType == ServerObjectType.Folder)
            {
                // Add this property to folder1.
                this.cprptCurrentType = ServerObjectType.Folder;
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
                this.RopSetProperties(cprptFolderHandle[0], taggedPropertyValueArray, true);

                // Add this property to folder2.
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.SecondObject);
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
                this.RopSetProperties(cprptFolderHandle[1], taggedPropertyValueArray, true);

                // Set PTagRulesData property for folder1 and folder2
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.ThirdObject);
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
                this.RopSetProperties(cprptFolderHandle[0], taggedPropertyValueArray, true);
                this.RopSetProperties(cprptFolderHandle[1], taggedPropertyValueArray, true);
            }

            if (objType == ServerObjectType.Logon)
            {
                // Add this property to logon.
                this.cprptCurrentType = ServerObjectType.Logon;
                taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
                this.RopSetProperties(this.cprptLogonHandle, taggedPropertyValueArray, true);
            }

            // Set current type.
            this.cprptCurrentType = objType;

            #endregion
        }

        /// <summary>
        /// The method is used to query an object for all the named properties. 
        /// </summary>
        /// <param name="queryFlags">Specifies QueryFlags parameter in request.</param>
        /// <param name="hasGuid">Indicates whether HasGuid is zero, 
        /// If the HasGUID field is non-zero then the PropertyGUID field MUST be included in the request. 
        /// If no PropertyGUID is specified, then properties from any GUID MUST be returned in the results.</param>
        /// <param name="isKind0x01Returned">True if the named properties of the response with the Kind field 
        /// ([MS-OXCDATA] section 2.6.1) set to 0x1 was returned.</param>
        /// <param name="isKind0x00Returned">True if the named properties of the response with the Kind field 
        /// ([MS-OXCDATA] section 2.6.1) set to 0x0 was returned.</param>
        /// <param name="isNamedPropertyGuidReturned">True if the named properties  with a GUID field ([MS-OXCDATA] 
        /// section 2.6.1) value that does not match the value of the PropertyGUID field was returned.</param>
        public void RopQueryNamedPropertiesMethod(
            QueryFlags queryFlags,
            bool hasGuid,
            out bool isKind0x01Returned,
            out bool isKind0x00Returned,
            out bool isNamedPropertyGuidReturned)
        {
            #region Initialize parameters property
            // The value of hasGuid in RopQueryNameProperties method.
            byte hasGuidRequest = 0;

            // The value of queryFlags in RopQueryNameProperties method.
            byte queryFlagRequest = Convert.ToByte(queryFlags);

            // The value of propertyGuid in RopQuerynameProperties method.
            byte[] propertyGuid = null;

            // Set default values for three parameters.
            isKind0x01Returned = false;
            isKind0x00Returned = false;
            isNamedPropertyGuidReturned = false;

            // When name property has GUID, set parameters.
            if (hasGuid)
            {
                hasGuidRequest = Convert.ToByte(true);
                PropertyName[] ptyName = new PropertyName[1];
                ptyName[0] = this.GetPropertyName((byte)Kind.NameField);
                propertyGuid = ptyName[0].Guid;

                this.RopGetPropertyIdsFromNames(this.cprptCurrentHandle, (byte)GetPropertyIdsFromNamesFlags.Create, ptyName, true);
                if (this.cprptCurrentType == ServerObjectType.Message)
                {
                    this.RopSaveChangesMessage(this.cprptCurrentHandle, true);
                }
                else if (this.cprptCurrentType == ServerObjectType.Attachment)
                {
                    this.RopSaveChangesAttachment(this.cprptCurrentHandle, true);
                }
            }
            #endregion
            // Call RopQueryNamedProperties.
            RopQueryNamedPropertiesResponse queryNamePtyRes = this.RopQueryNamedProperties(this.cprptCurrentHandle, queryFlagRequest, hasGuidRequest, propertyGuid);

            #region Verify isKind0x00Return and isKind0x00Return
            // Verify if the kind of NamedProperties in RopQueryNamedPropertiesResponse is 0x01.
            for (int i = 0; i < queryNamePtyRes.IdCount; i++)
            {
                if (queryNamePtyRes.PropertyNames[i].Kind == (byte)Kind.NameField)
                {
                    isKind0x01Returned = true;
                    break;
                }
            }

            // Verify if the kind of NamedProperties in RopQueryNamedPropertiesResponse is 0x00.
            for (int i = 0; i < queryNamePtyRes.IdCount; i++)
            {
                if (queryNamePtyRes.PropertyNames[i].Kind == (byte)Kind.LidField)
                {
                    isKind0x00Returned = true;
                    break;
                }
            }
            #endregion

            #region Verify isNamedPropertyGuidReturn
            // If the PropertyGUID field presents, named properties with a GUID field value that does not match the value of the PropertyGUID field MUST NOT be returned.
            for (int i = 0; i < queryNamePtyRes.IdCount; i++)
            {
                if (hasGuid && Common.CompareByteArray(queryNamePtyRes.PropertyNames[i].Guid, propertyGuid))
                {
                    isNamedPropertyGuidReturned = true;
                }
                else
                {
                    isNamedPropertyGuidReturned = false;
                    break;
                }
            }
            #endregion

            // Verify RopQueryNamedProperties.
            this.VerifyRopQueryNamedProperties(queryFlags, propertyGuid, queryNamePtyRes, this.cprptCurrentType, true);
        }

        /// <summary>
        /// This method is used to query for and return all of the property tags and values of properties that have been set. 
        /// </summary>
        /// <param name="isPropertySizeLimitZero">Indicates whether PropertySizeLimit parameter is zero.</param>
        /// <param name="isPropertyLagerThanLimit">Indicates whether request properties larger than limit
        /// When PropertySizeLimit is non-zero, Indicates whether request properties larger than PropertySizeLimit
        /// When PropertySizeLimit is zero, Indicates whether request properties larger than size of response.</param>
        /// <param name="isUnicode">Indicates whether the requested property is encoded in Unicode format in response buffer</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates whether returned value contains NotEnoughMemory error when the request properties are too large.</param>
        public void RopGetPropertiesAllMethod(bool isPropertySizeLimitZero, bool isPropertyLagerThanLimit, bool isUnicode, out bool isValueContainsNotEnoughMemory)
        {
            #region Initialize parameters
            // Set default value for isValueContainsNotEnoughMemory parameter.
            isValueContainsNotEnoughMemory = false;

            // Set value for wantUnidcode field.
            ushort wantUnicode = 0;
            if (isUnicode)
            {
                wantUnicode = 1;
            }

            // Set value for propertySizeLimit field.
            ushort propertySizeLimit = 0;

            // Set value for propertySizeLimit field.
            if (!isPropertySizeLimitZero)
            {
                propertySizeLimit = PropertySizeLimit;
            }

            // Set value for tagPtyValues array.
            TaggedPropertyValue[] tagPtyValues = new TaggedPropertyValue[1];
            tagPtyValues[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);

            // Set value as the larger one than the limited value.
            if (isPropertyLagerThanLimit)
            {
                List<byte> bytes = new List<byte>();
                foreach (byte b in tagPtyValues[0].Value)
                {
                    bytes.Add(b);
                }

                while (bytes.Count <= propertySizeLimit)
                {
                    bytes.Add(ConstValues.PropertyBytesFilling);
                }

                bytes.Add(ConstValues.PropertyBytesFilling);
                tagPtyValues[0].Value = bytes.ToArray();
            }

            if (tagPtyValues[0].PropertyTag.PropertyType == 0x0102)
            {
                tagPtyValues[0].Value = Common.AddInt16LengthBeforeBinaryArray(tagPtyValues[0].Value);
            }

            // Set Property before executing the query operation.
            this.RopSetProperties(this.cprptCurrentHandle, tagPtyValues, true);
            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                this.RopSaveChangesMessage(this.cprptCurrentHandle, true);
            }
            else if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                this.RopSaveChangesAttachment(this.cprptCurrentHandle, true);
            }
            #endregion

            RopGetPropertiesAllResponse getPropertiesAllResponse = this.RopGetPropertiesAll(this.cprptCurrentHandle, propertySizeLimit, wantUnicode);

            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                foreach (TaggedPropertyValue value in getPropertiesAllResponse.PropertyValues)
                {
                    byte[] notEnoughMemory = BitConverter.GetBytes((uint)CPRPTErrorCode.NotEnoughMemory);
                    if (Common.CompareByteArray(notEnoughMemory, value.Value))
                    {
                        isValueContainsNotEnoughMemory = true;
                    }
                }
            }

            if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                if (getPropertiesAllResponse.ReturnValue == (uint)CPRPTErrorCode.NotEnoughMemory)
                {
                    isValueContainsNotEnoughMemory = true;
                }
            }

            this.VerifyRopGetPropertiesAll(getPropertiesAllResponse, this.cprptCurrentType, isPropertySizeLimitZero, isPropertyLagerThanLimit, isUnicode, isValueContainsNotEnoughMemory);
        }

        /// <summary>
        /// This method is used to query for and returns all of the property tags for properties that have been set on an object.
        /// </summary>
        public void RopGetPropertiesListMethod()
        {
            // Call RopGetPropertiesList ROP.
            RopGetPropertiesListResponse getPropertiesListResponse = this.RopGetPropertiesList(this.cprptCurrentHandle);

            #region Verify response
            this.VerifyRopGetPropertiesList(getPropertiesListResponse, this.cprptCurrentType);

            this.Site.Assert.AreEqual((uint)CPRPTErrorCode.None, getPropertiesListResponse.ReturnValue, string.Format("Logon Failed! Error: 0x{0:X8}", getPropertiesListResponse.ReturnValue));
            #endregion
        }

        /// <summary>
        ///  This method is used to query for and return the values of properties specified in the PropertyTags field. 
        /// </summary>
        /// <param name="isTestOrder">Indicates whether to test returned PropertyNames order.</param>
        /// <param name="isPropertySizeLimitZero">Indicates whether PropertySizeLimit parameter is zero.</param>
        /// <param name="isPropertyLagerThanLimit">Indicates whether request properties larger than limit
        /// When PropertySizeLimit is non-zero, Indicates whether request properties larger than PropertySizeLimit
        /// When PropertySizeLimit is zero, Indicates whether request properties larger than size of response.</param>
        /// <param name="isValueContainsNotEnoughMemory">Indicates whether returned value contains NotEnoughMemory error when the request properties are too large.</param>
        public void RopGetPropertiesSpecificMethod(bool isTestOrder, bool isPropertySizeLimitZero, bool isPropertyLagerThanLimit, out bool isValueContainsNotEnoughMemory)
        {
            // Set default value for this parameter.
            isValueContainsNotEnoughMemory = false;

            // Set value for this parameter to specify whether the order in response is the same as the one in request.
            bool isOrder = true;

            #region Initialize parameters
            // Set value for wantUnicode in RopGetProertiesSpecific ROP.
            ushort wantUnicode = 0;

            // Set value for propertySizeLimit in RopGetProertiesSpecific ROP.
            ushort propertySizeLimit = 0;
            if (!isPropertySizeLimitZero)
            {
                propertySizeLimit = PropertySizeLimit;
            }

            // Set value for tagPtyValues in RopGetPropertiesSpecific.
            TaggedPropertyValue[] tagPtyValues = new TaggedPropertyValue[1];
            tagPtyValues[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            if (isPropertyLagerThanLimit)
            {
                List<byte> bytes = new List<byte>();
                foreach (byte b in tagPtyValues[0].Value)
                {
                    bytes.Add(b);
                }

                while (bytes.Count <= propertySizeLimit)
                {
                    bytes.Add(bytes[0]);
                }

                bytes.Add(bytes[0]);
                tagPtyValues[0].Value = bytes.ToArray();
            }

            if (tagPtyValues[0].PropertyTag.PropertyType == 0x0102)
            {
                tagPtyValues[0].Value = Common.AddInt16LengthBeforeBinaryArray(tagPtyValues[0].Value);
            }

            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                this.RopSaveChangesMessage(this.cprptCurrentHandle, true);
            }
            else if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                this.RopSaveChangesAttachment(this.cprptCurrentHandle, true);
            }

            // Set Property before executing the query operation.
            this.RopSetProperties(this.cprptCurrentHandle, tagPtyValues, true);

            PropertyTag[] ptyTags;
            TaggedPropertyValue[] tagPtyValues1 = new TaggedPropertyValue[1];
            if (isTestOrder)
            {
                // Set value for the second property.
                tagPtyValues1[0] = this.GetTaggedPropertyTag(ObjectToOperate.SecondObject);
                tagPtyValues1[0].Value = Common.AddInt16LengthBeforeBinaryArray(tagPtyValues1[0].Value);

                this.RopSetProperties(this.cprptCurrentHandle, tagPtyValues1, true);

                // Set values for property tags of properties for which the client is requesting. 
                ptyTags = new PropertyTag[2];
                ptyTags[0] = tagPtyValues[0].PropertyTag;
                ptyTags[1] = tagPtyValues1[0].PropertyTag;
            }
            else
            {
                ptyTags = new PropertyTag[1];
                ptyTags[0] = tagPtyValues[0].PropertyTag;
            }
            #endregion

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, propertySizeLimit, wantUnicode, ptyTags);

            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                foreach (PropertyValue value in getPropertiesSpecificResponse.RowData.PropertyValues)
                {
                    byte[] notEnoughMemory = BitConverter.GetBytes((uint)CPRPTErrorCode.NotEnoughMemory);
                    if (Common.CompareByteArray(notEnoughMemory, value.Value))
                    {
                        isValueContainsNotEnoughMemory = true;
                    }
                }
            }

            if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                if (getPropertiesSpecificResponse.ReturnValue == (uint)CPRPTErrorCode.NotEnoughMemory)
                {
                    isValueContainsNotEnoughMemory = true;
                }
            }

            #region Verify response and check order
            if (getPropertiesSpecificResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                if (isTestOrder)
                {
                    byte[] notEnoughMemory = BitConverter.GetBytes((uint)CPRPTErrorCode.NotEnoughMemory);
                    if (Common.CompareByteArray(notEnoughMemory, getPropertiesSpecificResponse.RowData.PropertyValues[0].Value) || Common.CompareByteArray(notEnoughMemory, getPropertiesSpecificResponse.RowData.PropertyValues[0].Value))
                    {
                        isOrder = true;
                    }
                    else
                    {
                        byte[] value0 = getPropertiesSpecificResponse.RowData.PropertyValues[0].Value;
                        byte[] value1 = getPropertiesSpecificResponse.RowData.PropertyValues[1].Value;
                        bool isValue0 = Common.CompareByteArray(value0, tagPtyValues[0].Value);
                        bool isValue1 = Common.CompareByteArray(value1, tagPtyValues1[0].Value);
                        if (isValue0 && isValue1)
                        {
                            isOrder = true;
                        }
                        else
                        {
                            isOrder = false;
                        }
                    }
                }
            }

            this.VerifyRopGetPropertiesSpecific(getPropertiesSpecificResponse, this.cprptCurrentType, isOrder, isPropertySizeLimitZero, isPropertyLagerThanLimit, isValueContainsNotEnoughMemory);
            #endregion
        }

        /// <summary>
        ///  This method is used to query for and return the values of properties specified in the PropertyTags field, which is related with unicode format.
        /// </summary>
        /// <param name="isUnicode">Indicates whether the requested property is encoded in Unicode format in response buffer</param>
        public void RopGetPropertiesSpecificForWantUnicode(bool isUnicode)
        {
            // The wantUnicode (0) means the server MUST return string properties that are requested without a specified type (PtypUnspecified) in MBCS format. 
            ushort wantUnicode = 0;

            // An integer that specifies the maximum size allowed for a property value. 
            ushort propertySizeLimit = 0;

            if (isUnicode)
            {
                wantUnicode = 1;
            }

            TaggedPropertyValue[] tagPtyValues = new TaggedPropertyValue[1];

            // Set value for a property, which is PtypString8 type
            tagPtyValues[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            tagPtyValues[0].PropertyTag.PropertyId++;
            tagPtyValues[0].PropertyTag.PropertyType = (ushort)0x001e;
            tagPtyValues[0].Value = Encoding.ASCII.GetBytes("test for unicode\0");

            this.RopSetProperties(this.cprptCurrentHandle, tagPtyValues, true);
            if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                this.RopSaveChangesAttachment(this.cprptCurrentHandle, true);
            }

            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                this.RopSaveChangesMessage(this.cprptCurrentHandle, true);
            }

            PropertyTag[] tags = new PropertyTag[1];
            tags[0].PropertyId = tagPtyValues[0].PropertyTag.PropertyId;
            tags[0].PropertyType = (ushort)0x0000;

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, propertySizeLimit, wantUnicode, tags);

            this.VerifyRopGetPropertiesSpecificForWantUnicode(getPropertiesSpecificResponse, isUnicode);
        }

        /// <summary>
        /// This method is used to query for and return the values of properties specified in the PropertyTags field, which is related with tagged properties. 
        /// </summary>
        public void RopGetPropertiesSpecificForTaggedProperties()
        {
            TaggedPropertyValue[] tagPtyValues = new TaggedPropertyValue[2];
            tagPtyValues[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            tagPtyValues[0].PropertyTag.PropertyId++;
            tagPtyValues[0].PropertyTag.PropertyType = (ushort)0x001e;
            tagPtyValues[0].Value = Encoding.ASCII.GetBytes("A property set by the client.\0");

            tagPtyValues[1] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            tagPtyValues[1].PropertyTag.PropertyId = 0x3016;
            tagPtyValues[1].PropertyTag.PropertyType = (ushort)0x000B;
            tagPtyValues[1].Value = BitConverter.GetBytes(true);

            // Call RopSetProperties operation to set a property.
            RopSetPropertiesResponse ropSetPropertiesResponse = this.RopSetProperties(this.cprptCurrentHandle, tagPtyValues, true);
            Site.Assert.AreEqual<uint>(0, ropSetPropertiesResponse.ReturnValue, "The properties should be set successfully.");

            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                RopSaveChangesMessageResponse ropSaveChangesMessageResponse = this.RopSaveChangesMessage(this.cprptCurrentHandle, true);
                Site.Assert.AreEqual<uint>(0, ropSaveChangesMessageResponse.ReturnValue, "The properties should be saved on the message object successfully.");
            }

            // Define an array of PropertyTags to retrieve.
            PropertyTag[] tags = new PropertyTag[3];
            tags[0].PropertyId = tagPtyValues[0].PropertyTag.PropertyId;
            tags[0].PropertyType = (ushort)PropertyTypeName.PtypString8;

            // PidTagRecordKey: ReadOnly server property.
            tags[1].PropertyId = (ushort)TaggedPropertyName.PidTagRecordKey;
            tags[1].PropertyType = (ushort)PropertyTypeName.PtypBinary;

            // PidTagConversationId: computed property.
            tags[2].PropertyId = (ushort)TaggedPropertyName.PidTagConversationId;
            tags[2].PropertyType = (ushort)PropertyTypeName.PtypBinary;

            // Call RopGetPropertiesSpecific to retrieve the targeted properties.
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, 0, 0, tags);

            this.VerifyRopGetPropertiesSpecificForTaggedProperties(getPropertiesSpecificResponse);
        }

        /// <summary>
        /// This method is used to map abstract, client-defined named properties to concrete 16-bit property IDs. 
        /// </summary>
        /// <param name="isTestOrder">Indicates whether to test returned PropertyNames order.</param>
        /// <param name="isCreateFlagSet">Indicates whether the "Create" Flags in request parameter is set.</param>
        /// <param name="isPropertyNameExisting">Indicates whether PropertyName is existing in object mapping.</param>
        /// <param name="specialPropertyName">Specifies PropertyName of request parameter</param>
        /// <param name="isCreatedEntryReturned">If Create Flags is set: If set, indicates that the server MUST create new
        /// entries for any name parameters that are not found in the existing mapping set, and return existing entries for any
        /// name parameters that are found in the existing mapping set.</param>
        /// <param name="error">Specifies the ErrorCode when server reached limit.</param>
        public void RopGetPropertyIdsFromNamesMethod(
            bool isTestOrder,
            bool isCreateFlagSet,
            bool isPropertyNameExisting,
            SpecificPropertyName specialPropertyName,
            out bool isCreatedEntryReturned,
            out CPRPTErrorCode error)
        {
            #region Initialize parameters
            // Specify whether the order in response is the same as that in request.
            bool isOrder = true;

            // Specify whether the property names and ids are mapped.
            bool isNamesAndIdsMapping = false;

            // Set default value for the out parameter.
            isCreatedEntryReturned = false;
            error = CPRPTErrorCode.None;

            // Set value for flags in RopGetPropertyIdsFromNames.
            byte flags = (byte)GetPropertyIdsFromNamesFlags.None;
            if (isCreateFlagSet)
            {
                flags = (byte)GetPropertyIdsFromNamesFlags.Create;
            }

            PropertyName[] propertyNames = null;

            if (isTestOrder)
            {
                propertyNames = new PropertyName[ConstValues.TestPropertyCount];
            }
            else
            {
                propertyNames = new PropertyName[1];
            }

            switch (specialPropertyName)
            {
                case SpecificPropertyName.Kind0x01:
                case SpecificPropertyName.NoConstraint:
                    propertyNames[0] = this.GetPropertyName((byte)Kind.NameField);
                    break;
                case SpecificPropertyName.PS_MAPIAndKind0x01:
                    propertyNames[0] = this.GetPropertyName((byte)Kind.NameField);
                    propertyNames[0].Guid = (new Guid(PSMAPI)).ToByteArray();
                    break;
                case SpecificPropertyName.PS_MAPIAndKind0x00:
                    propertyNames[0] = this.GetPropertyName((byte)Kind.LidField);
                    propertyNames[0].Guid = (new Guid(PSMAPI)).ToByteArray();
                    break;
                default:
                    Site.Assert.Fail("Invalid SpecificPropertyName enum value {0}.", specialPropertyName);
                    break;
            }

            if (!isPropertyNameExisting && !isCreateFlagSet)
            {
                propertyNames[0].Kind = (byte)Kind.NameField;
                propertyNames[0].Name = Encoding.Unicode.GetBytes(PropertyNameNotMapped + "\0");
                propertyNames[0].NameSize = (byte)propertyNames[0].Name.Length;
                propertyNames[0].Guid = (new Guid(PropertyNameNotMappedGuid)).ToByteArray();
            }

            // The existing property ID.
            ushort existId = 0;

            // If the property exists, create it.
            if (isPropertyNameExisting)
            {
                PropertyName[] propertyNameTmp = new PropertyName[1];
                propertyNameTmp[0] = propertyNames[0];
                RopGetPropertyIdsFromNamesResponse ptyNameExistingRes = this.RopGetPropertyIdsFromNames(this.cprptCurrentHandle, (byte)GetPropertyIdsFromNamesFlags.Create, propertyNameTmp, true);
                existId = ptyNameExistingRes.PropertyIds[0].ID;
            }

            // The second property ID.
            ushort secondId = 0;

            // The third property ID.
            ushort thirdId = 0;

            // If the PropertyNames order need to verify, create the second and the third properties.
            if (isTestOrder)
            {
                // Add second property.
                PropertyName[] propertyNameTmp = new PropertyName[1];
                propertyNameTmp[0] = this.GetPropertyName((byte)Kind.NameField);
                propertyNameTmp[0].Name[0]++;
                propertyNameTmp[0].Guid[0]++;
                propertyNames[1] = propertyNameTmp[0];
                RopGetPropertyIdsFromNamesResponse ptyNameExistingRes = this.RopGetPropertyIdsFromNames(this.cprptCurrentHandle, (byte)GetPropertyIdsFromNamesFlags.Create, propertyNameTmp, true);
                secondId = ptyNameExistingRes.PropertyIds[0].ID;

                // Add second property.
                propertyNameTmp[0] = this.GetPropertyName((byte)Kind.NameField);
                propertyNameTmp[0].Name[0]--;
                propertyNameTmp[0].Guid[0]--;
                propertyNames[2] = propertyNameTmp[0];
                ptyNameExistingRes = this.RopGetPropertyIdsFromNames(this.cprptCurrentHandle, (byte)GetPropertyIdsFromNamesFlags.Create, propertyNameTmp, true);
                thirdId = ptyNameExistingRes.PropertyIds[0].ID;
            }
            #endregion

            RopGetPropertyIdsFromNamesResponse getPropertyIdsFromNamesResponse = this.RopGetPropertyIdsFromNames(this.cprptCurrentHandle, flags, propertyNames, false);

            #region When property exists, check if names and ids are mapping
            if (isPropertyNameExisting)
            {
                RopGetNamesFromPropertyIdsResponse getPropertyNamesFromPropertyIdsResponse = this.RopGetNamesFromPropertyIds(this.cprptCurrentHandle, getPropertyIdsFromNamesResponse.PropertyIds);

                // Verify whether property name and id is mapping.
                for (int i = 0; i < getPropertyNamesFromPropertyIdsResponse.PropertyNameCount; i++)
                {
                    if (getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].Kind == (byte)Kind.LidField)
                    {
                        if ((getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].Guid.ToString() == propertyNames[i].Guid.ToString()) &&
                            (getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].Kind == propertyNames[i].Kind) &&
                            (getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].LID == propertyNames[i].LID))
                        {
                            isNamesAndIdsMapping = true;
                        }
                        else
                        {
                            isNamesAndIdsMapping = false;
                            break;
                        }
                    }
                    else if (getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].Kind == (byte)Kind.NameField)
                    {
                        if ((getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].Guid.ToString() == propertyNames[i].Guid.ToString()) &&
                            (getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].Kind == propertyNames[i].Kind) &&
                            (getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].NameSize == propertyNames[i].NameSize) &&
                            (getPropertyNamesFromPropertyIdsResponse.PropertyNames[i].Name.ToString() == propertyNames[i].Name.ToString()))
                        {
                            isNamesAndIdsMapping = true;
                        }
                        else
                        {
                            isNamesAndIdsMapping = false;
                            break;
                        }
                    }
                }
            }
            #endregion

            #region Test order
            if (getPropertyIdsFromNamesResponse.PropertyIds != null && propertyNames.Length == ConstValues.TestPropertyCount && getPropertyIdsFromNamesResponse.PropertyIds.Length == ConstValues.TestPropertyCount)
            {
                if (getPropertyIdsFromNamesResponse.PropertyIds[1].ID != secondId || getPropertyIdsFromNamesResponse.PropertyIds[2].ID != thirdId)
                {
                    isOrder = false;
                }
            }
            #endregion

            this.VerifyRopGetPropertyIdsFromNames((ushort)propertyNames.Length, propertyNames, flags, getPropertyIdsFromNamesResponse, isOrder, this.cprptCurrentType, isPropertyNameExisting, isNamesAndIdsMapping);

            #region Check error and isCreateEnrtyReturned

            switch ((CPRPTErrorCode)getPropertyIdsFromNamesResponse.ReturnValue)
            {
                case CPRPTErrorCode.None:

                    // Verify whether the created entry is returned.
                    if (isCreateFlagSet)
                    {
                        if (isPropertyNameExisting)
                        {
                            isCreatedEntryReturned = false;
                        }
                        else if (getPropertyIdsFromNamesResponse.PropertyIds[0].ID != existId)
                        {
                            isCreatedEntryReturned = true;
                        }
                    }

                    break;
                case CPRPTErrorCode.ecWarnWithErrors:
                    error = CPRPTErrorCode.ecWarnWithErrors;
                    break;
                case CPRPTErrorCode.NotEnoughMemory:
                    error = CPRPTErrorCode.NotEnoughMemory;
                    break;
                default:
                    Site.Assert.Fail("Unexpected RopGetPropertyIdsFromNames error code.Error: 0x{0:X8}", getPropertyIdsFromNamesResponse.ReturnValue);
                    break;
            }

            #endregion
        }

        /// <summary>
        /// This method is used to map concrete property IDs to abstract, client-defined named properties. 
        /// </summary>
        /// <param name="propertyIdType">Specifies different PropertyId type.</param>
        public void RopGetNamesFromPropertyIdsMethod(PropertyIdType propertyIdType)
        {
            // The property IDs needs to get in RopGetNamesFromPropertyIds ROP request.
            PropertyId[] propertyIds = new PropertyId[1];
            #region Set and save property
            if (propertyIdType == PropertyIdType.HaveAssociatedName || propertyIdType == PropertyIdType.LessThan0x8000)
            {
                TaggedPropertyValue[] tagPropertyValue = new TaggedPropertyValue[1];
                tagPropertyValue[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
                if (tagPropertyValue[0].PropertyTag.PropertyType == 0x0102)
                {
                    tagPropertyValue[0].Value = Common.AddInt16LengthBeforeBinaryArray(tagPropertyValue[0].Value);
                }

                this.RopSetProperties(this.cprptFirstHandle, tagPropertyValue, true);
                if (this.cprptCurrentType == ServerObjectType.Message)
                {
                    this.RopSaveChangesMessage(this.cprptFirstHandle, true);
                }
                else if (this.cprptCurrentType == ServerObjectType.Attachment)
                {
                    this.RopSaveChangesAttachment(this.cprptFirstHandle, true);
                }
            }
            #endregion

            RopGetNamesFromPropertyIdsResponse getNamesFromPropertyIdsResponse = this.RopGetNamesFromPropertyIds(this.cprptCurrentHandle, propertyIds);

            #region Verify response
            this.VerifyRopGetNamesFromPropertyIds((ushort)propertyIds.Length, propertyIds, getNamesFromPropertyIdsResponse, this.cprptCurrentType);

            this.Site.Assert.AreEqual((uint)CPRPTErrorCode.None, getNamesFromPropertyIdsResponse.ReturnValue, string.Format("RopGetNamesFromPropertyIds Failed! Error: 0x{0:X8}", getNamesFromPropertyIdsResponse.ReturnValue));

            #endregion
        }

        /// <summary>
        /// This method is used to set property values for an object without invoking replication. 
        /// </summary>
        /// <param name="isSameWithSetProperties">Indicates whether result is same as RopSetProperties.</param>
        public void RopSetPropertiesNoReplicateMethod(out bool isSameWithSetProperties)
        {
            #region Set parameters
            // The value of taggedPropertyValueArray used in RopSetPropertiesNoReplicate.
            TaggedPropertyValue[] taggedPropertyValueArray = new TaggedPropertyValue[1];
            taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
            isSameWithSetProperties = false;
            #endregion

            #region Store original property value
            string objKey;
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Folder:
                    objKey = "Folder1";
                    break;
                case ServerObjectType.Message:
                    objKey = "Message1";
                    break;
                case ServerObjectType.Attachment:
                    objKey = "Attachment1";
                    break;
                case ServerObjectType.Logon:
                    objKey = "Logon";
                    break;
                default:
                    objKey = string.Empty;
                    break;
            }

            // The second session gets its original value for test property.
            bool isPropertyFound;
            byte[] propertyValueBeforeAction = this.Session2GetPropertyData(objKey, taggedPropertyValueArray[0].PropertyTag, out isPropertyFound);

            // Store the original value for further verification.
            TaggedPropertyValue tagPropertyValue = new TaggedPropertyValue
            {
                PropertyTag = taggedPropertyValueArray[0].PropertyTag,
                Value = propertyValueBeforeAction
            };

            this.tagPropertyValuesToVerification.Clear();
            this.tagPropertyValuesToVerification.Add(objKey, tagPropertyValue);
            #endregion

            #region Get property value and set a new value
            TaggedPropertyValue pretag = this.SearchProperty(this.cprptCurrentHandle, taggedPropertyValueArray[0].PropertyTag.PropertyId);

            if (pretag != null)
            {
                if (this.cprptCurrentType != ServerObjectType.Logon)
                {
                    taggedPropertyValueArray[0].Value[3]++;
                }
                else
                {
                    bool value = BitConverter.ToBoolean(pretag.Value, 0);
                    taggedPropertyValueArray[0].Value = BitConverter.GetBytes(!value);
                }
            }
            #endregion

            RopSetPropertiesNoReplicateResponse setPtyNoReplicateResponse = this.RopSetPropertiesNoReplicate(this.cprptCurrentHandle, taggedPropertyValueArray);

            #region Check if it is the same value that was the one set previously
            TaggedPropertyValue aftertag = this.SearchProperty(this.cprptCurrentHandle, taggedPropertyValueArray[0].PropertyTag.PropertyId);
            if (Common.CompareByteArray(aftertag.Value, taggedPropertyValueArray[0].Value))
            {
                isSameWithSetProperties = true;
            }
            #endregion

            this.VerifyRopSetPropertiesNoReplicate(setPtyNoReplicateResponse, this.cprptCurrentType);
        }

        /// <summary>
        /// This method is used to delete property values from an object without invoking replication. 
        /// </summary>
        /// <param name="isSameWithDeleteProperties">Indicates whether result is same as RopDeleteProperties.</param>
        /// <param name="isChangedInDB">Indicates the value is changed in database or not.</param>
        public void RopDeletePropertiesNoReplicateMethod(out bool isSameWithDeleteProperties, out bool isChangedInDB)
        {
            #region Initialize parameters
            isSameWithDeleteProperties = false;
            isChangedInDB = false;

            // The value of propertyTags used in RopDeletePropertiesNoReplicate.
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject).PropertyTag;
            #endregion

            #region Store original property values
            string objKey;
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Folder:
                    objKey = "Folder1";
                    break;
                case ServerObjectType.Message:
                    objKey = "Message1";
                    break;
                case ServerObjectType.Attachment:
                    objKey = "Attachment1";
                    break;
                case ServerObjectType.Logon:
                    objKey = "Logon";
                    break;
                default:
                    objKey = string.Empty;
                    break;
            }

            // The second session gets its original value for test property.
            bool isPropertyFound;
            byte[] propertyValueBeforeAction = this.Session2GetPropertyData(objKey, propertyTags[0], out isPropertyFound);

            // Store original value for further verification.
            TaggedPropertyValue tagPropertyValue = new TaggedPropertyValue
            {
                PropertyTag = propertyTags[0],
                Value = propertyValueBeforeAction
            };

            this.tagPropertyValuesToVerification.Clear();
            this.tagPropertyValuesToVerification.Add(objKey, tagPropertyValue);
            #endregion

            #region Check if property is existing
            RopGetPropertiesSpecificResponse getSpecRes = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, propertyTags);
            byte[] expectedNotFoundError = BitConverter.GetBytes((uint)CPRPTErrorCode.NotFound);
            if (Common.CompareByteArray(getSpecRes.RowData.PropertyValues[0].Value, expectedNotFoundError))
            {
                this.Site.Assert.Fail("The deleted property does not exist.");
            }
            #endregion

            RopDeletePropertiesNoReplicateResponse deletePropertiesNoReplicateResponse = this.RopDeletePropertiesNoReplicate(this.cprptCurrentHandle, propertyTags);

            #region Check isSameWithDeleteProperties
            TaggedPropertyValue tag = this.SearchProperty(this.cprptCurrentHandle, propertyTags[0].PropertyId);

            if (tag == null || Common.CompareByteArray(tag.Value, BitConverter.GetBytes((uint)CPRPTErrorCode.NotFound)))
            {
                isSameWithDeleteProperties = true;
            }
            #endregion

            this.VerifyRopDeletePropertiesNoReplicate(deletePropertiesNoReplicateResponse, this.cprptCurrentType);

            if (deletePropertiesNoReplicateResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                RopGetPropertiesSpecificResponse getSpecResAfterDel = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, propertyTags);

                // In Exchange 2010, property value is equal to NotFound error code.
                bool resultOfRopGetPropertiesSpecific = Common.CompareByteArray(getSpecResAfterDel.RowData.PropertyValues[0].Value, expectedNotFoundError);

                TaggedPropertyValue existingValue = this.SearchProperty(this.cprptCurrentHandle, propertyTags[0].PropertyId);
                bool resultOfRopGetPropertiesAll = existingValue == null || Common.CompareByteArray(existingValue.Value, expectedNotFoundError);

                Site.Assert.IsTrue(resultOfRopGetPropertiesSpecific == resultOfRopGetPropertiesAll, "The return values of RopGetPropertiesSpecific and RopGetPropertiesAll should be equal.");

                isSameWithDeleteProperties = resultOfRopGetPropertiesSpecific;

                if (isSameWithDeleteProperties)
                {
                    byte[] propertyValue = getSpecResAfterDel.RowData.PropertyValues[0].Value;
                    this.VerifyPropertyValueAfterPropertyDeleted(propertyValue, false);
                }

                #region Check if it is changed in DB
                // Check if it is changed in DB.
                byte[] propertyValueSession2Got = this.Session2GetPropertyData(objKey, this.tagPropertyValuesToVerification[objKey].PropertyTag, out isPropertyFound);
                if (isPropertyFound)
                {
                    // For property deletion, if the property value that got from session2 is equal to the value before deleting it,
                    // it means the deleted property does not be submitted to DB.
                    isChangedInDB = !Common.CompareByteArray(propertyValueSession2Got, this.tagPropertyValuesToVerification[objKey].Value);
                }
                else
                {
                    isChangedInDB = true;
                }
                #endregion
            }
            else
            {
                this.Site.Assert.Fail("RopDeletePropertiesMethod Failed");
            }
        }

        /// <summary>
        /// This method is used to update the specified properties on an object. 
        /// </summary>
        /// <param name="isModifiedValueReturned">Indicates whether the modified value of a property can be returned use a same handle.</param>
        /// <param name="isChangedInDB">Indicates whether the modified value is submit to DB.
        /// For Message and Attachment object, it require another ROP for submit DB.
        /// For Logon and Folder object, it DO NOT need any other ROPs for submit.</param>
        public void RopSetPropertiesMethod(out bool isModifiedValueReturned, out bool isChangedInDB)
        {
            #region Initialize parameter
            isModifiedValueReturned = false;
            isChangedInDB = false;

            // The value of taggedPropertyValueArray used in RopSetProperties.
            TaggedPropertyValue[] taggedPropertyValueArray = new TaggedPropertyValue[1];
            taggedPropertyValueArray[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            if (this.cprptCurrentType != ServerObjectType.Logon)
            {
                taggedPropertyValueArray[0].Value = Common.AddInt16LengthBeforeBinaryArray(taggedPropertyValueArray[0].Value);
            }

            #region Store original property value
            string objKey;
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Folder:
                    objKey = "Folder1";
                    break;
                case ServerObjectType.Message:
                    objKey = "Message1";
                    break;
                case ServerObjectType.Attachment:
                    objKey = "Attachment1";
                    break;
                case ServerObjectType.Logon:
                    objKey = "Logon";
                    break;
                default:
                    objKey = string.Empty;
                    break;
            }

            // Session2 gets its original value for test property.
            bool isPropertyFound;
            byte[] propertyValueBeforeAction = this.Session2GetPropertyData(objKey, taggedPropertyValueArray[0].PropertyTag, out isPropertyFound);

            // Store original value for further verification.
            TaggedPropertyValue tagPropertyValue = new TaggedPropertyValue
            {
                PropertyTag = taggedPropertyValueArray[0].PropertyTag,
                Value = propertyValueBeforeAction
            };

            this.tagPropertyValuesToVerification.Clear();
            this.tagPropertyValuesToVerification.Add(objKey, tagPropertyValue);
            #endregion

            TaggedPropertyValue pretag = this.SearchProperty(this.cprptCurrentHandle, taggedPropertyValueArray[0].PropertyTag.PropertyId);
            if (pretag != null)
            {
                if (this.cprptCurrentType != ServerObjectType.Logon)
                {
                    taggedPropertyValueArray[0].Value[3]++;
                }
                else
                {
                    bool value = BitConverter.ToBoolean(pretag.Value, 0);
                    taggedPropertyValueArray[0].Value = BitConverter.GetBytes(!value);
                }
            }

            #endregion

            RopSetPropertiesResponse setPropertiesResponse;
            if (this.cprptCurrentType == ServerObjectType.Folder)
            {
                // Set taggedPropertyValueArray[1] to invalid value to verify PropertyProblem structure
                TaggedPropertyValue[] tempValues = new TaggedPropertyValue[2];
                tempValues[0] = taggedPropertyValueArray[0];
                tempValues[1] = this.GetTaggedPropertyTag(ObjectToOperate.SecondObject);
                tempValues[1].PropertyTag.PropertyId = NewCreatePropertyID;
                tempValues[1].Value = tempValues[0].Value;

                setPropertiesResponse = this.RopSetProperties(this.cprptCurrentHandle, tempValues, false);
            }
            else
            {
                setPropertiesResponse = this.RopSetProperties(this.cprptCurrentHandle, taggedPropertyValueArray, false);
            }

            this.VerifyRopSetProperties(setPropertiesResponse, this.cprptCurrentType);

            if (setPropertiesResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                #region Check isModifiedValueReturned
                TaggedPropertyValue tag = this.SearchProperty(this.cprptCurrentHandle, taggedPropertyValueArray[0].PropertyTag.PropertyId);
                isModifiedValueReturned = Common.CompareByteArray(taggedPropertyValueArray[0].Value, tag.Value);
                #endregion

                #region Check isChangedInDB
                byte[] propertyValueSession2Got = this.Session2GetPropertyData(objKey, this.tagPropertyValuesToVerification[objKey].PropertyTag, out isPropertyFound);
                isChangedInDB = !Common.CompareByteArray(this.tagPropertyValuesToVerification[objKey].Value, propertyValueSession2Got);
                #endregion
            }
        }

        /// <summary>
        /// This method is used to remove the specified properties from an object. 
        /// </summary>
        /// <param name="isNoValidValueReturnedForDeletedProperties">
        /// If the server returns success, it MUST NOT have a valid value to return to a client that asks for the value of this property.</param>
        /// <param name="isChangedInDB">
        /// Indicates whether the modified value is submit to DB For Message and Attachment object, it require another 
        /// ROP for submit DB.For Logon and Folder object, it DO NOT need any other ROPs for submit.</param>
        public void RopDeletePropertiesMethod(out bool isNoValidValueReturnedForDeletedProperties, out bool isChangedInDB)
        {
            #region Initialize parameters
            isNoValidValueReturnedForDeletedProperties = false;
            isChangedInDB = false;

            // The value of propertyTags used in RopGetPropertiesSpecific.
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject).PropertyTag;

            #region Store original property value
            string objKey;
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Folder:
                    objKey = "Folder1";
                    break;
                case ServerObjectType.Message:
                    objKey = "Message1";
                    break;
                case ServerObjectType.Attachment:
                    objKey = "Attachment1";
                    break;
                case ServerObjectType.Logon:
                    objKey = "Logon";
                    break;
                default:
                    objKey = string.Empty;
                    break;
            }

            // Session2 gets its original value for test property.
            bool isPropertyFound;
            byte[] propertyValueBeforeAction = this.Session2GetPropertyData(objKey, propertyTags[0], out isPropertyFound);

            // Store original value for further verification.
            TaggedPropertyValue tagPropertyValue = new TaggedPropertyValue
            {
                PropertyTag = propertyTags[0],
                Value = propertyValueBeforeAction
            };

            this.tagPropertyValuesToVerification.Clear();
            this.tagPropertyValuesToVerification.Add(objKey, tagPropertyValue);
            #endregion

            #region Check if property exists.
            RopGetPropertiesSpecificResponse getSpecRes = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, propertyTags);
            byte[] expectedNotFoundError = BitConverter.GetBytes((uint)CPRPTErrorCode.NotFound);
            if (Common.CompareByteArray(getSpecRes.RowData.PropertyValues[0].Value, expectedNotFoundError))
            {
                this.Site.Assert.Fail("The deleted property does not exist.");
            }
            #endregion
            #endregion

            RopDeletePropertiesResponse deletePropertiesResponse = this.RopDeleteProperties(this.cprptCurrentHandle, propertyTags, false);

            this.VerifyRopDeleteProperties(deletePropertiesResponse, this.cprptCurrentType);

            if (deletePropertiesResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                RopGetPropertiesSpecificResponse getSpecResAfterDel = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, propertyTags);

                // In Exchange 2010, property value will be equals to NotFound error code.
                isNoValidValueReturnedForDeletedProperties = Common.CompareByteArray(getSpecResAfterDel.RowData.PropertyValues[0].Value, expectedNotFoundError);

                if (isNoValidValueReturnedForDeletedProperties)
                {
                    byte[] propertyValue = getSpecResAfterDel.RowData.PropertyValues[0].Value;
                    this.VerifyPropertyValueAfterPropertyDeleted(propertyValue, true);
                }

                #region Check if it is changed in DB
                byte[] propertyValueSession2Got = this.Session2GetPropertyData(objKey, this.tagPropertyValuesToVerification[objKey].PropertyTag, out isPropertyFound);
                if (isPropertyFound)
                {
                    // For property deletion, if the property value that got by session2 is equal to the value before deleting,
                    // it means deleted property does not be submitted to DB.
                    isChangedInDB = !Common.CompareByteArray(propertyValueSession2Got, this.tagPropertyValuesToVerification[objKey].Value);
                }
                else
                {
                    isChangedInDB = true;
                }
                #endregion
            }
            else
            {
                this.Site.Assert.Fail("RopDeletePropertiesMethod Failed");
            }
        }

        /// <summary>
        /// This method is used to commit the changes made to a message. 
        /// </summary>
        /// <param name="isChangedInDB">Indicates whether changes of Message object submit to database 
        /// when [RopSetProperties] or [RopDeleteProperties].</param>
        public void RopSaveChangesMessageMethod(out bool isChangedInDB)
        {
            #region Initialize parameters
            isChangedInDB = false;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            #endregion

            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                saveChangesMessageResponse = this.RopSaveChangesMessage(this.cprptCurrentHandle, false);
            }
            else
            {
                saveChangesMessageResponse = this.RopSaveChangesMessage(cprptMessageHandle[0], false);
            }

            // Also need to save changes for destination object.
            this.RopSaveChangesMessage(cprptMessageHandle[1], true);

            this.Site.Assert.AreEqual((uint)CPRPTErrorCode.None, saveChangesMessageResponse.ReturnValue, string.Format("RopSaveChangesMessageMethod Failed! Error: 0x{0:X8}", saveChangesMessageResponse.ReturnValue));

            #region Check isChangedInDB
            // Check whether the property is changed in DB.
            // Get the first key of Dictionary.
            string objKey = string.Empty;
            foreach (string key in this.tagPropertyValuesToVerification.Keys)
            {
                objKey = key;
                break;
            }

            bool isPropertyFound;
            byte[] propertyValueAfterSave = this.Session2GetPropertyData(objKey, this.tagPropertyValuesToVerification[objKey].PropertyTag, out isPropertyFound);
            if (isPropertyFound)
            {
                isChangedInDB = !Common.CompareByteArray(propertyValueAfterSave, this.tagPropertyValuesToVerification[objKey].Value);
            }
            else
            {
                // Deleting property and saving it will cause property not found error for message object.
                isChangedInDB = true;
            }
            #endregion
        }

        /// <summary>
        /// This method is used to commit the changes made to an attachment. 
        /// </summary>
        /// <param name="isChangedInDB">Indicates whether changes of Message object submit to database 
        /// when [RopSetProperties] or [RopDeleteProperties].</param>
        public void RopSaveChangesAttachmentMethod(out bool isChangedInDB)
        {
            isChangedInDB = false;
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse = this.RopSaveChangesAttachment(this.cprptCurrentHandle, false);

            // Attachment will be saved after its parent message saved.
            // Also need to save changes for destination object.
            this.RopSaveChangesAttachment(cprptAttachmentHandle[1], true);
            this.RopSaveChangesMessage(cprptMessageHandle[1], true);

            this.Site.Assert.AreEqual((uint)CPRPTErrorCode.None, saveChangesAttachmentResponse.ReturnValue, string.Format("RopSaveChangesMessageMethod Failed! Error: 0x{0:X8}", saveChangesAttachmentResponse.ReturnValue));

            #region Check isChangedInDB
            // Check whether the property is changed in DB.
            string objKey = string.Empty;
            foreach (string key in this.tagPropertyValuesToVerification.Keys)
            {
                objKey = key;
                break;
            }

            bool isPropertyFound;
            byte[] propertyValueAfterSave = this.Session2GetPropertyData(objKey, this.tagPropertyValuesToVerification[objKey].PropertyTag, out isPropertyFound);
            if (isPropertyFound)
            {
                isChangedInDB = !Common.CompareByteArray(propertyValueAfterSave, this.tagPropertyValuesToVerification[objKey].Value);
            }
            else
            {
                // Deleting property and saving it will cause property not found error for message object.
                isChangedInDB = true;
            }
            #endregion
        }

        /// <summary>
        /// This method is used to open a property as a Stream object, enabling the client to perform various streaming operations on the property. 
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        /// <param name="openFlag">Specifies OpenModeFlags for [RopOpenStream].</param>
        /// <param name="isPropertyTagExist">Indicates whether request property exist.</param>
        /// <param name="isStreamSizeEqualToStream">Indicates whether StreamSize in response is 
        /// the same with the current number of BYTES in the stream.</param>
        /// <param name="error">If the property tag does not exist for the object and "Create" 
        /// is not specified in OpenModeFlags, NotFound error should be returned.</param>
        public void RopOpenStreamMethod(ObjectToOperate obj, OpenModeFlags openFlag, bool isPropertyTagExist, out bool isStreamSizeEqualToStream, out CPRPTErrorCode error)
        {
            #region Initialize properties and parameter
            error = CPRPTErrorCode.None;
            isStreamSizeEqualToStream = false;
            TaggedPropertyValue propertyTag = this.GetTaggedPropertyTag(obj);

            uint handle = 0;
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Folder:
                    handle = cprptFolderHandle[(int)obj];

                    // For folder object, open PTagRulesData property
                    propertyTag = this.GetTaggedPropertyTag(ObjectToOperate.ThirdObject);
                    break;
                case ServerObjectType.Message:
                    handle = cprptMessageHandle[(int)obj];
                    break;
                case ServerObjectType.Attachment:
                    handle = cprptAttachmentHandle[(int)obj];
                    break;
                default:
                    Site.Assert.Fail("Invalid server object type {0} for RopOpenStream method.", this.cprptCurrentType);
                    break;
            }

            byte openModeFlags = ConstValues.OpenModeFlagsReadOnly;
            switch (openFlag)
            {
                case OpenModeFlags.ReadOnly:
                    openModeFlags = ConstValues.OpenModeFlagsReadOnly;
                    break;
                case OpenModeFlags.ReadWrite:
                    openModeFlags = ConstValues.OpenModeFlagsReadWrite;
                    break;
                case OpenModeFlags.Create:
                    openModeFlags = ConstValues.OpenModeFlagsCreate;
                    break;
                case OpenModeFlags.BestAccess:
                    openModeFlags = ConstValues.OpenModeFlagsBestAccess;
                    break;
                default:
                    this.Site.Assert.Fail("Invalid OpenModeFlags enum value {0}.", openFlag);
                    break;
            }

            if (isPropertyTagExist)
            {
                #region Get original stream value of current property tag
                string objKey;
                switch (this.cprptCurrentType)
                {
                    case ServerObjectType.Folder:
                        if (obj == ObjectToOperate.FirstObject)
                        {
                            objKey = "Folder1";
                        }
                        else
                        {
                            objKey = "Folder2";
                        }

                        break;
                    case ServerObjectType.Message:
                        if (obj == ObjectToOperate.FirstObject)
                        {
                            objKey = "Message1";
                        }
                        else
                        {
                            objKey = "Message2";
                        }

                        break;
                    case ServerObjectType.Attachment:
                        if (obj == ObjectToOperate.FirstObject)
                        {
                            objKey = "Attachment1";
                        }
                        else
                        {
                            objKey = "Attachment2";
                        }

                        break;
                    default:
                        objKey = string.Empty;
                        break;
                }

                bool isPropertyFound;
                byte[] originalStreamValue = this.Session2ReadStream(objKey, propertyTag.PropertyTag, out isPropertyFound);

                // Store the value for further verification
                TaggedPropertyValue tagPropertyValue = new TaggedPropertyValue
                {
                    PropertyTag = propertyTag.PropertyTag,
                    Value = originalStreamValue
                };

                if (this.tagPropertyValuesToVerification.ContainsKey(objKey))
                {
                    this.tagPropertyValuesToVerification[objKey] = tagPropertyValue;
                }
                else
                {
                    this.tagPropertyValuesToVerification.Add(objKey, tagPropertyValue);
                }
                #endregion
            }
            else
            {
                propertyTag.PropertyTag.PropertyId = NewCreatePropertyID;
                propertyTag.PropertyTag.PropertyType = (ushort)PropertyTypeName.PtypBinary;
            }

            error = CPRPTErrorCode.None;
            isStreamSizeEqualToStream = false;
            #endregion

            RopOpenStreamResponse openStreamResponse;
            uint openHandle = this.RopOpenStream(handle, out openStreamResponse, propertyTag.PropertyTag, openModeFlags, false);

            this.VerifyRopOpenStream(openStreamResponse, this.cprptCurrentType, isPropertyTagExist, openFlag);

            #region Verify error, isStreamSizeEqualToStream and set handle
            if (openFlag != OpenModeFlags.Create && isPropertyTagExist == false && openStreamResponse.ReturnValue == (uint)CPRPTErrorCode.NotFound)
            {
                error = CPRPTErrorCode.NotFound;
            }
            else if (openStreamResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                error = CPRPTErrorCode.None;
                uint streamSize = openStreamResponse.StreamSize;
                switch (obj)
                {
                    case ObjectToOperate.FirstObject:
                        this.cprptCurrentHandle = this.cprptFirstHandle = openHandle;
                        break;
                    case ObjectToOperate.SecondObject:
                        this.cprptSecondHandle = openHandle;
                        break;
                    default:
                        break;
                }

                RopSeekStreamResponse res = this.RopSeekStream(openHandle, (byte)Origin.End, ConstValues.RopSeekStreamOffsetZero, false);
                if (res.ReturnValue == (uint)CPRPTErrorCode.None)
                {
                    if (streamSize == res.NewPosition)
                    {
                        isStreamSizeEqualToStream = true;
                    }
                }
            }
            #endregion
        }

        /// <summary>
        /// This method is used to open a different type of properties as a Stream object, enabling the client to perform various streaming operations on the property. 
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        /// <param name="propertyType">Specifies which type of property will be operated.</param>
        /// <param name="error">Returned error code.</param>
        public void RopOpenStreamWithDifferentPropertyType(ObjectToOperate obj, PropertyTypeName propertyType, out CPRPTErrorCode error)
        {
            #region Initialize properties and parameter
            error = CPRPTErrorCode.None;
            TaggedPropertyValue propertyTag = this.GetTaggedPropertyTag(obj);

            uint handle = 0;
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Folder:
                    int index = obj == ObjectToOperate.FifthObject ? (int)ObjectToOperate.FirstObject : (int)obj;
                    handle = cprptFolderHandle[index];
                    break;
                case ServerObjectType.Message:
                    handle = cprptMessageHandle[(int)obj];
                    break;
                case ServerObjectType.Attachment:
                    handle = cprptAttachmentHandle[(int)obj];
                    break;
                default:
                    break;
            }

            byte openModeFlags = ConstValues.OpenModeFlagsCreate;
            propertyTag.PropertyTag.PropertyType = (ushort)propertyType;

            if (this.cprptCurrentType != ServerObjectType.Folder)
            {
                propertyTag.PropertyTag.PropertyId++;
            }
            else
            {
                if (obj == ObjectToOperate.FifthObject)
                {
                    switch (propertyType)
                    {
                        case PropertyTypeName.PtypBinary:
                            propertyTag.PropertyTag.PropertyId = 0x0e09;
                            propertyTag.PropertyTag.PropertyType = 0x0102;
                            break;
                        case PropertyTypeName.PtypString:
                            propertyTag.PropertyTag.PropertyId = 0x3001;
                            propertyTag.PropertyTag.PropertyType = 0x001f;
                            break;
                        default:
                            break;
                    }

                    openModeFlags = ConstValues.OpenModeFlagsReadOnly;
                }
                else
                {
                    // Set PTagRulesData property for Folder object
                    propertyTag.PropertyTag.PropertyId = ushort.Parse(Common.GetConfigurationPropertyValue("FolderPropertyID2", this.Site));
                }
            }

            error = CPRPTErrorCode.Other;

            #endregion
            RopOpenStreamResponse openStreamResponse;
            uint openHandle = this.RopOpenStream(handle, out openStreamResponse, propertyTag.PropertyTag, openModeFlags, false);

            error = (CPRPTErrorCode)openStreamResponse.ReturnValue;

            #region Write stream to verify having write permission
            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                RopWriteStreamResponse writeStreamResponse = this.RopWriteStream(openHandle, WriteData, false);
                this.RopCommitStream(openHandle, false);

                this.VerifyRopWriteStreamWithCreatePermission(writeStreamResponse);
            }
            #endregion
        }

        /// <summary>
        /// This method is used to read the stream of bytes from a Stream object.
        /// </summary>
        /// <param name="isReadingFailed">Indicates whether reading stream get failure. E.g. object handle is not stream.</param>
        public void RopReadStreamMethod(bool isReadingFailed)
        {
            #region Initialize parameters
            // The value of maxinumByteCount used in RopReadStream.
            uint maxByteCount = 0;
            RopGetStreamSizeResponse getStreamRes = this.RopGetStreamSize(this.cprptCurrentHandle, true);

            // Set position to begin.
            this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Beginning, ConstValues.RopSeekStreamOffsetZero, true);

            // Set byte count in seekStreamRes.
            ushort byteCount = (ushort)getStreamRes.StreamSize;
            RopReadStreamResponse readStreamResponse;
            #endregion
            if (isReadingFailed)
            {
                readStreamResponse = this.RopReadStream(cprptFolderHandle[0], byteCount, maxByteCount, false);
            }
            else
            {
                readStreamResponse = this.RopReadStream(this.cprptCurrentHandle, byteCount, maxByteCount, false);
            }

            this.VerifyRopReadStream(readStreamResponse, isReadingFailed);
        }

        /// <summary>
        /// This method is used to read the stream of limited size bytes from a Stream object.
        /// </summary>
        /// <param name="byteCount">Indicates the size to be read.</param>
        /// <param name="maxByteCount">If byteCount is 0xBABE, use MaximumByteCount to determine the size to be read.</param>
        public void RopReadStreamWithLimitedSize(ushort byteCount, uint maxByteCount)
        {
            RopReadStreamResponse readStreamResponse = this.RopReadStream(this.cprptCurrentHandle, byteCount, maxByteCount, false);

            this.VerifyRopReadStreamWithLimitedSize(readStreamResponse, byteCount, maxByteCount);
        }

        /// <summary>
        /// This method is used to set the seek pointer to a new location, which is relative to the beginning of the stream, the end of the stream, or the location of the current seek pointer.
        /// </summary>
        /// <param name="condition">Specifies particular scenario of RopSeekStream.</param>
        /// <param name="isStreamExtended">Indicates whether a stream object is extended and zero filled to the new seek location.</param>
        /// <param name="error">Returned error code.</param>
        public void RopSeekStreamMethod(SeekStreamCondition condition, out bool isStreamExtended, out CPRPTErrorCode error)
        {
            error = CPRPTErrorCode.None;
            isStreamExtended = false;
            RopGetStreamSizeResponse sizeRes = this.RopGetStreamSize(this.cprptCurrentHandle, true);
            uint size = sizeRes.StreamSize;

            // The origin value in RopSeekStreamMethod.
            byte origin = (byte)Origin.Beginning;

            // The offset value in RopSeekStreamMethod.
            long offset = 0;
            switch (condition)
            {
                case SeekStreamCondition.MovedBeyondMaxStreamSize:
                    origin = (byte)Origin.Beginning;

                    // The maximum value of unit is larger than 2^31 bytes.
                    offset = uint.MaxValue;
                    break;
                case SeekStreamCondition.MovedBeyondEndOfStream:
                    origin = (byte)Origin.End;
                    offset++;
                    break;
                case SeekStreamCondition.Normal:
                    break;
                case SeekStreamCondition.OriginInvalid:
                    origin = (byte)Origin.Invalid;
                    break;
                default:
                    break;
            }

            RopSeekStreamResponse seekStreamResponse = this.RopSeekStream(this.cprptCurrentHandle, origin, offset, false);
            error = (CPRPTErrorCode)seekStreamResponse.ReturnValue;
            if (error != CPRPTErrorCode.StreamSeekError)
            {
                if (condition == SeekStreamCondition.MovedBeyondEndOfStream)
                {
                    this.RopWriteStream(this.cprptCurrentHandle, WriteData, true);
                }

                sizeRes = this.RopGetStreamSize(this.cprptCurrentHandle, true);
                if (size < sizeRes.StreamSize)
                {
                    isStreamExtended = true;
                }
            }

            this.VerifyRopSeekStream(seekStreamResponse, condition);
        }

        /// <summary>
        /// This method is used to lock a specified range of bytes in a Stream object.
        /// </summary>
        /// <param name="preState">Specifies the pre-state before call [RopLockRegionStream]</param>
        /// <param name="error">Return error
        /// 1. If there are previous locks that are not expired, the server MUST return an AccessDenied error.
        /// 2. If a session with an expired lock calls any ROP for this Stream object that would encounter the locked region, 
        /// the server MUST return a NetworkError.</param>
        public void RopLockRegionStreamMethod(PreStateBeforeLock preState, out CPRPTErrorCode error)
        {
            error = CPRPTErrorCode.None;
            RopLockRegionStreamResponse lockRegionStreamResponse;

            // The regionOffset value used in RopLockRegionStream.
            ulong regionOffset = 0;

            // The lockFlags value used in RopLockRegionStream.
            uint lockFlags = (uint)LockFlags.OtherValue;
            RopGetStreamSizeResponse getStreamSizeRes = this.RopGetStreamSize(this.cprptCurrentHandle, true);
            ulong regionSize = getStreamSizeRes.StreamSize;

            lockRegionStreamResponse = this.RopLockRegionStream(this.cprptCurrentHandle, regionOffset, regionSize, lockFlags, false);

            if (preState == PreStateBeforeLock.PreLockNotExpired)
            {
                error = (CPRPTErrorCode)lockRegionStreamResponse.ReturnValue;
                this.Site.Assert.AreEqual(CPRPTErrorCode.None, error, string.Format("RopLockRegionStream failed with error: 0x{0:X8}", lockRegionStreamResponse.ReturnValue));

                lockRegionStreamResponse = this.RopLockRegionStream(this.cprptCurrentHandle, regionOffset, regionSize, lockFlags, false);
                error = (CPRPTErrorCode)lockRegionStreamResponse.ReturnValue;
            }
            else if (preState == PreStateBeforeLock.WithExpiredLock)
            {
                // The byteCount variable in RopReadStream.
                ushort byteCount = 1;

                // The maxByteCount Variable in RopReadStream.
                uint maxByteCount = 1;
                this.RopReadStream(this.cprptCurrentHandle, byteCount, maxByteCount, false);
                error = (CPRPTErrorCode)lockRegionStreamResponse.ReturnValue;

                this.Site.Assert.AreNotEqual(CPRPTErrorCode.AccessDenied, error, string.Format("RopLockRegionStream failed with error: 0x{0:X8}", lockRegionStreamResponse.ReturnValue));
            }
            else if (preState == PreStateBeforeLock.Normal)
            {
                error = (CPRPTErrorCode)lockRegionStreamResponse.ReturnValue;

                this.Site.Assert.AreEqual((uint)CPRPTErrorCode.None, lockRegionStreamResponse.ReturnValue, string.Format("RopLockRegionStream failed with error: 0x{0:X8}", lockRegionStreamResponse.ReturnValue));
            }

            this.VerifyRopLockRegionStream(lockRegionStreamResponse);
        }

        /// <summary>
        /// This method is used to unlock a specified range of bytes in a Stream object.
        /// </summary>
        /// <param name="isPreviousLockExists">Indicates whether a previous lock exists and not owned by this session.
        /// If there are previous locks that are not owned by the session calling the ROP, the server MUST leave them unmodified.</param>
        public void RopUnlockRegionStreamMethod(bool isPreviousLockExists)
        {
            // The regionOffset used in RopUnlockRegionStream.
            ulong regionOffset = 0;

            // The lockFlags used in RopUnlockRegionStream.
            uint lockFlags = (uint)LockFlags.OtherValue;
            RopGetStreamSizeResponse getStreamSizeRes = this.RopGetStreamSize(this.cprptCurrentHandle, true);
            ulong regionSize = getStreamSizeRes.StreamSize;
            if (!isPreviousLockExists)
            {
                this.RopLockRegionStream(this.cprptCurrentHandle, regionOffset, regionSize, lockFlags, true);
            }

            RopUnlockRegionStreamResponse unlockRegionStreamResponse = this.RopUnlockRegionStream(this.cprptCurrentHandle, regionOffset, regionSize, lockFlags);

            this.VerifyRopUnlockRegionStream(unlockRegionStreamResponse);
        }

        /// <summary>
        /// This method is used to write the stream of bytes into a Stream object. 
        /// </summary>
        /// <param name="openFlag">Specifies the OpenModeFlags of the stream.</param>
        /// <param name="isExceedMax">Indicates whether the write will exceed the maximum stream size.</param>
        /// <param name="error"> Specifies the ErrorCode when WriteStream failed:STG_E_ACCESSDENIED
        /// 0x80030005 Write access is denied.When stream is opened with ReadOnly flag.</param>
        public void RopWriteStreamMethod(OpenModeFlags openFlag, bool isExceedMax, out CPRPTErrorCode error)
        {
            TaggedPropertyValue[] tagPtyValues = new TaggedPropertyValue[1];
            if (this.cprptCurrentType == ServerObjectType.Folder)
            {
                tagPtyValues[0] = this.GetTaggedPropertyTag(ObjectToOperate.ThirdObject);
            }
            else
            {
                tagPtyValues[0] = this.GetTaggedPropertyTag(this.cprptCurrentObj);
            }

            PropertyTag[] ptyTags;
            ptyTags = new PropertyTag[1];
            ptyTags[0] = tagPtyValues[0].PropertyTag;

            uint objHandle = 0;
            switch (this.cprptCurrentObj)
            {
                case ObjectToOperate.FirstObject:
                    objHandle = this.cprptFirstObjectHandle;
                    break;
                case ObjectToOperate.SecondObject:
                    objHandle = this.cprptSecondObjectHandle;
                    break;
                default:
                    break;
            }

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.RopGetPropertiesSpecific(objHandle, 0, 0, ptyTags);
            bool canBeRetrieval = false;
            bool isChangInDB = true;
            bool forErrorCode = false;
            if (isExceedMax)
            {
                this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Beginning, int.MaxValue, true);
            }

            RopSeekStreamResponse seekStreamResponse1 = this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Current, 0, true);
            RopWriteStreamResponse writeStreamResponse = this.RopWriteStream(this.cprptCurrentHandle, WriteData, false);
            RopSeekStreamResponse seekStreamResponse2 = this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Current, 0, true);
            bool isWriteSizeElemetRight = false;
            if (seekStreamResponse2.NewPosition - seekStreamResponse1.NewPosition == writeStreamResponse.WrittenSize)
            {
                isWriteSizeElemetRight = true;
            }

            if (writeStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.None))
            {
                this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Current, (long)(0 - WriteData.Length), true);
                RopReadStreamResponse readStreamResponse = this.RopReadStream(this.cprptCurrentHandle, (ushort)WriteData.Length, 0x70000000, false);
                if (WriteData == Encoding.ASCII.GetString(readStreamResponse.Data))
                {
                    canBeRetrieval = true;
                }

                RopGetPropertiesSpecificResponse getPropertiesSpecificResponse1 = this.RopGetPropertiesSpecific(objHandle, 0, 0, ptyTags);
                if (Common.CompareByteArray(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, getPropertiesSpecificResponse1.RowData.PropertyValues[0].Value))
                {
                    isChangInDB = false;
                }
            }
            else
            {
                forErrorCode = true;
            }

            this.VerifyRopWriteStream(writeStreamResponse, openFlag, WriteData, canBeRetrieval, isChangInDB, forErrorCode, isWriteSizeElemetRight);

            error = (CPRPTErrorCode)writeStreamResponse.ReturnValue;
        }

        /// <summary>
        /// This method is used to ensure that any changes made to a Stream object are persisted in storage for a Folder object. 
        /// </summary>
        /// <param name="openFlag">Indicates the OpenModeFlags when stream is opened.</param>
        /// <param name="isPropertyValueChanged">Indicates whether property value is changed.</param>
        public void RopCommitStreamMethod(OpenModeFlags openFlag, out bool isPropertyValueChanged)
        {
            isPropertyValueChanged = false;
            RopCommitStreamResponse commitStreamResponse = this.RopCommitStream(this.cprptCurrentHandle, false);

            // Message and attachment object can only be changed after saving.
            this.RopSaveChangesAttachment(cprptAttachmentHandle[0], true);
            this.RopSaveChangesMessage(cprptMessageHandle[0], true);

            this.VerifyRopCommitStream(commitStreamResponse);

            if (commitStreamResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                // Check whether it is changed in DB.
                string objKey = string.Empty;
                switch (this.cprptCurrentType)
                {
                    case ServerObjectType.Message:
                        if (this.cprptCurrentObj == ObjectToOperate.FirstObject)
                        {
                            objKey = "Message1";
                        }
                        else
                        {
                            objKey = "Message2";
                        }

                        break;
                    case ServerObjectType.Folder:
                        if (this.cprptCurrentObj == ObjectToOperate.FirstObject)
                        {
                            objKey = "Folder1";
                        }
                        else
                        {
                            objKey = "Folder2";
                        }

                        break;
                    case ServerObjectType.Attachment:
                        if (this.cprptCurrentObj == ObjectToOperate.FirstObject)
                        {
                            objKey = "Attachment1";
                        }
                        else
                        {
                            objKey = "Attachment2";
                        }

                        break;
                    default:
                        break;
                }

                bool isPropertyFound;
                byte[] propertyValueAfterWriteStream = this.Session2ReadStream(objKey, this.tagPropertyValuesToVerification[objKey].PropertyTag, out isPropertyFound);
                isPropertyValueChanged = !Common.CompareByteArray(propertyValueAfterWriteStream, this.tagPropertyValuesToVerification[objKey].Value);
            }
        }

        /// <summary>
        /// This method is used to release all resources associated with a Server object. 
        /// The client uses RopRelease ([MS-OXCROPS] section 2.2.14.3) after it is done with the Stream object.
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        /// <param name="isPropertyValueChanged">For Folder Object, this ROP should not change the value 
        /// in stream after RopWriteStream.For non-Folder Object, this ROP should change the value.</param>
        public void RopReleaseMethod(ObjectToOperate obj, out bool isPropertyValueChanged)
        {
            this.RopReleaseMethodNoVerify(obj);

            string objKey = string.Empty;

            // Save attachment and message to enforce changes.
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Message:
                    if (obj == ObjectToOperate.FirstObject)
                    {
                        objKey = "Message1";
                        this.RopSaveChangesMessage(cprptMessageHandle[0], true);
                    }
                    else
                    {
                        objKey = "Message2";
                        this.RopSaveChangesMessage(cprptMessageHandle[1], true);
                    }

                    break;
                case ServerObjectType.Attachment:
                    if (obj == ObjectToOperate.FirstObject)
                    {
                        objKey = "Attachment1";
                        this.RopSaveChangesAttachment(cprptAttachmentHandle[0], true);
                        this.RopSaveChangesMessage(cprptMessageHandle[0], true);
                    }
                    else
                    {
                        objKey = "Attachment2";
                        this.RopSaveChangesAttachment(cprptAttachmentHandle[1], true);
                        this.RopSaveChangesMessage(cprptMessageHandle[1], true);
                    }

                    break;
                case ServerObjectType.Folder:
                    if (obj == ObjectToOperate.FirstObject)
                    {
                        objKey = "Folder1";
                    }
                    else
                    {
                        objKey = "Folder2";
                    }

                    break;
                default:
                    break;
            }

            // Check whether it is changed in DB.
            bool isPropertyFound;
            byte[] propertyValueAfterWriteStream = this.Session2ReadStream(objKey, this.tagPropertyValuesToVerification[objKey].PropertyTag, out isPropertyFound);
            isPropertyValueChanged = !Common.CompareByteArray(propertyValueAfterWriteStream, this.tagPropertyValuesToVerification[objKey].Value);
        }

        /// <summary>
        /// This method is used to release all resources associated with a Server object. 
        /// The client uses RopRelease ([MS-OXCROPS] section 2.2.14.3) after it is done with the Stream object.
        /// </summary>
        /// <param name="obj">Specifies which object will be operated.</param>
        public void RopReleaseMethodNoVerify(ObjectToOperate obj)
        {
            // The handle to release.
            uint handle = 0;
            if (obj == ObjectToOperate.FirstObject)
            {
                handle = this.cprptFirstHandle;
            }
            else
            {
                handle = this.cprptSecondHandle;
            }

            this.RopRelease(handle);
        }

        /// <summary>
        /// This method is used to copy a specified number of bytes from the current seek pointer in the source stream to the current seek pointer in the destination stream. 
        /// </summary>
        /// <param name="isDestinationExist">Specified the whether the destination existed.</param>
        /// <param name="isReadWriteSuccess">When call success:The server MUST read the number of BYTES
        /// requested from the source Stream object, and write those bytes into the destination Stream object .</param>
        /// <param name="error">If Destination object does not exist .Expect DestinationNullObject error.</param>
        public void RopCopyToStreamMethod(bool isDestinationExist, out bool isReadWriteSuccess, out CPRPTErrorCode error)
        {
            isReadWriteSuccess = false;
            error = CPRPTErrorCode.None;

            // The source handle.
            uint firstHandle = this.cprptFirstHandle;

            // The destination handle.
            uint secondHandle = InvalidHandle;
            if (isDestinationExist)
            {
                secondHandle = this.cprptSecondHandle;
            }

            RopGetStreamSizeResponse getStreamSizeRes = this.RopGetStreamSize(firstHandle, true);
            ulong byteCount = getStreamSizeRes.StreamSize;

            // Set position of the source handle to the beginning of stream.
            this.RopSeekStream(firstHandle, (byte)Origin.Beginning, ConstValues.RopSeekStreamOffsetZero, true);
            if (isDestinationExist)
            {
                // Set position of the destination handle to the beginning of stream.
                this.RopSeekStream(secondHandle, (byte)Origin.Beginning, ConstValues.RopSeekStreamOffsetZero, true);
            }

            RopCopyToStreamResponse copyToStreamResponse = this.RopCopyToStream(firstHandle, secondHandle, (byte)HandleIndex.FirstIndex, (byte)HandleIndex.SecondIndex, byteCount);
            bool isCopyRightDataCount = false;
            if (isDestinationExist)
            {
                isCopyRightDataCount = byteCount == copyToStreamResponse.WrittenByteCount;
            }

            bool destHandleIsRight = false;
            if (copyToStreamResponse.ReturnValue.Equals((uint)CPRPTErrorCode.NullDestinationObject))
            {
                destHandleIsRight = copyToStreamResponse.DestHandleIndex == (uint)HandleIndex.SecondIndex;
            }

            this.VerifyRopCopyToStream(copyToStreamResponse, isDestinationExist, isCopyRightDataCount, destHandleIsRight);

            if (copyToStreamResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                isReadWriteSuccess = true;
            }

            if (copyToStreamResponse.ReturnValue == (uint)CPRPTErrorCode.NullDestinationObject)
            {
                isReadWriteSuccess = false;
                error = CPRPTErrorCode.NullDestinationObject;
            }
        }

        /// <summary>
        /// This method is used to copy or move properties from a source object to a destination object with error code returned. 
        /// </summary>
        /// <param name="condition">Specifies a special scenario of RopCopyTo.</param>
        public void RopCopyToMethodForErrorCodeTable(CopyToCondition condition)
        {
            // The source handle.
            uint sourceHandle = 0;

            // The destination handle.
            uint destHandle = 0;
            uint destMessageHandleFirst = 0;
            uint destMessageHandleSecond = 0;

            if ((condition == CopyToCondition.SourceMessageContainsDestMessage) || (condition == CopyToCondition.SourceMessageIndirectlyContainsDestMessage))
            {
                this.GetDestinationMessageHandle(cprptMessageHandle[0], out destMessageHandleFirst, out destMessageHandleSecond);
            }

            RopCreateFolderResponse createFolderResponse = new RopCreateFolderResponse();
            switch (condition)
            {
                case CopyToCondition.SourceDestNotCompatible:
                    sourceHandle = cprptMessageHandle[0];
                    destHandle = cprptFolderHandle[0];
                    break;
                case CopyToCondition.SourceContainsDest:
                    sourceHandle = cprptFolderHandle[2];
                    destHandle = this.RopCreateFolder(sourceHandle, out createFolderResponse, SubFolder, SubFolder, true);
                    break;
                case CopyToCondition.SourceDestHasSubObjWithSameDisplayName:
                    // Create sub folder for Folder1 and Folder2 with the same name.
                    this.RopCreateFolder(cprptFolderHandle[1], out createFolderResponse, SubFolder, SubFolder, true);
                    this.RopCreateFolder(cprptFolderHandle[2], out createFolderResponse, SubFolder, SubFolder, true);
                    sourceHandle = cprptFolderHandle[1];
                    destHandle = cprptFolderHandle[2];
                    break;
                case CopyToCondition.Normal:
                    sourceHandle = cprptMessageHandle[0];
                    destHandle = cprptMessageHandle[1];
                    break;
                case CopyToCondition.SourceMessageContainsDestMessage:
                    sourceHandle = cprptMessageHandle[0];
                    destHandle = destMessageHandleFirst;
                    break;
                case CopyToCondition.SourceMessageIndirectlyContainsDestMessage:
                    sourceHandle = cprptMessageHandle[0];
                    destHandle = destMessageHandleSecond;
                    break;
                default:
                    this.Site.Assert.Fail("Invalid CopyToCondition enum value {0}.", condition);
                    break;
            }

            // Copy subobject.
            byte wantSubObject = Convert.ToByte(true);

            RopCopyToResponse copyToRes = (RopCopyToResponse)this.RopCopyTo(sourceHandle, destHandle, (byte)HandleIndex.FirstIndex, (byte)HandleIndex.SecondIndex, 0, wantSubObject, (byte)RopCopyToCopyFlags.Move, null);

            this.VerifyRopCopyTo(copyToRes, condition, this.cprptCurrentType, CopyFlags.Move);
        }

        /// <summary>
        /// This method is used to copy or move properties from a source object to a destination object on public folder.
        /// </summary>
        public void RopCopyToForPublicFolder()
        {
            // Open the public folder and get the folder handle.
            ulong publicFolderId = this.folderIds[1];
            RopOpenFolderResponse openFolderResponse;
            uint publicFolderHandle = this.RopOpenFolder(this.cprptLogonHandle, out openFolderResponse, publicFolderId, true);

            // Open the existing public folder.
            string existingPublicFolderName = Common.GetConfigurationPropertyValue("PublicFolderName", this.Site);
            ulong existingPublicFolderId = this.GetSubfolderIDByName(publicFolderHandle, existingPublicFolderName);

            this.existingPublicFolderHandle = this.RopOpenFolder(publicFolderHandle, out openFolderResponse, existingPublicFolderId, true);

            // Create 2 sub-folders of the existing public folder.
            RopCreateFolderResponse createFolderResponse;
            for (int i = 0; i < 2; i++)
            {
                string displayName = FolderDisplayName + i.ToString();
                cprptFolderHandle[i] = this.RopCreateFolder(this.existingPublicFolderHandle, out createFolderResponse, displayName, "Comment: " + displayName, true);
                cprptFolderId[i] = createFolderResponse.FolderId;
            }

            // The source folder handle.
            uint sourceFolderHandle = cprptFolderHandle[0];

            // The destination folder handle.
            uint destinationFolderHandle = cprptFolderHandle[1];

            // Copy subobject.
            byte wantSubObject = Convert.ToByte(true);

            // Call RopCopyTo operation.
            RopCopyToResponse ropCopyToResponse = (RopCopyToResponse)this.RopCopyTo(sourceFolderHandle, destinationFolderHandle, (byte)HandleIndex.FirstIndex, (byte)HandleIndex.SecondIndex, 0, wantSubObject, (byte)RopCopyToCopyFlags.Move, null);

            this.VerifyRopCopyToForPublicFolder(ropCopyToResponse);
        }

        /// <summary>
        /// This method is used to write bytes to a stream and commits the stream. 
        /// </summary>
        /// <param name="error">This ROP MUST NOT be used on Stream objects opened on 
        /// properties on Folder objects which means it should be failed against Folder object.</param>
        public void RopWriteAndCommitStreamMethod(out CPRPTErrorCode error)
        {
            error = CPRPTErrorCode.None;
            bool isWriteCorrectDataToStream = false;
            RopSeekStreamResponse seekStreamResponse1 = this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Current, 0, true);
            RopWriteStreamResponse writeStreamResponse = this.RopWriteAndCommitStream(this.cprptCurrentHandle, WriteData, false);
            RopSeekStreamResponse seekStreamResponse2 = this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Current, 0, true);

            if (seekStreamResponse2.NewPosition - seekStreamResponse1.NewPosition == (ulong)WriteData.Length)
            {
                isWriteCorrectDataToStream = true;
            }

            // Message and attachment object can only be changed after saving.
            this.RopSaveChangesAttachment(cprptAttachmentHandle[0], true);
            this.RopSaveChangesMessage(cprptMessageHandle[0], true);

            this.VerifyRopWriteAndCommitStream(writeStreamResponse, isWriteCorrectDataToStream);

            if (this.cprptCurrentType == ServerObjectType.Folder && writeStreamResponse.ReturnValue != (uint)CPRPTErrorCode.None)
            {
                error = CPRPTErrorCode.NotFound;
            }
        }

        /// <summary>
        /// This method is used to create a new Stream object that is a clone of another Stream object. 
        /// </summary>
        public void RopCloneStreamMethod()
        {
            RopCloneStreamResponse cloneStreamResponse = this.RopCloneStream(this.cprptFirstHandle, this.cprptSecondHandle, true);

            this.VerifyRopCloneStream(cloneStreamResponse);
        }

        /// <summary>
        /// This method is used to retrieve the size of the stream. 
        /// </summary>
        public void RopGetStreamSizeMethod()
        {
            RopGetStreamSizeResponse getStreamSizeResponse = this.RopGetStreamSize(this.cprptCurrentHandle, false);

            this.VerifyRopGetStreamSize(getStreamSizeResponse);

            this.Site.Assert.AreEqual((uint)CPRPTErrorCode.None, getStreamSizeResponse.ReturnValue, string.Format("Get stream size error: 0x{0:X8}", getStreamSizeResponse.ReturnValue));
        }

        /// <summary>
        /// This method is used to set the size of a stream. 
        /// </summary>
        /// <param name="isSizeIncreased"> Indicates the new size is increased or decreased.</param>
        /// <param name="isExtendedValueZero">
        /// If the size of the stream is increased, then value of the extended stream MUST be zero.</param>
        /// <param name="isLost">
        /// If the size of the stream is decreased, the information that extends past the end of the new size is lost.</param>
        /// <param name="isIncrease">If the size of the stream is increased, set this value to true</param>
        public void RopSetStreamSizeMethod(bool isSizeIncreased, out bool isExtendedValueZero, out bool isLost, out bool isIncrease)
        {
            isExtendedValueZero = false;
            isIncrease = false;
            isLost = false;
            bool isStreamSizeEqualSetSize = false;
            RopGetStreamSizeResponse getStreamSizeRes = this.RopGetStreamSize(this.cprptCurrentHandle, true);

            // The size of current object
            uint orginSize = getStreamSizeRes.StreamSize;

            RopSetStreamSizeResponse setStreamSizeResponse = new RopSetStreamSizeResponse();
            if (isSizeIncreased)
            {
                setStreamSizeResponse = this.RopSetStreamSize(this.cprptCurrentHandle, orginSize + 1, false);
                if (setStreamSizeResponse.ReturnValue == 0)
                {
                    this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Beginning, 0, true);
                    RopReadStreamResponse readResponse = this.RopReadStream(this.cprptCurrentHandle, (ushort)(orginSize + 2), 0, true);
                    if (readResponse.DataSize == orginSize + 1)
                    {
                        isIncrease = true;
                        isStreamSizeEqualSetSize = true;
                    }
                }

                this.RopCommitStream(this.cprptCurrentHandle, true);
                this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Beginning, orginSize, true);

                RopReadStreamResponse readStreamRes = this.RopReadStream(this.cprptCurrentHandle, 1, 0, true);
                if (readStreamRes.DataSize > 0)
                {
                    if (readStreamRes.Data[0] == 0)
                    {
                        isExtendedValueZero = true;
                    }
                }
            }
            else
            {
                setStreamSizeResponse = this.RopSetStreamSize(this.cprptCurrentHandle, orginSize - 1, false);
                if (setStreamSizeResponse.ReturnValue == 0)
                {
                    this.RopSeekStream(this.cprptCurrentHandle, (byte)Origin.Beginning, 0, true);
                    RopReadStreamResponse readResponse = this.RopReadStream(this.cprptCurrentHandle, (ushort)orginSize, 0, true);
                    if (readResponse.DataSize == orginSize - 1)
                    {
                        isLost = true;
                        isStreamSizeEqualSetSize = true;
                    }
                }
            }

            this.VerifyRopSetStreamSize(setStreamSizeResponse, isStreamSizeEqualSetSize);
            this.Site.Assert.AreEqual((uint)CPRPTErrorCode.None, setStreamSizeResponse.ReturnValue, string.Format("Logon Failed! Error: 0x{0:X8}", setStreamSizeResponse.ReturnValue));
        }

        /// <summary>
        /// This method is used to copy or move one or more properties from one object to another. 
        /// </summary>
        /// <param name="copyFlag">Specifies the CopyFlags in the call request.</param>
        /// <param name="isWantAsynchronousZero">Indicates whether WantAsynchronous parameter in call request is zero.</param>
        /// <param name="isDestinationExist">Indicates whether destination object is exist for [RopCopyProperties]</param>
        /// <param name="isPropertiesDeleted">If CopyFlags is set to Move,Source object will be deleted after copy to.</param>
        /// <param name="isChangedInDB">Indicates whether the change is submit to DB.</param>
        /// <param name="isOverwriteDestination">If CopyFlags is set to NoOverWrite,Destination should not be overwritten.</param>
        /// <param name="isReturnedRopProgress">If this ROP is performed Asynchronously,RopProgress response returned 
        /// instead of RopCopyProperties response.</param>
        /// <param name="error">If destination object is not exist,NullDestinationObject error will be returned.</param>
        public void RopCopyPropertiesMethod(
            CopyFlags copyFlag,
            bool isWantAsynchronousZero,
            bool isDestinationExist,
            out bool isPropertiesDeleted,
            out bool isChangedInDB,
            out bool isOverwriteDestination,
            out bool isReturnedRopProgress,
            out CPRPTErrorCode error)
        {
            // Check if there are conflict SHOULD/MAY settings 
            this.CheckRequirementConflict(86701, 86502);
            this.CheckRequirementConflict(86704, 50101);
            this.CheckRequirementConflict(86705, 88001);

            #region Set parameter
            isPropertiesDeleted = false;
            isChangedInDB = false;
            isOverwriteDestination = false;
            isReturnedRopProgress = false;
            error = CPRPTErrorCode.None;

            byte copyFlags = (byte)RopCopyPropertiesCopyFlags.None;
            switch (copyFlag)
            {
                case CopyFlags.None:
                    copyFlags = (byte)RopCopyPropertiesCopyFlags.None;
                    break;
                case CopyFlags.Move:
                    copyFlags = (byte)RopCopyPropertiesCopyFlags.Move;
                    break;
                case CopyFlags.NoOverWriteAndDestPropNull:
                    copyFlags = (byte)RopCopyPropertiesCopyFlags.NoOverwrite;
                    break;
                case CopyFlags.NoOverWrite:
                    copyFlags = (byte)RopCopyPropertiesCopyFlags.NoOverwrite;
                    break;
                case CopyFlags.MoveAndNoOverWrite:
                    copyFlags = (byte)RopCopyPropertiesCopyFlags.MoveAndOverwrite;
                    break;
                case CopyFlags.Other:
                    copyFlags = (byte)4;
                    break;
                default:
                    this.Site.Assert.Fail("Invalid RopCopyPropertiesCopyFlags enum vaule {0}", copyFlag);
                    break;
            }

            byte wantAsynchronous = 0;
            if (!isWantAsynchronousZero)
            {
                wantAsynchronous = 1;
            }

            uint sourceHandle = this.cprptSecondHandle;
            uint destHandle = InvalidHandle;
            #endregion

            #region Set source and destination handle property
            TaggedPropertyValue[] sourceTag = new TaggedPropertyValue[1];
            sourceTag[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            sourceTag[0].Value = Common.AddInt16LengthBeforeBinaryArray(sourceTag[0].Value);

            TaggedPropertyValue[] destTag = new TaggedPropertyValue[1];
            destTag[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            destTag[0].Value = Common.AddInt16LengthBeforeBinaryArray(destTag[0].Value);
            destTag[0].Value[2]++;

            if (isDestinationExist)
            {
                destHandle = this.cprptFirstHandle;
                if (this.cprptCurrentType == ServerObjectType.Folder)
                {
                    RopCreateFolderResponse createFolderResponse;
                    string displayName = SubFolder;
                    sourceHandle = this.RopCreateFolder(sourceHandle, out createFolderResponse, displayName, "Comment: " + displayName, true);
                }

                this.RopSetProperties(destHandle, destTag, true);
                if (this.cprptCurrentType == ServerObjectType.Message)
                {
                    this.RopSaveChangesMessage(destHandle, true);
                }
                else if (this.cprptCurrentType == ServerObjectType.Attachment)
                {
                    this.RopSaveChangesAttachment(destHandle, true);
                    this.RopSaveChangesMessage(cprptMessageHandle[0], true);
                }

                #region Store destHandle property
                // Search property value for the first object
                TaggedPropertyValue tag = new TaggedPropertyValue
                {
                    PropertyTag = destTag[0].PropertyTag
                };
                RopGetPropertiesSpecificResponse destTagValue = this.RopGetPropertiesSpecific(destHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, new PropertyTag[] { tag.PropertyTag });
                tag.Value = destTagValue.RowData.PropertyValues[0].Value;

                // Store the TaggedPropertyValue for further verification
                this.tagPropertyValuesToVerification.Clear();
                switch (this.cprptCurrentType)
                {
                    case ServerObjectType.Folder:
                        this.tagPropertyValuesToVerification.Add(ConstValues.ServerObjectTypeFolderKey, tag);
                        break;
                    case ServerObjectType.Message:
                        this.tagPropertyValuesToVerification.Add(ConstValues.ServerObjectTypeMessageKey, tag);
                        break;
                    case ServerObjectType.Attachment:
                        this.tagPropertyValuesToVerification.Add(ConstValues.ServerObjectTypeAttachmentKey, tag);
                        break;
                    default:
                        break;
                }
                #endregion
            }

            this.RopSetProperties(sourceHandle, sourceTag, true);
            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                this.RopSaveChangesMessage(sourceHandle, true);
            }
            else if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                this.RopSaveChangesAttachment(sourceHandle, true);
                this.RopSaveChangesMessage(cprptMessageHandle[0], true);
            }
            #endregion

            // Add sourceTags[1] to verify PropertyProblems structure
            PropertyTag[] sourceTags = new PropertyTag[2];
            sourceTags[0] = sourceTag[0].PropertyTag;
            sourceTags[1] = sourceTag[0].PropertyTag;
            sourceTags[1].PropertyId--;

            object responseObj = this.RopCopyProperties(sourceHandle, destHandle, (byte)HandleIndex.FirstIndex, (byte)HandleIndex.SecondIndex, wantAsynchronous, copyFlags, sourceTags);

            if (responseObj is RopProgressResponse)
            {
                isReturnedRopProgress = true;
            }
            else
            {
                RopCopyPropertiesResponse copyPropertiesResponse = (RopCopyPropertiesResponse)responseObj;

                if (copyPropertiesResponse.ReturnValue != (uint)CPRPTErrorCode.None)
                {
                    #region Return error
                    switch ((CPRPTErrorCode)copyPropertiesResponse.ReturnValue)
                    {
                        case CPRPTErrorCode.NullDestinationObject:
                            error = CPRPTErrorCode.NullDestinationObject;
                            break;
                        case CPRPTErrorCode.NotSupported:
                            error = CPRPTErrorCode.NotSupported;
                            break;
                        case CPRPTErrorCode.InvalidParameter:
                            error = CPRPTErrorCode.InvalidParameter;
                            break;
                        default:
                            error = CPRPTErrorCode.Other;
                            break;
                    }
                    #endregion

                    this.VerifyRopCopyProperties(copyPropertiesResponse, this.cprptCurrentType, copyFlag, false);
                }
                else
                {
                    #region Check if the property is changed in DB
                    string objKey;
                    switch (this.cprptCurrentType)
                    {
                        case ServerObjectType.Folder:
                            objKey = ConstValues.ServerObjectTypeFolderKey;
                            break;
                        case ServerObjectType.Message:
                            objKey = ConstValues.ServerObjectTypeMessageKey;
                            break;
                        case ServerObjectType.Attachment:
                            objKey = ConstValues.ServerObjectTypeAttachmentKey;
                            break;
                        default:
                            objKey = string.Empty;
                            break;
                    }

                    bool isPropertyFound;
                    byte[] propertyValueBeforeSave = this.Session2GetPropertyData(objKey, destTag[0].PropertyTag, out isPropertyFound);

                    isChangedInDB = !Common.CompareByteArray(propertyValueBeforeSave, this.tagPropertyValuesToVerification[objKey].Value);
                    #endregion

                    this.VerifyRopCopyProperties(copyPropertiesResponse, this.cprptCurrentType, copyFlag, isPropertyFound);

                    #region Check if property is overwritten
                    if (isDestinationExist)
                    {
                        RopGetPropertiesSpecificResponse overWriteRes = this.RopGetPropertiesSpecific(destHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, new PropertyTag[] { destTag[0].PropertyTag });
                        isOverwriteDestination = Common.CompareByteArray(overWriteRes.RowData.PropertyValues[0].Value, sourceTag[0].Value);
                    }
                    #endregion

                    #region Check if property is deleted
                    // The message and attachment objects must apply the changes to database
                    if (this.cprptCurrentType == ServerObjectType.Message)
                    {
                        this.RopSaveChangesMessage(sourceHandle, true);
                    }
                    else if (this.cprptCurrentType == ServerObjectType.Attachment)
                    {
                        this.RopSaveChangesAttachment(sourceHandle, true);
                        this.RopSaveChangesMessage(cprptMessageHandle[1], true);
                    }

                    switch (this.cprptCurrentType)
                    {
                        case ServerObjectType.Folder:
                            objKey = "Folder2";
                            break;
                        case ServerObjectType.Message:
                            objKey = "Message2";
                            break;
                        case ServerObjectType.Attachment:
                            objKey = "Attachment2";
                            break;
                        default:
                            objKey = string.Empty;
                            break;
                    }

                    if (this.cprptCurrentType == ServerObjectType.Folder)
                    {
                        RopGetPropertiesSpecificResponse tagValueRes = this.RopGetPropertiesSpecific(sourceHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, new PropertyTag[] { sourceTag[0].PropertyTag });
                        if (tagValueRes.RowData.PropertyValues[0].Value == null)
                        {
                            isPropertiesDeleted = true;
                        }
                        else
                        {
                            byte[] expectedNotFoundError = new byte[] { 0x0f, 0x01, 0x04, 0x80 };
                            isPropertiesDeleted = Common.CompareByteArray(expectedNotFoundError, tagValueRes.RowData.PropertyValues[0].Value);
                        }
                    }
                    else
                    {
                        if (this.cprptCurrentType == ServerObjectType.Attachment && copyFlag == CopyFlags.Move && Common.IsRequirementEnabled(86704, this.Site))
                        {
                            isPropertiesDeleted = true;
                        }
                        else
                        {
                            byte[] propertyValueNow = this.Session2GetPropertyData(objKey, sourceTag[0].PropertyTag, out isPropertyFound);
                            if (!isPropertyFound)
                            {
                                isPropertiesDeleted = true;
                            }
                            else
                            {
                                byte[] expectedNotFoundError = new byte[] { 0x0f, 0x01, 0x04, 0x80 };
                                isPropertiesDeleted = Common.CompareByteArray(expectedNotFoundError, propertyValueNow);
                            }
                        }
                    }
                    #endregion
                }
            }
        }

        /// <summary>
        /// This method is used to copy or move all but a specified few properties from a source object to a destination object. 
        /// </summary>
        /// <param name="copyFlag">Specifies the CopyFlags in the call request.</param>
        /// <param name="isWantAsynchronousZero">Indicates whether WantAsynchronous parameter in call request is zero.</param>
        /// <param name="isWantSubObjectsZero">Indicates whether WantSubObjects parameter in call request is zero.</param>
        /// <param name="isDestinationExist">Indicates whether destination object is exist for [RopCopyTo].</param>
        /// <param name="isPropertiesDeleted">If CopyFlags is set to Move,Source object will be deleted after copy to.</param>
        /// <param name="isSubObjectCopied">Indicates whether sub-object properties is also be copied.</param>
        /// <param name="isOverwriteDestination">If CopyFlags is set to NoOverWrite,Destination should not be overwritten.</param>
        /// <param name="isReturnedRopProgress">If this ROP is performed Asynchronously,RopProgress response
        /// returned instead of RopCopyProperties response.</param>
        /// <param name="isChangedInDB">Indicates whether destination is changed in database.</param>
        /// <param name="error">If destination object is not exist,NullDestinationObject error will be returned.</param>
        public void RopCopyToMethod(
            CopyFlags copyFlag,
            bool isWantAsynchronousZero,
            bool isWantSubObjectsZero,
            bool isDestinationExist,
            out bool isPropertiesDeleted,
            out bool isSubObjectCopied,
            out bool isOverwriteDestination,
            out bool isReturnedRopProgress,
            out bool isChangedInDB,
            out CPRPTErrorCode error)
        {
            // Check if there are conflict SHOULD/MAY settings 
            this.CheckRequirementConflict(86702, 18402);
            this.CheckRequirementConflict(86708, 5070506);

            #region Set parameters
            isPropertiesDeleted = false;
            isSubObjectCopied = false;
            isOverwriteDestination = false;
            isReturnedRopProgress = false;
            isChangedInDB = false;
            error = CPRPTErrorCode.None;

            uint secondHandle = InvalidHandle;
            if (isDestinationExist)
            {
                secondHandle = this.cprptSecondHandle;
                if (this.cprptCurrentType == ServerObjectType.Folder)
                {
                    RopCreateFolderResponse createFolderResponse;
                    string displayName = SubFolder;
                    secondHandle = this.RopCreateFolder(secondHandle, out createFolderResponse, displayName, "Comment: " + displayName, true);
                }
            }

            byte copyFlags = (byte)RopCopyToCopyFlags.None;

            switch (copyFlag)
            {
                case CopyFlags.None:
                    copyFlags = (byte)RopCopyToCopyFlags.None;
                    break;
                case CopyFlags.Move:
                    copyFlags = (byte)RopCopyToCopyFlags.Move;
                    break;
                case CopyFlags.NoOverWrite:
                    copyFlags = (byte)RopCopyToCopyFlags.NoOverwrite;
                    break;
                case CopyFlags.MoveAndNoOverWrite:
                    copyFlags = (byte)RopCopyPropertiesCopyFlags.MoveAndOverwrite;
                    break;
                default:
                    copyFlags = (byte)0x05;
                    break;
            }

            // The value indicates whether the ROP is to be processed synchronously or asynchronously, 0 means synchronous operation. 
            byte wantAsynchronous = 0;

            byte wantSubObjects = 1;
            if (isWantSubObjectsZero)
            {
                wantSubObjects = 0;
            }

            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject).PropertyTag;
            #endregion

            #region Set the first handle property
            TaggedPropertyValue[] tagPropertyValue = new TaggedPropertyValue[2];
            tagPropertyValue[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            tagPropertyValue[0].Value = Common.AddInt16LengthBeforeBinaryArray(tagPropertyValue[0].Value);
            tagPropertyValue[1] = this.GetTaggedPropertyTag(ObjectToOperate.SecondObject);
            tagPropertyValue[1].Value = Common.AddInt16LengthBeforeBinaryArray(tagPropertyValue[1].Value);

            this.RopSetProperties(this.cprptFirstHandle, tagPropertyValue, true);
            if (this.cprptCurrentType == ServerObjectType.Message)
            {
                this.RopSaveChangesMessage(this.cprptFirstHandle, true);
            }
            else if (this.cprptCurrentType == ServerObjectType.Attachment)
            {
                this.RopSaveChangesAttachment(this.cprptFirstHandle, true);
                this.RopSaveChangesMessage(cprptMessageHandle[0], true);
            }
            #endregion

            #region Set the second handle property
            TaggedPropertyValue[] tagPropertyValue1 = new TaggedPropertyValue[2];
            tagPropertyValue1[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject);
            tagPropertyValue1[0].Value[0] = (byte)(tagPropertyValue1[0].Value[0] + 1);
            tagPropertyValue1[0].Value = Common.AddInt16LengthBeforeBinaryArray(tagPropertyValue1[0].Value);
            tagPropertyValue1[1] = this.GetTaggedPropertyTag(ObjectToOperate.SecondObject);
            tagPropertyValue1[1].Value[0] = (byte)(tagPropertyValue1[1].Value[0] + 1);
            tagPropertyValue1[1].Value = Common.AddInt16LengthBeforeBinaryArray(tagPropertyValue1[1].Value);
            if (isDestinationExist)
            {
                this.RopSetProperties(secondHandle, tagPropertyValue1, true);
                if (this.cprptCurrentType == ServerObjectType.Message)
                {
                    this.RopSaveChangesMessage(secondHandle, true);
                }
                else if (this.cprptCurrentType == ServerObjectType.Attachment)
                {
                    this.RopSaveChangesAttachment(secondHandle, true);
                    this.RopSaveChangesMessage(cprptMessageHandle[0], true);
                }

                // Store the TaggedPropertyValue for further verification
                this.tagPropertyValuesToVerification.Clear();
                switch (this.cprptCurrentType)
                {
                    case ServerObjectType.Folder:
                        this.tagPropertyValuesToVerification.Add(ConstValues.ServerObjectTypeFolderKey, tagPropertyValue1[0]);
                        break;
                    case ServerObjectType.Message:
                        this.tagPropertyValuesToVerification.Add(ConstValues.ServerObjectTypeMessageKey, tagPropertyValue1[0]);
                        break;
                    case ServerObjectType.Attachment:
                        this.tagPropertyValuesToVerification.Add(ConstValues.ServerObjectTypeAttachmentKey, tagPropertyValue1[0]);
                        break;
                    default:
                        break;
                }
            }
            #endregion

            if (!isWantAsynchronousZero)
            {
                wantAsynchronous = 1;
                if (copyFlag != CopyFlags.NoOverWrite && this.cprptCurrentType == ServerObjectType.Folder && !isWantSubObjectsZero)
                {
                    RopCreateMessageResponse createMessageResponse;
                    uint msgHandle = this.RopCreateMessage(cprptFolderHandle[0], cprptFolderId[0], ConstValues.RopCreateMessageAssociatedFlagNone, out createMessageResponse, true);
                    this.RopSaveChangesMessage(msgHandle, true);
                    this.RopRelease(msgHandle);
                }
            }

            // The property set in previous step should exist.
            TaggedPropertyValue existingValue = this.SearchProperty(this.cprptFirstHandle, tagPropertyValue[0].PropertyTag.PropertyId);
            Site.Assert.IsNotNull(existingValue, "The target property should not be null.");

            bool isExcluedPropertyCopied = true;
            PropertyTag[] propertyTag = new PropertyTag[1];
            propertyTag[0] = tagPropertyValue1[1].PropertyTag;
            object responseObj = this.RopCopyTo(this.cprptFirstHandle, secondHandle, (byte)HandleIndex.FirstIndex, (byte)HandleIndex.SecondIndex, wantAsynchronous, wantSubObjects, copyFlags, propertyTag);

            if (isDestinationExist)
            {
                RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = this.RopGetPropertiesSpecific(secondHandle, 0, 0, propertyTag);
                if (!Common.CompareByteArray(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, tagPropertyValue[0].Value))
                {
                    isExcluedPropertyCopied = false;
                }

                this.VerifyRopCopyToExcludeElement(isExcluedPropertyCopied);
            }

            if (responseObj is RopProgressResponse)
            {
                isReturnedRopProgress = true;
            }
            else
            {
                RopCopyToResponse ropCopyToResponse = (RopCopyToResponse)responseObj;

                this.VerifyRopCopyTo(ropCopyToResponse, CopyToCondition.Normal, this.cprptCurrentType, copyFlag);

                error = (CPRPTErrorCode)ropCopyToResponse.ReturnValue;

                if (!isReturnedRopProgress && isWantAsynchronousZero && ropCopyToResponse.ReturnValue == (uint)RopResponseType.SuccessResponse)
                {
                    #region Check if the property is changed in DB
                    string objKey;
                    switch (this.cprptCurrentType)
                    {
                        case ServerObjectType.Folder:
                            objKey = "Folder2";
                            break;
                        case ServerObjectType.Message:
                            objKey = "Message2";
                            break;
                        case ServerObjectType.Attachment:
                            objKey = "Attachment2";
                            break;
                        default:
                            objKey = string.Empty;
                            break;
                    }

                    bool isPropertyFound;
                    byte[] propertyValueBeforeSave = this.Session2GetPropertyData(objKey, tagPropertyValue[0].PropertyTag, out isPropertyFound);

                    isChangedInDB = !Common.CompareByteArray(propertyValueBeforeSave, tagPropertyValue1[0].Value);
                    #endregion

                    #region Check if the value is overwritten
                    TaggedPropertyValue overwriteValue = this.SearchProperty(secondHandle, tagPropertyValue1[0].PropertyTag.PropertyId);
                    if (Common.CompareByteArray(overwriteValue.Value, tagPropertyValue1[0].Value))
                    {
                        isOverwriteDestination = false;
                    }
                    else
                    {
                        isOverwriteDestination = true;
                    }
                    #endregion

                    #region Check if the sub-object is copied
                    if (!isWantSubObjectsZero && isDestinationExist)
                    {
                        if (this.cprptCurrentType == ServerObjectType.Folder)
                        {
                            TaggedPropertyValue ptagValue = this.SearchProperty(secondHandle, (ushort)TaggedPropertyName.PidTagContentCount);
                            if (ptagValue.Value[0] != 0 || ptagValue.Value[1] != 0 || ptagValue.Value[2] != 0 || ptagValue.Value[3] != 0)
                            {
                                isSubObjectCopied = true;
                            }
                        }
                        else if (this.cprptCurrentType == ServerObjectType.Message)
                        {
                            RopOpenAttachmentResponse openAttRes;
                            this.RopOpenAttachment(secondHandle, cprptAttachmentId[0], out openAttRes, true);
                            if (openAttRes.ReturnValue == (uint)CPRPTErrorCode.None)
                            {
                                isSubObjectCopied = true;
                            }
                        }
                    }
                    #endregion

                    #region Check if the value is deleted
                    // The message and attachment objects must apply the changes to database
                    if (this.cprptCurrentType == ServerObjectType.Message)
                    {
                        this.RopSaveChangesMessage(this.cprptFirstHandle, true);
                    }
                    else if (this.cprptCurrentType == ServerObjectType.Attachment)
                    {
                        this.RopSaveChangesAttachment(this.cprptFirstHandle, false);
                        this.RopSaveChangesMessage(cprptMessageHandle[0], false);
                    }

                    switch (this.cprptCurrentType)
                    {
                        case ServerObjectType.Folder:
                            objKey = ConstValues.ServerObjectTypeFolderKey;
                            break;
                        case ServerObjectType.Message:
                            objKey = ConstValues.ServerObjectTypeMessageKey;
                            break;
                        case ServerObjectType.Attachment:
                            objKey = ConstValues.ServerObjectTypeAttachmentKey;
                            break;
                        default:
                            objKey = string.Empty;
                            break;
                    }

                    if (this.cprptCurrentType == ServerObjectType.Folder)
                    {
                        TaggedPropertyValue deleteValue = this.SearchProperty(this.cprptFirstHandle, tagPropertyValue[0].PropertyTag.PropertyId);
                        if (deleteValue == null)
                        {
                            isPropertiesDeleted = true;
                        }
                    }
                    else
                    {
                        if (this.cprptCurrentType == ServerObjectType.Attachment && copyFlag == CopyFlags.Move && Common.IsRequirementEnabled(86707, this.Site))
                        {
                            isPropertiesDeleted = true;
                        }
                        else
                        {
                            byte[] propertyValueNow = this.Session2GetPropertyData(objKey, tagPropertyValue[0].PropertyTag, out isPropertyFound);
                            if (!isPropertyFound)
                            {
                                isPropertiesDeleted = true;
                            }
                            else
                            {
                                byte[] expectedNotFoundError = new byte[] { 0x0f, 0x01, 0x04, 0x80 };
                                isPropertiesDeleted = Common.CompareByteArray(expectedNotFoundError, propertyValueNow);
                            }
                        }
                    }
                    #endregion
                }
            }
        }

        /// <summary>
        /// This method is used to report the progress status of an asynchronous operation. 
        /// </summary>
        /// <param name="isOtherRopSent">Indicates whether other ROP is sent.</param>
        /// <param name="isWantCancel">Indicates whether WantCancel parameter is set to non-zero,any 
        /// non-zero value means client want cancel the original operation.</param>
        /// <param name="isOriginalOpsResponse">If original asynchronous ROPs are done or canceled,
        /// Response should be original ROPs response.Otherwise, it should be RopProgress response.</param>
        /// <param name="isOtherRopResponse">indicates the other ROP's response is returned.If the client sends a 
        /// ROP other than RopProgress to the server with the same logon before the asynchronous operation is 
        /// complete the server MUST abort the asynchronous operation and respond to the new ROP.</param>
        public void RopProgressMethod(bool isOtherRopSent, bool isWantCancel, out bool isOriginalOpsResponse, out bool isOtherRopResponse)
        {
            isOriginalOpsResponse = false;
            isOtherRopResponse = false;

            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0] = this.GetTaggedPropertyTag(ObjectToOperate.FirstObject).PropertyTag;
            if (isOtherRopSent)
            {
                this.RopGetPropertiesSpecific(this.cprptFirstHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, propertyTags);
                isOtherRopResponse = true;
            }
            else
            {
                byte wantCancel = 0;
                if (isWantCancel)
                {
                    wantCancel = 1;
                }

                object processRes = this.RopProgress(this.cprptFirstHandle, wantCancel);
                while (processRes is RopProgressResponse)
                {
                    processRes = this.RopProgress(this.cprptFirstHandle, wantCancel);
                }

                if (processRes is RopCopyPropertiesResponse || processRes is RopCopyToResponse)
                {
                    isOriginalOpsResponse = true;
                }
            }
        }

        /// <summary>
        ///   Get common object properties in order to test their type.
        /// </summary>
        /// <param name="commonProperty">The nine Common Object Properties defined in section 2.2.1.</param>
        public void GetCommonObjectProperties(CommonObjectProperty commonProperty)
        {
            PropertyTag[] proTag = new PropertyTag[1];

            proTag[0] = this.commonProperties[commonProperty];
            RopGetPropertiesSpecificResponse getPropertiesSecificResponse = this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, proTag);

            if (getPropertiesSecificResponse.ReturnValue == (uint)CPRPTErrorCode.None)
            {
                this.VerifyCommonProperties(proTag[0], getPropertiesSecificResponse, this.cprptCurrentType);
                this.VerifyCommonPropertiesInOXCDATA(proTag[0], getPropertiesSecificResponse);
                this.VerifyTypeOfCommonPropertiesInOXPROPS(proTag[0], getPropertiesSecificResponse);
            }
        }

        /// <summary>
        ///  Set common object properties in order to test whether each of them is read-only.
        /// </summary>
        /// <param name="commonProperty">The nine Common Object Properties defined in section 2.2.1</param>
        /// <param name="error">When a property is specified as "read-only for the client", the server MUST
        /// return an error and ignore any request to change the value of that property.</param>
        public void SetCommonObjectProperties(CommonObjectProperty commonProperty, out CPRPTErrorCode error)
        {
            PropertyTag[] proTag = new PropertyTag[1];

            proTag[0] = this.commonProperties[commonProperty];

            this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, proTag);

            error = CPRPTErrorCode.None;
            TaggedPropertyValue[] tags = new TaggedPropertyValue[1];
            tags[0] = new TaggedPropertyValue
            {
                PropertyTag = this.commonProperties[commonProperty]
            };

            switch ((PropertyTypeName)tags[0].PropertyTag.PropertyType)
            {
                case PropertyTypeName.PtypInteger32:
                    tags[0].Value = BitConverter.GetBytes(Convert.ToInt32(CommonValue));
                    break;
                case PropertyTypeName.PtypBinary:
                    tags[0].Value = Common.AddInt16LengthBeforeBinaryArray(Encoding.Unicode.GetBytes(CommonValue));
                    break;
                case PropertyTypeName.PtypTime:
                    tags[0].Value = new byte[ConstValues.PtypTimeSizeInByte];
                    for (int i = 0; i < ConstValues.PtypTimeSizeInByte; i++)
                    {
                        tags[0].Value[i] = (byte)i;
                    }

                    break;
                case PropertyTypeName.PtypString:
                    tags[0].Value = Encoding.Unicode.GetBytes(CommonValue + '\0');
                    break;
                default:
                    break;
            }

            RopSetPropertiesResponse setPtyRes = this.RopSetProperties(this.cprptCurrentHandle, tags, false);
            if (setPtyRes.ReturnValue != (uint)CPRPTErrorCode.None)
            {
                if (setPtyRes.ReturnValue == (uint)CPRPTErrorCode.AccessDenied)
                {
                    error = CPRPTErrorCode.AccessDenied;
                }
                else
                {
                    error = CPRPTErrorCode.Other;
                }
            }
            else
            {
                if (this.cprptCurrentType == ServerObjectType.Message)
                {
                    RopSaveChangesMessageResponse saveChangesMsgRes = this.RopSaveChangesMessage(this.cprptCurrentHandle, false);
                    if (saveChangesMsgRes.ReturnValue == (uint)CPRPTErrorCode.GeneralFailure)
                    {
                        error = CPRPTErrorCode.GeneralFailure;
                    }
                    else
                    {
                        this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, proTag);
                    }
                }
                else if (this.cprptCurrentType == ServerObjectType.Attachment)
                {
                    RopSaveChangesAttachmentResponse saveAttachmentRes = this.RopSaveChangesAttachment(this.cprptCurrentHandle, false);
                    if (saveAttachmentRes.ReturnValue == (uint)CPRPTErrorCode.GeneralFailure)
                    {
                        error = CPRPTErrorCode.GeneralFailure;
                    }
                    else
                    {
                        this.RopGetPropertiesSpecific(this.cprptCurrentHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, proTag);
                    }
                }
            }
        }

        /// <summary>
        /// Checks if the requirement is enabled in SHOULD/MAY configuration file.
        /// </summary>
        /// <param name="rsid">Requirement ID</param>
        /// <param name="enabled">True represents the requirement is enabled; false represents the requirement is disabled.</param>
        public void CheckRequirementEnabled(int rsid, out bool enabled)
        {
            enabled = Common.IsRequirementEnabled(rsid, this.Site);
        }

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        public void CheckMAPIHTTPTransportSupported(out bool isSupported)
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && !Common.IsRequirementEnabled(5000, this.Site))
            {
                isSupported = false;
                this.needDoCleanup = false;
                Site.Assert.Inconclusive("This case can not run on MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
            else 
            {
                isSupported = true;
            }
        }

        #endregion

        #region Private functions
        /// <summary>
        /// Test cleanup operation.
        /// </summary>
        private void CleanUpMethod()
        {
            if (this.needDoCleanup)
            {
                if (this.isPublicFolderLogonType)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        string subFolderName = FolderDisplayName + i.ToString();

                        // Get the subfolder ID according to its name.
                        ulong folderId = this.GetSubfolderIDByName(this.existingPublicFolderHandle, subFolderName);

                        // Cleanup the created public folder, including all messages and subfolders
                        this.RopDeleteFolder(this.existingPublicFolderHandle, folderId);
                    }
                }
                else
                {
                    // Cleanup the inbox folder, including all messages and subfolders
                    this.RopEmptyFolder(this.cprptInputBoxHandle, Convert.ToByte(false), true);
                }

                this.RopRelease(this.cprptLogonHandle);
            }

            bool isDisconnected = this.RpcDisconnect();
            this.Site.Assert.IsTrue(isDisconnected, "RPC disconnect should be executed successfully.");
        }

        /// <summary>
        /// Find a folder ID in the specified folder and with specified display name.
        /// </summary>
        /// <param name="openedFolderHandle">Handle of the parent folder.</param>
        /// <param name="folderName">The folder display name.</param>
        /// <returns>The folder ID.</returns>
        private ulong GetSubfolderIDByName(uint openedFolderHandle, string folderName)
        {
            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest();
            RopGetHierarchyTableResponse getHierarchyTableResponse = new RopGetHierarchyTableResponse();
            getHierarchyTableRequest.RopId = (byte)RopId.RopGetHierarchyTable;
            getHierarchyTableRequest.LogonId = 0x0;
            getHierarchyTableRequest.InputHandleIndex = 0x0;
            getHierarchyTableRequest.OutputHandleIndex = 0x1;
            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.Depth;

            this.responseSOHsValue = this.ProcessSingleRop(
                getHierarchyTableRequest,
                openedFolderHandle,
                ref this.responseValue,
                ref this.rawDataValue,
                RopResponseType.SuccessResponse);
            getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.responseValue;
            uint tableHandle = this.responseSOHsValue[0][getHierarchyTableResponse.OutputHandleIndex];

            PropertyTag[] properties = new PropertyTag[]
            {
                new PropertyTag()
                {
                    PropertyId = this.commonProperties[CommonObjectProperty.PidTagDisplayName].PropertyId,
                    PropertyType = (ushort)PropertyType.PtypString
                },
                new PropertyTag()
                {
                    PropertyId = (ushort)this.commonProperties[CommonObjectProperty.PidTagFolderId].PropertyId,
                    PropertyType = (ushort)PropertyType.PtypInteger64
                }
            };
            List<PropertyRow> propertyRows = this.GetTableRowValue(tableHandle, (ushort)getHierarchyTableResponse.RowCount, properties);

            ulong folderId = 0;
            foreach (PropertyRow propertyRow in propertyRows)
            {
                byte[] displayNameInBytes = propertyRow.PropertyValues[0].Value;
                string displayName = Encoding.Unicode.GetString(displayNameInBytes, 0, displayNameInBytes.Length).Replace("\0", string.Empty);
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
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                PropertyTagCount = (ushort)properties.Length,
                PropertyTags = properties,
                SetColumnsFlags = (byte)AsynchronousFlags.None
            };
            this.responseSOHsValue = this.ProcessSingleRop(
                    setColumnsRequest,
                    tableHandle,
                    ref this.responseValue,
                    ref this.rawDataValue,
                    RopResponseType.SuccessResponse);
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)this.responseValue;

            Site.Assert.AreEqual<uint>(
                SuccessReturnValue,
                setColumnsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #region The client calls RopQueryRows operation to query the folder which have the special properties.

            RopQueryRowsRequest queryRowsRequest = new RopQueryRowsRequest
            {
                RopId = (byte)RopId.RopQueryRows,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                RowCount = (ushort)rowCount,
                QueryRowsFlags = (byte)QueryRowsFlags.Advance,
                ForwardRead = 0x01
            };
            this.responseSOHsValue = this.ProcessSingleRop(
                    queryRowsRequest,
                    tableHandle,
                    ref this.responseValue,
                    ref this.rawDataValue,
                    RopResponseType.SuccessResponse);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.responseValue;

            Site.Assert.AreEqual<uint>(
                SuccessReturnValue,
                queryRowsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            return queryRowsResponse.RowData.PropertyRows;
        }

        /// <summary>
        /// Initialize the test environment.
        /// </summary>
        /// <param name="site">TestSite is used to initialization.</param>
        private void InitHandle(ITestSite site)
        {
            #region Open TestRootFolder
            this.Initialize(site);
            bool isConnected = this.RpcConnect(LogonType.Mailbox);
            this.Site.Assert.IsTrue(isConnected, "Rpc connect method should be executed successfully.");

            // Logon to the private mailbox.
            RopLogonResponse logonRes;
            uint logonHandle = this.RopLogon(LogonType.Mailbox, out logonRes, Common.GetConfigurationPropertyValue("UserEssdn", site), true);

            // Open inbox folder.
            RopOpenFolderResponse openRes;
            uint inboxHandle = this.RopOpenFolder(logonHandle, out openRes, logonRes.FolderIds[4], true);

            // Empty inbox folder.
            this.RopEmptyFolder(inboxHandle, Convert.ToByte(false), true);

            // Create root folder.
            RopCreateFolderResponse testFolderRes;
            this.cprptTestFolderName = RootFolderName;
            uint testRootFolderHandle = this.RopCreateFolder(inboxHandle, out testFolderRes, this.cprptTestFolderName, "Comment: " + this.cprptTestFolderName, true);
            rootFolderId = testFolderRes.FolderId;
            #endregion

            #region Create sub folder
            RopCreateFolderResponse createFolderResponse;
            for (int i = 0; i < 3; i++)
            {
                string displayName = FolderDisplayName + i.ToString();
                cprptFolderHandle[i] = this.RopCreateFolder(testRootFolderHandle, out createFolderResponse, displayName, "Comment: " + displayName, true);
                cprptFolderId[i] = createFolderResponse.FolderId;
                this.RopEmptyFolder(cprptFolderHandle[i], Convert.ToByte(false), true);
            }

            #endregion

            #region Create message and attachment
            RopCreateMessageResponse createMessageResponse;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            for (int i = 0; i < 3; i++)
            {
                // Create message.
                cprptMessageHandle[i] = this.RopCreateMessage(cprptFolderHandle[0], cprptFolderId[0], ConstValues.RopCreateMessageAssociatedFlagNone, out createMessageResponse, true);
                if (i == 0)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        // Create attachment in the first message.
                        RopCreateAttachmentResponse createAttachmentResponse;
                        cprptAttachmentHandle[j] = this.RopCreateAttachment(cprptMessageHandle[0], out createAttachmentResponse, true);
                        cprptAttachmentId[j] = createAttachmentResponse.AttachmentID;
                        this.RopSaveChangesAttachment(cprptAttachmentHandle[j], true);
                    }
                }

                saveChangesMessageResponse = this.RopSaveChangesMessage(cprptMessageHandle[i], true);
                cprptMessageId[i] = saveChangesMessageResponse.MessageId;
            }
            #endregion

            this.RopRelease(logonHandle);
            bool isDisconnected = this.RpcDisconnect();
            this.Site.Assert.IsTrue(isDisconnected, "RPC disconnect should be executed successfully.");
        }

        /// <summary>
        /// Initialize the test environment.
        /// </summary>
        /// <param name="site">TestSite is used to initialization.</param>
        /// <param name="logonType">This type specifies ongoing action on the mailbox or public folder.</param>
        private void InitHandle(ITestSite site, LogonType logonType)
        {
            this.Initialize(site);
            bool isConnected = this.RpcConnect(logonType);
            this.Site.Assert.IsTrue(isConnected, "Rpc connect method should be executed successfully.");

            // Log on to the specific folder.
            RopLogonResponse logonRes;
            this.cprptLogonHandle = this.logonHandle = this.RopLogon(logonType, out logonRes, Common.GetConfigurationPropertyValue("UserEssdn", site), true);
            this.folderIds = logonRes.FolderIds;
        }

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="logonType">The logon type.</param>
        /// <returns>Result of connecting.</returns>
        private bool RpcConnect(LogonType logonType)
        {
            ConnectionType connectionType = ConnectionType.PrivateMailboxServer;
            if (logonType == LogonType.PublicFolder)
            {
                connectionType = ConnectionType.PublicFolderServer;
            }

            bool retValue1 = this.oxcropsClient.Connect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    connectionType,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("UserName", this.Site),
                    Common.GetConfigurationPropertyValue("Password", this.Site));

            bool retValue2 = this.oxcropsClientSession2.Connect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    connectionType,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("UserName", this.Site),
                    Common.GetConfigurationPropertyValue("Password", this.Site));

            return retValue1 && retValue2;
        }

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        /// <returns>Result of disconnecting.</returns>
        private bool RpcDisconnect()
        {
            bool retValue1 = this.oxcropsClient.Disconnect();
            bool retValue2 = this.oxcropsClientSession2.Disconnect();

            return retValue1 && retValue2;
        }

        /// <summary>
        /// Method which executes single ROP.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <returns>Server objects handles in response.</returns>
        private List<List<uint>> ProcessSingleRop(
            ISerializable ropRequest,
            uint insideObjHandle,
            ref IDeserializable response,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };

            List<uint> requestSOH = new List<uint>
            {
                insideObjHandle
            };

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(DefaultOutputHandle);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            uint ret = this.oxcropsClient.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, MaxRgbOut);
            this.Site.Assert.AreEqual<uint>(OxcRpcErrorCode.ECNone, ret, "ROP call should return 0 for success, actually it returns {0}", ret);

            this.VerifyMAPITransport();

            if (responseRops != null)
            {
                if (responseRops.Count > 0)
                {
                    response = responseRops[0];
                }
            }
            else
            {
                response = null;
            }

            if (ropRequest.GetType() == typeof(RopReleaseRequest))
            {
                return responseSOHs;
            }

            try
            {
                string resName = response.GetType().Name;

                // The word "Response" takes 8 length.
                string ropName = resName.Substring(0, resName.Length - 8);
                Type adapterType = typeof(MS_OXCPRPTAdapter);

                // Call capture code using reflection mechanism. 
                // The code followed is to construct the verify method name of capture code and then call this method through reflection.
                string verifyMethodName = string.Empty;
                if (expectedRopResponseType == RopResponseType.SuccessResponse)
                {
                    verifyMethodName = "Verify" + ropName + "SuccessResponse";
                }
                else if (expectedRopResponseType == RopResponseType.FailureResponse)
                {
                    verifyMethodName = "Verify" + ropName + "FailureResponse";
                }
                else if (expectedRopResponseType == RopResponseType.Response)
                {
                    verifyMethodName = "Verify" + ropName + "Response";
                }
                else if (expectedRopResponseType == RopResponseType.NullDestinationFailureResponse)
                {
                    verifyMethodName = "Verify" + ropName + "NullDestinationFailureResponse";
                }
                else if (expectedRopResponseType == RopResponseType.RedirectResponse)
                {
                    verifyMethodName = "Verify" + ropName + "RedirectResponse";
                }

                Type reqType = ropRequest.GetType();
                MethodInfo method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                if (method == null)
                {
                    if (expectedRopResponseType == RopResponseType.SuccessResponse
                        || expectedRopResponseType == RopResponseType.FailureResponse)
                    {
                        verifyMethodName = "Verify" + ropName + "Response";
                        method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                    }
                }

                if (method != null)
                {
                    ParameterInfo[] paraInfos = method.GetParameters();
                    int paraNum = paraInfos.Length;
                    object[] paraObjects = new object[paraNum];
                    paraObjects[0] = response;
                    for (int i = 1; i < paraNum; i++)
                    {
                        FieldInfo fieldInReq = reqType.GetField(
                            paraInfos[i].Name,
                            BindingFlags.IgnoreCase
                            | BindingFlags.DeclaredOnly
                            | BindingFlags.Public
                            | BindingFlags.NonPublic
                            | BindingFlags.GetField
                            | BindingFlags.Instance);
                        paraObjects[i] = fieldInReq.GetValue(ropRequest);
                    }

                    method.Invoke(this, paraObjects);
                }
            }
            catch (TargetInvocationException invocationEx)
            {
                Site.Log.Add(LogEntryKind.Debug, invocationEx.Message);
                if (invocationEx.InnerException != null)
                {
                    throw invocationEx.InnerException;
                }
            }
            catch (NullReferenceException nullEx)
            {
                Site.Log.Add(LogEntryKind.Debug, nullEx.Message);
            }

            return responseSOHs;
        }

        /// <summary>
        /// Method which executes single ROP with multiple server objects.
        /// </summary>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="insideObjHandle">Server object handles in request.</param>
        /// <param name="response">ROP response object.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <returns>Server objects handles in response.</returns>
        private List<List<uint>> ProcessSingleRopWithMutipleServerObjects(
            ISerializable ropRequest,
            List<uint> insideObjHandle,
            ref IDeserializable response,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };

            List<uint> requestSOH = new List<uint>();
            for (int i = 0; i < insideObjHandle.Count; i++)
            {
                requestSOH.Add(insideObjHandle[i]);
            }

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(DefaultOutputHandle);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            uint ret = this.oxcropsClient.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, MaxRgbOut);
            this.Site.Assert.AreEqual<uint>(OxcRpcErrorCode.ECNone, ret, "ROP call should return 0 for success, actually it returns {0}", ret);
            if (responseRops != null)
            {
                if (responseRops.Count > 0)
                {
                    response = responseRops[0];
                }
            }
            else
            {
                response = null;
            }

            if (ropRequest is RopReleaseRequest)
            {
                return responseSOHs;
            }

            try
            {
                string resName = response.GetType().Name;

                // The word "Response" takes 8 length.
                string ropName = resName.Substring(0, resName.Length - 8);
                Type adapterType = typeof(MS_OXCPRPTAdapter);

                // Call capture code using reflection mechanism
                // The code followed is to construct the verify method name of capture code and then call this method through reflection.
                string verifyMethodName = string.Empty;
                if (expectedRopResponseType == RopResponseType.SuccessResponse)
                {
                    verifyMethodName = "Verify" + ropName + "SuccessResponse";
                }
                else if (expectedRopResponseType == RopResponseType.FailureResponse)
                {
                    verifyMethodName = "Verify" + ropName + "FailureResponse";
                }
                else if (expectedRopResponseType == RopResponseType.Response)
                {
                    verifyMethodName = "Verify" + ropName + "Response";
                }
                else if (expectedRopResponseType == RopResponseType.NullDestinationFailureResponse)
                {
                    verifyMethodName = "Verify" + ropName + "NullDestinationFailureResponse";
                }
                else if (expectedRopResponseType == RopResponseType.RedirectResponse)
                {
                    verifyMethodName = "Verify" + ropName + "RedirectResponse";
                }

                Type reqType = ropRequest.GetType();
                MethodInfo method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                if (method == null)
                {
                    if (expectedRopResponseType == RopResponseType.SuccessResponse || expectedRopResponseType == RopResponseType.FailureResponse)
                    {
                        verifyMethodName = "Verify" + ropName + "Response";
                        method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                    }
                }

                if (method != null)
                {
                    ParameterInfo[] paraInfos = method.GetParameters();
                    int paraNum = paraInfos.Length;
                    object[] paraObjects = new object[paraNum];
                    paraObjects[0] = response;
                    for (int i = 1; i < paraNum; i++)
                    {
                        FieldInfo fieldInReq = reqType.GetField(
                            paraInfos[i].Name,
                            BindingFlags.IgnoreCase
                            | BindingFlags.DeclaredOnly
                            | BindingFlags.Public
                            | BindingFlags.NonPublic
                            | BindingFlags.GetField
                            | BindingFlags.Instance);
                        paraObjects[i] = fieldInReq.GetValue(ropRequest);
                    }

                    method.Invoke(this, paraObjects);
                }
            }
            catch (TargetInvocationException invocationEx)
            {
                Site.Log.Add(LogEntryKind.Debug, invocationEx.Message);
                if (invocationEx.InnerException != null)
                {
                    throw invocationEx.InnerException;
                }
            }
            catch (NullReferenceException nullEx)
            {
                Site.Log.Add(LogEntryKind.Debug, nullEx.Message);
            }

            return responseSOHs;
        }

        /// <summary>
        /// Obtain Property value via session2.
        /// </summary>
        /// <param name="objKey">String key to Indicates which object to be acted on.</param>
        /// <param name="propertyTag">Property Tag.</param>
        /// <param name="isPropertyFound">Indicates whether the property exists.</param>
        /// <returns>Property value.</returns>
        private byte[] Session2GetPropertyData(string objKey, PropertyTag propertyTag, out bool isPropertyFound)
        {
            Dictionary<string, uint> objHandles = new Dictionary<string, uint>();
            uint logonHandle;
            uint folderHandle1;
            uint folderHandle2;
            uint msgHandle1;
            uint msgHandle2;
            uint attachmentHandle1;
            uint attachmentHandle2;

            OxcropsClient tempClient = this.oxcropsClient;
            this.oxcropsClient = this.oxcropsClientSession2;

            #region Preparation for getting object handles
            RopLogonResponse logonRes;
            RopOpenFolderResponse openFolderRes;
            RopOpenMessageResponse openMsgRes;
            RopOpenAttachmentResponse openAttRes;

            // Logon to mailbox.
            logonHandle = this.RopLogon(LogonType.Mailbox, out logonRes, Common.GetConfigurationPropertyValue("UserEssdn", this.Site), true);
            objHandles.Add("Logon", logonHandle);

            // Open inbox folder.
            uint inboxHandle = this.RopOpenFolder(logonHandle, out openFolderRes, logonRes.FolderIds[4], true);

            // Open test root folder.
            uint testRootHandle = this.RopOpenFolder(inboxHandle, out openFolderRes, rootFolderId, true);

            // Open testFolder1.
            folderHandle1 = this.RopOpenFolder(testRootHandle, out openFolderRes, cprptFolderId[0], true);
            objHandles.Add("Folder1", folderHandle1);

            // Open testFolder2.
            folderHandle2 = this.RopOpenFolder(testRootHandle, out openFolderRes, cprptFolderId[1], true);
            objHandles.Add("Folder2", folderHandle2);

            // For first folder open message1.
            msgHandle1 = this.RopOpenMessage(folderHandle1, cprptFolderId[0], cprptMessageId[0], out openMsgRes, true);
            objHandles.Add("Message1", msgHandle1);

            // For first folder open message2.
            msgHandle2 = this.RopOpenMessage(folderHandle1, cprptFolderId[0], cprptMessageId[1], out openMsgRes, true);
            objHandles.Add("Message2", msgHandle2);

            // For first message open attachment1.
            attachmentHandle1 = this.RopOpenAttachment(msgHandle1, cprptAttachmentId[0], out openAttRes, true);
            objHandles.Add("Attachment1", attachmentHandle1);

            // For first message open attachment2.
            attachmentHandle2 = this.RopOpenAttachment(msgHandle1, cprptAttachmentId[1], out openAttRes, true);
            objHandles.Add("Attachment2", attachmentHandle2);
            #endregion

            if (!objHandles.ContainsKey(objKey))
            {
                this.oxcropsClient = tempClient;
                this.Site.Assert.Fail("Invalid Object key for getting its Property value");
            }

            uint objHandle = objHandles[objKey];

            // The RopGetPropertiesAll is used to obtain all properties.
            RopGetPropertiesSpecificResponse getPropSpecRes = this.RopGetPropertiesSpecific(objHandle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo, new PropertyTag[] { propertyTag });

            byte[] expectedNotFoundError = BitConverter.GetBytes((uint)CPRPTErrorCode.NotFound);
            if (Common.CompareByteArray(getPropSpecRes.RowData.PropertyValues[0].Value, expectedNotFoundError))
            {
                isPropertyFound = false;
            }
            else
            {
                isPropertyFound = true;
            }

            this.oxcropsClient = tempClient;

            return getPropSpecRes.RowData.PropertyValues[0].Value;
        }

        /// <summary>
        /// Session get property value via RopReadStream.
        /// </summary>
        /// <param name="objKey">String key to Indicates which object to be acted on.</param>
        /// <param name="propertyTag">Property Tag.</param>
        /// <param name="isPropertyFound">Indicates whether the property exists.</param>
        /// <returns>Property value in a stream.</returns>
        private byte[] Session2ReadStream(string objKey, PropertyTag propertyTag, out bool isPropertyFound)
        {
            Dictionary<string, uint> objHandles = new Dictionary<string, uint>();
            uint logonHandle;
            uint folderHandle1;
            uint folderHandle2;
            uint msgHandle1;
            uint msgHandle2;
            uint attachmentHandle1;
            uint attachmentHandle2;

            OxcropsClient tempClient = this.oxcropsClient;
            this.oxcropsClient = this.oxcropsClientSession2;

            try
            {
                #region Preparation for getting object handles
                RopLogonResponse logonRes;
                RopOpenFolderResponse openFolderRes;
                RopOpenMessageResponse openMsgRes;
                RopOpenAttachmentResponse openAttRes;

                // Logon to mailbox.
                logonHandle = this.RopLogon(LogonType.Mailbox, out logonRes, Common.GetConfigurationPropertyValue("UserEssdn", this.Site), true);
                objHandles.Add("Logon", logonHandle);

                // Open inbox folder.
                uint inboxHandle = this.RopOpenFolder(logonHandle, out openFolderRes, logonRes.FolderIds[4], true);

                // Open testRoot folder.
                uint testRootHandle = this.RopOpenFolder(inboxHandle, out openFolderRes, rootFolderId, true);

                // Open testFolder1.
                folderHandle1 = this.RopOpenFolder(testRootHandle, out openFolderRes, cprptFolderId[0], true);
                objHandles.Add("Folder1", folderHandle1);

                // Open testFolder2.
                folderHandle2 = this.RopOpenFolder(testRootHandle, out openFolderRes, cprptFolderId[1], true);
                objHandles.Add("Folder2", folderHandle2);

                // For first folder open message1.
                msgHandle1 = this.RopOpenMessage(folderHandle1, cprptFolderId[0], cprptMessageId[0], out openMsgRes, true);
                objHandles.Add("Message1", msgHandle1);

                // For first folder open message2.
                msgHandle2 = this.RopOpenMessage(folderHandle1, cprptFolderId[0], cprptMessageId[1], out openMsgRes, true);
                objHandles.Add("Message2", msgHandle2);

                // For first message open attachment1.
                attachmentHandle1 = this.RopOpenAttachment(msgHandle1, cprptAttachmentId[0], out openAttRes, true);
                objHandles.Add("Attachment1", attachmentHandle1);

                // For first message open attachment2.
                attachmentHandle2 = this.RopOpenAttachment(msgHandle1, cprptAttachmentId[1], out openAttRes, true);
                objHandles.Add("Attachment2", attachmentHandle2);
                #endregion

                if (!objHandles.ContainsKey(objKey))
                {
                    this.oxcropsClient = tempClient;
                    Site.Assert.Fail("Invalid Object key for getting its Property value");
                }

                uint objHandle = objHandles[objKey];

                RopGetPropertiesAllResponse getAllRes;
                getAllRes = this.RopGetPropertiesAll(objHandle, 0, 0);
                foreach (TaggedPropertyValue taggedValue in getAllRes.PropertyValues)
                {
                    if (taggedValue.PropertyTag.PropertyId == propertyTag.PropertyId)
                    {
                        isPropertyFound = true;

                        // Open stream.
                        RopOpenStreamResponse openStreamRes;
                        uint streamHandle = this.RopOpenStream(objHandle, out openStreamRes, propertyTag, ConstValues.OpenModeFlagsReadWrite, true);

                        // Read stream.
                        RopReadStreamResponse readStreamRes;
                        readStreamRes = this.RopReadStream(streamHandle, ConstValues.StreamLengthByteCountIndicator, 0x70000000, true);
                        return readStreamRes.Data;
                    }
                }

                isPropertyFound = false;
                return new byte[0];
            }
            catch (System.IO.IOException exception)
            {
                this.Site.Log.Add(LogEntryKind.Comment, exception.Message);
                isPropertyFound = false;
                return null;
            }
            finally
            {
                this.oxcropsClient = tempClient;
            }
        }

        /// <summary>
        /// Get folders, messages and attachments handle.
        /// </summary>
        private void GetHandles()
        {
            #region Get folder, message and attachment handle
            // Logon and save logon handle.
            RopLogonResponse logonResponse;
            this.cprptCurrentHandle = this.cprptLogonHandle = this.RopLogon(LogonType.Mailbox, out logonResponse, Common.GetConfigurationPropertyValue("UserEssdn", this.Site), true);

            // Open inbox and save inbox handle.
            RopOpenFolderResponse openFolderResponse;
            this.cprptInputBoxFolderId = logonResponse.FolderIds[4];
            this.cprptInputBoxHandle = this.RopOpenFolder(this.cprptLogonHandle, out openFolderResponse, this.cprptInputBoxFolderId, true);

            // Open root folder and save folder handle.
            this.cprptTestRootFolderHandle = this.RopOpenFolder(this.cprptInputBoxHandle, out openFolderResponse, rootFolderId, true);

            // Open subfolder and save subfolder handle.
            for (int i = 0; i < 3; i++)
            {
                cprptFolderHandle[i] = this.RopOpenFolder(this.cprptTestRootFolderHandle, out openFolderResponse, cprptFolderId[i], true);
            }

            // Open message and save message handle.
            RopOpenMessageResponse openMessageResponse;
            for (int i = 0; i < 3; i++)
            {
                cprptMessageHandle[i] = this.RopOpenMessage(cprptFolderHandle[0], cprptFolderId[0], cprptMessageId[i], out openMessageResponse, true);
            }

            // Open attachment and save attachment handle.
            RopOpenAttachmentResponse openAttachmentResponse;
            for (int i = 0; i < 3; i++)
            {
                cprptAttachmentHandle[i] = this.RopOpenAttachment(cprptMessageHandle[0], cprptAttachmentId[i], out openAttachmentResponse, true);
            }
            #endregion
        }

        /// <summary>
        /// Get TaggedProperty from ptfconfig file.
        /// </summary>
        /// <param name="obj">The order of object to get.</param>
        /// <returns>Return TaggedProperty.</returns>
        private TaggedPropertyValue GetTaggedPropertyTag(ObjectToOperate obj)
        {
            TaggedPropertyValue propertyTag = new TaggedPropertyValue();
            switch (this.cprptCurrentType)
            {
                case ServerObjectType.Logon:
                    if (obj == ObjectToOperate.FirstObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("LogonPropertyID0", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("LogonPropertyType0", this.Site));
                        propertyTag.Value = BitConverter.GetBytes(Convert.ToBoolean(Common.GetConfigurationPropertyValue("LogonPropertyValue0", this.Site)));
                    }
                    else if (obj == ObjectToOperate.SecondObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("LogonPropertyID1", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("LogonPropertyType1", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("LogonPropertyValue1", this.Site) + "\0");
                    }

                    break;

                case ServerObjectType.Attachment:
                    if (obj == ObjectToOperate.FirstObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("AttachmentPropertyID0", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("AttachmentPropertyType0", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("AttachmentPropertyValue0", this.Site) + "\0");
                    }
                    else if (obj == ObjectToOperate.SecondObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("AttachmentPropertyID1", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("AttachmentPropertyType1", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("AttachmentPropertyValue1", this.Site) + "\0");
                    }

                    break;

                case ServerObjectType.Folder:
                    if (obj == ObjectToOperate.FirstObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("FolderPropertyID0", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("FolderPropertyType0", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("FolderPropertyValue0", this.Site) + "\0");
                    }
                    else if (obj == ObjectToOperate.SecondObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("FolderPropertyID1", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("FolderPropertyType1", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("FolderPropertyValue1", this.Site) + "\0");
                    }
                    else if (obj == ObjectToOperate.ThirdObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("FolderPropertyID2", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("FolderPropertyType2", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("FolderPropertyValue2", this.Site) + "\0");
                    }

                    break;

                case ServerObjectType.Message:
                    if (obj == ObjectToOperate.FirstObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("MessagePropertyID0", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("MessagePropertyType0", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("MessagePropertyValue0", this.Site) + "\0");
                    }
                    else if (obj == ObjectToOperate.SecondObject)
                    {
                        propertyTag.PropertyTag.PropertyId = Convert.ToUInt16(Common.GetConfigurationPropertyValue("MessagePropertyID1", this.Site));
                        propertyTag.PropertyTag.PropertyType = Convert.ToUInt16(Common.GetConfigurationPropertyValue("MessagePropertyType1", this.Site));
                        propertyTag.Value = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("MessagePropertyValue1", this.Site) + "\0");
                    }

                    break;

                default:
                    Site.Assert.Fail("Invalid ServerObjectType enum value {0}.", this.cprptCurrentType);
                    break;
            }

            return propertyTag;
        }

        /// <summary>
        /// Get PropertyName from ptfconfig file.
        /// </summary>
        /// <param name="kind">The kind of PropertyName.</param>
        /// <returns>An instance of PropertyName.</returns>
        private PropertyName GetPropertyName(byte kind)
        {
            PropertyName propertyName = new PropertyName
            {
                Kind = kind
            };

            if (propertyName.Kind == (byte)Kind.LidField)
            {
                propertyName.LID = PropertyName00Lid;
                propertyName.Guid = (new Guid(PropertyName00Guid)).ToByteArray();
            }
            else if (propertyName.Kind == (byte)Kind.NameField)
            {
                propertyName.Name = Encoding.Unicode.GetBytes(PropertyName01Name + "\0");
                propertyName.NameSize = (byte)propertyName.Name.Length;
                propertyName.Guid = (new Guid(PropertyName01Guid)).ToByteArray();
            }
            else if (propertyName.Kind == (byte)Kind.NoAssociated)
            {
                propertyName.Guid = (new Guid(PropertyNameFFGuid)).ToByteArray();
            }

            return propertyName;
        }

        /// <summary>
        /// Search TaggedPropertyValue by property id.
        /// </summary>
        /// <param name="handle">Handle to be search.</param>
        /// <param name="propertyId">Property id.</param>
        /// <returns>Result in TaggedPropertyValue format.</returns>
        private TaggedPropertyValue SearchProperty(uint handle, ushort propertyId)
        {
            TaggedPropertyValue tagValue = null;
            RopGetPropertiesAllResponse getAllRes = this.RopGetPropertiesAll(handle, ConstValues.PropertySizeLimitNone, ConstValues.WantUnicodeNo);

            if (getAllRes.ReturnValue != (uint)RopResponseType.SuccessResponse)
            {
                return null;
            }

            for (int i = 0; i < getAllRes.PropertyValues.Length; i++)
            {
                if (getAllRes.PropertyValues[i].PropertyTag.PropertyId == propertyId)
                {
                    tagValue = getAllRes.PropertyValues[i];
                }
            }

            return tagValue;
        }

        /// <summary>
        /// Get destination message handles.
        /// </summary>
        /// <param name="handle">Handle to be operated.</param>
        /// <param name="firstMessageHandle">The first message handle.</param>
        /// <param name="secondMessageHandle">The second message handle.</param>
        private void GetDestinationMessageHandle(uint handle, out uint firstMessageHandle, out uint secondMessageHandle)
        {
            #region Call RopCreateAttachment to create an embedded attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentHandle;
            attachmentHandle = this.RopCreateAttachment(handle, out createAttachmentResponse, false);
            #endregion

            #region Call RopSetProperties to set PidTagAttachMethod property, that is the attachment is the embedded attachment.
            RopSetPropertiesResponse setPropertiesResponse;
            TaggedPropertyValue[] taggedPropertyValues = new TaggedPropertyValue[1];
            TaggedPropertyValue tag = new TaggedPropertyValue();
            tag.PropertyTag.PropertyId = (ushort)TaggedPropertyName.PidTagAttachMethod;
            tag.PropertyTag.PropertyType = (ushort)PropertyType.PtypInteger32;
            tag.Value = BitConverter.GetBytes(0x00000005);
            taggedPropertyValues[0] = tag;
            setPropertiesResponse = this.RopSetProperties(attachmentHandle, taggedPropertyValues, false);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment changes.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            saveChangesAttachmentResponse = this.RopSaveChangesAttachment(attachmentHandle, true);
            #endregion

            #region Call RopSaveChangesMessage to save the newly created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.RopSaveChangesMessage(handle, true);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x02 to create the attachment if it doesn't exist, and expect to get a successful response
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponse;
            uint embededMessageHandle;
            embededMessageHandle = this.RopOpenEmbeddedMessage(attachmentHandle, out openEmbeddedMessageResponse, true);
            firstMessageHandle = embededMessageHandle;
            #endregion

            #region Call RopCreateAttachment to create the second embedded attachment.
            RopCreateAttachmentResponse createAttachmentResponseSecond;
            uint attachmentHandleSecond;
            attachmentHandleSecond = this.RopCreateAttachment(handle, out createAttachmentResponseSecond, false);
            #endregion

            #region Call RopSetProperties to set PidTagAttachMethod property, that is the second attachment is the embedded attachment.
            RopSetPropertiesResponse setPropertiesResponseSecond;
            setPropertiesResponseSecond = this.RopSetProperties(attachmentHandleSecond, taggedPropertyValues, false);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment changes.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponseSecond;
            saveChangesAttachmentResponseSecond = this.RopSaveChangesAttachment(attachmentHandleSecond, true);
            #endregion

            #region Call RopSaveChangesMessage to save the newly created message.
            RopSaveChangesMessageResponse saveChangesMessageResponseSecond;
            saveChangesMessageResponseSecond = this.RopSaveChangesMessage(handle, true);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x02 to create the attachment if it doesn't exist, and expect to get a successful response
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponseSecond;
            uint embededMessageHandleSecond;
            embededMessageHandleSecond = this.RopOpenEmbeddedMessage(attachmentHandleSecond, out openEmbeddedMessageResponseSecond, true);
            secondMessageHandle = embededMessageHandleSecond;
            #endregion
        }

        /// <summary>
        /// To verify if there are conflict settings in SHOULD/MAY configuration file
        /// </summary>
        /// <param name="rsid">The requirement IDs to be checked</param>
        private void CheckRequirementConflict(params int[] rsid)
        {
            int enabledReqFound = 0;
            string requirementStr = string.Empty;
            foreach (int i in rsid)
            {
                requirementStr = requirementStr + i + " ";
                if (Common.IsRequirementEnabled(i, this.Site))
                {
                    enabledReqFound++;
                }
            }

            if (enabledReqFound == 0)
            {
                Site.Assert.Fail("Must implement one of optional behaviors in {0}", requirementStr);
            }
            else if (enabledReqFound > 1)
            {
                Site.Assert.Fail("Optional behavior conflict found in {0}", requirementStr);
            }
        }
        #endregion
    }
}