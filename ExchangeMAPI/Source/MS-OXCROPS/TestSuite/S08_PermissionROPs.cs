namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Permission ROPs. 
    /// </summary>
    [TestClass]
    public class S08_PermissionROPs : TestSuiteBase
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
        /// This method tests the ROP buffers of RopGetPermission and RopModifyPermission.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S08_TC01_TestRopGetAndModifyPermissions()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Open folder and create a subfolder, then get it's handle.
            #region Common operations

            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);
            uint folderHandle = GetFolderObjectHandle(ref logonResponse);

            #endregion

            // Step 2: Construct RopGetPermissionsTable request.
            #region RopGetPermissionsTable request

            RopGetPermissionsTableRequest getPermissionsTableRequest;

            getPermissionsTableRequest.RopId = (byte)RopId.RopGetPermissionsTable;

            getPermissionsTableRequest.LogonId = TestSuiteBase.LogonId;
            getPermissionsTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            getPermissionsTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            getPermissionsTableRequest.TableFlags = (byte)PermTableFlags.IncludeFreeBusy;

            #endregion

            // Step 3: Construct RopSetColumns request.
            #region RopSetColumns request

            // Call CreatePermissionPropertyTags method to create Permission PropertyTags.
            PropertyTag[] propertyTags = this.CreatePermissionPropertyTags();
            RopSetColumnsRequest setColumnsRequest;

            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;
            setColumnsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;
            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;

            #endregion

            // Step 4: Construct RopQueryRows request.
            #region RopQueryRows request

            RopQueryRowsRequest queryRowsRequest;

            queryRowsRequest.RopId = (byte)RopId.RopQueryRows;
            queryRowsRequest.LogonId = TestSuiteBase.LogonId;
            queryRowsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;
            queryRowsRequest.QueryRowsFlags = (byte)QueryRowsFlags.Advance;

            // TRUE: read the table forwards.
            queryRowsRequest.ForwardRead = TestSuiteBase.NonZero;

            // Maximum number of rows to be returned
            queryRowsRequest.RowCount = TestSuiteBase.RowCount;

            #endregion

            // Step 5: Send the Multiple ROPs request and verify the success response.
            #region Multiple ROPs

            List<ISerializable> requests = new List<ISerializable>();
            List<IDeserializable> ropResponses = new List<IDeserializable>();
            List<uint> handleList = new List<uint>
            {
                folderHandle,
                TestSuiteBase.DefaultFolderHandle
            };
            requests.Add(getPermissionsTableRequest);
            requests.Add(setColumnsRequest);
            requests.Add(queryRowsRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the Multiple ROPs request.");

            // Send the Multiple ROPs request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(
                requests,
                handleList,
                ref ropResponses,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopGetPermissionsTableResponse getPermissionsTableResponse = (RopGetPermissionsTableResponse)ropResponses[0];
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)ropResponses[1];
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)ropResponses[2];

            // Send the RopModifyPermissions request and verify response.
            #region RopModifyPermissions request

            RopModifyPermissionsRequest modifyPermissionsRequest;
            PermissionData[] permissionsDataArray = this.GetPermissionDataArray();

            modifyPermissionsRequest.RopId = (byte)RopId.RopModifyPermissions;
            modifyPermissionsRequest.LogonId = TestSuiteBase.LogonId;
            modifyPermissionsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            modifyPermissionsRequest.ModifyFlags = (byte)ModifyFlags.IncludeFreeBusy;
            modifyPermissionsRequest.ModifyCount = (ushort)permissionsDataArray.Length;
            modifyPermissionsRequest.PermissionsData = permissionsDataArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopModifyPermissions request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                modifyPermissionsRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)response;

            #endregion

            bool isRopsSuccess = (getPermissionsTableResponse.ReturnValue == TestSuiteBase.SuccessReturnValue)
                                 && (setColumnsResponse.ReturnValue == TestSuiteBase.SuccessReturnValue)
                                 && (queryRowsResponse.ReturnValue == TestSuiteBase.SuccessReturnValue)
                                 && (modifyPermissionsResponse.ReturnValue == TestSuiteBase.SuccessReturnValue);
            Site.Assert.IsTrue(isRopsSuccess, "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion
        }

        #endregion

        #region Common methods

        /// <summary>
        /// Create Permission PropertyTags
        /// </summary>
        /// <returns>Return propertyTag array</returns>
        private PropertyTag[] CreatePermissionPropertyTags()
        {
            // The following sample tags is from MS-OXCPERM 4.1
            PropertyTag[] propertyTags = new PropertyTag[4];

            // PidTagMemberId
            propertyTags[0] = this.propertyDictionary[PropertyNames.PidTagMemberId];

            // PidTagMemberName
            propertyTags[1] = this.propertyDictionary[PropertyNames.PidTagMemberName];

            // PidTagMemberRights
            propertyTags[2] = this.propertyDictionary[PropertyNames.PidTagMemberRights];

            // PidTagEntryId
            propertyTags[3] = this.propertyDictionary[PropertyNames.PidTagEntryId];

            return propertyTags;
        }

        /// <summary>
        /// Get GetPermissionData Array for modify permissions
        /// </summary>
        /// <returns>Return GetPermissionData array</returns>
        private PermissionData[] GetPermissionDataArray()
        {
            // Get PropertyValues
            TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue();
            TaggedPropertyValue[] propertyValues = new TaggedPropertyValue[2];

            // PidTagMemberId
            taggedPropertyValue.PropertyTag.PropertyId = this.propertyDictionary[PropertyNames.PidTagMemberId].PropertyId;
            taggedPropertyValue.PropertyTag.PropertyType = this.propertyDictionary[PropertyNames.PidTagMemberId].PropertyType;

            // Anonymous Client: The server MUST use the permissions specified in PidTagMemberRights for 
            // any anonymous users that have not been authenticated with user credentials.
            taggedPropertyValue.Value = BitConverter.GetBytes(TestSuiteBase.TaggedPropertyValueForPidTagMemberId);
            propertyValues[0] = taggedPropertyValue;

            // PidTagMemberRights
            taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagMemberRights].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagMemberRights].PropertyType
                },
                Value = BitConverter.GetBytes(TestSuiteBase.TaggedPropertyValueForPidTagMemberRights)
            };

            // CreateSubFolder 
            propertyValues[1] = taggedPropertyValue;

            PermissionData[] permissionsDataArray = new PermissionData[1];
            permissionsDataArray[0].PermissionDataFlags = (byte)PermissionDataFlags.ModifyRow;
            permissionsDataArray[0].PropertyValueCount = (ushort)propertyValues.Length;
            permissionsDataArray[0].PropertyValues = propertyValues;
            return permissionsDataArray;
        }

        #endregion
    }
}