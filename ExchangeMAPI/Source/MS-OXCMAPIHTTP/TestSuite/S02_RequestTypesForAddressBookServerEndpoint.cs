namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the request types for address book server endpoint.
    /// </summary>
    [TestClass]
    public class S02_RequestTypesForAddressBookServerEndpoint : TestSuiteBase
    {
        #region Variable
        /// <summary>
        /// A Boolean value indicating whether Address Book Member has been deleted.
        /// </summary>
        private bool isAddressBookMemberDeleted = true;

        /// <summary>
        /// A Boolean value indicating whether Address Book Public Delegate have been deleted.
        /// </summary>
        private bool isAddressBookPublicDelegateDeleted = true;

        /// <summary>
        /// The minimal ID for delete the address book member.
        /// </summary>
        private uint minimalIDForDeleteAddressBookMember;

        /// <summary>
        /// The entry ID for delete the address book member.
        /// </summary>
        private byte[] entryIDBufferForDeleteAddressBookMember;

        /// <summary>
        /// The minimal ID for delete the address book public delegates.
        /// </summary>
        private uint minimalIDForDeleteAddressBookPublicDelegate;

        /// <summary>
        /// The entry ID for delete the address book public delegates.
        /// </summary>
        private byte[] entryIDBufferForDeleteAddressBookPublicDelegate;
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

        #region Test Cases
        /// <summary>
        /// This case is designed to verify the requirements related to Bind and Unbind request types returning success.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC01_BindAndUnbind()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type with flags field set to 0x0 to established a session context with the address book server.

            // A STAT structure ([MS-OXNSPI] section 2.3.7) that specifies the state of a specific address book container. 
            STAT stat = new STAT();
            stat.InitiateStat();

            // A set of bit flags that specify the authentication type for the connection. The server MUST ignore values other than the bit flag fAnonymousLogin (0x00000020).
            uint flags = 0x0;
            BindRequestBody bindRequestBody = this.BuildBindRequestBody(stat, flags);

            int responseCode;
            BindResponseBody bindResponseBody = this.Adapter.Bind(bindRequestBody, out responseCode);
            Site.Assert.AreEqual<uint>(0, bindResponseBody.ErrorCode, "Bind should succeed and 0 is expected to be returned. The returned value is {0}.", bindResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R324");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R324
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                bindResponseBody.ErrorCode,
                324,
                @"[In Bind Request Type] The Bind request type is used by the client to establish a Session Context with the server, as specified in section 3.1.4.1.1 and section 3.1.4.1.1");
            
            #endregion

            #region Call Bind request type with flags field set to another value other than fAnonymousLogin (0x00000020) to check that server will return the same result.

            // A set of bit flags that specify the authentication type for the connection. The server MUST ignore values other than the bit flag fAnonymousLogin (0x00000020).
            flags = 0x01;
            bindRequestBody.Flags = flags;
            BindResponseBody bindResponseBodyWithDifferentFlags = this.Adapter.Bind(bindRequestBody, out responseCode);
            Site.Assert.AreEqual<uint>(0, bindResponseBody.ErrorCode, "Bind should succeed and 0 is expected to be returned. The returned value is {0}.", bindResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1320");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1320
            this.Site.CaptureRequirementIfAreEqual<uint>(
                bindResponseBody.ErrorCode,
                bindResponseBodyWithDifferentFlags.ErrorCode,
                1320,
                @"[In Bind Request Type Request Body] If this field [Flags] is set to different values other than the bit flag fAnonymousLogin (0x00000020), the server will return the same result.");
            #endregion

            #region Send a PING request to address book server endpoint.
            List<string> metatagsFromMailbox = new List<string>();
            WebHeaderCollection headers = new WebHeaderCollection();
            uint pingResponseCode = this.Adapter.PING(ServerEndpoint.AddressBookServerEndpoint, out metatagsFromMailbox, out headers);
            Site.Assert.AreEqual<uint>(0, pingResponseCode, "PING should succeed and 0 is expected to be returned. The returned Value is {0}.", pingResponseCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            UnbindRequestBody unbindRequest = this.BuildUnbindRequestBody();
            UnbindResponseBody unbindResponseBody = this.Adapter.Unbind(unbindRequest);
            Site.Assert.AreEqual<uint>(1, unbindResponseBody.ErrorCode, "Unbind should succeed and 0 is expected to be returned. The returned value is {0}.", unbindResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1438");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1438
            // The above assert ensures that the Unbind request type executes successfully, so R1250 can be verified if the Unbind response body is not null.
            this.Site.CaptureRequirementIfIsNotNull(
                unbindResponseBody,
                1438,
                @"[In Responding to a Disconnect or Unbind Request Type Request] If successful, the server's response includes the Unbind request type success response body, as specified in section 2.2.5.2.2.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R353");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R353
            this.Site.CaptureRequirementIfAreEqual<uint>(
                1,
                unbindResponseBody.ErrorCode,
                353,
                @"[In Unbind Request Type] The Unbind request type is used by the client to delete a Session Context with the server, as specified in section 3.1.4.1.2.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1441");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1441
            // The cookies returned from Bind request type is used by Unbind request type, so if the Unbind request type executes successfully and the Bind response body is not null, R1441 can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                bindResponseBody,
                1441,
                @"[In Responding to a Connect or Bind Request Type Request] If successful, the server's response includes the Bind request type response body as specified in section 2.2.5.1.2.");
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to GetMailboxUrl request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC02_GetMailboxUrl()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call the GetMailboxUrl request type to get the Uniform Resource Locator (URL) of the specified mailbox server endpoint.

            string serverDn = Common.GetConfigurationPropertyValue("ServerDN", this.Site);

            // The client MUST set this field to 0x00000000 and the server MUST ignore this field.
            uint flagsOfGetMailboxUrl = 0x00000000;
            GetMailboxUrlRequestBody getMailboxUrlRequestBody = new MS_OXCMAPIHTTP.GetMailboxUrlRequestBody();
            getMailboxUrlRequestBody.Flags = flagsOfGetMailboxUrl;
            getMailboxUrlRequestBody.ServerDn = serverDn;
            getMailboxUrlRequestBody.AuxiliaryBuffer = new byte[] { };
            getMailboxUrlRequestBody.AuxiliaryBufferSize = 0;

            GetMailboxUrlResponseBody getMailboxUrlResponseBody = this.Adapter.GetMailboxUrl(getMailboxUrlRequestBody);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1089");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1089
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getMailboxUrlResponseBody.ErrorCode,
                1089,
                @"[In GetMailboxUrl Request Type] The GetMailboxUrl request type is used by the client to get the Uniform Resource Locator (URL) of the specified mailbox server endpoint (4).");
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to GetAddressBookUrl request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC03_GetAddressBookUrl()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region cALL GetAddressBookUrl request type to get the URL of the specified address book server endpoint.
            // The client MUST set this field to 0x00000000 and the server MUST ignore this field.
            uint flagsOfGetAddressBookUrl = 0x00000000;
            GetAddressBookUrlRequestBody getAddressBookUrlRequestBody = new GetAddressBookUrlRequestBody();
            getAddressBookUrlRequestBody.Flags = flagsOfGetAddressBookUrl;
            getAddressBookUrlRequestBody.UserDn = this.AdminUserDN;
            getAddressBookUrlRequestBody.AuxiliaryBuffer = new byte[] { };
            getAddressBookUrlRequestBody.AuxiliaryBufferSize = 0;

            GetAddressBookUrlResponseBody getAddressBookUrlResponseBody = this.Adapter.GetAddressBookUrl(getAddressBookUrlRequestBody);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1106");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1106
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getAddressBookUrlResponseBody.ErrorCode,
                1106,
                @"[In GetAddressBookUrl Request Type] The GetAddressBookUrl request type is used by the client to the URL of the specified address book server endpoint (4).");
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify requirements related to DnToMinId request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC04_DnToMinId()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call DnToMinId request type to map a set of DNs to a set of Minimal Entry IDs.
            // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
            uint reserved = 0;
            byte[] auxIn = new byte[] { };
            DNToMinIdRequestBody requestBodyOfDNToMId = new DNToMinIdRequestBody()
            {
                Reserved = reserved,
                HasNames = true,
                Names = new StringArray_r
                {
                    CValues = 1,
                    LppzA = new string[]
                    {
                        this.AdminUserDN
                    }
                },
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            DnToMinIdResponseBody responseBodyOfDNToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
            Site.Assert.AreEqual<uint>(0, responseBodyOfDNToMinId.ErrorCode, "DnToMinId should succeed and 0 is expected to be returned. The returned value is {0}.", responseBodyOfDNToMinId.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R405");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R405
            // The client defines a set of DNs, if there is MinimalIds, it indicates that the client maps a set of DNs (1) to a set of Minimal Entry IDs successfully.
            this.Site.CaptureRequirementIfIsNotNull(
                responseBodyOfDNToMinId.MinimalIds,
                405,
                @"[In DnToMinId Request Type] The DnToMinId request type is used by the client to map a set of DNs (1) to a set of Minimal Entry IDs.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R436, the value of MinimalIdCount is {0}, the value of MinimalIds is {1}.",
                responseBodyOfDNToMinId.MinimalIdCount,
                responseBodyOfDNToMinId.MinimalIds);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R436
            this.Site.CaptureRequirementIfIsTrue(
                responseBodyOfDNToMinId.MinimalIdCount == requestBodyOfDNToMId.Names.CValues && responseBodyOfDNToMinId.MinimalIds != null,
                436,
                @"[In DnToMinId Request Type Success Response Body] MinimalIds (optional) (variable): An array of MinimalEntryID structures ([MS-OXNSPI] section 2.3.8.1), each of which specifies a Minimal Entry ID that matches a requested distinguished name (DN) (1).");
            #endregion

            #region Call DnToMinId request body with HasNames set to false.
            requestBodyOfDNToMId = new DNToMinIdRequestBody()
            {
                Reserved = reserved,
                HasNames = false,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            responseBodyOfDNToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
            Site.Assert.AreEqual<uint>(0, responseBodyOfDNToMinId.ErrorCode, "DnToMinId should succeed and 0 is expected to be returned. The returned value is {0}.", responseBodyOfDNToMinId.ErrorCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        ///  This case is designed to test the requirements related to CompareMinIds request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC05_CompareMinIds()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call DnToMinId request type to map a set of DNs to a set of Minimal Entry IDs.
            // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
            uint reserved = 0;
            byte[] auxIn = new byte[] { };
            DNToMinIdRequestBody requestBodyOfDNToMId = new DNToMinIdRequestBody()
            {
                Reserved = reserved,
                HasNames = true,
                Names = new StringArray_r
                {
                    CValues = 2,
                    LppzA = new string[]
                    {
                        this.AdminUserDN,
                        Common.GetConfigurationPropertyValue("GeneralUserEssdn", this.Site)
                    }
                },
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            DnToMinIdResponseBody responseBodyOfDNToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
            Site.Assert.AreEqual<uint>(0, responseBodyOfDNToMinId.ErrorCode, "DnToMinId should succeed and 0 is expected to be returned. The returned value is {0}.", responseBodyOfDNToMinId.ErrorCode);
            #endregion

            #region Call CompareMinIds to compare the positions of two objects in an address book container.
            STAT stat = new STAT();
            stat.InitiateStat();
            CompareMinIdsRequestBody compareMinIdsRequestBody = new CompareMinIdsRequestBody()
            {
                // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
                Reserved = 0,
                HasState = true,
                State = stat,
                MinimalId1 = responseBodyOfDNToMinId.MinimalIds[0],
                MinimalId2 = responseBodyOfDNToMinId.MinimalIds[1],
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            CompareMinIdsResponseBody compareMinIdsResponseBody = this.Adapter.CompareMinIds(compareMinIdsRequestBody);
            Site.Assert.AreEqual<uint>(0, compareMinIdsResponseBody.ErrorCode, "CompareMinIds should succeed and 0 is expected to be returned. The returned value is {0}.", compareMinIdsResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R372");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R372
            // If the position of the object specified by MId1 comes before the position of the object specified by MId2 in the table, the server MUST return a value less than 0; 
            // If the position of the object specified by MId1 comes after the position of the object specified by MId2 in the table, the server MUST return a value greater than 0; 
            // If the position of the object specified by MId1 is the same as the position of the object specified by MId2 in the table, the server MUST return a value of 0.
            // So if the value type of result is integer, R372 is verified. 
            this.Site.CaptureRequirementIfIsInstanceOfType(
                compareMinIdsResponseBody.Result,
                typeof(int),
                372,
                @"[In CompareMinIds Request Type] The CompareMinIds request type is used by the client to compare the positions of two objects in an address book container.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2111");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2111
            this.Site.CaptureRequirementIfIsTrue(
                compareMinIdsResponseBody.Result < 0,
                2111,
                @"[In CompareMinIds Request Type Success Response Body] [Result] A value less than 0 (zero): The position of the object specified by the MinimalId1 field of the request body precedes the position of the object specified by the MinimalId2 field.");

            compareMinIdsRequestBody = new CompareMinIdsRequestBody()
            {
                // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
                Reserved = 0,
                HasState = true,
                State = stat,
                MinimalId1 = responseBodyOfDNToMinId.MinimalIds[1],
                MinimalId2 = responseBodyOfDNToMinId.MinimalIds[0],
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            compareMinIdsResponseBody = this.Adapter.CompareMinIds(compareMinIdsRequestBody);
            Site.Assert.AreEqual<uint>(0, compareMinIdsResponseBody.ErrorCode, "CompareMinIds should succeed and 0 is expected to be returned. The returned value is {0}.", compareMinIdsResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2112");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2112
            this.Site.CaptureRequirementIfIsTrue(
                compareMinIdsResponseBody.Result > 0,
                2112,
                @"[In CompareMinIds Request Type Success Response Body] [Result] A value greater than 0 (zero): The position of the object specified by the MinimalId1 field of the request body succeeds the position of the object specified by the MinimalId2 field.");

            compareMinIdsRequestBody = new CompareMinIdsRequestBody()
            {
                // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
                Reserved = 0,
                HasState = true,
                State = stat,
                MinimalId1 = responseBodyOfDNToMinId.MinimalIds[0],
                MinimalId2 = responseBodyOfDNToMinId.MinimalIds[0],
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            compareMinIdsResponseBody = this.Adapter.CompareMinIds(compareMinIdsRequestBody);
            Site.Assert.AreEqual<uint>(0, compareMinIdsResponseBody.ErrorCode, "CompareMinIds should succeed and 0 is expected to be returned. The returned value is {0}.", compareMinIdsResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2109");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2109
            this.Site.CaptureRequirementIfIsTrue(
                compareMinIdsResponseBody.Result == 0,
                2109,
                @"[In CompareMinIds Request Type Success Response Body] [Result] Value 0 (zero): The position of the object specified by the MinimalId1 field of the request body is the same as the position of the object specified by the MinimalId2 field. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2110");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2110
            this.Site.CaptureRequirementIfAreEqual<uint>(
                compareMinIdsRequestBody.MinimalId1,
                compareMinIdsRequestBody.MinimalId2,
                2110,
                @"[In CompareMinIds Request Type Success Response Body] [Result] Value 0 (zero): That is, the two fields specify the same object.");
            #endregion

            #region Call CompareMinIds request body with HasState set to false to compare the positions of two objects in an address book container.
            compareMinIdsRequestBody = new CompareMinIdsRequestBody()
            {
                // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
                Reserved = 0,
                HasState = false,
                MinimalId1 = responseBodyOfDNToMinId.MinimalIds[0],
                MinimalId2 = responseBodyOfDNToMinId.MinimalIds[1],
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            compareMinIdsResponseBody = this.Adapter.CompareMinIds(compareMinIdsRequestBody);
            Site.Assert.AreEqual<uint>(0, compareMinIdsResponseBody.StatusCode, "CompareMinIds should succeed and 0 is expected to be returned. The returned value is {0}.", compareMinIdsResponseBody.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        ///  This case is designed to test the requirements related to ResortRestriction request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC06_ResortRestriction()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call DNToMId to get a set of Minimal Entry IDs. These IDs will be used as the parameter of ResortRestriction.
            StringArray_r names = new StringArray_r
            {
                CValues = 2,
                LppzA = new string[]
                    {
                        this.AdminUserDN,
                        Common.GetConfigurationPropertyValue("GeneralUserEssdn", this.Site)
                    }
            };

            DNToMinIdRequestBody requestBodyOfDNToMId = this.BuildDNToMinIdRequestBody(true, names);

            DnToMinIdResponseBody responseBodyOfDNToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
            Site.Assert.AreEqual<uint>(0, responseBodyOfDNToMinId.ErrorCode, "DnToMinId should succeed and 0 is expected to be returned. The returned value is {0}.", responseBodyOfDNToMinId.ErrorCode);
            #endregion

            #region Call ResortRestriction to sort the objects in a restricted address book container.
            STAT stat = new STAT();
            stat.InitiateStat();

            ResortRestrictionRequestBody resortRestrictionRequestBody = this.BuildResortRestriction(true, stat, true, responseBodyOfDNToMinId.MinimalIdCount.Value, responseBodyOfDNToMinId.MinimalIds);

            ResortRestrictionResponseBody resortRestrictionResponseBody = this.Adapter.ResortRestriction(resortRestrictionRequestBody);
            Site.Assert.AreEqual<uint>(0, resortRestrictionResponseBody.ErrorCode, "ResortRestriction should succeed and 0 is expected to be returned. The returned value is {0}.", resortRestrictionResponseBody.ErrorCode);
            #endregion

            #region Call QueryRows method to query rows which contain the specific properties.
            LargePropertyTagArray columns = new LargePropertyTagArray
            {
                PropertyTagCount = 1,
                PropertyTags = new PropertyTag[1]
                {
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypString,
                        PropertyId = (ushort)PropertyID.PidTagDisplayName
                    }
                }
            };

            QueryRowsRequestBody queryRowsRequestBody = this.BuildQueryRowsRequestBody(
                true,
                resortRestrictionResponseBody.State.Value,
                resortRestrictionResponseBody.MinimalIdCount.Value,
                resortRestrictionResponseBody.MinimalIds,
                resortRestrictionResponseBody.MinimalIdCount.Value,
                true,
                columns);

            QueryRowsResponseBody queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>(0, queryRowsResponseBody.ErrorCode, "QueryRows should succeed and 0 is expected to be returned. The returned value is {0}.", queryRowsResponseBody.ErrorCode);

            bool isCorrectOrder = true;
            AddressBookPropertyRow[] rowData = queryRowsResponseBody.RowData;
            for (int i = 0; i < queryRowsResponseBody.RowCount - 1; i++)
            {
                List<AddressBookPropertyValue> valueArray1 = new List<AddressBookPropertyValue>(rowData[i].ValueArray);
                List<AddressBookPropertyValue> valueArray2 = new List<AddressBookPropertyValue>(rowData[i + 1].ValueArray);
                if (string.Compare(valueArray1[0].Value.ToString(), valueArray2[0].Value.ToString()) > 0)
                {
                    isCorrectOrder = false;
                    break;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R958");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R958
            this.Site.CaptureRequirementIfIsTrue(
                isCorrectOrder,
                958,
                @"[In ResortRestriction Request Type] The ResortRestriction request type is used by the client to sort the objects in a restricted address book container.");
            #endregion

            #region Call ResortRestriction without state or minimal IDs.
            ResortRestrictionRequestBody resortRestrictionRequestBody2 = this.BuildResortRestriction(false, stat, false, responseBodyOfDNToMinId.MinimalIdCount.Value, responseBodyOfDNToMinId.MinimalIds);

            ResortRestrictionResponseBody resortRestrictionResponseBody2 = this.Adapter.ResortRestriction(resortRestrictionRequestBody2);
            Site.Assert.AreEqual<uint>(0, resortRestrictionResponseBody2.StatusCode, "ResortRestriction should succeed and 0 is expected to be returned. The returned value is {0}.", resortRestrictionResponseBody2.StatusCode);
            #endregion

            #region Call ResortRestriction with state and without minimal IDs.
            ResortRestrictionRequestBody resortRestrictionRequestBody3 = this.BuildResortRestriction(true, stat, false, responseBodyOfDNToMinId.MinimalIdCount.Value, responseBodyOfDNToMinId.MinimalIds);

            ResortRestrictionResponseBody resortRestrictionResponseBody3 = this.Adapter.ResortRestriction(resortRestrictionRequestBody3);
            Site.Assert.AreEqual<uint>(0, resortRestrictionResponseBody3.StatusCode, "ResortRestriction should succeed and 0 is expected to be returned. The returned value is {0}.", resortRestrictionResponseBody3.StatusCode);
            #endregion

            #region Call ResortRestriction without state but with minimal IDs.
            ResortRestrictionRequestBody resortRestrictionRequestBody4 = this.BuildResortRestriction(false, stat, true, responseBodyOfDNToMinId.MinimalIdCount.Value, responseBodyOfDNToMinId.MinimalIds);

            ResortRestrictionResponseBody resortRestrictionResponseBody4 = this.Adapter.ResortRestriction(resortRestrictionRequestBody4);
            Site.Assert.AreEqual<uint>(0, resortRestrictionResponseBody4.StatusCode, "ResortRestriction should succeed and 0 is expected to be returned. The returned value is {0}.", resortRestrictionResponseBody4.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        ///  This case is designed to test the requirements related to UpdateStat request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC07_UpdateStat()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call UpdateStat to update the STAT structure to reflect the client's changes.
            STAT stat = new STAT();
            stat.InitiateStat();
            int delta = 1;
            stat.Delta = delta;

            UpdateStatRequestBody updateStatRequestBody = this.BuildUpdateStatRequestBody(true, stat, true);
            UpdateStatResponseBody updateStatResponseBody = this.Adapter.UpdateStat(updateStatRequestBody);
            Site.Assert.AreEqual<uint>(0, updateStatResponseBody.ErrorCode, "UpdateStat should succeed and 0 is expected to be returned. The returned value is {0}.", updateStatResponseBody.ErrorCode);

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R1059. After calling UpdateStat, the value of each field (CodePage, ContainerID, CurrentRec, Delta, NumPos, SortLocale, SortType, TemplateLocale, TotalRecs)"
                + "in STAT structures is {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}.",
                 updateStatResponseBody.State.Value.CodePage,
                 updateStatResponseBody.State.Value.ContainerID,
                 updateStatResponseBody.State.Value.CurrentRec,
                 updateStatResponseBody.State.Value.Delta,
                 updateStatResponseBody.State.Value.NumPos,
                 updateStatResponseBody.State.Value.SortLocale,
                 updateStatResponseBody.State.Value.SortType,
                 updateStatResponseBody.State.Value.TemplateLocale,
                 updateStatResponseBody.State.Value.TotalRecs);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1059
            bool isVerifiedR1059 = stat.CodePage != updateStatResponseBody.State.Value.CodePage ||
                stat.ContainerID != updateStatResponseBody.State.Value.ContainerID ||
                stat.CurrentRec != updateStatResponseBody.State.Value.CurrentRec ||
                stat.Delta != updateStatResponseBody.State.Value.Delta ||
                stat.NumPos != updateStatResponseBody.State.Value.NumPos ||
                stat.SortLocale != updateStatResponseBody.State.Value.SortLocale ||
                stat.SortType != updateStatResponseBody.State.Value.SortType ||
                stat.TemplateLocale != updateStatResponseBody.State.Value.TemplateLocale ||
                stat.TotalRecs != updateStatResponseBody.State.Value.TotalRecs;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1059,
                1059,
                @"[In UpdateStat Request Type] The UpdateStat request type is used by the client to update the STAT structure ([MS-OXNSPI] section 2.3.7) to reflect the client's changes.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1083");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1083
            this.Site.CaptureRequirementIfAreEqual<int>(
                delta,
                updateStatResponseBody.Delta.Value,
                1083,
                @"[In UpdateStat Request Type Success Response Body] Delta: A signed integer that specifies the movement within the address book container that was specified in the State field of the request.");
            #endregion
            #endregion

            #region Call UpdateStat with HasState field set to false and DeltaRequested set to true.
            UpdateStatRequestBody updateStatRequestBody1 = this.BuildUpdateStatRequestBody(false, stat, true);
            UpdateStatResponseBody updateStatResponseBody1 = this.Adapter.UpdateStat(updateStatRequestBody1);
            Site.Assert.AreEqual<uint>(0, updateStatResponseBody1.StatusCode, "UpdateStat should succeed and 0 is expected to be returned. The returned value is {0}.", updateStatResponseBody1.StatusCode);
            #endregion

            #region Call UpdateStat with fields HasState and DeltaRequested set to false.
            UpdateStatRequestBody updateStatRequestBody2 = this.BuildUpdateStatRequestBody(false, stat, false);
            UpdateStatResponseBody updateStatResponseBody2 = this.Adapter.UpdateStat(updateStatRequestBody2);
            Site.Assert.AreEqual<uint>(0, updateStatResponseBody2.StatusCode, "UpdateStat should succeed and 0 is expected to be returned. The returned value is {0}.", updateStatResponseBody2.StatusCode);
            #endregion

            #region Call UpdateStat with HasState field set to true and DeltaRequested set to false.
            UpdateStatRequestBody updateStatRequestBody3 = this.BuildUpdateStatRequestBody(true, stat, false);
            UpdateStatResponseBody updateStatResponseBody3 = this.Adapter.UpdateStat(updateStatRequestBody3);
            Site.Assert.AreEqual<uint>(0, updateStatResponseBody3.StatusCode, "UpdateStat should succeed and 0 is expected to be returned. The returned value is {0}.", updateStatResponseBody3.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        ///  This case is designed to test the requirements related to GetSpecialTable request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC08_GetSpecialTable()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call GetSpecialTable request type to get an address creation table.
            uint flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiAddressCreationTemplates;
            uint version = 0;
            STAT stat = new STAT();
            stat.InitiateStat();
            stat.CodePage = (uint)RequiredCodePages.CP_TELETEX;

            GetSpecialTableRequestBody getSpecialTableRequestBody = this.BuildGetSpecialTableRequestBody(flagsOfGetSpecialTable, true, stat, true, version);
            GetSpecialTableResponseBody getSpecialTableResponseBody = this.Adapter.GetSpecialTable(getSpecialTableRequestBody);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R631");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R631
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getSpecialTableResponseBody.ErrorCode,
                631,
                @"[In GetSpecialTable Request Type] The GetSpecialTable request type is used by the client to get a special table, which can be either an address book hierarchy table or an address creation table.");

            for (int i = 0; i < getSpecialTableResponseBody.RowCount; i++)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R677");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R677
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    getSpecialTableResponseBody.Rows[i],
                    typeof(AddressBookPropertyValueList),
                    677,
                    @"[In GetSpecialTable Request Type  Success Response Body] Rows (optional) (variable): An array of AddressBookPropertyValueList structures, each of which contains a row of the table that the client requested.");
            }

            #endregion

            #region Call GetSpecialTable request type with flags field set to 0x00000000.
            getSpecialTableRequestBody.Flags = 0x00000000;
            stat.TemplateLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            stat.CodePage = (uint)RequiredCodePages.CP_TELETEX;
            getSpecialTableRequestBody.State = stat;
            GetSpecialTableResponseBody getSpecialTableResponseBodyFlags0 = this.Adapter.GetSpecialTable(getSpecialTableRequestBody);
            Site.Assert.AreEqual<uint>(0, getSpecialTableResponseBodyFlags0.ErrorCode, "GetSpecialTable operation should succeed and 0 is expected to be returned. The return value is {0}.", getSpecialTableResponseBodyFlags0.ErrorCode);
            #endregion

            #region Call GetSpecialTable request type with flags field set to another value other than 0x0000000.
            getSpecialTableRequestBody.Flags = 0x00000001;
            stat.TemplateLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            stat.CodePage = (uint)RequiredCodePages.CP_TELETEX;
            getSpecialTableRequestBody.State = stat;
            GetSpecialTableResponseBody getSpecialTableResponseBodyFlags1 = this.Adapter.GetSpecialTable(getSpecialTableRequestBody);
            Site.Assert.AreEqual<uint>(0, getSpecialTableResponseBodyFlags1.ErrorCode, "GetSpecialTable operation should succeed and 0 is expected to be returned. The return value is {0}.", getSpecialTableResponseBodyFlags1.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1322");

            this.Site.Assert.AreEqual<uint>(getSpecialTableResponseBodyFlags0.RowCount.Value, getSpecialTableResponseBodyFlags1.RowCount.Value, "The two response bodies' RowCount field {0}, {1} should be equal.", getSpecialTableResponseBodyFlags0.RowCount.Value, getSpecialTableResponseBodyFlags1.RowCount.Value);
            for (int i = 0; i < getSpecialTableResponseBodyFlags0.RowCount; i++)
            {
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1322
                this.Site.CaptureRequirementIfIsTrue(
                    AdapterHelper.AreTwoAddressBookPropValueListEqual(getSpecialTableResponseBodyFlags0.Rows[i], getSpecialTableResponseBodyFlags1.Rows[i]),
                    1322,
                    @"[In GetSpecialTable Request Type Request Body] If this field [Flags] is set to different values other than the bit flags NspiAddressCreationTemplates (0x00000002) and NspiUnicodeStrings (0x00000004), the server will return the same result.");
            }

            #endregion

            #region Call GetSpecialTable request type without state.

            GetSpecialTableRequestBody getSpecialTableRequestBodyWithoutState = this.BuildGetSpecialTableRequestBody(flagsOfGetSpecialTable, false, stat, true, version);
            GetSpecialTableResponseBody getSpecialTableResponseBodyWithoutState = this.Adapter.GetSpecialTable(getSpecialTableRequestBodyWithoutState);
            this.Site.Assert.AreEqual<uint>(0, getSpecialTableResponseBodyWithoutState.StatusCode, "GetSpecialTable should be accepted by the Address Book Server, the status code is {0}.", getSpecialTableResponseBodyWithoutState.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        ///  This case is designed to test the requirements related to GetTemplateInfo request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC09_GetTemplateInfo()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call GetTemplateInfo to get information about a template that is used by the address book without DN.
            STAT stat = new STAT();
            stat.InitiateStat();
            uint flagsOfGetTemplateInfo = (uint)NspiGetTemplateInfoFlags.TI_TEMPLATE;
            uint type = (uint)DisplayTypeValues.DT_MAILUSER;
            string dn = null;
            uint codePage = stat.CodePage;
            uint locateID = stat.TemplateLocale;

            GetTemplateInfoRequestBody getTemplateInfoRequestBody = this.BuildGetTemplateInfoRequestBody(flagsOfGetTemplateInfo, type, false, dn, codePage, locateID);
            GetTemplateInfoResponseBody getTemplateInfoResponseBody = this.Adapter.GetTemplateInfo(getTemplateInfoRequestBody);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R683");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R683
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getTemplateInfoResponseBody.ErrorCode,
                683,
                @"[In GetTemplateInfo Request Type] The GetTemplateInfo request type is used by the client to get information about a template that is used by the address book.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R718");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R718
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getTemplateInfoResponseBody.Row,
                typeof(AddressBookPropertyValueList),
                718,
                @"[In GetTemplateInfo Request Type Success Response Body] Row (optional) (variable): An AddressBookPropertyValueList structure (section 2.2.1.1) that specifies the information that the client requested.");

            #endregion

            #region Call GetTemplateInfo method with flags 0x00000000 other than the bit flags TI_HELPFILE_NAME(0x00000020) or TI_HELPFILE_CONTENTS(0x00000040).

            getTemplateInfoRequestBody.Flags = 0x00000000;
            GetTemplateInfoResponseBody getTemplateInfoResponseBodyFlags0 = this.Adapter.GetTemplateInfo(getTemplateInfoRequestBody);
            this.Site.Assert.AreEqual<uint>(0, getTemplateInfoResponseBodyFlags0.ErrorCode, "GetTemplateInfo request should be executed successfully. The returned error code is {0}.", getTemplateInfoResponseBodyFlags0.ErrorCode);
            #endregion

            #region Call GetTemplateInfo method with flags 0x00000002 other than the bit flags TI_HELPFILE_NAME(0x00000020) or TI_HELPFILE_CONTENTS(0x00000040).

            getTemplateInfoRequestBody.Flags = 0x00000002;
            GetTemplateInfoResponseBody getTemplateInfoResponseBodyFlags1 = this.Adapter.GetTemplateInfo(getTemplateInfoRequestBody);
            this.Site.Assert.AreEqual<uint>(0, getTemplateInfoResponseBodyFlags1.ErrorCode, "GetTemplateInfo request should be executed successfully. The returned error code is {0}.", getTemplateInfoResponseBodyFlags1.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1324");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1324
            bool isVerifiedR1324 = AdapterHelper.AreTwoAddressBookPropValueListEqual(getTemplateInfoResponseBodyFlags0.Row.Value, getTemplateInfoResponseBodyFlags1.Row.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1324,
                1324,
                @"[In GetTemplateInfo Request Type Request Body] If this field [Flags] is set to different values other than the bit flags TI_HELPFILE_NAME (0x00000020), TI_HELPFILE_CONTENTS (0x00000040), TI_SCRIPT (0x00000004), TI_TEMPLATE (0x00000001), and TI_EMT (0x00000010), the server will return the same result.");

            #endregion

            #region Call GetTemplateInfo method with DisplayType set to CP_WINUNICODE.
            getTemplateInfoRequestBody.DisplayType = (uint)RequiredCodePages.CP_WINUNICODE;
            getTemplateInfoResponseBody = this.Adapter.GetTemplateInfo(getTemplateInfoRequestBody);
            this.Site.Assert.AreEqual<uint>(0, getTemplateInfoResponseBody.StatusCode, "GetTemplateInfo request should be executed successfully. The returned status code is {0}.", getTemplateInfoResponseBody.StatusCode);
            #endregion

            #region Call GetSpecialTable to get the template DN.
            uint flagsOfGetSpecialTable = (uint)NspiGetSpecialTableFlags.NspiAddressCreationTemplates;
            uint version = 0;
            stat.InitiateStat();
            stat.CodePage = (uint)RequiredCodePages.CP_TELETEX;
            GetSpecialTableRequestBody getSpecialTableRequestBody = this.BuildGetSpecialTableRequestBody(flagsOfGetSpecialTable, true, stat, true, version);
            GetSpecialTableResponseBody getSpecialTableResponseBody = this.Adapter.GetSpecialTable(getSpecialTableRequestBody);
            this.Site.Assert.AreEqual<uint>(0, getSpecialTableResponseBody.ErrorCode, "GetSpecialTable request should be executed successfully, the error code is {0}.", getSpecialTableResponseBody.ErrorCode);

            // Parse and record the template DN.
            string templateDN = string.Empty;

            foreach (AddressBookPropertyValueList row in getSpecialTableResponseBody.Rows)
            {
                AddressBookTaggedPropertyValue[] propertyValue = row.PropertyValues;
                for (int i = 0; i < propertyValue.Length; i++)
                {
                    if (propertyValue[i].PropertyType == 0x0102 && propertyValue[i].PropertyId == 0x0FFF)
                    {
                        PermanentEntryID permanentEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(propertyValue[i].Value);
                        templateDN = permanentEntryID.DistinguishedName;

                        break;
                    }
                }

                if (templateDN != string.Empty)
                {
                    break;
                }
            }
            #endregion

            #region Call GetTemplateInfo method with DN set to a non-null value.

            flagsOfGetTemplateInfo = (uint)NspiGetTemplateInfoFlags.TI_SCRIPT;
            type = (uint)DisplayTypeValues.DT_MAILUSER;
            codePage = (uint)RequiredCodePages.CP_TELETEX;
            locateID = stat.TemplateLocale;

            GetTemplateInfoRequestBody getTemplateInfoRequestBodyWithDN = this.BuildGetTemplateInfoRequestBody(flagsOfGetTemplateInfo, type, true, templateDN, codePage, locateID);
            GetTemplateInfoResponseBody getTemplateInfoResponseBodyWithDN = this.Adapter.GetTemplateInfo(getTemplateInfoRequestBodyWithDN);
            this.Site.Assert.AreEqual<uint>(0, getTemplateInfoResponseBodyWithDN.ErrorCode, "GetTemplateInfo request should be executed successfully, the returned error code is {0}", getTemplateInfoResponseBodyWithDN.ErrorCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to GetMatches request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC10_GetMatchesRequestType()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call GetMatches request type without MinimalIds, PropertyNameGuid or PropertyNameId.
            STAT stat = new STAT();
            stat.InitiateStat();

            PropertyTag pidTagDisplayNameTag = new PropertyTag((ushort)PropertyID.PidTagDisplayName, (ushort)PropertyTypeValues.PtypString);

            LargePropertyTagArray propTagArray = new LargePropertyTagArray();
            propTagArray.PropertyTagCount = 1;
            propTagArray.PropertyTags = new PropertyTag[]
            {
                pidTagDisplayNameTag
            };

            ExistRestriction existRestriction = new ExistRestriction()
            {
                PropTag = new PropertyTag((ushort)PropertyID.PidTagEntryId, (ushort)PropertyTypeValues.PtypBinary)
            };

            byte[] filter = existRestriction.Serialize();
            Guid propertyGuid = new Guid();
            uint propertyNameId = new uint();
            GetMatchesRequestBody getMatchRequestBody = this.BuildGetMatchRequestBody(true, stat, false, 0, null, true, filter, false, propertyGuid, propertyNameId, ConstValues.GetMatchesRequestedRowNumber, true, propTagArray);

            GetMatchesResponseBody getMatchesResponseBody = this.Adapter.GetMatches(getMatchRequestBody);
            this.Site.Assert.AreEqual<uint>(0, getMatchesResponseBody.ErrorCode, "GetMatches request should be executed successfully, the returned error code is {0}", getMatchesResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R443, the return count of rows in GetMatchs response is {0}", getMatchesResponseBody.RowCount);

            // If ErrorCode field is 0x0000000 and the RowsCount has values, then server returns the Explicit Table.
            // So R443 will be verified.
            this.Site.CaptureRequirementIfIsTrue(
                getMatchesResponseBody.ErrorCode == 0 && getMatchesResponseBody.RowData != null,
                443,
                @"[In GetMatches Request Type] The GetMatches request type is used by the client to get an Explicit Table, in which the rows are determined by the specified criteria.");
            #endregion

            #region Call GetMatches request type without State.
            STAT emptyStat = new STAT();
            getMatchRequestBody = this.BuildGetMatchRequestBody(false, emptyStat, false, 0, null, true, filter, false, propertyGuid, propertyNameId, 1, true, propTagArray);
            getMatchesResponseBody = this.Adapter.GetMatches(getMatchRequestBody);
            this.Site.Assert.AreEqual<uint>(0, getMatchesResponseBody.StatusCode, "GetMatches request should be executed successfully, the returned error code is {0}", getMatchesResponseBody.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to QueryRows request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC11_QueryRowsRequestType()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call QueryRows with all optional fields exist and the type of all properties specified.
            uint tableCount = 0;
            uint[] table = null;
            uint rowCount = ConstValues.QueryRowsRequestedRowNumber;
            STAT stat = new STAT();
            stat.InitiateStat();

            LargePropertyTagArray columns = new LargePropertyTagArray()
            {
                PropertyTagCount = 2,
                PropertyTags = new PropertyTag[]
                {
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyType.PtypString,
                        PropertyId = (ushort)PropertyID.PidTagDisplayName
                    }, 
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyType.PtypInteger32,
                        PropertyId = (ushort)PropertyID.PidTagDisplayType
                    }
                }
            };

            QueryRowsRequestBody queryRowsRequestBody = this.BuildQueryRowsRequestBody(true, stat, tableCount, table, rowCount, true, columns);
            QueryRowsResponseBody queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.ErrorCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.ErrorCode);

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R802");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R802
            this.Site.CaptureRequirementIfAreEqual<uint>(
                columns.PropertyTagCount,
                queryRowsResponseBody.Columns.Value.PropertyTagCount,
                802,
                @"[In QueryRows Request Type] The QueryRows request type is used by the client to get a number of rows from the specified Explicit Table.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R827");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R827
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoLargePropertyTagArrayEqual(queryRowsRequestBody.Columns, queryRowsResponseBody.Columns.Value),
                827,
                @"[In QueryRows Request Type Request Body] Columns (optional) (variable): A LargePropTagArray structure (section 2.2.1.3) that specifies the properties that the client requires for each row returned.");

            AddressBookPropertyRow[] rowData = queryRowsResponseBody.RowData;
            for (int i = 0; i < rowData.Length; i++)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R455");

                int lengthString = rowData[i].ValueArray[0].Value.Length;
                bool result = (rowData[i].ValueArray[0].Value[lengthString - 1] == 0) && (rowData[i].ValueArray[0].Value[lengthString - 2] == 0);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R455
                this.Site.CaptureRequirementIfIsTrue(
                    result,
                    "MS-OXCDATA",
                    455,
                    @"[In PropertyValue Structure] PropertyValue (variable): If the property value being passed is a string, the data includes the terminating null characters.");

                List<AddressBookPropertyValue> valueArray = new List<AddressBookPropertyValue>(rowData[i].ValueArray);

                for (int j = 0; j < valueArray.Count; j++)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R18");

                    bool isVerifyR18 = rowData[i].Flag == 0x0 || rowData[i].Flag == 0x1;

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R18
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifyR18,
                        18,
                        @"[In AddressBookPropertyRow Structure] [Flags] The flag MUST be set to one of the values [0x0 and 0x1] in the following table.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2036");

                    // MS-OXCMAPIHTTP_R18 is verified, and AddressBookPropertyRow structure in QueryRows response body is parsed successfully,
                    // so MS-OXCMAPIHTTP_R2036 can be verified directly if code can reach here.
                    this.Site.CaptureRequirement(
                        2036,
                        @"[In AddressBookPropertyRow Structure] [ValueArray] Each structure in the array MUST be interpreted based on the Flag field.");

                    if (rowData[i].Flag == 0x0)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R450");

                        // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R450
                        // In this step, the type of all properties required to be returned is not PtypUnspecified, that is, their type is specified.
                        // So if the type of each value returned in the response is AddressBookPropertyValue, MS-OXCMAPIHTTP_R450 can be verified.
                        this.Site.CaptureRequirementIfIsInstanceOfType(
                            valueArray[j],
                            typeof(AddressBookPropertyValue),
                            450,
                            @"[In AddressBookPropertyRow Structure] [Flags] If the value of the Flags field is set to 0x0: The ValueArray field contains either an AddressBookPropertyValue structure, see section 2.2.1.1, if the type of the property is specified.");

                        if (valueArray[j].HasValue == 0xFF)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2006");

                            // Verify MS-OXCMAPIHTTP_R2006
                            this.Site.CaptureRequirementIfIsNotNull(
                                valueArray[j].Value,
                                2006,
                                @"[In AddressBookPropertyValue Structure] [HasValue] A TRUE value means that the PropertyValue field is present.");
                        }
                        else
                        {
                            if(valueArray[j].HasValue == null)
                            {
                                // Add the debug information
                                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2010");

                                // Verify MS-OXCMAPIHTTP_R2010
                                this.Site.CaptureRequirementIfIsNotNull(
                                    valueArray[j].Value,
                                    2010,
                                    @"[In AddressBookPropertyValue Structure] [PropertyValue] This field is always present when HasValue is not present.");
                            }

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2009");

                            // Verify MS-OXCMAPIHTTP_R2009
                            this.Site.CaptureRequirementIfIsNotNull(
                                valueArray[j].Value,
                                2009,
                                @"[In AddressBookPropertyValue Structure] PropertyValue (optional) (variable): A PropertyValue structure ([MS-OXCDATA] section 2.11.2.1), unless HasValue is present with a value of FALSE (0x00).");
                        }

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R454");

                        // Verify MS-OXCDATA requirement: MS-OXCDATA_R454
                        // In this step, the type of all properties required to be returned is not PtypUnspecified, that is, their type is specified.
                        // So if the type of each value returned in the response is AddressBookPropertyValue, MS-OXCMAPIHTTP_R454 can be verified.
                        this.Site.CaptureRequirementIfIsInstanceOfType(
                            valueArray[j],
                            typeof(AddressBookPropertyValue),
                            "MS-OXCDATA",
                            454,
                            @"[In PropertyValue Structure] PropertyValue (variable):  For multivalue types, the first element in the ROP buffer is a 16-bit integer specifying the number of entries.");
                    }
                }
            }

            // The list of property tags is already specified in columns field in QueryRows request body, and AddressBookPropertyRow structure in QueryRows response body is parsed successfully,
            // so MS-OXCMAPIHTTP_R16 can be verified directly if code can reach here.
            this.Site.CaptureRequirement(
                16,
                @"[In AddressBookPropertyRow Structure] [AddressBookPropertyRow] This structure is used when the list of property tags is known in advance.");
            #endregion
            #endregion

            #region Call QueryRows with Flags set to one value other than fEphID (0x00000002) and fSkipObjects (0x00000001).
            queryRowsRequestBody.Flags = 0x3;
            QueryRowsResponseBody queryRowsResponseBody1 = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody1.ErrorCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody1.ErrorCode);
            #endregion

            #region Call QueryRows with Flags set to another value other than fEphID (0x00000002) and fSkipObjects (0x00000001).
            queryRowsRequestBody.Flags = 0x4;
            QueryRowsResponseBody queryRowsResponseBody2 = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody2.ErrorCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody2.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1326.");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1326
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoAddressBookPropertyRowEqual(queryRowsResponseBody1.RowData, queryRowsResponseBody2.RowData),
                1326,
                @"[In QueryRows Request Type Request Body] If this field [Flags] is set to different values other than the bit flag fEphID (0x00000002) and fSkipObjects (0x00000001), the server will return the same result.");
            #endregion

            #region Call QueryRows with HasState field set to false and HasColumns set to true.
            queryRowsRequestBody = this.BuildQueryRowsRequestBody(false, stat, tableCount, table, rowCount, true, columns);
            queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.StatusCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.StatusCode);
            #endregion

            #region Call QueryRows with HasState field set to false and HasColumns set to false.
            queryRowsRequestBody = this.BuildQueryRowsRequestBody(false, stat, tableCount, table, rowCount, false, columns);
            queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.StatusCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.StatusCode);
            #endregion

            #region Call QueryRows with HasState field set to true and HasColumns set to false.
            queryRowsRequestBody = this.BuildQueryRowsRequestBody(true, stat, tableCount, table, rowCount, false, columns);
            queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.ErrorCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.ErrorCode);

            AddressBookPropertyRow[] rowData2 = queryRowsResponseBody.RowData;
            for (int i = 0; i < rowData2.Length; i++)
            {
                List<AddressBookPropertyValue> valueArray = new List<AddressBookPropertyValue>(rowData2[i].ValueArray);

                for (int j = 0; j < valueArray.Count; j++)
                {
                    if (rowData2[i].Flag == 0x01)
                    {                       
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R453");

                        // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R453
                        // In this step, the type of all properties required to be returned is not PtypUnspecified, that is, their type is specified.
                        // So if the type of each value returned in the response is AddressBookFlaggedPropertyValue, MS-OXCMAPIHTTP_R451 can be verified.
                        this.Site.CaptureRequirementIfIsInstanceOfType(
                            valueArray[j],
                            typeof(AddressBookFlaggedPropertyValue),
                            453,
                            @"[In AddressBookPropertyRow Structure] [Flags] If the value of the Flags field is set to 0x1: The ValueArray field contains either an AddressBookFlaggedPropertyValue structure, see section 2.2.1.5, if the type of the property is specified.");

                        AddressBookFlaggedPropertyValue flaggedPropertyValueForMapiHTTP = (AddressBookFlaggedPropertyValue)valueArray[j];

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R473");

                        // Verify MS-OXCMAPIHTTP requirement: MS-OXCDATA_R473
                        // In this step, the type of all properties required to be returned is not PtypUnspecified, that is, their type is specified.
                        // So if the flag type of each value returned in the response is byte, MS-OXCMAPIHTTP_R473 can be verified.
                        this.Site.CaptureRequirementIfIsInstanceOfType(
                            flaggedPropertyValueForMapiHTTP.Flag,
                            typeof(byte),
                            "MS-OXCDATA",
                            473,
                            @"[In FlaggedPropertyValue Structure] Flag (1 byte): An unsigned integer.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R475");

                        // Verify MS-OXCMAPIHTTP requirement: MS-OXCDATA_R475
                        // In this step, the type of all properties required to be returned is not PtypUnspecified, that is, their type is specified.
                        // So if the value of Flag in the response is one of the values [0x0, 0x1, 0xA], MS-OXCDATA_R475 can be verified.
                        bool isVerifiedR475 = flaggedPropertyValueForMapiHTTP.Flag == 0x0 ||
                                         flaggedPropertyValueForMapiHTTP.Flag == 0x1 ||
                                         flaggedPropertyValueForMapiHTTP.Flag == 0xA;

                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifiedR475,
                            "MS-OXCDATA",
                            475,
                            @"[In FlaggedPropertyValue Structure] Flag (1 byte): The flag MUST be set to one of the values [0x0, 0x1, 0xA] in the following table.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2021");

                        //Since MS-OXCDATA_R475 is verified, MS-OXCMAPIHTTP_R2021 can be captured directly.
                        this.Site.CaptureRequirement(
                            2021,
                            @"[In AddressBookFlaggedPropertyValue Structure] [Flag]The flag MUST be set to one of the values [0x0, 0x1, 0xA] in the following table.");

                        AddressBookFlaggedPropertyValue propertyValue = (AddressBookFlaggedPropertyValue)valueArray[j];

                        if (propertyValue.Flag != 0x01)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R479.");

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R479
                            // If the value of property exists, it indicates that the PropertyValue field exists.
                            this.Site.CaptureRequirementIfIsNotNull(
                                propertyValue.Value,
                                "MS-OXCDATA",
                                479,
                                @"[In FlaggedPropertyValue Structure] PropertyValue (optional) (variable): A PropertyValue structure, as specified in section 2.11.2.1, unless the Flag field is set to 0x1.");

                            if (propertyValue.Flag == 0x0)
                            {
                                // Add the debug information
                                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R476.");

                                // Verify MS-OXCDATA requirement: MS-OXCDATA_R476
                                // The parser code parses the PropertyValue field according to the property type. So R476 can be verified if the returned property value is not null.
                                this.Site.CaptureRequirementIfIsNotNull(
                                    propertyValue.Value,
                                    "MS-OXCDATA",
                                    476,
                                    @"[In FlaggedPropertyValue Structure] The Flag value 0x0 means the PropertyValue field will be a PropertyValue structure containing a value compatible with the property type implied by the context.");

                                // Add the debug information
                                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2022.");

                                // Verify MS-OXCDATA requirement: MS-OXCMAPIHTTP_R2022
                                // The parser code parses the PropertyValue field according to the property type. So R2022 can be verified if the returned property value is not null.
                                this.Site.CaptureRequirementIfIsNotNull(
                                    propertyValue.Value,
                                    2022,
                                    @"[In AddressBookFlaggedPropertyValue Structure] Flag value 0x0 meaning The PropertyValue field will be an AddressBookPropertyValue structure (section 2.2.1.1) containing a value compatible with the property type implied by the context.");
                            }
                        }
                    }
                }
            }
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to ResolveNames request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC12_ResolveNamesRequestType()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call ResolveNames request type to perform ANR.
            STAT stat = new STAT();
            stat.InitiateStat();
            byte[] auxIn = new byte[] { };

            ResolveNamesRequestBody resolveNamesRequestBody = new ResolveNamesRequestBody();
            resolveNamesRequestBody.Reserved = 0;
            resolveNamesRequestBody.HasState = true;
            resolveNamesRequestBody.State = stat;
            resolveNamesRequestBody.HasPropertyTags = true;
            LargePropertyTagArray largePropertyTagArray = new LargePropertyTagArray()
            {
                PropertyTagCount = 1,
                PropertyTags = new PropertyTag[]
                {
                    new PropertyTag 
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypString,
                        PropertyId = (ushort)PropertyID.PidTagDisplayName
                    }
                }
            };
            resolveNamesRequestBody.PropertyTags = largePropertyTagArray;
            resolveNamesRequestBody.HasNames = true;
            WStringsArray_r stringsArray = new WStringsArray_r();
            stringsArray.CValues = 5;
            stringsArray.LppszW = new string[stringsArray.CValues];
            stringsArray.LppszW[0] = this.AdminUserDN;
            stringsArray.LppszW[1] = string.Empty;
            stringsArray.LppszW[2] = Common.GetConfigurationPropertyValue("AmbiguousName", this.Site);
            stringsArray.LppszW[3] = null;
            stringsArray.LppszW[4] = "XXXXXX";

            resolveNamesRequestBody.Names = stringsArray;
            resolveNamesRequestBody.AuxiliaryBuffer = auxIn;
            resolveNamesRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            ResolveNamesResponseBody resolveNamesResponseBody = this.Adapter.ResolveNames(resolveNamesRequestBody);
            Site.Assert.AreEqual<uint>(0, resolveNamesResponseBody.ErrorCode, "ResolveNames should succeed and 0 is expected to be returned. The returned value is {0}.", resolveNamesResponseBody.ErrorCode);

            #region Capture code.
            bool isMinimalIdsCorrect = false;
            for (int i = 0; i < resolveNamesResponseBody.MinimalIds.Length; i++)
            {
                isMinimalIdsCorrect = resolveNamesResponseBody.MinimalIds[i] == (uint)ANRMinEntryIDs.MID_AMBIGUOUS ||
                                      resolveNamesResponseBody.MinimalIds[i] == (uint)ANRMinEntryIDs.MID_RESOLVED ||
                                      resolveNamesResponseBody.MinimalIds[i] == (uint)ANRMinEntryIDs.MID_UNRESOLVED;
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R895");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R895
            this.Site.CaptureRequirementIfIsTrue(
                isMinimalIdsCorrect,
                895,
                @"[In ResolveNames Request Type] The ResolveNames request type is used by the client to perform ambiguous name resolution (ANR), as specified in [MS-OXNSPI] section 3.1.4.7.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R913");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R913
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoLargePropertyTagArrayEqual(resolveNamesRequestBody.PropertyTags, resolveNamesResponseBody.PropertyTags.Value),
                913,
                @"[In ResolveNames Request Type Request Body] PropertyTags (optional) (variable): A LargePropTagArray structure (section 2.2.1.3) that specifies the properties that client requires for the rows returned.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R946, the value of PropertyTags is {0}, the value of RowData is {1}.",
                resolveNamesResponseBody.PropertyTags.Value.PropertyTags,
                resolveNamesResponseBody.RowData);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R946
            bool isVerifiedR946 = resolveNamesResponseBody.PropertyTags.GetType().Equals(typeof(LargePropertyTagArray)) &&
                                  resolveNamesResponseBody.PropertyTags.Value.PropertyTags != null &&
                                  resolveNamesResponseBody.RowData != null;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR946,
                946,
                @"[In ResolveNames Request Type Success Response Body] PropertyTags (optional) (variable): A LargePropTagArray structure (section 2.2.1.3) that specifies the properties returned for the rows in the RowData field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R942");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R942
            bool isVerifiedR942 = resolveNamesResponseBody.MinimalIds.GetType().IsArray && isMinimalIdsCorrect;

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R942
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR942,
                942,
                @"[In ResolveNames Request Type Success Response Body] MinimalIds (optional) (variable): An array of MinimalEntryID structures ([MS-OXNSPI] section 2.3.8.1), each of which specifies a Minimal Entry ID matching a name requested by the client.");

            #endregion
            #endregion

            #region Call ResolveNames request type with HasState, HasPropertyTags and HasNames set to false to perform ambiguous name resolution (ANR).
            resolveNamesRequestBody.HasState = false;
            resolveNamesRequestBody.HasPropertyTags = false;
            resolveNamesRequestBody.HasNames = false;

            resolveNamesResponseBody = this.Adapter.ResolveNames(resolveNamesRequestBody);
            Site.Assert.AreEqual<uint>(0, resolveNamesResponseBody.StatusCode, "ResolveNames should succeed and 0 is expected to be returned. The returned value is {0}.", resolveNamesResponseBody.StatusCode);
            #endregion

            #region Call ResolveNames request type with HasState and HasPropertyTags set to true, HasNames set to false to perform ambiguous name resolution (ANR).
            resolveNamesRequestBody.HasState = true;
            resolveNamesRequestBody.HasPropertyTags = true;
            resolveNamesRequestBody.HasNames = false;

            resolveNamesResponseBody = this.Adapter.ResolveNames(resolveNamesRequestBody);
            Site.Assert.AreEqual<uint>(0, resolveNamesResponseBody.StatusCode, "ResolveNames should succeed and 0 is expected to be returned. The returned value is {0}.", resolveNamesResponseBody.StatusCode);
            #endregion

            #region Call ResolveNames request type with HasState and HasNames set to true, HasPropertyTags set to false to perform ambiguous name resolution (ANR).
            resolveNamesRequestBody.HasState = true;
            resolveNamesRequestBody.HasPropertyTags = false;
            resolveNamesRequestBody.HasNames = true;

            resolveNamesResponseBody = this.Adapter.ResolveNames(resolveNamesRequestBody);
            Site.Assert.AreEqual<uint>(0, resolveNamesResponseBody.ErrorCode, "ResolveNames should succeed and 0 is expected to be returned. The returned value is {0}.", resolveNamesResponseBody.ErrorCode);
            #endregion

            #region Call ResolveNames request type with HasState set to false, HasPropertyTags and HasNames set to true to perform ambiguous name resolution (ANR).
            resolveNamesRequestBody.HasState = false;
            resolveNamesRequestBody.HasPropertyTags = true;
            resolveNamesRequestBody.HasNames = true;

            resolveNamesResponseBody = this.Adapter.ResolveNames(resolveNamesRequestBody);
            Site.Assert.AreEqual<uint>(0, resolveNamesResponseBody.StatusCode, "ResolveNames should succeed and 0 is expected to be returned. The returned value is {0}.", resolveNamesResponseBody.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to SeekEntries request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC13_SeekEntriesRequestType()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call QueryRows request type with propTags set to null.
            STAT stat = new STAT();
            stat.InitiateStat();
            LargePropertyTagArray columns = new LargePropertyTagArray();
            QueryRowsRequestBody queryRowsRequestBody = this.BuildQueryRowsRequestBody(true, stat, 0, null, ConstValues.QueryRowsRequestedRowNumber, false, columns);
            QueryRowsResponseBody queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.ErrorCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.ErrorCode);
            uint[] outMids = new uint[queryRowsResponseBody.RowCount.Value];
            for (int i = 0; i < queryRowsResponseBody.RowCount; i++)
            {
                outMids[i] = BitConverter.ToUInt32(queryRowsResponseBody.RowData[i].ValueArray[0].Value, 0);
            }

            stat = queryRowsResponseBody.State.Value;
            #endregion

            #region Call UpdateStat request type to update the STAT block that represents the position in a table to reflect positioning changes requested by the client.
            stat.Delta = 1;

            byte[] auxIn = new byte[] { };
            UpdateStatRequestBody updateStatRequestBody = this.BuildUpdateStatRequestBody(true, stat, true);
            UpdateStatResponseBody updateStatResponseBody = this.Adapter.UpdateStat(updateStatRequestBody);
            Site.Assert.AreEqual<uint>(0, updateStatResponseBody.ErrorCode, "UpdateStat should succeed and 0 is expected to be returned. The returned value is {0}.", updateStatResponseBody.ErrorCode);
            #endregion

            #region Call SeekEntries request type which contains fields State, Trage and Columns.
            PropertyTag pidTagDisplayNameTag = new PropertyTag((ushort)PropertyID.PidTagDisplayName, (ushort)PropertyTypeValues.PtypString);
            string displayName = Common.GetConfigurationPropertyValue("GeneralUserName", this.Site) + "\0";

            PropertyValue_r target = new PropertyValue_r()
            {
                PropTag = pidTagDisplayNameTag,
                Reserved = 0,
                Value = System.Text.Encoding.Unicode.GetBytes(displayName)
            };

            columns = new LargePropertyTagArray()
            {
                PropertyTagCount = 1,
                PropertyTags = new PropertyTag[1]
                {
                    target.PropTag
                }
            };

            SeekEntriesRequestBody seekEntriesRequestBody = this.BuildSeekEntriesRequestBody(true, updateStatResponseBody.State.Value, true, target, true, (uint)outMids.Length, outMids, true, columns);

            SeekEntriesResponseBody seekEntriesResponseBody = this.Adapter.SeekEntries(seekEntriesRequestBody);
            Site.Assert.AreEqual<uint>(0, seekEntriesResponseBody.ErrorCode, "SeekEntries should succeed and 0 is expected to be returned. The returned value is {0}.", seekEntriesResponseBody.ErrorCode);
            #endregion
         
            #region Call QueryRows request type using a new list of minimal entry ids as input parameter according to the parameters of SeekEntries request type response.
            uint[] newTable = new uint[outMids.Length - (int)seekEntriesResponseBody.State.Value.NumPos];
            Array.Copy(outMids, (int)seekEntriesResponseBody.State.Value.NumPos, newTable, 0, newTable.Length);

            queryRowsRequestBody = this.BuildQueryRowsRequestBody(true, stat, (uint)newTable.Length, newTable, ConstValues.QueryRowsRequestedRowNumber, true, columns);
            queryRowsRequestBody.Flags = (uint)RetrievePropertyFlags.fEphID;
            queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.StatusCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.StatusCode);
  
            if (seekEntriesResponseBody.HasColumnsAndRows == false)
            {
                seekEntriesResponseBody = this.Adapter.SeekEntries(seekEntriesRequestBody);
                Site.Assert.AreEqual<uint>(0, seekEntriesResponseBody.ErrorCode, "SeekEntries should succeed and 0 is expected to be returned. The returned value is {0}.", seekEntriesResponseBody.ErrorCode);
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1013");

            // If the RowData is not Null, then client retrieves information about rows in an Explicit Table by the SeekEntries request type.
            // So R1013 will be verified.
            this.Site.CaptureRequirementIfIsTrue(
                seekEntriesResponseBody.ErrorCode == 0x00000000 && seekEntriesResponseBody.RowData != null,
                1013,
                @"[In SeekEntries Request Type] Optionally, the SeekEntries request type can also be used to retrieve information about rows in an Explicit Table.");

            bool isTheValueGreaterThanOrEqualTo = false;
            for (int i = 0; i < seekEntriesResponseBody.RowData.Length; i++)
            {
                // Check whether the returned property value is greater than or equal to the input property PidTagDisplayName.
                string dispalyNameReturnedFromServer = System.Text.Encoding.UTF8.GetString(seekEntriesResponseBody.RowData[i].ValueArray[0].Value);

                // Value Condition greater than zero indicates the returned property value follows the input PidTagDisplayName property value.
                int result = dispalyNameReturnedFromServer.CompareTo(displayName);
                if (result < 0)
                {
                    this.Site.Log.Add(LogEntryKind.Debug, "The display name returned from server is {0}, the expected display name is {1}", dispalyNameReturnedFromServer.Replace("\0", string.Empty), displayName.Replace("\0", string.Empty));

                    isTheValueGreaterThanOrEqualTo = false;
                    break;
                }
                else
                {
                    isTheValueGreaterThanOrEqualTo = true;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1012");

            this.Site.CaptureRequirementIfIsTrue(
                isTheValueGreaterThanOrEqualTo,
                1012,
                @"[In SeekEntries Request Type] The SeekEntries request type is used by the client to search for and set the logical position in a specific table to the first entry greater than or equal to a specified value.");

            bool isLargePropertyTagArrayEqual = AdapterHelper.AreTwoLargePropertyTagArrayEqual(seekEntriesRequestBody.Columns, seekEntriesResponseBody.Columns.Value);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1047");

            this.Site.CaptureRequirementIfIsTrue(
                isLargePropertyTagArrayEqual,
                1047,
                @"[In SeekEntries Request Type Success Response Body] Columns (optional) (variable): A LargePropTagArray structure (section 2.2.1.3) that specifies the columns used for the rows returned.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1053");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1053
            this.Site.CaptureRequirementIfIsInstanceOfType(
                seekEntriesResponseBody.RowData,
                typeof(AddressBookPropertyRow[]),
                1053,
                @"[In SeekEntries Request Type Success Response Body] RowData (optional) (variable): An array of AddressBookPropertyRow structures (section 2.2.1.2), each of which specifies the row data for the entries queried.");
            #endregion

            #region Call SeekEntries request type without State field.
            seekEntriesRequestBody = this.BuildSeekEntriesRequestBody(false, stat, true, target, false, 0, null, true, columns);

            seekEntriesResponseBody = this.Adapter.SeekEntries(seekEntriesRequestBody);
            Site.Assert.AreEqual<uint>(0, seekEntriesResponseBody.StatusCode, "SeekEntries should succeed and 0 is expected to be returned. The returned value is {0}.", seekEntriesResponseBody.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify requirements related to GetPropList request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC14_GetPropList()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call GetPropList method with Flags field set to 0x1 and MinimalId set to 0 to get the default global address book object.
            uint mid = 0;
            GetPropListRequestBody getPropListRequestBody = this.BuildGetPropListRequestBody(0x1, mid);
            GetPropListResponseBody getPropListResponseBody = this.Adapter.GetPropList(getPropListRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropListResponseBody.ErrorCode, "GetPropList operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropListResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R583");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R583
            // In this step, MinimalId is set to 0 to get the default global address book object. So if the returned propertyTags is not null, 
            // it indicates that this field contains the property tags of properties that have values on the requested object.
            this.Site.CaptureRequirementIfIsNotNull(
                getPropListResponseBody.PropertyTags,
                583,
                @"[In GetPropList Request Type Success Response Body] PropertyTags (optional) (variable): A LargePropertyTagArray structure (section 2.2.1.3) that contains the property tags of properties that have values on the requested object.");
            #endregion

            #region Call QueryColumns to get all of the properties that exist in the address book.
            uint mapiFlags = 0x80000000;
            QueryColumnsRequestBody queryColumnsRequestBody = this.BuildQueryColumnsRequestBody(mapiFlags);
            QueryColumnsResponseBody queryColumnsResponseBody = this.Adapter.QueryColumns(queryColumnsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryColumnsResponseBody.ErrorCode, "QueryColumns operation should succeed and 0 is expected to be returned. The return value is {0}.", queryColumnsResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R558");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R558
            this.Site.CaptureRequirementIfAreEqual<uint>(
                queryColumnsResponseBody.Columns.Value.PropertyTagCount,
                getPropListResponseBody.PropertyTags.Value.PropertyTagCount,
                558,
                @"[In GetPropList Request Type] The GetPropList request type is used by the client to get a list of all of the properties that have values on an object.");
            #endregion

            #region Call GetPropList method with Flags field set to 0x0 and MinimalId set to 0 to get the default global address book object.
            getPropListRequestBody = this.BuildGetPropListRequestBody(0x0, mid);
            GetPropListResponseBody getPropListResponseBody1 = this.Adapter.GetPropList(getPropListRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropListResponseBody1.ErrorCode, "GetPropList operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropListResponseBody.ErrorCode);
            #endregion

            #region Call GetPropList method with Flags field set to 0x2 to check that server will ignore other flag values rather than 0x1.
            getPropListRequestBody = this.BuildGetPropListRequestBody(0x2, mid);
            GetPropListResponseBody getPropListResponseBody2 = this.Adapter.GetPropList(getPropListRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropListResponseBody2.ErrorCode, "GetPropList operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropListResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1321");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1321
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoLargePropertyTagArrayEqual(getPropListResponseBody1.PropertyTags.Value, getPropListResponseBody2.PropertyTags.Value),
                1321,
                @"[In GetPropList Request Type Request Body] If this field [Flags] is set to different values other than the bit flag fSkipObjects (0x00000001), the server will return the same result.");
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify requirements related to QueryColumns request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC15_QueryColumns()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call QueryColumns to get all of the properties that exist in the address book.
            QueryColumnsRequestBody queryColumnsRequestBody = this.BuildQueryColumnsRequestBody((uint)NspiQueryColumnsFlag.NspiUnicodeProptypes);
            QueryColumnsResponseBody queryColumnsResponseBody = this.Adapter.QueryColumns(queryColumnsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryColumnsResponseBody.ErrorCode, "QueryColumns operation should succeed and 0 is expected to be returned. The return value is {0}.", queryColumnsResponseBody.ErrorCode);
            #endregion

            #region Call GetPropList method with Flags field set to 0x1 and MinimalId set to 0 to get the default global address book object.
            uint mid = 0;
            GetPropListRequestBody getPropListRequestBody = this.BuildGetPropListRequestBody(0x1, mid);
            GetPropListResponseBody getPropListResponseBody = this.Adapter.GetPropList(getPropListRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropListResponseBody.ErrorCode, "GetPropList operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropListResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R865");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R865
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getPropListResponseBody.PropertyTags.Value.PropertyTagCount,
                queryColumnsResponseBody.Columns.Value.PropertyTagCount,
                865,
                @"[In QueryColumns Request Type] The QueryColumns request type is used by the client to get a list of all of the properties that exist in the address book.");
            #endregion

            #region Call QueryColumns with MapiFlags set to one value rather than 0x80000000.
            uint mapiFlags1 = 0x0;
            QueryColumnsRequestBody queryColumnsRequestBody1 = this.BuildQueryColumnsRequestBody(mapiFlags1);
            QueryColumnsResponseBody queryColumnsResponseBody1 = this.Adapter.QueryColumns(queryColumnsRequestBody1);
            Site.Assert.AreEqual<uint>((uint)0, queryColumnsResponseBody1.ErrorCode, "QueryColumns operation should succeed and 0 is expected to be returned. The return value is {0}.", queryColumnsResponseBody1.ErrorCode);
            #endregion

            #region Call QueryColumns with MapiFlags set to another value rather than 0x80000000.
            uint mapiFlags2 = 0x1;
            QueryColumnsRequestBody queryColumnsRequestBody2 = this.BuildQueryColumnsRequestBody(mapiFlags2);
            QueryColumnsResponseBody queryColumnsResponseBody2 = this.Adapter.QueryColumns(queryColumnsRequestBody2);
            Site.Assert.AreEqual<uint>((uint)0, queryColumnsResponseBody2.ErrorCode, "QueryColumns operation should succeed and 0 is expected to be returned. The return value is {0}.", queryColumnsResponseBody2.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1412");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1412
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoLargePropertyTagArrayEqual(queryColumnsResponseBody1.Columns.Value, queryColumnsResponseBody2.Columns.Value),
                1412,
                @"[In QueryColumns Request Type Request Body] If this field [MapiFlags] is set to different values other than the bit flag NspiUnicodeProptypes (0x80000000), the server will return the same result.");
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify requirements related to GetProps request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC16_GetProps()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call UpdateStat to update the STAT block to make CurrentRec point to the first row of the table.
            STAT stat = new STAT();
            stat.InitiateStat();
            UpdateStatRequestBody updateStatRequestBody = this.BuildUpdateStatRequestBody(true, stat, true);

            UpdateStatResponseBody updateStatResponseBody = this.Adapter.UpdateStat(updateStatRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, updateStatResponseBody.ErrorCode, "UpdateStat operation should succeed and 0 is expected to be returned. The return value is {0}.", updateStatResponseBody.ErrorCode);
            #endregion

            #region Call GetProps request type with Flags set to fEphID.
            LargePropertyTagArray largePropTagArray = new LargePropertyTagArray();
            largePropTagArray.PropertyTagCount = 1;

            // PidTagDisplayName property.
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)PropertyID.PidTagDisplayName;
            propertyTags[0].PropertyType = (ushort)PropertyTypeValues.PtypString;
            largePropTagArray.PropertyTags = propertyTags;

            GetPropsRequestBody getPropertyRequestBody = this.BuildGetPropsRequestBody((uint)RetrievePropertyFlags.fEphID, true, updateStatResponseBody.State, true, largePropTagArray);
            GetPropsResponseBody getPropertyResponseBody = this.Adapter.GetProps(getPropertyRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropertyResponseBody.ErrorCode, "GetProps operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropertyResponseBody.ErrorCode);

            AddressBookTaggedPropertyValue[] propertyValues = getPropertyResponseBody.PropertyValues.Value.PropertyValues;

            bool isVerifiedR624 = false;

            if (getPropertyResponseBody.PropertyValues.Value.PropertyValueCount != largePropTagArray.PropertyTagCount)
            {
                isVerifiedR624 = false;
            }
            else
            {
                for (int i = 0; i < getPropertyResponseBody.PropertyValues.Value.PropertyValueCount; i++)
                {
                    if (propertyValues[i].PropertyId == propertyTags[i].PropertyId && propertyValues[i].PropertyType == propertyTags[i].PropertyType)
                    {
                        isVerifiedR624 = true;
                    }
                    else
                    {
                        isVerifiedR624 = false;
                        break;
                    }
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R624. The returned property numbers is {0}.", getPropertyResponseBody.PropertyValues.Value.PropertyValueCount);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R624
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR624,
                624,
                @"[In GetProps Request Type Success Response Body] PropertyValues (optional) (variable): An AddressBookPropertyValueList structure (section 2.2.1.1) that contains the values of the properties requested.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R590. The returned property number is {0}, the returned property ID is {1}, the returned property type is {2}.",
                getPropertyResponseBody.PropertyValues.Value.PropertyValueCount,
                propertyValues[0].PropertyId,
                propertyValues[0].PropertyType);
            #endregion

            #region Call GetProps request type with Flags set to fSkipObjects.
            getPropertyRequestBody = this.BuildGetPropsRequestBody((uint)RetrievePropertyFlags.fSkipObjects, true, updateStatResponseBody.State, true, largePropTagArray);
            getPropertyResponseBody = this.Adapter.GetProps(getPropertyRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropertyResponseBody.ErrorCode, "GetProps operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropertyResponseBody.ErrorCode);
            #endregion

            #region Call GetProps request type with Flags set to one value other than fEphID (0x2) and fSkipObjects (0x1).
            uint flags1 = 0x0;
            getPropertyRequestBody = this.BuildGetPropsRequestBody(flags1, true, updateStatResponseBody.State, true, largePropTagArray);
            GetPropsResponseBody getPropertyResponseBody1 = this.Adapter.GetProps(getPropertyRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropertyResponseBody1.ErrorCode, "GetProps operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropertyResponseBody1.ErrorCode);
            #endregion

            #region Call GetProps request type with Flags set to one value other than fEphID (0x2) and fSkipObjects (0x1).
            uint flags2 = 0x8;
            getPropertyRequestBody = this.BuildGetPropsRequestBody(flags2, true, updateStatResponseBody.State, true, largePropTagArray);
            GetPropsResponseBody getPropertyResponseBody2 = this.Adapter.GetProps(getPropertyRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropertyResponseBody2.ErrorCode, "GetProps operation should succeed and 0 is expected to be returned. The return value is {0}.", getPropertyResponseBody2.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1323");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1323
            this.Site.CaptureRequirementIfIsTrue(
                AdapterHelper.AreTwoAddressBookPropValueListEqual(getPropertyResponseBody1.PropertyValues.Value, getPropertyResponseBody2.PropertyValues.Value),
                1323,
                @"[In GetProps Request Type Request Body] If this field [Flags] is set to different values other than the bit flags fEphID (0x00000002) and fSkipObjects (0x00000001), the server will return the same result.");
            #endregion

            #region Call GetProps request type with setting fields HasState and HasPropertyTags to false.
            getPropertyRequestBody = this.BuildGetPropsRequestBody((uint)RetrievePropertyFlags.fEphID, false, updateStatResponseBody.State, false, largePropTagArray);
            getPropertyResponseBody = this.Adapter.GetProps(getPropertyRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropertyResponseBody.StatusCode, "GetProps request should be executed successfully and 0 is expected to be returned. The return value is {0}.", getPropertyResponseBody.StatusCode);
            #endregion

            #region Call GetProps request type with HasState field set to false and HasPropertyTags set to true.
            getPropertyRequestBody = this.BuildGetPropsRequestBody((uint)RetrievePropertyFlags.fEphID, false, updateStatResponseBody.State, true, largePropTagArray);
            getPropertyResponseBody = this.Adapter.GetProps(getPropertyRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, getPropertyResponseBody.StatusCode, "GetProps request should be executed successfully and 0 is expected to be returned. The return value is {0}.", getPropertyResponseBody.StatusCode);
            #endregion

            #region Call GetProps request type with Flags set to fEphID.
            largePropTagArray = new LargePropertyTagArray();
            largePropTagArray.PropertyTagCount = 2;

            // PidTagDisplayName property.
            propertyTags = new PropertyTag[2];
            propertyTags[0].PropertyId = (ushort)PropertyID.PidTagDisplayName;
            propertyTags[0].PropertyType = (ushort)PropertyTypeValues.PtypString;
            propertyTags[1].PropertyId = (ushort)PropertyID.PidTagAddressBookMember;
            propertyTags[1].PropertyType = (ushort)PropertyTypeValues.PtypEmbeddedTable;
            largePropTagArray.PropertyTags = propertyTags;

            getPropertyRequestBody = this.BuildGetPropsRequestBody((uint)RetrievePropertyFlags.fEphID, true, updateStatResponseBody.State, true, largePropTagArray);
            uint responseCodeHeader = 0;
            getPropertyResponseBody = this.Adapter.GetProps(getPropertyRequestBody, out responseCodeHeader);

            if (Common.IsRequirementEnabled(2237, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2237.");

                // Verify MS-OXCDATA requirement: MS-OXCMAPIHTTP_R2237
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    responseCodeHeader,
                    2237,
                    @"[In Appendix A: Product Behavior] If the type of the returned property is PtypObject or PtypEmbeddedTable ([MS-OXCDATA] section 2.11.1), implementation will return value 1 for the X-ResponseCode header.  (Exchange 2013 SP1 follows this behavior.)");
            }
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify requirements related to ModProps request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC17_ModProps()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call DNToMId to get the MIDs of specified user.
            STAT stat = new STAT();
            stat.InitiateStat();

            byte[] auxIn = new byte[] { };
            uint reserved = 0;
            string userESSDN = this.AdminUserDN;
            DNToMinIdRequestBody requestBodyOfDNToMId = new DNToMinIdRequestBody()
            {
                Reserved = reserved,
                HasNames = true,
                Names = new StringArray_r
                {
                    CValues = 1,
                    LppzA = new string[]
                    {
                        userESSDN
                    }
                },
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            DnToMinIdResponseBody responseBodyOfdnToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
            this.Site.Assert.AreEqual((uint)0, responseBodyOfdnToMinId.ErrorCode, "DnToMinId method should be succeed and 0 is expected to be returned. The returned value is {0}.", responseBodyOfdnToMinId.ErrorCode);
            #endregion

            #region Call GetMatches to get the specific PidTagAddressBookX509Certificate property to be modified.
            stat.CurrentRec = responseBodyOfdnToMinId.MinimalIds[0];
            uint[] minimalIds = new uint[1]
            {
                (uint)responseBodyOfdnToMinId.MinimalIds[0]
            };

            byte[] filter = new byte[] { };
            Guid propertyGuid = new Guid();
            uint propertyNameId = new uint();
            LargePropertyTagArray columns = new LargePropertyTagArray
            {
                PropertyTagCount = 1,
                PropertyTags = new PropertyTag[]
                    {
                        new PropertyTag
                        {
                            PropertyType = (ushort)PropertyTypeValues.PtypMultipleBinary,
                            PropertyId = (ushort)PropertyID.PidTagAddressBookX509Certificate
                        }
                    }
            };
            GetMatchesRequestBody getMatchRequestBody = this.BuildGetMatchRequestBody(true, stat, true, 1, minimalIds, false, filter, false, propertyGuid, propertyNameId, 1, true, columns);

            GetMatchesResponseBody getMatchResponseBody = this.Adapter.GetMatches(getMatchRequestBody);
            this.Site.Assert.AreEqual((uint)0, getMatchResponseBody.ErrorCode, "GetMatches method should be succeed and 0 is expected to be returned. The returned value is {0}.", getMatchResponseBody.ErrorCode);
            #endregion

            #region Call ModProps request type with all optional fields.
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)PropertyID.PidTagAddressBookX509Certificate;
            propertyTags[0].PropertyType = (ushort)PropertyTypeValues.PtypMultipleBinary;

            LargePropertyTagArray propertyTagsToRemove = new LargePropertyTagArray()
            {
                PropertyTagCount = 1,
                PropertyTags = new PropertyTag[] 
                {
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypMultipleBinary,
                        PropertyId = (ushort)PropertyID.PidTagAddressBookX509Certificate
                    }
                }
            };

            ModPropsRequestBody modPropsRequestBody = this.BuildModPropsRequestBody(true, getMatchResponseBody.State.Value, true, propertyTags[0], true, propertyTagsToRemove);

            ModPropsResponseBody modPropsResponseBody = this.Adapter.ModProps(modPropsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, modPropsResponseBody.StatusCode, "ModProps request should succeed and 0 is expected to be returned. The return value is {0}.", modPropsResponseBody.StatusCode);
            #endregion

            #region Call ModProps request type without all optional fields.
            ModPropsRequestBody modPropsRequestBody1 = this.BuildModPropsRequestBody(false, getMatchResponseBody.State.Value, false, propertyTags[0], false, propertyTagsToRemove);

            ModPropsResponseBody modPropsResponseBody1 = this.Adapter.ModProps(modPropsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, modPropsResponseBody1.StatusCode, "ModProps request should succeed and 0 is expected to be returned. The return value is {0}.", modPropsResponseBody1.StatusCode);
            #endregion

            #region Call ModProps request type without PropertyTagsToRemove field.
            ModPropsRequestBody modPropsRequestBody2 = this.BuildModPropsRequestBody(true, getMatchResponseBody.State.Value, true, propertyTags[0], false, propertyTagsToRemove);

            ModPropsResponseBody modPropsResponseBody2 = this.Adapter.ModProps(modPropsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, modPropsResponseBody2.StatusCode, "ModProps request should succeed and 0 is expected to be returned. The return value is {0}.", modPropsResponseBody2.StatusCode);
            #endregion

            #region Call ModProps request type without PropertyValues field.
            ModPropsRequestBody modPropsRequestBody3 = this.BuildModPropsRequestBody(true, getMatchResponseBody.State.Value, false, propertyTags[0], true, propertyTagsToRemove);

            ModPropsResponseBody modPropsResponseBody3 = this.Adapter.ModProps(modPropsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, modPropsResponseBody3.StatusCode, "ModProps request should succeed and 0 is expected to be returned. The return value is {0}.", modPropsResponseBody3.StatusCode);
            #endregion

            #region Call ModProps request type without PropertyValues and state fields.
            ModPropsRequestBody modPropsRequestBody4 = this.BuildModPropsRequestBody(false, getMatchResponseBody.State.Value, false, propertyTags[0], true, propertyTagsToRemove);

            ModPropsResponseBody modPropsResponseBody4 = this.Adapter.ModProps(modPropsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, modPropsResponseBody4.StatusCode, "ModProps request should succeed and 0 is expected to be returned. The return value is {0}.", modPropsResponseBody4.StatusCode);
            #endregion

            #region Call ModProps request type without PropertyTagsToRemove and state fields.
            ModPropsRequestBody modPropsRequestBody5 = this.BuildModPropsRequestBody(false, getMatchResponseBody.State.Value, true, propertyTags[0], false, propertyTagsToRemove);

            ModPropsResponseBody modPropsResponseBody5 = this.Adapter.ModProps(modPropsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, modPropsResponseBody5.StatusCode, "ModProps request should succeed and 0 is expected to be returned. The return value is {0}.", modPropsResponseBody5.StatusCode);
            #endregion

            #region Call the Unbind request type to destroy the session context.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to ModLinkAtt request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC18_ModLinkAttRequestType()
        {
            this.CheckMapiHttpIsSupported();

            byte[] auxIn = new byte[] { };
   
            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call QueryRows request type to get a set of valid rows used to matched entry ID as the input paramter of ModLinkAtt method.
            STAT stat = new STAT();
            stat.InitiateStat();

            uint tableCount = 0;
            uint[] table = null;
            LargePropertyTagArray largePropTagArray = new LargePropertyTagArray()
            {
                PropertyTagCount = 4,
                PropertyTags = new PropertyTag[]
                {
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypString,
                        PropertyId = (ushort)PropertyID.PidTagDisplayName
                    }, 
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypBinary,
                        PropertyId = (ushort)PropertyID.PidTagEntryId
                    },
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypInteger32,
                        PropertyId = (ushort)PropertyID.PidTagDisplayType
                    },
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypMultipleString8,
                        PropertyId = (ushort)PropertyID.PidTagAddressBookMember
                    }
                }
            };

            QueryRowsRequestBody queryRowsRequestBody = new QueryRowsRequestBody();
            queryRowsRequestBody.Flags = (uint)RetrievePropertyFlags.fSkipObjects;
            queryRowsRequestBody.HasState = true;
            queryRowsRequestBody.State = stat;
            queryRowsRequestBody.ExplicitTableCount = tableCount;
            queryRowsRequestBody.ExplicitTable = table;
            queryRowsRequestBody.RowCount = ConstValues.QueryRowsRequestedRowNumber;
            queryRowsRequestBody.HasColumns = true;
            queryRowsRequestBody.Columns = largePropTagArray;
            queryRowsRequestBody.AuxiliaryBuffer = auxIn;
            queryRowsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            QueryRowsResponseBody queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>(0, queryRowsResponseBody.ErrorCode, "QueryRows request should be executed successfully, the returned value {0}.", queryRowsResponseBody.ErrorCode);
            #endregion

            #region Capture code
            AddressBookPropertyRow[] rowData = queryRowsResponseBody.RowData;
            for (int i = 0; i < rowData.Length; i++)
            {
                List<AddressBookPropertyValue> valueArray = new List<AddressBookPropertyValue>(rowData[i].ValueArray);

                for (int j = 0; j < valueArray.Count; j++)
                {
                    if (largePropTagArray.PropertyTags[j].PropertyType == 0x001F)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2001");

                        // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2001
                        this.Site.CaptureRequirementIfIsInstanceOfType(
                            valueArray[j].HasValue,
                            typeof(byte),
                            2001,
                            @"[In AddressBookPropertyValue Structure] HasValue (optional) (1 byte): An unsigned integer when the PropertyType ([MS-OXCDATA] section 2.11.1) is known to be PtypString.");
                    }

                    if (largePropTagArray.PropertyTags[j].PropertyType == 0x0102)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2003");

                        // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2003
                        this.Site.CaptureRequirementIfIsInstanceOfType(
                            valueArray[j].HasValue,
                            typeof(byte),
                            2003,
                            @"[In AddressBookPropertyValue Structure] HasValue (optional) (1 byte): An unsigned integer when the PropertyType ([MS-OXCDATA] section 2.11.1) is known to be PtypBinary.");
                    }
                }
            }
            #endregion

            #region Call ModLinkAtt with flags 0x00000000 to add the specified PidTagAddressBookMember value.
            uint flagsOfModLinkAtt = 0;
            PropertyTag propTagOfModLinkAtt = new PropertyTag
            {
                PropertyType = (ushort)PropertyTypeValues.PtypEmbeddedTable,
                PropertyId = (ushort)PropertyID.PidTagAddressBookMember
            };
            uint midOfModLinkAtt = 0;
            byte[] entryId = null;
            GetPropsRequestBody getPropsRequestBodyForAddressBookMember = null;
            uint lengthOfErrorCodeValue = sizeof(uint);
            string dlistName = Common.GetConfigurationPropertyValue("DistributionListName", this.Site);
            string memberName = Common.GetConfigurationPropertyValue("GeneralUserName", this.Site);

            for (int i = 0; i < queryRowsResponseBody.RowCount; i++)
            {
                string name = System.Text.Encoding.Unicode.GetString(queryRowsResponseBody.RowData[i].ValueArray[0].Value);
                if (name.ToLower().Contains(dlistName.ToLower()))
                {
                    PermanentEntryID entryID = AdapterHelper.ParsePermanentEntryIDFromBytes(queryRowsResponseBody.RowData[i].ValueArray[1].Value);

                    DNToMinIdRequestBody requestBodyOfDNToMId = new DNToMinIdRequestBody()
                    {
                        Reserved = 0,
                        HasNames = true,
                        Names = new StringArray_r
                        {
                            CValues = 1,
                            LppzA = new string[]
                            {
                                entryID.DistinguishedName
                            }
                        },
                        AuxiliaryBuffer = auxIn,
                        AuxiliaryBufferSize = (uint)auxIn.Length
                    };
                    DnToMinIdResponseBody responseBodyOfDNToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
                    midOfModLinkAtt = responseBodyOfDNToMinId.MinimalIds[0];

                    stat.CurrentRec = midOfModLinkAtt;
                    STAT? statForGetProps = stat;
                    LargePropertyTagArray propertyTagForGetProps = new LargePropertyTagArray()
                    {
                        PropertyTagCount = 2,
                        PropertyTags = new PropertyTag[] 
                        {
                            new PropertyTag
                            {
                                PropertyType = (ushort)PropertyTypeValues.PtypInteger32,
                                PropertyId = (ushort)PropertyID.PidTagDisplayType
                            },
                            new PropertyTag
                            {
                                PropertyType = (ushort)PropertyTypeValues.PtypMultipleString8,
                                PropertyId = (ushort)PropertyID.PidTagAddressBookMember
                            },
                        }
                    };

                    getPropsRequestBodyForAddressBookMember = this.BuildGetPropsRequestBody((uint)0, true, statForGetProps, true, propertyTagForGetProps);
                    GetPropsResponseBody getPropsResponseBody = this.Adapter.GetProps(getPropsRequestBodyForAddressBookMember);
                    this.Site.Assert.AreEqual<uint>(0, getPropsResponseBody.StatusCode, "The GetProps request should be executed successfully, the returned status code is {0}", getPropsResponseBody.StatusCode);
                    this.Site.Assert.AreEqual<uint>(lengthOfErrorCodeValue, (uint)getPropsResponseBody.PropertyValues.Value.PropertyValues[1].Value.Length, "The length of the property value should be equal the ErrorCodeValue: Not Found(0x8004010F)");
                    uint propertyValue = BitConverter.ToUInt32(getPropsResponseBody.PropertyValues.Value.PropertyValues[1].Value, 0);
                    this.Site.Assert.AreEqual<uint>((uint)ErrorCodeValue.NotFound, propertyValue, "The property value of AddressBookMember should be Not Found(0x8004010F), actual value is {0}", propertyValue);
                }
                else if (name.ToLower().Contains(memberName.ToLower()))
                {
                    entryId = queryRowsResponseBody.RowData[i].ValueArray[1].Value;
                }

                if (midOfModLinkAtt != 0 && entryId != null)
                {
                    break;
                }
            }

            ModLinkAttRequestBody modLinkAttRequestBody = new ModLinkAttRequestBody();
            modLinkAttRequestBody.Flags = flagsOfModLinkAtt;
            modLinkAttRequestBody.PropertyTag = propTagOfModLinkAtt;
            modLinkAttRequestBody.MinimalId = midOfModLinkAtt;
            modLinkAttRequestBody.HasEntryIds = true;
            modLinkAttRequestBody.EntryIdCount = 1;
            modLinkAttRequestBody.EntryIDs = new byte[][] { entryId };
            modLinkAttRequestBody.AuxiliaryBuffer = auxIn;
            modLinkAttRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;
            this.minimalIDForDeleteAddressBookMember = midOfModLinkAtt;
            this.entryIDBufferForDeleteAddressBookMember = entryId;
            this.isAddressBookMemberDeleted = false;

            ModLinkAttResponseBody modLinkAttResponseBodyOfAdd = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
            GetPropsResponseBody getPropsResponseBodyForAddAddressBookMember = this.Adapter.GetProps(getPropsRequestBodyForAddressBookMember);
            uint modifyDisplayType = BitConverter.ToUInt32(getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[0].Value, 0);
            int propertyValuelength = getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[1].Value.Length;

            if (getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[1].PropertyType == 0x101E)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2258");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2258
                this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[1].HasValue,
                typeof(byte),
                2258,
                @"[In AddressBookPropertyValue Structure] HasValue (optional) (1 byte): An unsigned integer when the PropertyType ([MS-OXCDATA] section 2.11.1) is known to be PtypMultipleString8 ([MS-OXCDATA] section 2.11.1).");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R724");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R724
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                modLinkAttResponseBodyOfAdd.ErrorCode,
                724,
                @"[In ModLinkAtt Request Type] The ModLinkAtt request type is used by the client to modify a specific property of a row in the address book.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1333, the DisplayTypeValues should be {0}, the length of the property value was changed from {1} to {2}.", DisplayTypeValues.DT_DISTLIST.ToString(), lengthOfErrorCodeValue, propertyValuelength);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1333
            bool isVerifiedR1333 = (modifyDisplayType == (uint)DisplayTypeValues.DT_DISTLIST) && (propertyValuelength != lengthOfErrorCodeValue);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1333,
                1333,
                @"[In ModLinkAtt Request Type Request Body] The PidTagAddressBookMember property ([MS-OXOABK] section 2.2.6.1) can be modified on an Address Book object that has a display type of DT_DISTLIST.");
            #endregion

            #region Call ModLinkAtt to delete the specified PidTagAddressBookMember value.
            modLinkAttRequestBody.Flags = 1;
            ModLinkAttResponseBody modLinkAttResponseBodyOfDelete = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
            Site.Assert.AreEqual<uint>(0, modLinkAttResponseBodyOfDelete.ErrorCode, "ModLinkAtt request should be executed successfully, the returned error code is {0}.", modLinkAttResponseBodyOfDelete.ErrorCode);
            this.isAddressBookMemberDeleted = true;
            #endregion

            #region Call ModLinkAtt with flags 0x00000002 to add the specified PidTagAddressBookMember value.
            this.Site.Assert.IsTrue(this.isAddressBookMemberDeleted, "The previous added address book member should be deleted, actual is {0}", this.isAddressBookMemberDeleted);
            modLinkAttRequestBody.Flags = 2;
            ModLinkAttResponseBody modLinkAttResponseBodyFlags2 = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
            Site.Assert.AreEqual<uint>(0, modLinkAttResponseBodyFlags2.ErrorCode, "ModLinkAtt request should be executed successfully, the returned error code is {0}.", modLinkAttResponseBodyFlags2.ErrorCode);
            this.isAddressBookMemberDeleted = false;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1325");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1325
            this.Site.CaptureRequirementIfAreEqual<uint>(
                modLinkAttResponseBodyFlags2.ErrorCode,
                modLinkAttResponseBodyOfAdd.ErrorCode,
                1325,
                @"[In ModLinkAtt Request Type Request Body] If this field [Flags] is set to different values other than the bit flag fDelete (0x00000001), the server will return the same result.");
            #endregion

            #region Call ModLinkAtt to delete the specified PidTagAddressBookMember value.
            modLinkAttRequestBody.Flags = 1;
            ModLinkAttResponseBody modLinkAttResponseBodyOfDeleteWithFlags2 = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
            Site.Assert.AreEqual<uint>(0, modLinkAttResponseBodyOfDeleteWithFlags2.ErrorCode, "The ModLinkAtt request to delete the address book member should be executed successfully, the actual returned error code is {0}.", modLinkAttResponseBodyOfDeleteWithFlags2.ErrorCode);
            this.isAddressBookMemberDeleted = true;
            #endregion

            #region Call ModLinkAtt request to add the specified PidTagAddressBookPublicDelegates value
            midOfModLinkAtt = 0;
            entryId = null;
            GetPropsRequestBody getPropsRequestBodyForPublicDelegates = null;
            memberName = Common.GetConfigurationPropertyValue("GeneralUserName", this.Site);

            PropertyTag pidTagAddressBookPublicDelegates = new PropertyTag();
            pidTagAddressBookPublicDelegates.PropertyType = (ushort)PropertyTypeValues.PtypMultipleString8;
            pidTagAddressBookPublicDelegates.PropertyId = (ushort)PropertyID.PidTagAddressBookPublicDelegates;

            for (int i = 0; i < queryRowsResponseBody.RowCount; i++)
            {
                string name = System.Text.Encoding.Unicode.GetString(queryRowsResponseBody.RowData[i].ValueArray[0].Value);
                if (name.ToLower().Contains(memberName.ToLower()))
                {
                    PermanentEntryID entryID = AdapterHelper.ParsePermanentEntryIDFromBytes(queryRowsResponseBody.RowData[i].ValueArray[1].Value);

                    DNToMinIdRequestBody requestBodyOfDNToMId = new DNToMinIdRequestBody()
                    {
                        Reserved = 0,
                        HasNames = true,
                        Names = new StringArray_r
                        {
                            CValues = 1,
                            LppzA = new string[]
                            {
                                entryID.DistinguishedName
                            }
                        },
                        AuxiliaryBuffer = auxIn,
                        AuxiliaryBufferSize = (uint)auxIn.Length
                    };

                    DnToMinIdResponseBody responseBodyOfDNToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
                    midOfModLinkAtt = responseBodyOfDNToMinId.MinimalIds[0];

                    stat.CurrentRec = midOfModLinkAtt;
                    STAT? statForGetProps = stat;
                    LargePropertyTagArray propertyTagForGetProps = new LargePropertyTagArray()
                    {
                        PropertyTagCount = 2,
                        PropertyTags = new PropertyTag[] 
                        {
                            new PropertyTag
                            {
                                PropertyType = (ushort)PropertyTypeValues.PtypInteger32,
                                PropertyId = (ushort)PropertyID.PidTagDisplayType
                            },
                            new PropertyTag
                            {
                                PropertyType = (ushort)PropertyTypeValues.PtypMultipleString8,
                                PropertyId = (ushort)PropertyID.PidTagAddressBookPublicDelegates
                            },
                        }
                    };

                    getPropsRequestBodyForPublicDelegates = this.BuildGetPropsRequestBody((uint)0, true, statForGetProps, true, propertyTagForGetProps);
                    GetPropsResponseBody getPropsResponseBody = this.Adapter.GetProps(getPropsRequestBodyForPublicDelegates);
                    this.Site.Assert.AreEqual<uint>(0, getPropsResponseBody.StatusCode, "The GetProps request should be executed successfully, the returned status code is {0}", getPropsResponseBody.StatusCode);
                    this.Site.Assert.AreEqual<uint>(lengthOfErrorCodeValue, (uint)getPropsResponseBody.PropertyValues.Value.PropertyValues[1].Value.Length, "The length of the property value should be equal the ErrorCodeValue: Not Found(0x8004010F)");
                    uint propertyValue = BitConverter.ToUInt32(getPropsResponseBody.PropertyValues.Value.PropertyValues[1].Value, 0);
                    this.Site.Assert.AreEqual<uint>((uint)ErrorCodeValue.NotFound, propertyValue, "The property value of AddressBookPublicDelegates should be Not Found(0x8004010F), actual value is {0}", propertyValue);
                }
                else if (name.ToLower().Contains(this.AdminUserName.ToLower()))
                {
                    entryId = queryRowsResponseBody.RowData[i].ValueArray[1].Value;
                }

                if (midOfModLinkAtt != 0 && entryId != null)
                {
                    break;
                }
            }

            ModLinkAttRequestBody modLinkAttRequestBodyForPublicDelegates = new ModLinkAttRequestBody();
            modLinkAttRequestBodyForPublicDelegates.Flags = 0;
            modLinkAttRequestBodyForPublicDelegates.PropertyTag = pidTagAddressBookPublicDelegates;
            modLinkAttRequestBodyForPublicDelegates.MinimalId = midOfModLinkAtt;
            modLinkAttRequestBodyForPublicDelegates.HasEntryIds = true;
            modLinkAttRequestBodyForPublicDelegates.EntryIdCount = 1;
            modLinkAttRequestBodyForPublicDelegates.EntryIDs = new byte[][] { entryId };
            modLinkAttRequestBodyForPublicDelegates.AuxiliaryBuffer = auxIn;
            modLinkAttRequestBodyForPublicDelegates.AuxiliaryBufferSize = (uint)auxIn.Length;
            this.minimalIDForDeleteAddressBookPublicDelegate = midOfModLinkAtt;
            this.entryIDBufferForDeleteAddressBookPublicDelegate = entryId;
            this.isAddressBookPublicDelegateDeleted = false;

            ModLinkAttResponseBody modLinkAttResponseBodyForAddPublicDelegates = this.Adapter.ModLinkAtt(modLinkAttRequestBodyForPublicDelegates);

            GetPropsResponseBody getPropsResponseBodyForAddPublicDelegates = this.Adapter.GetProps(getPropsRequestBodyForPublicDelegates);
            modifyDisplayType = BitConverter.ToUInt32(getPropsResponseBodyForAddPublicDelegates.PropertyValues.Value.PropertyValues[0].Value, 0);
            propertyValuelength = getPropsResponseBodyForAddPublicDelegates.PropertyValues.Value.PropertyValues[1].Value.Length;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1335, the DisplayTypeValues should be {0}, the length of the property value was changed from {1} to {2}.", DisplayTypeValues.DT_MAILUSER.ToString(), lengthOfErrorCodeValue, propertyValuelength);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1335
            bool isVerifiedR1335 = (modifyDisplayType == (uint)DisplayTypeValues.DT_MAILUSER) && (propertyValuelength != lengthOfErrorCodeValue);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1335,
                1335,
                @"[In ModLinkAtt Request Type Request Body] The PidTagAddressBookPublicDelegates property ([MS-OXOABK] section 2.2.5.5) can be modified on an Address Book object that has a display type of DT_MAILUSER.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R740, the minimal ID should be {0}, the length of the property value was changed from {1} to {2}.", midOfModLinkAtt, lengthOfErrorCodeValue, propertyValuelength);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R740
            bool isVerifiedR740 = (midOfModLinkAtt == (uint)getPropsRequestBodyForPublicDelegates.State.CurrentRec) && ((uint)propertyValuelength != lengthOfErrorCodeValue);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR740,
                740,
                @"[In ModLinkAtt Request Type Request Body] MinimalId (4 bytes): A MinimalEntryID structure ([MS-OXNSPI] section 2.3.8.1) that specifies the Minimal Entry ID of the address book row to be modified.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R746.");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R746, the length of the property value is different, means the property was modified.
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                lengthOfErrorCodeValue,
                (uint)propertyValuelength,
                746,
                @"[In ModLinkAtt Request Type Request Body] EntryIds: Each entry ID in the array specifies an address book row in which the specified property is to be modified.");
            #endregion

            #region Call ModLinkAtt request to delete the specified PidTagAddressBookPublicDelegates value
            modLinkAttRequestBodyForPublicDelegates.Flags = 1;
            ModLinkAttResponseBody modLinkAttResponseBodyForDeletePublicDelegates = this.Adapter.ModLinkAtt(modLinkAttRequestBodyForPublicDelegates);
            this.Site.Assert.AreEqual<uint>((uint)0, modLinkAttResponseBodyForDeletePublicDelegates.ErrorCode, "The ModLinkAtt request to delete address book public delegates should be executed successfully, the returned value is {0}", modLinkAttResponseBodyForDeletePublicDelegates.ErrorCode);
            this.isAddressBookPublicDelegateDeleted = true;
            #endregion

            #region Call Unbind request to destroy the session between the client and the server.
            this.Unbind();
            #endregion
        }
        
        /// <summary>
        /// This case is designed to verify Flag is 0xA.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC19_TestFlagWithPtypErrorCode()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call QueryRows with all optional fields exist and the type of all properties specified.
            uint tableCount = 0;
            uint[] table = null;
            uint rowCount = ConstValues.QueryRowsRequestedRowNumber;
            STAT stat = new STAT();
            stat.InitiateStat();

            LargePropertyTagArray columns = new LargePropertyTagArray()
            {
                PropertyTagCount = 1,
                PropertyTags = new PropertyTag[]
                {
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyType.PtypErrorCode,
                        PropertyId = (ushort)PropertyID.PidTagDisplayName
                    }
                }
            };

            QueryRowsRequestBody queryRowsRequestBody = this.BuildQueryRowsRequestBody(true, stat, tableCount, table, rowCount, true, columns);
            QueryRowsResponseBody queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.ErrorCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.ErrorCode);
            #endregion

            #region Capture code
            AddressBookPropertyRow[] rowData = queryRowsResponseBody.RowData;
            for (int i = 0; i < rowData.Length; i++)
            {
                List<AddressBookPropertyValue> valueArray = new List<AddressBookPropertyValue>(rowData[i].ValueArray);

                for (int j = 0; j < valueArray.Count; j++)
                {
                    AddressBookFlaggedPropertyValue propertyValue = (AddressBookFlaggedPropertyValue)valueArray[j];

                    if (propertyValue.Flag == 0xA)
                    {
                        bool isVerifyR2024 = propertyValue.Value != null && columns.PropertyTags[j].PropertyType == 10;
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2024.");

                        // Verify MS-OXCDATA requirement: MS-OXCMAPIHTTP_R2024
                        this.Site.CaptureRequirementIfIsTrue(
                            isVerifyR2024,
                            2024,
                            @"[In AddressBookFlaggedPropertyValue Structure] Flag value 0xA meaning The PropertyValue field will be an AddressBookPropertyValue structure containing a value of PtypErrorCode, as specified in [MS-OXCDATA] section 2.11.1. ");
                    }
                }
            }
            #endregion

            #region Call Unbind request to destroy the session between the client and the server.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify HasValue with PropertyType PtypString8.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC20_TestHasValueWithPropertyTypePtypString8()
        {
            this.CheckMapiHttpIsSupported();

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call QueryRows with all optional fields exist and the type of all properties specified.
            uint tableCount = 0;
            uint[] table = null;
            uint rowCount = ConstValues.QueryRowsRequestedRowNumber;
            STAT stat = new STAT();
            stat.InitiateStat();

            LargePropertyTagArray columns = new LargePropertyTagArray()
            {
                PropertyTagCount = 1,
                PropertyTags = new PropertyTag[]
                {
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypString8,
                        PropertyId = (ushort)PropertyID.PidTagDisplayName
                    }
                }
            };

            QueryRowsRequestBody queryRowsRequestBody = this.BuildQueryRowsRequestBody(true, stat, tableCount, table, rowCount, true, columns);
            QueryRowsResponseBody queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>((uint)0, queryRowsResponseBody.ErrorCode, "QueryRows operation should succeed and 0 is expected to be returned. The return value is {0}.", queryRowsResponseBody.ErrorCode);
            #endregion

            #region Capture code
            AddressBookPropertyRow[] rowData = queryRowsResponseBody.RowData;
            for (int i = 0; i < rowData.Length; i++)
            {
                List<AddressBookPropertyValue> valueArray = new List<AddressBookPropertyValue>(rowData[i].ValueArray);

                for (int j = 0; j < valueArray.Count; j++)
                {
                    if (columns.PropertyTags[j].PropertyType == 0x1E)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2002");

                        // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2002
                        this.Site.CaptureRequirementIfIsInstanceOfType(
                            valueArray[j].HasValue,
                            typeof(byte),
                            2002,
                            @"[In AddressBookPropertyValue Structure] HasValue (optional) (1 byte): An unsigned integer when the PropertyType ([MS-OXCDATA] section 2.11.1) is known to be PtypString8.");
                    }
                }
            }
            #endregion

            #region Call Unbind request to destroy the session between the client and the server.
            this.Unbind();
            #endregion
        }

        /// <summary>
        /// This case is designed to verify HasValue with PropertyType PtypMultipleString.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S02_TC21_TestHasValueWithPropertyTypePtypMultipleString()
        {
            this.CheckMapiHttpIsSupported();

            byte[] auxIn = new byte[] { };

            #region Call Bind request type to established a session context with the address book server.
            this.Bind();
            #endregion

            #region Call QueryRows request type to get a set of valid rows used to matched entry ID as the input paramter of ModLinkAtt method.
            STAT stat = new STAT();
            stat.InitiateStat();

            uint tableCount = 0;
            uint[] table = null;
            LargePropertyTagArray largePropTagArray = new LargePropertyTagArray()
            {
                PropertyTagCount = 4,
                PropertyTags = new PropertyTag[]
                {
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypString,
                        PropertyId = (ushort)PropertyID.PidTagDisplayName
                    }, 
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypBinary,
                        PropertyId = (ushort)PropertyID.PidTagEntryId
                    },
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypInteger32,
                        PropertyId = (ushort)PropertyID.PidTagDisplayType
                    },
                    new PropertyTag
                    {
                        PropertyType = (ushort)PropertyTypeValues.PtypMultipleString,
                        PropertyId = (ushort)PropertyID.PidTagAddressBookMember
                    }
                }
            };

            QueryRowsRequestBody queryRowsRequestBody = new QueryRowsRequestBody();
            queryRowsRequestBody.Flags = (uint)RetrievePropertyFlags.fSkipObjects;
            queryRowsRequestBody.HasState = true;
            queryRowsRequestBody.State = stat;
            queryRowsRequestBody.ExplicitTableCount = tableCount;
            queryRowsRequestBody.ExplicitTable = table;
            queryRowsRequestBody.RowCount = ConstValues.QueryRowsRequestedRowNumber;
            queryRowsRequestBody.HasColumns = true;
            queryRowsRequestBody.Columns = largePropTagArray;
            queryRowsRequestBody.AuxiliaryBuffer = auxIn;
            queryRowsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            QueryRowsResponseBody queryRowsResponseBody = this.Adapter.QueryRows(queryRowsRequestBody);
            Site.Assert.AreEqual<uint>(0, queryRowsResponseBody.ErrorCode, "QueryRows request should be executed successfully, the returned value {0}.", queryRowsResponseBody.ErrorCode);
            #endregion

            #region Call ModLinkAtt with flags 0x00000000 to add the specified PidTagAddressBookMember value.
            uint flagsOfModLinkAtt = 0;
            PropertyTag propTagOfModLinkAtt = new PropertyTag
            {
                PropertyType = (ushort)PropertyTypeValues.PtypEmbeddedTable,
                PropertyId = (ushort)PropertyID.PidTagAddressBookMember
            };
            uint midOfModLinkAtt = 0;
            byte[] entryId = null;
            GetPropsRequestBody getPropsRequestBodyForAddressBookMember = null;
            uint lengthOfErrorCodeValue = sizeof(uint);
            string dlistName = Common.GetConfigurationPropertyValue("DistributionListName", this.Site);
            string memberName = Common.GetConfigurationPropertyValue("GeneralUserName", this.Site);

            for (int i = 0; i < queryRowsResponseBody.RowCount; i++)
            {
                string name = System.Text.Encoding.Unicode.GetString(queryRowsResponseBody.RowData[i].ValueArray[0].Value);
                if (name.ToLower().Contains(dlistName.ToLower()))
                {
                    PermanentEntryID entryID = AdapterHelper.ParsePermanentEntryIDFromBytes(queryRowsResponseBody.RowData[i].ValueArray[1].Value);

                    DNToMinIdRequestBody requestBodyOfDNToMId = new DNToMinIdRequestBody()
                    {
                        Reserved = 0,
                        HasNames = true,
                        Names = new StringArray_r
                        {
                            CValues = 1,
                            LppzA = new string[]
                            {
                                entryID.DistinguishedName
                            }
                        },
                        AuxiliaryBuffer = auxIn,
                        AuxiliaryBufferSize = (uint)auxIn.Length
                    };
                    DnToMinIdResponseBody responseBodyOfDNToMinId = this.Adapter.DnToMinId(requestBodyOfDNToMId);
                    midOfModLinkAtt = responseBodyOfDNToMinId.MinimalIds[0];

                    stat.CurrentRec = midOfModLinkAtt;
                    STAT? statForGetProps = stat;
                    LargePropertyTagArray propertyTagForGetProps = new LargePropertyTagArray()
                    {
                        PropertyTagCount = 2,
                        PropertyTags = new PropertyTag[] 
                        {
                            new PropertyTag
                            {
                                PropertyType = (ushort)PropertyTypeValues.PtypInteger32,
                                PropertyId = (ushort)PropertyID.PidTagDisplayType
                            },
                            new PropertyTag
                            {
                                PropertyType = (ushort)PropertyTypeValues.PtypMultipleString,
                                PropertyId = (ushort)PropertyID.PidTagAddressBookMember
                            },
                        }
                    };

                    getPropsRequestBodyForAddressBookMember = this.BuildGetPropsRequestBody((uint)0, true, statForGetProps, true, propertyTagForGetProps);
                    GetPropsResponseBody getPropsResponseBody = this.Adapter.GetProps(getPropsRequestBodyForAddressBookMember);
                    this.Site.Assert.AreEqual<uint>(0, getPropsResponseBody.StatusCode, "The GetProps request should be executed successfully, the returned status code is {0}", getPropsResponseBody.StatusCode);
                    this.Site.Assert.AreEqual<uint>(lengthOfErrorCodeValue, (uint)getPropsResponseBody.PropertyValues.Value.PropertyValues[1].Value.Length, "The length of the property value should be equal the ErrorCodeValue: Not Found(0x8004010F)");
                    uint propertyValue = BitConverter.ToUInt32(getPropsResponseBody.PropertyValues.Value.PropertyValues[1].Value, 0);
                    this.Site.Assert.AreEqual<uint>((uint)ErrorCodeValue.NotFound, propertyValue, "The property value of AddressBookMember should be Not Found(0x8004010F), actual value is {0}", propertyValue);
                }
                else if (name.ToLower().Contains(memberName.ToLower()))
                {
                    entryId = queryRowsResponseBody.RowData[i].ValueArray[1].Value;
                }

                if (midOfModLinkAtt != 0 && entryId != null)
                {
                    break;
                }
            }

            ModLinkAttRequestBody modLinkAttRequestBody = new ModLinkAttRequestBody();
            modLinkAttRequestBody.Flags = flagsOfModLinkAtt;
            modLinkAttRequestBody.PropertyTag = propTagOfModLinkAtt;
            modLinkAttRequestBody.MinimalId = midOfModLinkAtt;
            modLinkAttRequestBody.HasEntryIds = true;
            modLinkAttRequestBody.EntryIdCount = 1;
            modLinkAttRequestBody.EntryIDs = new byte[][] { entryId };
            modLinkAttRequestBody.AuxiliaryBuffer = auxIn;
            modLinkAttRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;
            this.minimalIDForDeleteAddressBookMember = midOfModLinkAtt;
            this.entryIDBufferForDeleteAddressBookMember = entryId;
            this.isAddressBookMemberDeleted = false;

            ModLinkAttResponseBody modLinkAttResponseBodyOfAdd = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
            GetPropsResponseBody getPropsResponseBodyForAddAddressBookMember = this.Adapter.GetProps(getPropsRequestBodyForAddressBookMember);
            uint modifyDisplayType = BitConverter.ToUInt32(getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[0].Value, 0);
            int propertyValuelength = getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[1].Value.Length;

            if (getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[1].PropertyType == 0x101F)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2257");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2257
                this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBodyForAddAddressBookMember.PropertyValues.Value.PropertyValues[1].HasValue,
                typeof(byte),
                2257,
                @"[In AddressBookPropertyValue Structure] HasValue (optional) (1 byte): An unsigned integer when the PropertyType ([MS-OXCDATA] section 2.11.1) is known to be PtypMultipleString ([MS-OXCDATA] section 2.11.1).");
            }
            #endregion

            #region Call ModLinkAtt to delete the specified PidTagAddressBookMember value.
            modLinkAttRequestBody.Flags = 1;
            ModLinkAttResponseBody modLinkAttResponseBodyOfDelete = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
            Site.Assert.AreEqual<uint>(0, modLinkAttResponseBodyOfDelete.ErrorCode, "ModLinkAtt request should be executed successfully, the returned error code is {0}.", modLinkAttResponseBodyOfDelete.ErrorCode);
            this.isAddressBookMemberDeleted = true;
            #endregion

            #region Call Unbind request to destroy the session between the client and the server.
            this.Unbind();
            #endregion
        }
        #endregion Test Cases

        /// <summary>
        /// Clean up the test suite.
        /// </summary>
        protected override void TestCleanup()
        {
            // Flags of ModLinkAtt for delete the address book member or address book public delegates.
            uint flagsOfModLinkAtt = 1;
            PropertyTag propTagOfAddressBookMember = new PropertyTag
            {
                PropertyType = (ushort)PropertyTypeValues.PtypEmbeddedTable,
                PropertyId = (ushort)PropertyID.PidTagAddressBookMember
            };

            PropertyTag propTagOfAddressBookPublicDelegate = new PropertyTag
            {
                PropertyType = (ushort)PropertyTypeValues.PtypEmbeddedTable,
                PropertyId = (ushort)PropertyID.PidTagAddressBookPublicDelegates
            };
            byte[] auxIn = new byte[] { };

            ModLinkAttRequestBody modLinkAttRequestBody = new ModLinkAttRequestBody();
            modLinkAttRequestBody.Flags = flagsOfModLinkAtt;
            modLinkAttRequestBody.HasEntryIds = true;
            modLinkAttRequestBody.EntryIdCount = 1;
            modLinkAttRequestBody.AuxiliaryBuffer = auxIn;
            modLinkAttRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            if (!this.isAddressBookMemberDeleted)
            {
                modLinkAttRequestBody.PropertyTag = propTagOfAddressBookMember;
                modLinkAttRequestBody.MinimalId = this.minimalIDForDeleteAddressBookMember;
                modLinkAttRequestBody.EntryIDs = new byte[][] { this.entryIDBufferForDeleteAddressBookMember };

                ModLinkAttResponseBody modLinkAttResponseBodyOfDelete = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
                Site.Assert.AreEqual<uint>(0, modLinkAttResponseBodyOfDelete.ErrorCode, "ModLinkAtt request to delete the address book member should be executed successfully, the returned error code is {0}.", modLinkAttResponseBodyOfDelete.ErrorCode);
            }

            if (!this.isAddressBookPublicDelegateDeleted)
            {
                modLinkAttRequestBody.PropertyTag = propTagOfAddressBookPublicDelegate;
                modLinkAttRequestBody.MinimalId = this.minimalIDForDeleteAddressBookPublicDelegate;
                modLinkAttRequestBody.EntryIDs = new byte[][] { this.entryIDBufferForDeleteAddressBookPublicDelegate };

                ModLinkAttResponseBody modLinkAttResponseBodyForDeletePublicDelegates = this.Adapter.ModLinkAtt(modLinkAttRequestBody);
                this.Site.Assert.AreEqual<uint>((uint)0, modLinkAttResponseBodyForDeletePublicDelegates.ErrorCode, "The ModLinkAtt request to delete address book public delegates should be executed successfully, the returned value is {0}", modLinkAttResponseBodyForDeletePublicDelegates.ErrorCode);
            }

            base.TestCleanup();
        }

        #region Private methods

        /// <summary>
        /// Build the GetMatches request body.
        /// </summary>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="state">A STAT structure that specifies the state of a specific address book container.</param>
        /// <param name="hasMinimalIds">A Boolean value that specifies whether the MinimalIdCount and MinimalIds fields are present.</param>
        /// <param name="minimalIdCount">An unsigned integer that specifies the number of structures present in the MinimalIds field.</param>
        /// <param name="minimalIds">An array of MinimalEntryID structures that constitute an Explicit Table.</param>
        /// <param name="hasFilter">A Boolean value that specifies whether the Filter field is present.</param>
        /// <param name="filter">A restriction that is to be applied to the rows in the address book container.</param>
        /// <param name="hasPropertyName">A Boolean value that specifies whether the PropertyNameGuid and PropertyNameId fields are present.</param>
        /// <param name="propertyNameGuid">The GUID of the property to be opened.</param>
        /// <param name="propertyNameId">A 4-byte value that specifies the ID of the property to be opened.</param>
        /// <param name="rowCount">An unsigned integer that specifies the number of rows the client is requesting.</param>
        /// <param name="hasColumns">A Boolean value that specifies whether the Columns field is present.</param>
        /// <param name="columns">A LargePropertyTagArray structure that specifies the columns that the client is requesting.</param>
        /// <returns>The GetMatches request body.</returns>
        private GetMatchesRequestBody BuildGetMatchRequestBody(bool hasState, STAT state, bool hasMinimalIds, uint minimalIdCount, uint[] minimalIds, bool hasFilter, byte[] filter, bool hasPropertyName, Guid propertyNameGuid, uint propertyNameId, uint rowCount, bool hasColumns, LargePropertyTagArray columns)
        {
            GetMatchesRequestBody getMatchRequestBody = new GetMatchesRequestBody();

            getMatchRequestBody.Reserved = 0x0;
            getMatchRequestBody.HasState = hasState;
            if (hasState)
            {
                getMatchRequestBody.State = state;
            }

            getMatchRequestBody.HasMinimalIds = hasMinimalIds;
            if (hasMinimalIds)
            {
                getMatchRequestBody.MinimalIdCount = minimalIdCount;
                getMatchRequestBody.MinimalIds = minimalIds;
            }

            getMatchRequestBody.InterfaceOptionFlags = 0x0;
            getMatchRequestBody.HasFilter = hasFilter;
            if (hasFilter)
            {
                getMatchRequestBody.Filter = filter;
            }

            getMatchRequestBody.HasPropertyName = hasPropertyName;
            if (hasPropertyName)
            {
                getMatchRequestBody.PropertyNameGuid = propertyNameGuid;
                getMatchRequestBody.PropertyNameId = propertyNameId;
            }

            getMatchRequestBody.RowCount = rowCount;
            getMatchRequestBody.HasColumns = hasColumns;
            if (hasColumns)
            {
                getMatchRequestBody.Columns = columns;
            }

            byte[] auxIn = new byte[] { };
            getMatchRequestBody.AuxiliaryBuffer = auxIn;
            getMatchRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return getMatchRequestBody;
        }

        /// <summary>
        /// Build the DNToMinId request body.
        /// </summary>
        /// <param name="hasNames">A Boolean value that specifies whether the Names field is present.</param>
        /// <param name="names">A StringsArray_r structure that contains the list of distinguished names (DNs) (1) to be mapped to Minimal Entry IDs.</param>
        /// <returns>The DNToMinId request body.</returns>
        private DNToMinIdRequestBody BuildDNToMinIdRequestBody(bool hasNames, StringArray_r names)
        {
            DNToMinIdRequestBody requestBodyOfDNToMId = new DNToMinIdRequestBody();

            requestBodyOfDNToMId.Reserved = 0x0;
            requestBodyOfDNToMId.HasNames = hasNames;
            if (hasNames)
            {
                requestBodyOfDNToMId.Names = names;
            }

            byte[] auxIn = new byte[] { };
            requestBodyOfDNToMId.AuxiliaryBuffer = auxIn;
            requestBodyOfDNToMId.AuxiliaryBufferSize = (uint)auxIn.Length;

            return requestBodyOfDNToMId;
        }

        /// <summary>
        /// Build the QueryRows request body.
        /// </summary>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="state">A STAT structure that specifies the state of a specific address book container.</param>
        /// <param name="explicitTableCount">An unsigned integer that specifies the number of structures present in the ExplicitTable field.</param>
        /// <param name="explicitTable">An array of MinimalEntryID structures that constitute the Explicit Table.</param>
        /// <param name="rowCount">An unsigned integer that specifies the number of rows the client is requesting.</param>
        /// <param name="hasColumns">A Boolean value that specifies whether the Columns field is present.</param>
        /// <param name="columns">A LargePropTagArray structure that specifies the properties that the client requires for each row returned.</param>
        /// <returns>Returns the QueryRows request body.</returns>
        private QueryRowsRequestBody BuildQueryRowsRequestBody(bool hasState, STAT state, uint explicitTableCount, uint[] explicitTable, uint rowCount, bool hasColumns, LargePropertyTagArray columns)
        {
            QueryRowsRequestBody queryRowsRequestBody = new QueryRowsRequestBody();

            queryRowsRequestBody.Flags = (uint)RetrievePropertyFlags.fSkipObjects;
            queryRowsRequestBody.HasState = hasState;
            if (hasState)
            {
                queryRowsRequestBody.State = state;
            }

            queryRowsRequestBody.ExplicitTableCount = explicitTableCount;
            queryRowsRequestBody.ExplicitTable = explicitTable;
            queryRowsRequestBody.RowCount = rowCount;

            queryRowsRequestBody.HasColumns = hasColumns;
            if (hasColumns)
            {
                queryRowsRequestBody.Columns = columns;
            }

            byte[] auxIn = new byte[] { };
            queryRowsRequestBody.AuxiliaryBuffer = auxIn;
            queryRowsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return queryRowsRequestBody;
        }

        /// <summary>
        /// Build the ResortRestriction request body.
        /// </summary>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="state">A STAT structure that specifies the state of a specific address book container.</param>
        /// <param name="hasMinimalIds">A Boolean value that specifies whether the MinimalIdCount and MinimalIds fields are present.</param>
        /// <param name="minimalIdCount">An unsigned integer that specifies the number of structures in the MinimalIds field.</param>
        /// <param name="minimalIds">An array of MinimalEntryID structures that compose a restricted address book container.</param>
        /// <returns>The ResortRestriction request body.</returns>
        private ResortRestrictionRequestBody BuildResortRestriction(bool hasState, STAT state, bool hasMinimalIds, uint minimalIdCount, uint[] minimalIds)
        {
            ResortRestrictionRequestBody resortRestrictionRequestBody = new ResortRestrictionRequestBody();

            byte[] auxIn = new byte[] { };

            resortRestrictionRequestBody.Reserved = 0x0;
            resortRestrictionRequestBody.HasState = hasState;
            if (hasState)
            {
                resortRestrictionRequestBody.State = state;
            }

            resortRestrictionRequestBody.HasMinimalIds = hasMinimalIds;
            if (hasMinimalIds)
            {
                resortRestrictionRequestBody.MinimalIdCount = minimalIdCount;
                resortRestrictionRequestBody.MinimalIds = minimalIds;
            }

            resortRestrictionRequestBody.AuxiliaryBuffer = auxIn;
            resortRestrictionRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return resortRestrictionRequestBody;
        }

        /// <summary>
        /// Initialize bind request body.
        /// </summary>
        /// <param name="stat">The state of the request body.</param>
        /// <param name="flags">The flags of the request body.</param>
        /// <returns>The bind request body.</returns>
        private BindRequestBody BuildBindRequestBody(STAT stat, uint flags)
        {
            BindRequestBody bindRequestBody = new BindRequestBody();
            bindRequestBody.State = stat;
            bindRequestBody.Flags = flags;
            bindRequestBody.HasState = true;
            byte[] auxIn = new byte[] { };
            bindRequestBody.AuxiliaryBuffer = auxIn;
            bindRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return bindRequestBody;
        }

        /// <summary>
        /// Initialize the unbind request body.
        /// </summary>
        /// <returns>The Unbind request body</returns>
        private UnbindRequestBody BuildUnbindRequestBody()
        {
            UnbindRequestBody unbindRequest = new UnbindRequestBody();
            unbindRequest.Reserved = 0x00000000;
            byte[] auxIn = new byte[] { };
            unbindRequest.AuxiliaryBuffer = auxIn;
            unbindRequest.AuxiliaryBufferSize = (uint)auxIn.Length;

            return unbindRequest;
        }

        /// <summary>
        /// Initialize ModProps request body.
        /// </summary>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="stat">A STAT structure that specifies the state of a specific address book container.</param>
        /// <param name="hasPropertyValues">A Boolean value that specifies whether the PropertyValues field is present.</param>
        /// <param name="propertyTag">A property tag both identifies a property and gives the data type its value.</param>
        /// <param name="hasPropertyTagsToRemove">A Boolean value that specifies whether the PropertyTagsToRemove field is present.</param>
        /// <param name="propertyTagsToRemove">A LargePropTagArray structure that specifies the properties that the client is requesting to be removed. </param>
        /// <returns>Returns the ModProps request body.</returns>
        private ModPropsRequestBody BuildModPropsRequestBody(bool hasState, STAT stat, bool hasPropertyValues, PropertyTag propertyTag, bool hasPropertyTagsToRemove, LargePropertyTagArray propertyTagsToRemove)
        {
            ModPropsRequestBody modPropsRequestBody = new ModPropsRequestBody();
            modPropsRequestBody.Reserved = 0x0;
            byte[] auxIn = new byte[] { };
            modPropsRequestBody.AuxiliaryBuffer = auxIn;
            modPropsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            modPropsRequestBody.HasState = hasState;
            if (hasState)
            {
                modPropsRequestBody.State = stat;
            }

            modPropsRequestBody.HasPropertyValues = hasPropertyValues;
            if (hasPropertyValues)
            {
                AddressBookPropertyValueList addressBookProperties = new AddressBookPropertyValueList();

                addressBookProperties.PropertyValueCount = 1;

                AddressBookTaggedPropertyValue[] taggedPropertyValues = new AddressBookTaggedPropertyValue[1];
                AddressBookTaggedPropertyValue taggedPropertyValue = new AddressBookTaggedPropertyValue();

                taggedPropertyValue.PropertyType = propertyTag.PropertyType;
                taggedPropertyValue.PropertyId = propertyTag.PropertyId;
                taggedPropertyValue.Value = new byte[] { 0x00, 0x00 };
                taggedPropertyValues[0] = taggedPropertyValue;
                addressBookProperties.PropertyValues = taggedPropertyValues;

                modPropsRequestBody.PropertyVaules = addressBookProperties;
            }

            modPropsRequestBody.HasPropertyTagsToRemove = hasPropertyTagsToRemove;
            if (hasPropertyTagsToRemove)
            {
                modPropsRequestBody.PropertyTagsToRemove = propertyTagsToRemove;
            }

            return modPropsRequestBody;
        }

        /// <summary>
        /// Build the GetProps request body.
        /// </summary>
        /// <param name="flags">A set of bit flags that specify options to the server.</param>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="stat">A STAT structure that specifies the state of a specific address book container.</param>
        /// <param name="hasPropertyTags">A Boolean value that specifies whether the PropertyTags field is present.</param>
        /// <param name="propetyTags">A LargePropertyTagArray structure that contains the property tags of the properties that the client is requesting.</param>
        /// <returns>The GetProps request body.</returns>
        private GetPropsRequestBody BuildGetPropsRequestBody(uint flags, bool hasState, STAT? stat, bool hasPropertyTags, LargePropertyTagArray propetyTags)
        {
            GetPropsRequestBody getPropertyRequestBody = new GetPropsRequestBody();
            getPropertyRequestBody.Flags = flags;
            byte[] auxIn = new byte[] { };
            getPropertyRequestBody.AuxiliaryBuffer = auxIn;
            getPropertyRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            getPropertyRequestBody.HasState = hasState;
            if (hasState)
            {
                getPropertyRequestBody.State = (STAT)stat;
            }

            getPropertyRequestBody.HasPropertyTags = hasPropertyTags;
            if (hasPropertyTags)
            {
                getPropertyRequestBody.PropertyTags = propetyTags;
            }

            return getPropertyRequestBody;
        }

        /// <summary>
        /// Build UpdateStat request body.
        /// </summary>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="stat">A STAT structure that specifies the state of a specific address book container.</param>
        /// <param name="deltaRequested">A Boolean value that specifies whether the client is requesting a value to be returned in the Delta field of the response.</param>
        /// <returns>The UpdateStat request body.</returns>
        private UpdateStatRequestBody BuildUpdateStatRequestBody(bool hasState, STAT stat, bool deltaRequested)
        {
            UpdateStatRequestBody updateStatRequestBody = new UpdateStatRequestBody();
            updateStatRequestBody.Reserved = 0x0;
            updateStatRequestBody.HasState = hasState;
            if (hasState)
            {
                updateStatRequestBody.State = stat;
            }

            updateStatRequestBody.DeltaRequested = deltaRequested;

            byte[] auxIn = new byte[] { };
            updateStatRequestBody.AuxiliaryBuffer = auxIn;
            updateStatRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;
            return updateStatRequestBody;
        }

        /// <summary>
        /// Build the QueryColumns request body.
        /// </summary>
        /// <param name="flag">A set of bit flags that specify options to the server.</param>
        /// <returns>The QueryColumns request body.</returns>
        private QueryColumnsRequestBody BuildQueryColumnsRequestBody(uint flag)
        {
            QueryColumnsRequestBody queryColumnsRequestBody = new QueryColumnsRequestBody();
            queryColumnsRequestBody.MapiFlags = flag;
            queryColumnsRequestBody.Reserved = 0x0;
            byte[] auxIn = new byte[] { };
            queryColumnsRequestBody.AuxiliaryBuffer = auxIn;
            queryColumnsRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return queryColumnsRequestBody;
        }

        /// <summary>
        /// Build the GetPropList request body.
        /// </summary>
        /// <param name="flag">A set of bit flags that specify options to the server.</param>
        /// <param name="mid">A unsigned integer that specifies the object for which to return properties.</param>
        /// <returns>The GetPropList request body.</returns>
        private GetPropListRequestBody BuildGetPropListRequestBody(uint flag, uint mid)
        {
            GetPropListRequestBody getPropListRequestBody = new GetPropListRequestBody();
            getPropListRequestBody.Flags = flag;
            getPropListRequestBody.CodePage = (uint)RequiredCodePages.CP_TELETEX;
            getPropListRequestBody.MinmalId = mid;
            byte[] auxIn = new byte[] { };
            getPropListRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;
            getPropListRequestBody.AuxiliaryBuffer = auxIn;

            return getPropListRequestBody;
        }

        /// <summary>
        /// Initiate a session between the client and the server.
        /// </summary>
        private void Bind()
        {
            STAT stat = new STAT();
            stat.InitiateStat();

            uint flags = 0x0;
            BindRequestBody bindRequestBody = new BindRequestBody();
            bindRequestBody.State = stat;
            bindRequestBody.Flags = flags;
            bindRequestBody.HasState = true;
            byte[] auxIn = new byte[] { };
            bindRequestBody.AuxiliaryBuffer = auxIn;
            bindRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            int responseCode;
            BindResponseBody bindResponseBody = this.Adapter.Bind(bindRequestBody, out responseCode);
            Site.Assert.AreEqual<uint>((uint)0, bindResponseBody.ErrorCode, "Bind operation should succeed and 0 is expected to be returned. The return value is {0}.", bindResponseBody.ErrorCode);
        }

        /// <summary>
        /// Destroy the session between the client and the server.
        /// </summary>
        private void Unbind()
        {
            UnbindRequestBody unbindRequest = new UnbindRequestBody();
            unbindRequest.Reserved = 0x00000000;
            byte[] auxIn = new byte[] { };
            unbindRequest.AuxiliaryBuffer = auxIn;
            unbindRequest.AuxiliaryBufferSize = (uint)auxIn.Length;
            UnbindResponseBody unbindResponseBody = this.Adapter.Unbind(unbindRequest);
            Site.Assert.AreEqual<uint>((uint)1, unbindResponseBody.ErrorCode, "Unbind method should succeed and the expected value is 1. The return value is {0}.", unbindResponseBody.ErrorCode);
        }

        /// <summary>
        /// Build the GetSpecialTable request body.
        /// </summary>
        /// <param name="flags">A set of bit flags that specify options to the server.</param>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="state">A STAT structure that specifies the state of a specific address book container.</param>
        /// <param name="hasVersion">A Boolean value that specifies whether the Version field is present.</param>
        /// <param name="version">A unsigned integer that specifies the version number of the address book hierarchy table that the client has.</param>
        /// <returns>The GetSpecialTable request body.</returns>
        private GetSpecialTableRequestBody BuildGetSpecialTableRequestBody(uint flags, bool hasState, STAT state, bool hasVersion, uint version)
        {
            byte[] auxIn = new byte[] { };
            GetSpecialTableRequestBody getSpecialTableRequestBody = new GetSpecialTableRequestBody()
            {
                Flags = flags,
                HasState = hasState,
                HasVersion = hasVersion,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            if (hasState)
            {
                getSpecialTableRequestBody.State = state;
            }

            if (hasVersion)
            {
                getSpecialTableRequestBody.Version = version;
            }

            return getSpecialTableRequestBody;
        }

        /// <summary>
        /// Build the GetTemplateInfo request body.
        /// </summary>
        /// <param name="flags">A set of bit flags that specify options to the server.</param>
        /// <param name="displayType">An unsigned integer that specifies the display type of the template for which information is requested.</param>
        /// <param name="hasTemplateDn">A Boolean value indicating whether the TemplateDN field is present.</param>
        /// <param name="templateDn">A string that specifies the distinguished name of the template requested.</param>
        /// <param name="codePage">An unsigned integer that specifies the code page of template for which information is requested.</param>
        /// <param name="locateId">An unsigned integer that specifies the language code identifier(LCID) of the template for which information is requested.</param>
        /// <returns>The GetTemplateInfo request body.</returns>
        private GetTemplateInfoRequestBody BuildGetTemplateInfoRequestBody(uint flags, uint displayType, bool hasTemplateDn, string templateDn, uint codePage, uint locateId)
        {
            byte[] auxIn = new byte[] { };
            GetTemplateInfoRequestBody getTemplateInfoRequestBody = new GetTemplateInfoRequestBody()
            {
                Flags = flags,
                DisplayType = displayType,
                HasTemplateDn = hasTemplateDn,
                TemplateDn = templateDn,
                CodePage = codePage,
                LocaleId = locateId,
                AuxiliaryBuffer = auxIn,
                AuxiliaryBufferSize = (uint)auxIn.Length
            };

            return getTemplateInfoRequestBody;
        }

        /// <summary>
        /// Build the GetMatches request body.
        /// </summary>
        /// <param name="hasState">A Boolean value that specifies whether the State field is present.</param>
        /// <param name="state">A STAT structure that specifies the state of a specific address book container. </param>
        /// <param name="hasTarget">A Boolean value that specifies whether the Target field is present.</param>
        /// <param name="target">A PropertyValue_r structure that specifies the property value being sought.</param>
        /// <param name="hasExplicitTable">A Boolean value that specifies whether the ExplicitTableCount and ExplicitTable fields are present.</param>
        /// <param name="explicitableCount">An unsigned integer that specifies the number of structures present in the ExplicitTable field.</param>
        /// <param name="explicitTable">An array of unsigned integer that constitute an Explicit Table.</param>
        /// <param name="hasColumns">A Boolean value that specifies whether the Columns field is present.</param>
        /// <param name="columns">A LargePropTagArray structure that specifies the columns that the client is requesting.</param>
        /// <returns>Return an instance of SeekEntriesRequestBody class.</returns>
        private SeekEntriesRequestBody BuildSeekEntriesRequestBody(bool hasState, STAT state, bool hasTarget, PropertyValue_r target, bool hasExplicitTable, uint explicitableCount, uint[] explicitTable, bool hasColumns, LargePropertyTagArray columns)
        {
            SeekEntriesRequestBody requestBody = new SeekEntriesRequestBody();

            // Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
            requestBody.Reserved = 0x00000000;
            requestBody.HasState = hasState;
            if (hasState)
            {
                requestBody.State = state;
            }

            requestBody.HasTarget = hasTarget;
            if (hasTarget)
            {
                requestBody.Target = target;
            }

            requestBody.HasExplicitTable = hasExplicitTable;
            if (hasExplicitTable)
            {
                requestBody.ExplicitableCount = explicitableCount;
                requestBody.ExplicitTable = explicitTable;
            }

            requestBody.HasColumns = hasColumns;

            if (hasColumns)
            {
                requestBody.Columns = columns;
            }

            requestBody.AuxiliaryBufferSize = 0;
            requestBody.AuxiliaryBuffer = new byte[] { };

            return requestBody;
        }
        #endregion
    }
}