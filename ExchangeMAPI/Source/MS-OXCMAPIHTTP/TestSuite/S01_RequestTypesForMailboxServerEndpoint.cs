namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test HTTP header, common response format and the request types for mailbox server endpoint.
    /// </summary>
    [TestClass]
    public class S01_RequestTypesForMailboxServerEndpoint : TestSuiteBase
    {
        #region Variable
        /// <summary>
        /// A Boolean indicates whether Receive a new mail. 
        /// </summary>
        private bool isReceiveNewMail = false;

        /// <summary>
        /// Declares a delegate for a method that returns an unsigned integer.
        /// </summary>
        /// <returns>Returns a delegate object.</returns>
        private delegate uint MethodCaller();
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
        /// This case is designed to verify the requirements related to create a Session Context successfully.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC01_ConnectToMailboxServerSucceeded()
        {
            this.CheckMapiHttpIsSupported();

            WebHeaderCollection headers = new WebHeaderCollection();
            MailboxResponseBodyBase responseBody;

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
            uint connectResult = AdapterHelper.GetFinalResponseCode(headers["X-ResponseCode"]);
            string firstGUIDPortionOfRequestId = headers["X-RequestId"].Substring(0, headers["X-RequestId"].IndexOf(":"));
            int firstCounterOfRequestId = int.Parse(headers["X-RequestId"].Substring(headers["X-RequestId"].IndexOf(":") + 1));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R215: the DnPrefix field in Connect response is {0}.", connectResponse.DNPrefix.Replace("\0", string.Empty));
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R215
            this.Site.CaptureRequirementIfIsTrue(
                connectResponse.DNPrefix.EndsWith("\0"),
                215,
                @"[In Connect Request Type Success Response Body] DnPrefix (variable): A null-terminated ASCII string that specifies the DN (1) prefix to be used for building message recipients (1).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R216: the DisplayName field in Connect success response should is {0} and actual is  {1}", this.AdminUserName.Replace("\0", string.Empty), connectResponse.DisplayName.Replace("\0", string.Empty));

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R216
            this.Site.CaptureRequirementIfIsTrue(
                string.Compare(this.AdminUserName, connectResponse.DisplayName, true) == 0,
                216,
                @"[In Connect Request Type Success Response Body] DisplayName (variable): A null-terminated Unicode string that specifies the display name of the user who is specified in the UserDn field of the Connect request type request body.");

            Site.Assert.IsNotNull(
                headers["Set-Cookie"],
                "The Set-Cookie header should be returned from server if establishing a Session Context with server successfully");

            Site.Assert.IsNotNull(
                connectResponse,
                "The Connect request type success response body should be return from server if establish a Session Context with server successful.");
            
            // According to steps above, the Set-Cookie and Connect request type successful response body are included in the response.
            // So the requirement R1219 is verified.
            this.Site.CaptureRequirement(
                1219,
                @"[In Responding to a Connect or Bind Request Type Request] The server creates a new Session Context and associates it with a session context cookie.");

            // According to steps above, the server returned the Connect success response body are included in the response.
            // So the requirement R2224 is verified.
            this.Site.CaptureRequirement(
                2224,
                @"[In Responding to a Connect or Bind Request Type Request] If successful, the server's response includes the Connect request type success response body, as specified in section 2.2.4.1.2.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2226");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2226
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                connectResult,
                2226,
                @"[In Responding to All Request Type Requests] After the request has been authorized, authenticated, parsed, and validated, the server MUST return a X-ResponseCode with a value of 0 (zero) to indicate that the request has been accepted.");
            #endregion

            #region Send an Execute request type that is used to send Logon ROP to server.
            WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            string expectClientInfo = executeHeaders["X-ClientInfo"];

            ExecuteRequestBody requestBody = this.InitializeExecuteRequestBody(this.GetRopLogonRequest());
            List<string> metaTags = new List<string>();
            ExecuteSuccessResponseBody executeSuccessResponse = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;
            uint executeResult = AdapterHelper.GetFinalResponseCode(executeHeaders["X-ResponseCode"]);
            string actualClientInfo = executeHeaders["X-ClientInfo"];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R181: the X-ResponseCode header in Connect response is {0} and the X-ResponseCode header in Execute response is {1}.", connectResult, executeResult);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R181
            bool isVerifiedR181 = connectResult == 0 && executeResult == 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR181,
                181,
                @"[In Connect Request Type] The Connect request type is used to establish a Session Context with the server, as specified in section 3.1.5.1 and section 3.1.5.7.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R156");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R156
            this.Site.CaptureRequirementIfAreEqual<string>(
                expectClientInfo,
                actualClientInfo,
                156,
                @"[In X-ClientInfo Header Field] The server MUST return this header [X-ClientInfo] with the same information in the response back to the client.");

            string secondGUIDPortionOfRequestId = executeHeaders["X-RequestId"].Substring(0, executeHeaders["X-RequestId"].IndexOf(":"));
            int secondCounterOfRequestId = int.Parse(executeHeaders["X-RequestId"].Substring(executeHeaders["X-RequestId"].IndexOf(":") + 1));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1432");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1432
            this.Site.CaptureRequirementIfAreEqual<string>(
                firstGUIDPortionOfRequestId,
                secondGUIDPortionOfRequestId,
                1432,
                @"[In X-RequestId Header Field] The GUID portion of the X-RequestId header is same in two different responses in one Session Contexts.");

            Site.Assert.AreEqual<int>(firstCounterOfRequestId + 1, secondCounterOfRequestId, "The counter portion of X-RequesutId header value should increase with every new HTTP request.");

            // If code can reach here, then the GUID portion of the X-RequestId header is same and counter portion of X-RequesutId header is increased. 
            // So R1330 is verified.
            this.Site.CaptureRequirement(
                1330,
                @"[In X-RequestId Header Field] In one client instance, the value of X-RequestId header field is different for two different responses and in the format of a GUID followed by an increasing decimal counter which MUST increase with every new HTTP request.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1147: the Set-Cookie header in Connect response is {0}.", headers["Set-Cookie"]);
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1147
            bool isVerifiedR1147 = headers["Set-Cookie"] != null && executeSuccessResponse.ErrorCode == 0;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1147,
                1147,
                @"[[In Creating a Session Context by Using the Connect or Bind Request Type] As specified in section 3.2.5.1, the server returns cookies used to identify the Session Context that has been created.");
            #endregion

            #region Send a Disconnect request type request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            DisconnectSuccessResponseBody disconnectResponseBody = (DisconnectSuccessResponseBody)responseBody;
            Site.Assert.AreEqual<uint>(0, disconnectResponseBody.ErrorCode, "Disconnect should succeed and 0 is expected to be returned. The returned value is {0}.", disconnectResponseBody.ErrorCode);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1250");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1250
            // The above assert ensures that the Disconnect request type executes successfully, so R1250 can be verified if the Disconnect response body is not null.
            this.Site.CaptureRequirementIfIsNotNull(
                responseBody,
                1250,
                @"[In Responding to a Disconnect or Unbind Request Type Request] The server sends a response, as specified in section 2.2.2.2, to a Disconnect request type or Unbind request type request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2234");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2234
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody,
                typeof(DisconnectSuccessResponseBody),
                2234,
                @"[In Responding to a Disconnect or Unbind Request Type Request] If successful, the server's response includes the Disconnect request type success response body, as specified in section 2.2.4.3.2.");
            #endregion

            #region Send an Execute request after disconnect with server.
            uint responseCodeAfterDisconenct = this.ExecuteLogonROP(AdapterHelper.SessionContextCookies);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R304");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R304
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                responseCodeAfterDisconenct,
                304,
                @"[In Disconnect Request Type] The Disconnect request type is used by the client to delete a Session Context with the server, as specified in section 3.1.5.4.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R146");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R146
            this.Site.CaptureRequirementIfAreEqual<uint>(
                10,
                responseCodeAfterDisconenct,
                146,
                @"[In X-ResponseCode Header Field] Context Not Found (10): The Session Context is not found.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1225");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1225
            this.Site.CaptureRequirementIfAreEqual<uint>(
                10,
                responseCodeAfterDisconenct,
                1225,
                @"[In Responding to a Connect or Bind Request Type Request] If the authentication context differs, the server MUST fail the request with a value of 10 (""Context Not Found"" error) in the X-ResponseCode header.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1253");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1253
            this.Site.CaptureRequirementIfAreEqual<uint>(
                10,
                responseCodeAfterDisconenct,
                1253,
                @"[In Responding to a Disconnect or Unbind Request Type Request] If the client attempts to use an invalid session context cookie [which is released by server] in a request, the server MUST fail the request to indicate to the client that the Session Context is not valid.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2236");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2236
            this.Site.CaptureRequirementIfAreEqual<uint>(
                10,
                responseCodeAfterDisconenct,
                2236,
                @"[In Responding to a Disconnect or Unbind Request Type Request] Once the session context cookie has been invalidated by the server, the client cannot use it in subsequent requests.");
            #endregion

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse2 = this.ConnectToServer(out headers);
            string thirdGUIDPortionOfRequestId = headers["X-RequestId"].Substring(0, headers["X-RequestId"].IndexOf(":"));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1464");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1464
            this.Site.CaptureRequirementIfAreNotEqual<string>(
                firstGUIDPortionOfRequestId,
                thirdGUIDPortionOfRequestId,
                1464,
                @"[In X-RequestId Header Field] The GUID portion of the X-RequestId header is different for different Session Contexts.");
            #endregion

            #region Send a Disconnect request type request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            disconnectResponseBody = (DisconnectSuccessResponseBody)responseBody;
            Site.Assert.AreEqual<uint>(0, disconnectResponseBody.ErrorCode, "Disconnect should succeed and 0 is expected to be returned. The returned value is {0}.", disconnectResponseBody.ErrorCode);
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to reconnecting and establishing a new Session Context.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC02_ReconnectToMailboxServer()
        {
            this.CheckMapiHttpIsSupported();
            
            MailboxResponseBodyBase responseBody;

            #region Send a valid Connect request type to establish a Session Context with the server.
            WebHeaderCollection connectHeaders = new WebHeaderCollection();
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out connectHeaders);
            #endregion

            #region Wait for the Session Context expired.
            int expireTime = int.Parse(connectHeaders["X-ExpirationInfo"]);
            System.Threading.Thread.Sleep(expireTime);
            
            List<string> metatagsFromMailbox = new List<string>();
            WebHeaderCollection pingHeaders = new WebHeaderCollection();
            uint responseCode = this.Adapter.PING(ServerEndpoint.MailboxServerEndpoint, out metatagsFromMailbox, out pingHeaders);

            if (responseCode != 10)
            {
                expireTime = int.Parse(pingHeaders["X-ExpirationInfo"]);
                System.Threading.Thread.Sleep(expireTime);
            }
            #endregion

            #region Send an Execute request that includes a Logon ROP to server after the Session Context expired.
            uint responseCodeAfterSessionContextExpired = this.ExecuteLogonROP(AdapterHelper.SessionContextCookies);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1283");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1283
            this.Site.CaptureRequirementIfAreEqual<uint>(
                10,
                responseCodeAfterSessionContextExpired,
                1283,
                @"[In Reconnecting and Establishing a New Session Context Server] If the Session Context has expired, is no longer valid, or is not valid for the server in which the mailbox currently resides, then the server fails the request with an X-ResponseCode value of 10, as specified in section 2.2.3.3.3.");
            #endregion

            #region Send a Connect request type which includes a valid cookie to reconnect with server.
            connectHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Connect, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            CookieCollection cookies = AdapterHelper.SessionContextCookies;
            HttpStatusCode httpStatusCode;
            this.Adapter.Connect(this.AdminUserName, this.AdminUserPassword, this.AdminUserDN, ref cookies, out responseBody, ref connectHeaders, out httpStatusCode);
            uint responseCodeHasSequenceCookie = AdapterHelper.GetFinalResponseCode(connectHeaders["X-ResponseCode"]);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1280");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1280
            // If the Set-Cookie header is not null, then server passes the cookies to the client.
            this.Site.CaptureRequirementIfIsNotNull(
                connectHeaders["Set-Cookie"],
                1280,
                @"[In Reconnecting and Establishing a New Session Context Server] The response from the server uses the Set-Cookie header, as specified in section 2.2.3.2.3, to pass any required cookies to the client.");
            #endregion

            #region Send an Execute request that includes a Logon ROP to server.
            uint responseCodeAfterReconnect = this.ExecuteLogonROP(AdapterHelper.SessionContextCookies);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1278");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1278
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseCodeAfterReconnect,
                1278,
                @"[In Reconnecting and Establishing a New Session Context Server] The server returns a new session context cookie that is associated with a new Session Context.");
            #endregion

            #region Send a Connect request type which includes a valid Session Context cookie and doesn't include sequence cookie to reconnect with server.
            connectHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Connect, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            CookieCollection cookieNotSequence = new CookieCollection();
            cookieNotSequence.Add(AdapterHelper.SessionContextCookies[0]);

            this.Adapter.Connect(this.AdminUserName, this.AdminUserPassword, this.AdminUserDN, ref cookieNotSequence, out responseBody, ref connectHeaders, out httpStatusCode);
            uint responseCodeNotSequenceCookie = AdapterHelper.GetFinalResponseCode(connectHeaders["X-ResponseCode"]);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1329: The X-Response header is {0} when client sends the sequence cookie and the X-Response header is {1} when client does not send the sequence cookie", responseCodeHasSequenceCookie, responseCodeNotSequenceCookie);
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1329
            bool isVerifiedR1329 = responseCodeNotSequenceCookie == 0 && responseCodeHasSequenceCookie == 0;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1329,
                1329,
                @"[In Reconnecting and Establishing a New Session Context Server] Server will reconnect successfully no matter client sends the sequence validation cookie or not.");
            #endregion

            #region Send a Disconnect request type request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            #endregion
        }
        
        /// <summary>
        /// This case is used to verify the requirements related to the X-ResponseCode header's values that return from server.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC03_ResponseCodeHeader()
        {
            this.CheckMapiHttpIsSupported();
            WebHeaderCollection headers = new WebHeaderCollection();
            MailboxResponseBodyBase responseBody;

            #region Send a Connect request to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R133");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R133
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                AdapterHelper.GetFinalResponseCode(headers["X-ResponseCode"]),
                133,
                @"[In X-ResponseCode Header Field] An X-ResponseCode of 0 (zero) means success from the perspective of the protocol transport.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R136");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R136
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                AdapterHelper.GetFinalResponseCode(headers["X-ResponseCode"]),
                136,
                @"[In X-ResponseCode Header Field] Success (0): The request was properly formatted and accepted.");

            Site.Assert.IsNotNull(connectResponse, "The response body should be parsed succeeded if the X-ResponseCode header is 0");

            // According to steps above, when the X-ResponseCode is 0, server return the valid response body that is parsed by client.
            // So MS-OXCMAPIHTTP_R55 can be verified.
            this.Site.CaptureRequirement(
                55,
                @"[In Common Response Format] An X-ResponseCode of 0 (zero) means success from the perspective of the protocol transport, and the client parses the RESPONSE BODY based on the request that was issued.");
            #endregion

            #region Send an Execute request which includes an invalid X-RequestType header.
            WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, Guid.NewGuid().ToString(), 1);
            executeHeaders.Set("X-RequestType", "invalidValue");

            ExecuteRequestBody requestBody = this.InitializeExecuteRequestBody(this.GetRopLogonRequest());
            List<string> metaTags = new List<string>();
            responseBody = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R141");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R141
            this.Site.CaptureRequirementIfAreEqual<uint>(
                5,
                AdapterHelper.GetFinalResponseCode(executeHeaders["X-ResponseCode"]),
                141,
                @"[In X-ResponseCode Header Field] Invalid Request Type (5): The request has an invalid X-RequestType header.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2232");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2232
            this.Site.CaptureRequirementIfIsNull(
                responseBody,
                2232,
                @"[In Responding to All Request Type Requests] If an additional X-ResponseCode header is returned and if it indicates a failure, then the response body can be empty or can include failure data in a format that is specified by an additional Content-Type header as specified in section 2.2.3.2.2.");

            #endregion

            #region Send an Execute request which misses cookie.
            uint responseCodeMissCookie = this.ExecuteLogonROP(new CookieCollection());

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R149");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R149
            this.Site.CaptureRequirementIfAreEqual<uint>(
                13,
                responseCodeMissCookie,
                149,
                @"[In X-ResponseCode Header Field] Missing Cookie (13): The request is missing a required cookie.");
            #endregion

            #region Send a Disconnect request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            #endregion
        }

        /// <summary>
        /// This case is used to test the requirements related to the request that include an invalid cookie.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC04_SendInvalidContextCookie()
        {
            this.CheckMapiHttpIsSupported();
            WebHeaderCollection headers = new WebHeaderCollection();
            MailboxResponseBodyBase responseBody;

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
            Site.Assert.AreEqual<uint>(0, connectResponse.StatusCode, "The server should return a Status 0 in X-ResponseCode header if client connect to server succeeded.");
            Site.Assert.IsTrue(
                AdapterHelper.SessionContextCookies.Count > 0,
                "The server should return Session Context cookie if server creates new Session Context successfully. The count of the cookie is {0}",
                AdapterHelper.SessionContextCookies.Count);
            #endregion

            #region Send an Execute request which includes an invalid session context cookie.
            CookieCollection invalidCookies = new CookieCollection();
            for (int i = 0; i < AdapterHelper.SessionContextCookies.Count; i++)
            {
                Cookie cookie = new Cookie()
                {
                    Domain = Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Name = AdapterHelper.SessionContextCookies[i].Name,
                    Value = "Invalid cookie"
                };
                invalidCookies.Add(cookie);
            }

            uint responseCodeInvalidCookie = this.ExecuteLogonROP(invalidCookies);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R142");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R142
            this.Site.CaptureRequirementIfAreEqual<uint>(
                6,
                responseCodeInvalidCookie,
                142,
                @"[In X-ResponseCode Header Field] Invalid Context Cookie (6): The request has an invalid session context cookie.");
            #endregion

            #region Send a Disconnect request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            #endregion
        }

        /// <summary>
        /// This case is used to test the requirements related to the request that include an invalid request body.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC05_SendInvalidRequestBody()
        {
            this.CheckMapiHttpIsSupported();
            WebHeaderCollection headers = new WebHeaderCollection();
            MailboxResponseBodyBase responseBody;

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
            Site.Assert.AreEqual<uint>(0, connectResponse.StatusCode, "The server should return a Status 0 in X-ResponseCode header if client connect to server succeeded.");
            #endregion

            #region Send an Execute request includes an invalid request body.
            WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);

            byte[] ropBuffer = RopBufferHelper.BuildRequestBuffer(this.GetRopLogonRequest(), 0);

            // Construct an invalid request body that RopBuffer is null and the RopBufferSize is not zero.
            ExecuteRequestBody invalidRequestBody = new ExecuteRequestBody();
            invalidRequestBody.Flags = 0x00000003;
            invalidRequestBody.RopBufferSize = (uint)ropBuffer.Length;
            invalidRequestBody.RopBuffer = new byte[] { };
            invalidRequestBody.MaxRopOut = 0x10008;
            invalidRequestBody.AuxiliaryBufferSize = 0;
            invalidRequestBody.AuxiliaryBuffer = new byte[] { };

            List<string> metaTags = new List<string>();
            responseBody = this.SendExecuteRequest(invalidRequestBody, ref executeHeaders, out metaTags);
           
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R148");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R148
            this.Site.CaptureRequirementIfAreEqual<uint>(
                12,
                AdapterHelper.GetFinalResponseCode(executeHeaders["X-ResponseCode"]),
                148,
                @"[In X-ResponseCode Header Field] Invalid Request Body (12): The request body is invalid.");
            #endregion

            #region Send a Disconnect request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            #endregion
        }

        /// <summary>
        /// This case is used to test the requirements related to PING request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC06_PINGRequestType()
        {
            this.CheckMapiHttpIsSupported();

            WebHeaderCollection headers = new WebHeaderCollection();
            MailboxResponseBodyBase responseBody;

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
            Site.Assert.AreEqual<uint>(0, connectResponse.StatusCode, "The server should return a Status 0 in X-ResponseCode header if client connect to server succeeded.");
           #endregion

            #region Send a PING request to Mailbox server endpoint.
            List<string> metatagsFromMailbox = new List<string>();
            uint responseCode = this.Adapter.PING(ServerEndpoint.MailboxServerEndpoint, out metatagsFromMailbox, out headers);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1124");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1124
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseCode,
                1124,
                @"[In PING Request Type] The PING request type allows a client to determine whether a server's endpoint (4) is reachable and operational.");
            #endregion

            #region Send a Disconnect request type request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            #endregion
        }

        /// <summary>
        /// This case is used to test the requirements related to client issues simultaneous requests within a Session Context.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC07_SimultaneousRequestWithinSameSessionContext()
        {
            this.CheckMapiHttpIsSupported();
            MailboxResponseBodyBase responseBody;

            #region Send a valid Connect request type to establish a Session Context with the server.
            WebHeaderCollection connectHeaders = new WebHeaderCollection();
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out connectHeaders);
            #endregion

            #region Send two Execute request that includes Logon ROP to server simultaneously
            uint firstResponseCode;
            uint secondResponseCode;
            CookieCollection firstCookies = new CookieCollection();
            CookieCollection secondCookies = new CookieCollection();
            foreach (Cookie cookie in AdapterHelper.SessionContextCookies)
            {
                Cookie c = new Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain);
                firstCookies.Add(c);
                secondCookies.Add(c);
            }

            MethodCaller asyncThread = new MethodCaller(
                () =>
                {
                    return this.ExecuteLogonROP(firstCookies);
                });

            IAsyncResult result = asyncThread.BeginInvoke(null, null);
            secondResponseCode = this.ExecuteLogonROP(secondCookies);
            firstResponseCode = asyncThread.EndInvoke(result);

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug, 
                "Verify MS-OXCMAPIHTTP_R1228: When the client has issued simultaneous requests within a Session Context,the first X-Response header is {0} and the second X-Response header is {1}.",
                firstResponseCode,
                secondResponseCode);
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1228
            bool isVerifiedR1228 = firstResponseCode == 15 || secondResponseCode == 15;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1228,
                1228,
                @"[In Responding to a Connect or Bind Request Type Request] If the server detects that the client has issued simultaneous requests within a Session Context, the server MUST fail every subsequent request with a value of 15 (""Invalid Sequence"" error) in the X-ResponseCode header.");

            // If the R1228 has been verified, then server return the X-ResponseCode 15 when the request has violated the sequencing requirement of one request at a time per Session Context.
            // So R151 will be verified.
            this.Site.CaptureRequirement(
                151,
                @"[In X-ResponseCode Header Field] Invalid Sequence (15): The request has violated the sequencing requirement of one request at a time per Session Context.");
            #endregion

            #region Send a Disconnect request type request to destroy the Session Context.
            this.Adapter.Disconnect(out responseBody);
            #endregion
        }

        /// <summary>
        ///  This case is used to test the requirements related to Execute request type.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC08_ExecuteRequestType()
        {
            this.CheckMapiHttpIsSupported();
            WebHeaderCollection headers = new WebHeaderCollection();

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
            Site.Assert.AreEqual<uint>(0, connectResponse.StatusCode, "The server should return a Status 0 in X-ResponseCode header if client connect to server succeeded.");
            #endregion

            #region Send an Execute request that includes Logon ROP to server.
            WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);

            ExecuteRequestBody requestBody = this.InitializeExecuteRequestBody(this.GetRopLogonRequest());
            List<string> metaTags = new List<string>();
            ExecuteSuccessResponseBody executeSuccessResponse = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1134");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1134
            // According to the Open Specification, the server must respond immediately to a request while the request is being queued and the initial response includes the PROCESSING meta-tag.
            // So MS-OXCMAPIHTTP_R1134 can be verified if the first meta-tag is PROCESSING.
            this.Site.CaptureRequirementIfAreEqual<string>(
                "PROCESSING",
                metaTags[0],
                1134,
                @"[In Response Meta-Tags] PROCESSING: The server has queued the request to be processed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1136");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1136
            // According to the Open Specification, the final response must include the DONE meta-tag. And the response body is parsed according to the description of this requirement.
            // So if the last meta-tag is DONE and code can reach here, MS-OXCMAPIHTTP_R1136 can be verified.
            this.Site.CaptureRequirementIfAreEqual<string>(
                "DONE",
                metaTags[metaTags.Count - 1],
                1136,
                @"[In Response Meta-Tags] DONE: The server has completed the processing of the request and additional response headers and the response body follow the DONE meta-tag.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1266. The header value of Transfer-Encoding is {0} and Content-Length is {1}.", executeHeaders["Transfer-Encoding"], executeHeaders["Content-Length"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1266
            bool isVerifiedR1266 = string.IsNullOrEmpty(executeHeaders["Content-Length"]) && executeHeaders["Transfer-Encoding"].ToLower() == "chunked".ToLower();

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1266,
                2227,
                @"[In Responding to All Request Type Requests] If the server is using the ""chunked"" transfer coding, it MUST flush this to the client (being careful to make sure it disables any internal Nagle algorithms, as described in [RFC896], that might attempt to buffer response data).");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R1230. The first meta-Tag is {0}, the value of header Transfer-Encoding is {1}.",
                metaTags[0],
                executeHeaders["Transfer-Encoding"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1230
            // Because the server will spend some time to process the Logon request. So the entire response is not readily available.
            // The first meta tag is not "DONE" indicates the server is processing the request and not done.
            bool isVerifiedR1230 = metaTags[0] != "DONE" && executeHeaders["Transfer-Encoding"].ToLower() == "chunked".ToLower();

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1230,
                1230,
                @"[In Responding to All Request Type Requests] If the entire response is not readily available, the server MUST use the Transfer-Encoding header, as specified in section 2.2.3.2.5, with a value of ""chunked"".");
            
            // R1230 ensures that the "chunked" transfer encoding is used and server can return data to the client while the request is still being processed.
            this.Site.CaptureRequirement(
                1174,
                @"[In Handling a Chunked Response] By using ""chunked"" transfer encoding, the server is able to return data to the client while the request is still being processed.");

            // R1230 ensures that the "chunked" transfer encoding is used and a positive connection between the server and client is established.
            this.Site.CaptureRequirement(
                1173,
                @"[In Handling a Chunked Response] To facilitate a positive connection between the server and client, the server uses the Transfer-Encoding header, as specified in section 2.2.3.2.5, with ""chunked"" transfer encoding, as specified in [RFC2616].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1244");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1244
            // The X-ResponseCode header is parsed after the DONE meta-tag, and Execute request type response body is parsed after the header X-ResponseCode. 
            // So if the last meta-tag is DONE and code can reach here, this requirement can be verified.
            this.Site.CaptureRequirementIfAreEqual<string>(
                metaTags[metaTags.Count - 1],
                "DONE",
                1244,
                @"[In Responding to All Request Type Requests] After the server finishes processing the request it finishes with the DONE meta-tag, as specified in section 2.2.7; followed by any additional response headers.");

            #endregion
            #endregion

            #region Send an Execute request to open one folder to check the logon operation succeeds.
            RPC_HEADER_EXT[] rpcHeaderExts;
            byte[][] rops;
            uint[][] serverHandleObjectsTables;

            RopBufferHelper ropBufferHelper = new RopBufferHelper(Site);
            ropBufferHelper.ParseResponseBuffer(executeSuccessResponse.RopBuffer, out rpcHeaderExts, out rops, out serverHandleObjectsTables);
            RopLogonResponse logonResponse = new RopLogonResponse();
            logonResponse.Deserialize(rops[0], 0);
            uint logonHandle = serverHandleObjectsTables[0][logonResponse.OutputHandleIndex];

            RopOpenFolderRequest openFolderRequest = this.OpenFolderRequest(logonResponse.FolderIds[4]);

            ExecuteRequestBody openFolderRequestBody = this.InitializeExecuteRequestBody(openFolderRequest, logonHandle);
            executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            executeSuccessResponse = this.SendExecuteRequest(openFolderRequestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;

            Site.Assert.AreEqual<uint>((uint)0, executeSuccessResponse.StatusCode, "Execute method should succeed.");
            ropBufferHelper.ParseResponseBuffer(executeSuccessResponse.RopBuffer, out rpcHeaderExts, out rops, out serverHandleObjectsTables);
            RopOpenFolderResponse openFolderResponse = new RopOpenFolderResponse();
            openFolderResponse.Deserialize(rops[0], 0);

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R226. The error code of Execute request type is {0}, the return value of RopOpenFolder is {1}.",
                executeSuccessResponse.ErrorCode,
                openFolderResponse.ReturnValue);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R226
            // If the return value of RopOpenFolder is 0, it indicates that the Logon handle is valid, so the Execute request type can be used to send the remote operation requests.
            bool isVerifiedR226 = executeSuccessResponse.ErrorCode == 0 && openFolderResponse.ReturnValue == 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR226,
                226,
                @"[In Execute Request Type] The Execute request type is used by the client to send remote operation requests to the server.");
            #endregion

            #region Send a Disconnect request to destroy the Session Context.
            MailboxResponseBodyBase response;
            this.Adapter.Disconnect(out response);
            #endregion
        }

        /// <summary>
        /// This case is used to test the NotificationWait request type with pending event.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC09_NotificationWaitWithPendingEvent()
        {
            this.CheckMapiHttpIsSupported();
            WebHeaderCollection headers = new WebHeaderCollection();

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
            Site.Assert.AreEqual<uint>(0, connectResponse.StatusCode, "The server should return a Status 0 in X-ResponseCode header if client connect to server succeeded.");
            #endregion

            #region Send an Execute request that includes Logon ROP to server.
            WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);

            ExecuteRequestBody requestBody = this.InitializeExecuteRequestBody(this.GetRopLogonRequest());
            List<string> metaTags = new List<string>();

            ExecuteSuccessResponseBody executeSuccessResponse = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;

            ulong folderId;
            RopLogonResponse logonResponse = new RopLogonResponse();
            uint logonHandle = this.ParseLogonResponse(executeSuccessResponse.RopBuffer, out folderId, out logonResponse);
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "Client logon to the server should be succeed and 0 is expected to be returned for the RopLogon. The returned value is {0}.", logonResponse.ReturnValue);

            #endregion

            #region Call RopRegisterNotification to register an event on server.
            executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            requestBody = this.InitializeExecuteRequestBody(this.RegisterNotificationRequest(folderId));
            metaTags = new List<string>();

            executeSuccessResponse = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;
            Site.Assert.AreEqual<uint>((uint)0, executeSuccessResponse.StatusCode, "Execute method should succeed.");
            #endregion

            #region Send a new mail to trigger the event.
            bool isSendSuccess = this.SUTControlAdapter.SendMailItem();
            this.Site.Assert.IsTrue(isSendSuccess, "Send a mail should successfully.");
            this.isReceiveNewMail = isSendSuccess;
            #endregion

            #region Call NotificationWait Request Type to get the pending event.
            NotificationWaitRequestBody notificationWaitRequestBody = this.NotificationWaitRequest();
            WebHeaderCollection notificationWaitWebHeaderCollection = AdapterHelper.InitializeHTTPHeader(RequestType.NotificationWait, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            MailboxResponseBodyBase responseBody;
            Dictionary<string, string> addtionalHeaders = new Dictionary<string, string>();

            uint result = this.Adapter.NotificationWait(notificationWaitRequestBody, ref notificationWaitWebHeaderCollection, out responseBody, out metaTags, out addtionalHeaders);

            NotificationWaitSuccessResponseBody notificationWaitResponseBody = new NotificationWaitSuccessResponseBody();
            notificationWaitResponseBody = (NotificationWaitSuccessResponseBody)responseBody;
            Site.Assert.AreEqual<uint>((uint)0, notificationWaitResponseBody.StatusCode, "NotificationWait method should succeed and 0 is expected to be returned. The returned value is {0}.", notificationWaitResponseBody.StatusCode);

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1371");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1371
            // A pending event is registered in step 3 and triggered in step 4, so MS-OXCMAPIHTTP_R1371 can be verified if the value of EventPending flag is 0x00000001.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000001,
                notificationWaitResponseBody.EventPending,
                1371,
                @"[In NotificationWait Request Type Success Response Body] [EventPending] The value 0x00000001 indicates that an event is pending.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1375");

            // The pending event is registered and triggered in the above steps, so MS-OXCMAPIHTTP_R1258 can be verified if code can reach here.
            this.Site.CaptureRequirement(
                1258,
                @"[In Responding to a NotificationWait Request Type Request] This response [NotificationWait request type response] is not sent until the current server event completes.");

            #endregion
            #endregion

            #region Call Execute Request Type with no ROP in rgbIn to get the notify information.
            executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            byte[] ropBuffer = this.RopBufferHelper.BuildRequestBufferWithoutRop();
            metaTags = new List<string>();

            requestBody = new ExecuteRequestBody();
            requestBody.Flags = 0x00000003;
            requestBody.RopBufferSize = (uint)ropBuffer.Length;
            requestBody.RopBuffer = ropBuffer;
            requestBody.MaxRopOut = 0x10008;
            requestBody.AuxiliaryBufferSize = 0;
            requestBody.AuxiliaryBuffer = new byte[] { };

            executeSuccessResponse = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;
            Site.Assert.AreEqual<uint>((uint)0, executeSuccessResponse.StatusCode, "Execute method should succeed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R318");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R318
            // No ROP is included in the request, so if the RopBuffer in the response is not null, it indicates that the event details are included in the Execute request type response buffer.
            this.Site.CaptureRequirementIfIsNotNull(
                executeSuccessResponse.RopBuffer,
                318,
                @"[In NotificationWait Request Type Success Response Body] [EventPending] The server will return the event details in the Execute request type response body.");
            #endregion

            #region Send a Disconnect request to destroy the Session Context.
            MailboxResponseBodyBase response;
            this.Adapter.Disconnect(out response);
            #endregion
        }

        /// <summary>
        /// This case is used to test the NotificationWait request type without pending event.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC10_NotificationWaitWithoutPendingEvent()
        {
            this.CheckMapiHttpIsSupported();
            WebHeaderCollection headers = new WebHeaderCollection();

            #region Send a valid Connect request type to establish a Session Context with the server.
            ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
            Site.Assert.AreEqual<uint>(0, connectResponse.StatusCode, "The server should return a Status 0 in X-ResponseCode header if client connect to server succeeded.");
            #endregion

            #region Send an Execute request that includes Logon ROP to server.
            WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);

            ExecuteRequestBody requestBody = this.InitializeExecuteRequestBody(this.GetRopLogonRequest());
            List<string> metaTags = new List<string>();

            ExecuteSuccessResponseBody executeSuccessResponse = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;
            #region Capture code
            string pendingPeriodHeader = executeHeaders["X-PendingPeriod"];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1182");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1182
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(pendingPeriodHeader),
                1182,
                @"[In Handling a Chunked Response] The immediate response includes an X-PendingPeriod header, specified in section 2.2.3.3.5, to tell the client the number of milliseconds to be expected between keep-alive PENDING meta-tags in the response stream during the time a request is currently being processed on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1183");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1183
            this.Site.CaptureRequirementIfAreEqual<string>(
                "15000",
                pendingPeriodHeader,
                1183,
                @"[In Handling a Chunked Response] The default value for the keep-alive interval is 15 seconds, until the request is done.");
            #endregion
            #endregion

            #region Call NotificationWait Request Type to get the pending event.
            NotificationWaitRequestBody notificationWaitRequestBody = this.NotificationWaitRequest();
            WebHeaderCollection notificationWaitWebHeaderCollection = AdapterHelper.InitializeHTTPHeader(RequestType.NotificationWait, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            MailboxResponseBodyBase responseBody;
            Dictionary<string, string> addtionalHeaders = new Dictionary<string, string>();
            DateTime startTime = DateTime.Now;
            uint result = this.Adapter.NotificationWait(notificationWaitRequestBody, ref notificationWaitWebHeaderCollection, out responseBody, out metaTags, out addtionalHeaders);
            DateTime endTime = DateTime.Now;
            TimeSpan interval = endTime.Subtract(startTime);

            NotificationWaitSuccessResponseBody notificationWaitResponseBody = new NotificationWaitSuccessResponseBody();
            notificationWaitResponseBody = (NotificationWaitSuccessResponseBody)responseBody;
            Site.Assert.AreEqual<uint>((uint)0, notificationWaitResponseBody.StatusCode, "NotificationWait method should succeed and 0 is expected to be returned. The returned value is {0}.", notificationWaitResponseBody.StatusCode);

            #region Capture code
            bool isPendingMetaTag = false;
            for (int i = 1; i < metaTags.Count - 1; i++)
            {
                if (string.Compare(metaTags[i], "PENDING", true) == 0)
                {
                    isPendingMetaTag = true;
                }
                else
                {
                    this.Site.Log.Add(LogEntryKind.Debug, "Expect the Meta-Tag is PENDING, actually the Meta-Tag is {0}", metaTags[i]);
                    isPendingMetaTag = false;
                    break;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1135");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1135
            // No pending event exists in this case, so NotificationWait needs to wait 5 minutes and the PENDING meta-tag will be returned before the DONE meta-tag is returned.
            this.Site.CaptureRequirementIfIsTrue(
                isPendingMetaTag,
                1135,
                @"[In Response Meta-Tags] PENDING: The server is processing the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1178");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1178
            // Server keeps the response alive by sending response to client including PENDING meta-tag. So if the PENDING meta-tag is included in the returned meta-tags, R1178 can be verified.
            this.Site.CaptureRequirementIfIsTrue(
                isPendingMetaTag,
                1178,
                @"[In Handling a Chunked Response] The keep-alive response contains the PENDING meta-tag.");

            // R1178 ensures that the keep-alive response includes the PENDING meta-tags, so R1241 can be verified directly.
            this.Site.CaptureRequirement(
                1241,
                @"[In Responding to All Request Type Requests] The keep-alive response includes the PENDING meta-tag, as specified in section 2.2.7.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R1185. The first meta-tag is {0}, the meta-tag before the last one is {1}, the last one is {2}.",
                metaTags[0],
                metaTags[metaTags.Count - 2],
                metaTags[metaTags.Count - 1]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1185
            // Server uses the "PENDING" meta-tag to keep the transmission alive with client. So R1185 can be verified if all the three kinds of meta-tags exist.
            bool isVerifiedR1185 = metaTags[0] == "PROCESSING" && metaTags[metaTags.Count - 2] == "PENDING" && metaTags[metaTags.Count - 1] == "DONE";
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1185,
                1185,
                @"[In Handling a Chunked Response] The initial response, plus the intermediate keep-alive transmissions, and the final response body are all part of the inner response stream.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1372");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1372
            // No pending event is registered in this case. So MS-OXCMAPIHTTP_R1372 can be verified if the value of EventPending field is 0x00000000.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                notificationWaitResponseBody.EventPending,
                1372,
                @"[In NotificationWait Request Type Success Response Body] [EventPending] The value 0x00000000 indicates that no event is pending.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1256");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1256
            this.Site.CaptureRequirementIfIsNotNull(
                notificationWaitResponseBody,
                1256,
                @"[In Responding to a NotificationWait Request Type Request] The server creates a NotificationWait request type response, as specified in section 2.2.2.2, including the NotificationWait request type response body as specified in section 2.2.4.4.2 if the request was successful.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1259. No pending event exists on the server. The NotificationWait request type execute time is {0}.", interval.TotalMinutes);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1259
            // Because above step does not trigger any event, so if EventPending is false and the execute time of NotificationWait is larger than 5 minutes, R1930 will be verified.
            bool isVerifiedR1259 = interval.TotalSeconds >= 300 && notificationWaitResponseBody.EventPending == 0x00000000;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1259,
                1259,
                @"[In Responding to a NotificationWait Request Type Request] This response [NotificationWait request type response] is not sent until the 5-minute maximum time limit expires.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R308");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R308
            // If the response body is not null, it indicates that method NotificationWait has been completed because server will return the event details in the ROP response buffer.
            this.Site.CaptureRequirementIfIsNotNull(
                notificationWaitResponseBody,
                308,
                @"[In NotificationWait Request Type] The NotificationWait request type is used by the client to request that the server notify the client when a processing request that takes an extended amount of time completes.");
            #endregion
            #endregion

            #region Send a Disconnect request to destroy the Session Context.
            MailboxResponseBodyBase response;
            this.Adapter.Disconnect(out response);
            #endregion
        }

        /// <summary>
        /// This case is designed to verify the requirements related to HTTP status code.
        /// </summary>
        [TestCategory("MSOXCMAPIHTTP"), TestMethod]
        public void MSOXCMAPIHTTP_S01_TC11_HTTPStatusCode()
        {
            this.CheckMapiHttpIsSupported();

            WebHeaderCollection headers = new WebHeaderCollection();
            CookieCollection cookies = new CookieCollection();
            MailboxResponseBodyBase response;

            #region Send a Connect request type request which misses X-RequestId header.
            HttpStatusCode httpStatusCodeFailed;
            AdapterHelper.ClientInstance = Guid.NewGuid().ToString();
            AdapterHelper.Counter = 1;
            headers = AdapterHelper.InitializeHTTPHeader(RequestType.Connect, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            headers.Remove("X-RequestId");
            this.Adapter.Connect(this.AdminUserName, this.AdminUserPassword, this.AdminUserDN, ref cookies, out response, ref headers, out httpStatusCodeFailed);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R143");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R143
            this.Site.CaptureRequirementIfAreEqual<uint>(
                7,
                AdapterHelper.GetFinalResponseCode(headers["X-ResponseCode"]),
                143,
                @"[In X-ResponseCode Header Field] Missing Header (7): The request has a missing required header.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R49");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R49
            this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                HttpStatusCode.OK,
                httpStatusCodeFailed,
                49,
                @"[In Common Response Format] The server returns a ""200 OK"" HTTP status code when the request failed for all non-exceptional failures.");
            #endregion

            #region Send a valid Connect request type to estabish a Session Context with the server.
            HttpStatusCode httpStatusCodeSucceed;
            AdapterHelper.ClientInstance = Guid.NewGuid().ToString();
            AdapterHelper.Counter = 1;
            headers = AdapterHelper.InitializeHTTPHeader(RequestType.Connect, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            string requestIdValue = headers["X-RequestId"];
            cookies = new CookieCollection();
            this.Adapter.Connect(this.AdminUserName, this.AdminUserPassword, this.AdminUserDN, ref cookies, out response, ref headers, out httpStatusCodeSucceed);
            Site.Assert.AreEqual<uint>(0, uint.Parse(headers["X-ResponseCode"]), "The server should return a Status 0 in X-ResponseCode header if client connect to server succeeded.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R48");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R48
            this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                HttpStatusCode.OK,
                httpStatusCodeSucceed,
                48,
                @"[In Common Response Format] The server returns a ""200 OK"" HTTP status code when the request succeeds.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2040");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2040
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestIdValue,
                headers["X-RequestId"],
                2040,
                @"[In X-RequestId Header Field] the server MUST return this header [X-RequestId] with the same information in the response back to the client.");
            #endregion

            #region Send a Disconnect request type request to destroy the Session Context.
            this.Adapter.Disconnect(out response);
            #endregion

            #region Send an anonymous request that includes a Connect request type.

            System.Threading.Thread.Sleep(60000);

            HttpStatusCode httpStatusCodeUnauthorized;
            AdapterHelper.ClientInstance = Guid.NewGuid().ToString();
            AdapterHelper.Counter = 1;
            headers = AdapterHelper.InitializeHTTPHeader(RequestType.Connect, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            cookies = new CookieCollection();

            this.Adapter.Connect(string.Empty, string.Empty, string.Empty, ref cookies, out response, ref headers, out httpStatusCodeUnauthorized);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1331");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1331
            this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                HttpStatusCode.Unauthorized,
                httpStatusCodeUnauthorized,
                1331,
                @"[In Common Response Format] The server deviates from a ""200 OK"" HTTP status code for authentication (""401 Access Denied"").");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R36");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R36
            this.Site.CaptureRequirementIfAreNotEqual<HttpStatusCode>(
                HttpStatusCode.OK,
                httpStatusCodeUnauthorized,
                36,
                @"[In POST Method] No anonymous requests are allowed.");
            #endregion
        }
        #endregion Test Cases

        /// <summary>
        /// Clean up the test.
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.isReceiveNewMail == true)
            {
                WebHeaderCollection headers = new WebHeaderCollection();
                MailboxResponseBodyBase response;

                #region Send a valid Connect request type to establish a Session Context with the server.
                ConnectSuccessResponseBody connectResponse = this.ConnectToServer(out headers);
                #endregion

                #region Send an Execute request that incluldes Logon ROP to server.
                WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);

                ExecuteRequestBody requestBody = this.InitializeExecuteRequestBody(this.GetRopLogonRequest());
                List<string> metaTags = new List<string>();

                ExecuteSuccessResponseBody executeSuccessResponse = this.SendExecuteRequest(requestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;

                ulong folderId;
                RopLogonResponse logonResponse = new RopLogonResponse();
                uint logonHandle = this.ParseLogonResponse(executeSuccessResponse.RopBuffer, out folderId, out logonResponse);
                #endregion

                #region Send an Execute request to open inbox folder.
                RPC_HEADER_EXT[] rpcHeaderExts;
                byte[][] rops;
                uint[][] serverHandleObjectsTables;
                RopOpenFolderRequest openFolderRequest = this.OpenFolderRequest(logonResponse.FolderIds[4]);

                ExecuteRequestBody openFolderRequestBody = this.InitializeExecuteRequestBody(openFolderRequest, logonHandle);
                executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
                executeSuccessResponse = this.SendExecuteRequest(openFolderRequestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;

                RopBufferHelper ropBufferHelper = new RopBufferHelper(Site);
                ropBufferHelper.ParseResponseBuffer(executeSuccessResponse.RopBuffer, out rpcHeaderExts, out rops, out serverHandleObjectsTables);
                RopOpenFolderResponse openFolderResponse = new RopOpenFolderResponse();
                openFolderResponse.Deserialize(rops[0], 0);
                 uint folderHandle = serverHandleObjectsTables[0][openFolderResponse.OutputHandleIndex];
                #endregion

                #region Send an Execute request type to hard delete messages in inbox folder.
                RopHardDeleteMessagesAndSubfoldersRequest hardDeleteRequest;
                hardDeleteRequest.RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders;
                hardDeleteRequest.LogonId = ConstValues.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                hardDeleteRequest.InputHandleIndex = 0;
                hardDeleteRequest.WantAsynchronous = 0x00; // Synchronously
                hardDeleteRequest.WantDeleteAssociated = 0xFF; // TRUE: delete all messages and subfolders
                ExecuteRequestBody hardDeleteRequestBody = this.InitializeExecuteRequestBody(hardDeleteRequest, folderHandle);
                executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
                executeSuccessResponse = this.SendExecuteRequest(hardDeleteRequestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;
                RopHardDeleteMessagesAndSubfoldersResponse hardDeleteMessagesAndSubfoldersResponse = new RopHardDeleteMessagesAndSubfoldersResponse();
                hardDeleteMessagesAndSubfoldersResponse.Deserialize(rops[0], 0);
                #endregion

                #region Send an Execute request to open sent items folder.
                openFolderRequest = this.OpenFolderRequest(logonResponse.FolderIds[6]);

                openFolderRequestBody = this.InitializeExecuteRequestBody(openFolderRequest, logonHandle);
                executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
                executeSuccessResponse = this.SendExecuteRequest(openFolderRequestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;

                ropBufferHelper = new RopBufferHelper(Site);
                ropBufferHelper.ParseResponseBuffer(executeSuccessResponse.RopBuffer, out rpcHeaderExts, out rops, out serverHandleObjectsTables);
                openFolderResponse = new RopOpenFolderResponse();
                openFolderResponse.Deserialize(rops[0], 0);
                folderHandle = serverHandleObjectsTables[0][openFolderResponse.OutputHandleIndex];
                #endregion

                #region Send an Execute request type to hard delete messages in sent items folder.
                hardDeleteRequest.RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders;
                hardDeleteRequest.LogonId = ConstValues.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                hardDeleteRequest.InputHandleIndex = 0;
                hardDeleteRequest.WantAsynchronous = 0x00; // Synchronously
                hardDeleteRequest.WantDeleteAssociated = 0xFF; // TRUE: delete all messages and subfolders
                hardDeleteRequestBody = this.InitializeExecuteRequestBody(hardDeleteRequest, folderHandle);
                executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
                executeSuccessResponse = this.SendExecuteRequest(hardDeleteRequestBody, ref executeHeaders, out metaTags) as ExecuteSuccessResponseBody;
                hardDeleteMessagesAndSubfoldersResponse = new RopHardDeleteMessagesAndSubfoldersResponse();
                hardDeleteMessagesAndSubfoldersResponse.Deserialize(rops[0], 0);
                #endregion

                #region Send a Disconnect request to destroy the Session Context.
                this.Adapter.Disconnect(out response);
                #endregion

                this.isReceiveNewMail = false;
            }

            base.TestCleanup();
        }

        #region Private methods
        /// <summary>
        /// Send the Logon ROP to server via Execute request type request.
        /// </summary>
        /// <param name="cookies">The Session Context cookie.</param>
        /// <returns>The value of the X-ResponseCode header.</returns>
        private uint ExecuteLogonROP(CookieCollection cookies)
        {
            MailboxResponseBodyBase responseBody;
            WebHeaderCollection executeHeaders = AdapterHelper.InitializeHTTPHeader(RequestType.Execute, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            ExecuteRequestBody requestBody = this.InitializeExecuteRequestBody(this.GetRopLogonRequest());
            List<string> metaTags = new List<string>();

            this.Adapter.Execute(requestBody, cookies, ref executeHeaders, out responseBody, out metaTags);

            return AdapterHelper.GetFinalResponseCode(executeHeaders["X-ResponseCode"]);
        }

        /// <summary>
        /// Parse the logon response.
        /// </summary>
        /// <param name="logonRop">An array of bytes that constitute the ROP response payload.</param>
        /// <param name="logonFolderId">The folder ID which client logons on to it.</param>
        /// <param name="logonResponse">The logon response parsed from the ROP response payload.</param>
        /// <returns>The logon handle.</returns>
        private uint ParseLogonResponse(byte[] logonRop, out ulong logonFolderId, out RopLogonResponse logonResponse)
        {
            RPC_HEADER_EXT[] rpcHeaderExts;
            byte[][] rops;
            uint[][] serverHandleObjectsTables;

            RopBufferHelper.ParseResponseBuffer(logonRop, out rpcHeaderExts, out rops, out serverHandleObjectsTables);
            RopLogonResponse response = new RopLogonResponse();
            response.Deserialize(rops[0], 0);
            uint logonHandle = serverHandleObjectsTables[0][response.OutputHandleIndex];

            // The folder Id of inbox folder. 
            logonFolderId = response.FolderIds[4]; 
            logonResponse = response;

            return logonHandle;
        }

        /// <summary>
        /// Build the NotificationWait request body.
        /// </summary>
        /// <returns>The NotificationWait request body.</returns>
        private NotificationWaitRequestBody NotificationWaitRequest()
        {
            NotificationWaitRequestBody notificationWaitRequestBody = new NotificationWaitRequestBody();
            notificationWaitRequestBody.Flags = ConstValues.ReserveDefault;
            byte[] auxIn = new byte[] { };
            notificationWaitRequestBody.AuxiliaryBuffer = auxIn;
            notificationWaitRequestBody.AuxiliaryBufferSize = (uint)auxIn.Length;

            return notificationWaitRequestBody;
        }

        /// <summary>
        /// Build the RopRegisterNotification request.
        /// </summary>
        /// <param name="folderId">The folder used to register notification.</param>
        /// <returns>The RopRegisterNotification request.</returns>
        private RopRegisterNotificationRequest RegisterNotificationRequest(ulong folderId)
        {
            RopRegisterNotificationRequest registerNotificationRequest = new RopRegisterNotificationRequest()
            {
                RopId = (byte)RopId.RopRegisterNotification,
                LogonId = ConstValues.LogonId,

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                InputHandleIndex = 0,

                // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle,
                // for the output Server object will be stored.
                OutputHandleIndex = 1,

                // The server MUST send notifications to the client when new object events occur within the scope of interest.
                NotificationTypes = (byte)NotificationTypes.NewMail,

                // This field is reserved. The field value MUST be 0x00.
                Reserved = 0x00, 
                
                // If the scope for notifications is the entire database, the value of wantWholeStore is true; otherwise, FALSE (0x00).
                WantWholeStore = 0x00,

                // Set the value of the specified folder ID. 
                FolderId = folderId,
                MessageId = 0
            };

            return registerNotificationRequest;
        }

        /// <summary>
        /// Build the open folder request.
        /// </summary>
        /// <param name="folderId">The folder ID which will be opened.</param>
        /// <returns>The RopOpenFolder request.</returns>
        private RopOpenFolderRequest OpenFolderRequest(ulong folderId)
        {
            RopOpenFolderRequest openFolderRequest;
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = ConstValues.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input server object is stored.
            openFolderRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle,
            // for the output server object will be stored.
            openFolderRequest.OutputHandleIndex = 1;
            openFolderRequest.FolderId = folderId;
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            return openFolderRequest;
        }

        /// <summary>
        /// Initialize an Execute request Body.
        /// </summary>
        /// <param name="rop">The ROP which is included in RopBuffer field in Execute request Body.</param>
        /// <returns>An instance of ExecuteRequestBody.</returns>
        private ExecuteRequestBody InitializeExecuteRequestBody(ISerializable rop)
        {
            return this.InitializeExecuteRequestBody(rop, 0);
        }

        /// <summary>
        /// Initialize an Execute request Body.
        /// </summary>
        /// <param name="rop">The ROP which is included in RopBuffer field in Execute request Body.</param>
        /// <param name="insideObjectHandle">The server object handle in request.</param>
        /// <returns>The Execute request body.</returns>
        private ExecuteRequestBody InitializeExecuteRequestBody(ISerializable rop, uint insideObjectHandle)
        {
            byte[] ropBuffer = this.RopBufferHelper.BuildRequestBuffer(rop, insideObjectHandle);

            ExecuteRequestBody requestBody = new ExecuteRequestBody();

            // The Flags field is 0x00000003 meaning the server must not compress and not obfuscate the ROP response payload.
            requestBody.Flags = 0x00000003;
            requestBody.RopBufferSize = (uint)ropBuffer.Length;
            requestBody.RopBuffer = ropBuffer;

            // An unsigned integer that specifies the maximum size for the RopBuffer field of the Execute request type success response body.
            requestBody.MaxRopOut = 0x10008;
            requestBody.AuxiliaryBufferSize = 0;
            requestBody.AuxiliaryBuffer = new byte[] { };

            return requestBody;
        }

        /// <summary>
        /// Compose a Logon ROP request.
        /// </summary>
        /// <returns>Return a Logon ROP request.</returns>
        private RopLogonRequest GetRopLogonRequest()
        {
            RopLogonRequest logonRop = new RopLogonRequest();
            logonRop.RopId = (byte)RopId.RopLogon; // RopId 0XFE indicates RopLogon.
            logonRop.LogonId = ConstValues.LogonId; // The logonId 0x00 is associated with this operation.
            logonRop.OutputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            string userDN = this.AdminUserDN + "\0"; // A null-terminated string that specifies the DN of the AdminUser who is requesting the connection.
            logonRop.StoreState = 0; // A flags structure. This field MUST be set to 0x00000000.
            logonRop.LogonFlags = 0x01; // Logon to a private mailbox.
            logonRop.OpenFlags = 0x01000000; // For a private-mailbox logon, the USE_PER_MDB_REPLID_MAPPING flag should be set. 
            logonRop.EssdnSize = (ushort)System.Text.Encoding.ASCII.GetByteCount(userDN);
            logonRop.Essdn = System.Text.Encoding.ASCII.GetBytes(userDN);

            return logonRop;
        }

        /// <summary>
        /// Send a valid Connect request type to establish a Session Context with the server.
        /// </summary>
        /// <param name="headers">The HTTP headers in response</param>
        /// <returns>Return a Connect request type successful response body.</returns>
        private ConnectSuccessResponseBody ConnectToServer(out WebHeaderCollection headers)
        {
            CookieCollection cookies = new CookieCollection();
            AdapterHelper.ClientInstance = Guid.NewGuid().ToString();
            AdapterHelper.Counter = 1;
            headers = AdapterHelper.InitializeHTTPHeader(RequestType.Connect, AdapterHelper.ClientInstance, AdapterHelper.Counter);
            MailboxResponseBodyBase response;
            HttpStatusCode httpStatusCode;
            this.Adapter.Connect(this.AdminUserName, this.AdminUserPassword, this.AdminUserDN, ref cookies, out response, ref headers, out httpStatusCode);
            Site.Assert.AreEqual<uint>(0, uint.Parse(headers["X-ResponseCode"]), "The server should return a Status 0 in X-ResponseCode header if client connects to server successfully.");

            return response as ConnectSuccessResponseBody;
        }

        /// <summary>
        /// Send an Execute request type that includes a ROP to server.
        /// </summary>
        /// <param name="requestBody">The Execute request body.</param>
        /// <param name="headers">The HTTP header in request and response.</param>
        /// <param name="metatags">The meta-tag in response.</param>
        /// <returns>Return an Execute successful response body.</returns>
        private MailboxResponseBodyBase SendExecuteRequest(ExecuteRequestBody requestBody, ref WebHeaderCollection headers, out List<string> metatags)
        {
            MailboxResponseBodyBase response;
            this.Adapter.Execute(requestBody, AdapterHelper.SessionContextCookies, ref headers, out response, out metatags);

            return response;
        }
        #endregion
    }
}