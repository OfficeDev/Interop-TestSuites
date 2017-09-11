namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Server role and both role Adapter requirements capture code for MS-OXCMAPIHTTP.
    /// </summary>
    public partial class MS_OXCMAPIHTTPAdapter
    {
        #region Define parameters
        /// <summary>
        /// The list of X-ResponseType values.
        /// </summary>
        private List<string> requestTypeList = new List<string> { "Connect", "Execute", "Disconnect", "NotificationWait", "PING", "Bind", "Unbind", "CompareMIds", "DNToMId", "GetMatches", "GetPropList", "GetProps", "GetSpecialTable", "GetTemplateInfo", "ModLinkAtt", "ModProps", "QueryColumns", "QueryRows", "ResolveNames", "ResortRestriction", "SeekEntries", "UpdateStat", "GetMailboxUrl", "GetAddressBookUrl" };

        /// <summary>
        /// The list of meta-tag values.
        /// </summary>
        private List<string> metaTagsList = new List<string> { "PROCESSING", "PENDING", "DONE" };

        #endregion

        #region Verify transport and authentication
        /// <summary>
        /// Verify the requirements related to HTTPS transport.
        /// </summary>
        /// <param name="response">The HttpWebResponse to be verified.</param>
        private void VerifyHTTPS(HttpWebResponse response)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2: the HTTP status code in response is {0}", response.StatusCode.ToString());

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2
            bool isVerifiedR2 = string.Compare("HTTPS", response.ResponseUri.Scheme, true) == 0 && response.StatusCode == HttpStatusCode.OK;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR2,
                2,
                @"[In Transport] The protocol MUST use HTTPS secure requests using version 1.1 of HTTP, as specified in [RFC2616] and [RFC2818].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R456");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R456
            this.Site.CaptureRequirementIfAreEqual<string>(
                "POST",
                response.Method,
                456,
                @"[In Transport] POST supports uploading a request block and returning a response block.");
        }

        /// <summary>
        /// Verify the requirements related to authentication.
        /// </summary>
        /// <param name="response">The HttpWebResponse to be verified.</param>
        private void VerifyAuthentication(HttpWebResponse response)
        {
            // If the StatusCode in response is OK, then the request has been authenticated by server.
            if (response.StatusCode == HttpStatusCode.OK)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R35");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R35
                // If the StatusCode in response is OK, then the request has been authenticated by server.
                // If the response includes the X-RequestType header and X-ResponseCode header, then the response follows MS-OXCMAPIHTTP.
                // So when the request has been authenticated and the response follows MS-OXCMAPIHTTP, server will return a correct response and R35 will be verified.
                bool isVerifiedR35 = response.Headers["X-RequestType"] != null && response.Headers["X-ResponseCode"] != null;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR35,
                    35,
                    @"[In POST Method] All requests MUST be authenticated prior to being processed by server.");
            }
        }
        #endregion

        #region Verify AutoDiscover
        /// <summary>
        /// Verify the requirements related to AutoDiscover.
        /// </summary>
        /// <param name="httpStatusCode">The HttpStatusCode to be verified.</param>
        /// <param name="serverEndpoint">The value of server endpoint.</param>
        private void VerifyAutoDiscover(HttpStatusCode httpStatusCode, ServerEndpoint serverEndpoint)
        {
            // The status code in response is 200 OK which means client accesses server successfully.
            if (httpStatusCode == HttpStatusCode.OK)
            {
                if (serverEndpoint == ServerEndpoint.MailboxServerEndpoint)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1436: the URL {0} is returned in Autodiscover for mailbox server point.", this.mailStoreUrl.Replace("\0", string.Empty));

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1436
                    // The value of mailStoreUri is returned from MailStore element in response of Autodiscover.
                    // So if the value of mailStoreUri follows the URI format, R1436 will be verified.
                    bool isVerifiedR1436 = Uri.IsWellFormedUriString(this.mailStoreUrl, UriKind.RelativeOrAbsolute);

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR1436,
                        1436,
                        @"[In POST Method] A separate URI is returned in Autodiscover for mailbox server point.");

                    // Because the URI includes the information about destination server, request path and optional parameters.
                    // So if R1436 has been verified, the AutoDiscover has returned the URIs for accessing a given mailbox and R28 will be verified.
                    this.Site.CaptureRequirement(
                        28,
                        @"[In POST Method] The destination server, request path, and optional parameters for accessing a given mailbox are returned in URIs from Autodiscover.");

                    // Because client uses this URI to access mailbox server.
                    // So if the status code in response is 200 OK, then client uses the URI to access mailbox server.
                    // So MS-OXCMAPIHTTP_R30 will be verified.
                    this.Site.CaptureRequirement(
                        30,
                        @"[In POST Method] The URI is used by the server to route a request to the appropriate mailbox server.");
                }

                if (serverEndpoint == ServerEndpoint.AddressBookServerEndpoint)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1437: the URI {0} is returned in Autodiscover for address book server point.", this.addressBookUrl.Replace("\0", string.Empty));

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1437
                    // The value of addressBookUri is returned from MailStore element in response of Autodiscover.
                    // So if the value of addressBookUri follows the URI format, then R1437 will be verified.
                    bool isVerifiedR1437 = Uri.IsWellFormedUriString(this.addressBookUrl, UriKind.RelativeOrAbsolute);

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR1437,
                        1437,
                        @"[In POST Method] A separate URI is returned in Autodiscover for address book server point.");
                }

                // In the upper two if statements, one verified Mailbox server endpoint and another one verified AddressBook server endpoint.
                // The validity of URI is verified in all the two endpoints, so if code run to here, R459 is verified.
                this.Site.CaptureRequirement(
                    459,
                    @"[In Transport] The Autodiscover response, as specified in [MS-OXDSCLI], contains a URI that the client will use to access the two endpoints (4) used by this protocol: the mailbox sever endpoint (same as that used for the EMSMDB interface) and the address book server endpoint (same as that used for the NSPI interface).");
            }
        }
        #endregion

        #region Verify HTTP headers
        /// <summary>
        /// Verify the requirements related to HTTP header.
        /// </summary>
        /// <param name="headers">The collection of HTTP headers.</param>
        private void VerifyHTTPHeaders(WebHeaderCollection headers)
        {
            // If the response includes the Transfer-Encoding header, then the response is a chunked response.
            if (!string.IsNullOrEmpty(headers["Transfer-Encoding"]))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R62");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R62
                // If the response includes Transfer-Encoding header, then this response is a chunked response.
                // If Content-Length header is not included in response, then R62 is verified.
                this.Site.CaptureRequirementIfIsNull(
                    headers["Content-Length"],
                    62,
                    @"[In Content-Length Header Field] This header [Content-Length Header] is not used in chunked requests and chunked responses.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R76");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R76
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "chunked",
                    headers["Transfer-Encoding"],
                    76,
                    @"[In Transfer-Encoding Header Field] The Transfer-Encoding header field contains the string ""chunked"" transfer coding, as specified in [RFC2616] section 3.6.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R132: the X-ResponseCode header is {0}.", headers["X-ResponseCode"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R132
            int responseCodeValue;

            this.Site.CaptureRequirementIfIsTrue(
                int.TryParse(headers["X-ResponseCode"], out responseCodeValue),
                132,
                @"[In X-ResponseCode Header Field] The X-ResponseCode header contains a numerical value that represents the specific result that occurred on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R155: the X-ClientInfo header is {0}.", headers["X-ClientInfo"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R155
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(headers["X-ClientInfo"]),
                155,
                @"[In X-ClientInfo Header Field] The X-ClientInfo header field MUST be a combination of a globally unique value in the format of a GUID followed by a decimal counter (for example, ""{2EF33C39-49C8-421C-B876-CDF7F2AC3AA0}:123"").");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R164: the X-ServerApplication header is {0}.", headers["X-ServerApplication"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R164
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(headers["X-ServerApplication"]),
                164,
                @"[In X-ServerApplication Header Field] On every response, the server includes the X-ServerApplication header to indicate to the client what server version is being used.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R165: the X-ServerApplication header is {0}.", headers["X-ServerApplication"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R165
            this.Site.CaptureRequirementIfIsTrue(
                Regex.IsMatch(headers["X-ServerApplication"], @"^Exchange/15.\d{2}.\d{4}.\d{3}$"),
                165,
                @"[In X-ServerApplication Header Field] The value of this header field [X-ServerApplication] has the following format: ""Exchange/15.xx.xxxx.xxx"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R167: the X-ExpirationInfo header is {0}.", headers["X-ExpirationInfo"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R167
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(headers["X-ExpirationInfo"]),
                167,
                @"[In X-ExpirationInfo Header Field] The X-ExpirationInfo header is returned by the server in every response to notify the client of the number of milliseconds before the server times-out the Session Context.");

            string setCookieHeader = headers["Set-Cookie"];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R68: the Set-Cookie is {0}.", setCookieHeader);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R68
            bool isVerifiedR68 = Regex.IsMatch(setCookieHeader, "^.*=.*$");

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR68,
                68,
                @"[In Set-Cookie Header Field] The Set-Cookie header field contains an opaque value of the form <cookie name>=<opaque string>.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1236");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1236
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(headers["Set-Cookie"]),
                1236,
                @"[In Responding to All Request Type Requests] The response includes all Set-Cookie headers as specified in section 2.2.3.2.4 associated with the Session Context.");

            string pendingPeriodHeader = headers["X-PendingPeriod"];
            int pendingPeriod = 0;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R157");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R157
            this.Site.CaptureRequirementIfIsTrue(
                pendingPeriodHeader != null && int.TryParse(pendingPeriodHeader, out pendingPeriod),
                157,
                @"[In X-PendingPeriod Header Field] The X-PendingPeriod header field, returned by the server, specifies the number of milliseconds to be expected between keep-alive PENDING meta-tags in the response stream while the server is executing the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R158");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R158
            this.Site.CaptureRequirementIfAreEqual<int>(
                15000,
                pendingPeriod,
                158,
                @"[In X-PendingPeriod Header Field] The default value of this header [X-PendingPeriod] is 15000 milliseconds (15 seconds).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1242");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1242
            this.Site.CaptureRequirementIfIsTrue(
                pendingPeriodHeader != null && int.TryParse(pendingPeriodHeader, out pendingPeriod),
                1242,
                @"[In Responding to All Request Type Requests] Since the keep-alive interval is configurable or auto-adjusted, the server MUST return the X-PendingPeriod header, specified in section 2.2.3.3.3, within the immediate response to tell the client the number of milliseconds to be expected between keep-alive responses from the server during the time a request is currently being executed on the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1243");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1243
            this.Site.CaptureRequirementIfAreEqual<int>(
                15000,
                pendingPeriod,
                1243,
                @"[In Responding to All Request Type Requests] The default value of the X-PendingPeriod header is 15 seconds.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2050");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2050
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(headers["X-DeviceInfo"]),
                2050,
                @"[In X-DeviceInfo Header Field] The server MUST not send this header [X-DeviceInfo] in a response to a client endpoint. ");
        }

        /// <summary>
        /// Verify the requirements related to Content-Type header.
        /// </summary>
        /// <param name="headers">The collection of HTTP headers.</param>
        private void VerifyContentTypeHeader(WebHeaderCollection headers)
        {
            string contentType = headers["Content-Type"];
            int responseCodeValue;
            int.TryParse(headers["X-ResponseCode"], out responseCodeValue);
            if (responseCodeValue == 0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2037");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2037
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "application/mapi-http",
                    contentType,
                    2037,
                    @"[In Content-Type Header Field] The Content-Type header MUST contain the string ""application/mapi-http"" on responses with X-ResponseCode header of 0.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2038");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2038
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "text/html",
                    contentType,
                    2038,
                    @"[In Content-Type Header Field] If X-ResponseCode is non-zero, the Content-Type header MUST contain the string ""text/html"".");
            }
        }

        /// <summary>
        /// Verify the requirements related to additional header.
        /// </summary>
        /// <param name="additionalHeaders">The additional headers.</param>
        private void VerifyAdditionalHeaders(Dictionary<string, string> additionalHeaders)
        {
            string elapsedTimeHeader = additionalHeaders["X-ElapsedTime"];
            int elapsedTimeHeaderValue;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R169");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R169
            // Because the additional headers are parsed from the final response according to Open Specification. 
            // So if the X-ElapsedTime is not null, then R169 will be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                elapsedTimeHeader,
                169,
                @"[In X-ElapsedTime Header Field] This header [X-ElapsedTime] is returned by the server as an additional header in the final response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R168: the X-ElapsedTime header is {0}", elapsedTimeHeader);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R168
            this.Site.CaptureRequirementIfIsTrue(
                int.TryParse(elapsedTimeHeader, out elapsedTimeHeaderValue),
                168,
                @"[In X-ElapsedTime Header Field] The X-ElapsedTime header specifies the amount of time, in milliseconds, that the server took to process the request.");

            string startTimeHeader = additionalHeaders["X-StartTime"];
            DateTime startTime;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R171");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R171
            // Because the additional headers are parsed from the final response according to Open Specification. 
            // So if the X-StartTime is not null, then R171 will be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                startTimeHeader,
                171,
                @"[In X-StartTime Header Field] This header [X-StartTime] is returned by the server as an additional header in the final response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R170");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R170
            // Because the additional headers are parsed from the final response according to Open Specification. 
            // So if the X-StartTime is not null and the value is a valid time, then R170 will be verified.
            this.Site.CaptureRequirementIfIsTrue(
                startTimeHeader != null && DateTime.TryParse(startTimeHeader, out startTime),
                170,
                @"[In X-StartTime Header Field] The X-StartTime header specifies the time that the server started processing the request.");
        }
        #endregion

        #region Verify PING request type
        /// <summary>
        /// Verify the PING request type related requirements.
        /// </summary>
        /// <param name="commonResponse">The CommonResponse to be verified.</param>
        /// <param name="endpoint">The value of server endpoint.</param>
        /// <param name="responseCode">The value of X-ResponseCode header.</param>
        private void VerifyPINGRequestType(CommonResponse commonResponse, ServerEndpoint endpoint, uint responseCode)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1460");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1460
            this.Site.CaptureRequirementIfAreEqual<int>(
                0,
                commonResponse.ResponseBodyRawData.Length,
                1460,
                @"[In PING Request Type] The PING request type has no response body.");

            if (responseCode == 0 && endpoint == ServerEndpoint.MailboxServerEndpoint)
            {
                // If the code can reach here, that means the Mailbox server endpoint executes successfully, R1127 is verified successfully.
                this.Site.CaptureRequirement(
                    1127,
                    @"[In PING Request Type] The PING request type is supported by the mailbox server endpoint (4).");
            }
            else if (responseCode == 0 && endpoint == ServerEndpoint.AddressBookServerEndpoint)
            {
                // If the code can reach here, that means the AddressBook server endpoint executes successfully, R1128 is verified successfully.
                this.Site.CaptureRequirement(
                    1128,
                    @"[In PING Request Type] The PING request type is supported by the address book server endpoint (4).");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1247");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1247
            // If the response is not null, that means the server responds to a PING request type.
            this.Site.CaptureRequirementIfIsNotNull(
                commonResponse,
                1247,
                @"[In Responding to a PING Request Type] The server responds to a PING request type by returning a PING request type response, as specified in section 2.2.6.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1249");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1249
            this.Site.CaptureRequirementIfIsTrue(
                commonResponse.MetaTags.Count != 0 && commonResponse.ResponseBodyRawData.Length == 0,
                1249,
                @"[In Responding to a PING Request Type] Meta-tags can be returned, but no response body is returned.");
        }

        #endregion

        #region Verify Response Meta-Tags
        /// <summary>
        /// Verify the Response Meta-Tags related requirements.
        /// </summary>
        /// <param name="metaTags">The meta-tag value to be verified.</param>
        private void VerifyResponseMetaTags(List<string> metaTags)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1130");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1130
            bool isVerifiedR1130 = false;

            for (int i = 0; i < metaTags.Count; i++)
            {
                if (this.metaTagsList.Contains(metaTags[i]))
                {
                    isVerifiedR1130 = true;
                }
                else
                {
                    isVerifiedR1130 = false;
                    break;
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1130,
                1130,
                @"[In Response Meta-Tags] The protocol defines three meta-tags [PROCESSING, PENDING, DONE] that are used to inform the client as to the state of processing a request on the server.");

            // The structure of response body is parsed according to the request of this Specification, if the code run to here and passed, 
            // that means R1131 is verified successfully.
            this.Site.CaptureRequirement(
                1131,
                @"[In Response Meta-Tags] A meta-tag is returned at the beginning of the response body.");
        }

        #endregion

        #region Verify AddressBookPropertyValue Structure
        /// <summary>
        /// Verify the AddressBookPropertyValue structure related requirements.
        /// </summary>
        /// <param name="addressBookPropertyValue">The AddressBookPropertyValue value to be verified.</param>
        private void VerifyAddressBookPropertyValueStructure(AddressBookPropertyValue addressBookPropertyValue)
        {
            // Since Test Suite parsed AddressBookPropertyValue structure according to Specification, if the program run to here, R2000 can be verified directly.
            this.Site.CaptureRequirement(
                2000,
                @"[In AddressBookPropertyValue Structure] The AddressBookPropertyValue structure includes a property value.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2004");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2005");

            bool isVerifyR2005 = addressBookPropertyValue.HasValue == 0xFF || addressBookPropertyValue.HasValue == 0x00;

            if (addressBookPropertyValue.HasValue != null)
            {
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2005
                this.Site.CaptureRequirementIfIsTrue(
                    isVerifyR2005,
                    2005,
                    @"[In AddressBookPropertyValue Structure] [HasValue] This field MUST contain either TRUE (0xFF) or FALSE (0x00). ");
            }
        }
        #endregion

        #region Verify AddressBookTaggedPropertyValue Structure
        /// <summary>
        /// Verify the AddressBookTaggedPropertyValue structure related requirements.
        /// </summary>
        /// <param name="addressBookTaggedPropertyValue">The AddressBookTaggedPropertyValue value to be verified.</param>
        private void VerifyAddressBookTaggedPropertyValueStructure(AddressBookTaggedPropertyValue addressBookTaggedPropertyValue)
        {
            // Since Test Suite parsed AddressBookTaggedPropertyValue structure according to Specification, if the program run to here, R2011 can be verified directly.
            this.Site.CaptureRequirement(
                2011,
                @"[In AddressBookTaggedPropertyValue Structure] The AddressBookTaggedPropertyValue structure includes property type, property identifier and property value.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2012");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2012
            this.Site.CaptureRequirementIfIsInstanceOfType(
                addressBookTaggedPropertyValue.PropertyType,
                typeof(ushort),
                2012,
                @"[In AddressBookTaggedPropertyValue Structure] PropertyType (2 bytes): An unsigned integer that identifies the data type of the property value ([MS-OXCDATA] section 2.11.1).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2013");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2013
            this.Site.CaptureRequirementIfIsInstanceOfType(
                addressBookTaggedPropertyValue.PropertyId,
                typeof(ushort),
                2013,
                @"[In AddressBookTaggedPropertyValue Structure] PropertyId (2 bytes): An unsigned integer that identifies the property.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2014");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2014
            this.Site.CaptureRequirementIfIsInstanceOfType(
                addressBookTaggedPropertyValue,
                typeof(AddressBookPropertyValue),
                2014,
                @"[In AddressBookTaggedPropertyValue Structure] PropertyValue (variable): An AddressBookPropertyValue structure, see section 2.2.1.1.");

        }
        #endregion

        #region Verify AddressBookFlaggedPropertyValue Structure
        /// <summary>
        /// Verify the AddressBookFlaggedPropertyValue structure related requirements.
        /// </summary>
        /// <param name="addressBookFlaggedPropertyValue">The AddressBookFlaggedPropertyValue value to be verified.</param>
        private void VerifyAddressBookFlaggedPropertyValueStructure(AddressBookFlaggedPropertyValue addressBookFlaggedPropertyValue)
        {
            // Since Test Suite parsed AddressBookFlaggedPropertyValue structure according to Specification, if the program run to here, R2018 can be verified directly.
            this.Site.CaptureRequirement(
                2018,
                @"[In AddressBookFlaggedPropertyValue Structure] The AddressBookFlaggedPropertyValue structure includes a flag to indicate whether the value was successfully retrieved or not.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2019");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2019
            this.Site.CaptureRequirementIfIsInstanceOfType(
                addressBookFlaggedPropertyValue.Flag,
                typeof(byte),
                2019,
                @"[In AddressBookFlaggedPropertyValue Structure] Flag (1 byte): An unsigned integer. ");

            if (addressBookFlaggedPropertyValue.Flag != 0x1)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2027");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2027
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    addressBookFlaggedPropertyValue,
                    typeof(AddressBookPropertyValue),
                    2027,
                    @"[In AddressBookFlaggedPropertyValue Structure] PropertyValue (optional) (variable): An AddressBookPropertyValue structure, as specified in section 2.2.1.1, unless the Flag field is set to 0x1.");
            }
        }
        #endregion

        #region Verify AddressBookPropertyRow Structure
        /// <summary>
        /// Verify the AddressBookPropertyRow structure related requirements.
        /// </summary>
        /// <param name="addressBookPropertyRow">The AddressBookPropertyRow value to be verified.</param>
        private void VerifyAddressBookPropertyRowStructure(AddressBookPropertyRow addressBookPropertyRow)
        {
            // Since Test Suite parsed AddressBookPropertyRow structure according to Specification, if the program run to here, R15 can be verified directly.
            this.Site.CaptureRequirement(
                15,
                @"[In AddressBookPropertyRow Structure] The AddressBookPropertyRow structure a list of property values without including the property tags that correspond to the property values.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R17");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R17
            this.Site.CaptureRequirementIfIsInstanceOfType(
                addressBookPropertyRow.Flag,
                typeof(byte),
                17,
                @"[In AddressBookPropertyRow Structure] Flags (1 byte): An unsigned integer that indicates whether all property values are present and without error in the ValueArray field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2035");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2035
            this.Site.CaptureRequirementIfIsInstanceOfType(
                addressBookPropertyRow.ValueArray,
                typeof(AddressBookPropertyValue[]),
                2035,
                @"[In AddressBookPropertyRow Structure] ValueArray (variable): An array of variable-sized structures.");
        }
        #endregion

        #region Verify LargePropTagArray structure
        /// <summary>
        /// Verify the LargePropTagArray structure related requirements.
        /// </summary>
        /// <param name="largePropTagArray">The LargePropTagArray value to be verified.</param>
        private void VerifyLargePropertyTagArrayStructure(LargePropertyTagArray largePropTagArray)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R24");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R24
            this.Site.CaptureRequirementIfIsInstanceOfType(
                largePropTagArray.PropertyTagCount,
                typeof(uint),
                24,
                @"[In LargePropertyTagArray Structure] PropertyTagCount (4 bytes): An unsigned integer that specifies the number of structures contained in the PropertyTags field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R25, the value of PropertyTagCount is {0}.", largePropTagArray.PropertyTagCount);
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R25
            this.Site.CaptureRequirementIfIsTrue(
                largePropTagArray.PropertyTagCount <= 100000,
                25,
                @"[In LargePropertyTagArray Structure] [PropertyTagCount (4 bytes)] The number is limited to 100,000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R26");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R26
            this.Site.CaptureRequirementIfIsInstanceOfType(
                largePropTagArray.PropertyTags,
                typeof(PropertyTag[]),
                26,
                @"[In LargePropertyTagArray Structure] PropertyTags (variable): An array of PropertyTag structures ([MS-OXCDATA] section 2.9), each of which contains a property tag that specifies a property.");
        
            PropertyTag[] propertytags = largePropTagArray.PropertyTags;
            foreach (PropertyTag propertyTag in propertytags)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R181");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCDATA_R181
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    propertyTag.PropertyType,
                    typeof(ushort),
                    "MS-OXCDATA",
                    181,
                    @"[In PropertyTag Structure] PropertyType (2 bytes): An unsigned integer that identifies the data type of the property value, as specified by the table in section 2.11.1.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R182");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCDATA_R182
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    propertyTag.PropertyId,
                    typeof(ushort),
                    "MS-OXCDATA",
                    182,
                    @"[In PropertyTag Structure] PropertyId (2 bytes): An unsigned integer that identifies the property.");
            }

            // Since Test Suite parsed largePropTagArray structure according to Specification, and the PropertyTag has been walked in the upper step,
            // so if the program run to here, R23 can be verified directly.
            this.Site.CaptureRequirement(
                23,
                @"[In LargePropertyTagArray Structure] The LargePropertyTagArray structure contains a list of property tags.");
        }
        #endregion

        #region Verify responding to all request types
        /// <summary>
        /// Verify requirements related to the responding of all request type.
        /// </summary>
        /// <param name="response">The HTTP response returned from server.</param>
        /// <param name="commonResponse">The response format parsed by TestSuite.</param>
        /// <param name="responseCode">The value of X-ResponseCode.</param>
        private void VerifyRespondingToAllRequestTypeRequests(HttpWebResponse response, CommonResponse commonResponse, uint responseCode)
        {
            List<string> headerValus = response.Headers.AllKeys.ToList<string>();

            if (!string.IsNullOrEmpty(response.Headers["X-ResponseCode"]) && responseCode == 0)
            {
                // If the response includes the Transfer-Encoding header, then the response is a chunked response.
                if (!string.IsNullOrEmpty(response.Headers["Transfer-Encoding"]))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1229");

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1229
                    // If the response is not null, that means the server can return response to corresponding request type requests. 
                    this.Site.CaptureRequirementIfIsNotNull(
                        response,
                        1229,
                        @"[In Responding to All Request Type Requests] The server can respond to all request type requests with a chunked response, as specified in section 2.2.2.2.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1232");

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1232
                    // If the code can reach here, Transfer-Encoding header is exist. 
                    // If Content-Length header is null, that means the Transfer-Encoding header is used instead of the Content-Length header field.
                    this.Site.CaptureRequirementIfIsNull(
                        response.Headers["Content-Length"],
                        1232,
                        @"[In Responding to All Request Type Requests] The Transfer-Encoding header is used instead of the Content-Length header field as specified in section 2.2.3.2.1.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1238");

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1238
                    this.Site.CaptureRequirementIfAreEqual<string>(
                        "PROCESSING",
                        commonResponse.MetaTags[0],
                        1238,
                        @"[In Responding to All Request Type Requests] The initial response includes the PROCESSING meta-tag, as specified in section 2.2.7.");

                    // The structure of response body is parsed according to the request of this Specification, if the code run to here and passed, 
                    // that means R41 is invoked successfully.
                    this.Site.CaptureRequirement(
                        41,
                        @"[In Common Response Format] The common server response across all endpoints (4) used in this protocol has the following formats, depending on whether the response is chunked.
Chunked response:
 HTTP/1.1 200 OK
 Transfer-Encoding: chunked
 Content-Type: application/mapi-http
 X-RequestType: <?>
 X-ResponseCode: <?>
 X-RequestId: <?>
 X-ServerApplication:<server version>
 
 <META-TAGS>
 <ADDITIONAL HEADERS>
 <RESPONSE BODY>");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1300");

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1300
                    this.Site.CaptureRequirementIfIsTrue(
                        headerValus.Contains("Transfer-Encoding"),
                        1300,
                        @"[In Common Response Format] In addition to headers [Host, Content-Type, X-RequestType and X-ResponseCode], a chunked response contains the Transfer-Encoding header, specified in section 2.2.3.2.5.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2231");

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2231
                    // The structure of response body is parsed according to the request of this Specification, if the code run to here and passed, 
                    // that means R2231 is invoked successfully.
                    this.Site.CaptureRequirement(
                        2231,
                        @"[In Responding to All Request Type Requests] The request type response body for the initiating request type will immediately follow any additional headers preceded by a CRLF on an empty line.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2228");

                    // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R2228
                    // The structure of response body is parsed according to the request of this Specification, if the code run to here and passed, 
                    // that means R2228 is invoked successfully.
                    this.Site.CaptureRequirement(
                        2228,
                        @"[In Responding to All Request Type Requests] The X-ElapsedTime and X-StartTime are present after the DONE meta-tag, as specified in section 2.2.3.3.9 and 2.2.3.3.10 respectively.");

                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R42");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R42
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "POST",
                    response.Method,
                    42,
                    @"[In Common Response Format] The first line of all server responses begins with the POST method response specified in [RFC2616], ""HTTP/1.1"".");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R43");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R43
                // If StatusCode is not null, that means server response contains an HTTP status code.
                this.Site.CaptureRequirementIfIsNotNull(
                    response.StatusCode,
                    43,
                    @"[In Common Response Format] It [The first line of all server responses] also contains an HTTP status code, as described in this section.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R54");

                // Since all the succeed and failure situation is judged according to X-ResponseCode header, if the code executes to here, R54 is verified successfully.
                this.Site.CaptureRequirement(
                    54,
                    @"[In Common Response Format] This header [X-ResponseCode] contains a numerical value that indicates the specific failure that occurred on the server.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1301");

                // If the code execute to here, that means the server returned an X-ResponseCode header for success condition.
                this.Site.CaptureRequirement(
                    1301,
                    @"[In Common Response Format] For success condition, the server will return an X-ResponseCode header.");
            }
            else if (!string.IsNullOrEmpty(response.Headers["X-ResponseCode"]) && responseCode != 0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1459");

                // If the code execute to here, that means the server returned an X-ResponseCode header for non-exceptional condition (most failures).
                this.Site.CaptureRequirement(
                    1459,
                    @"[In Common Response Format] For non-exceptional condition (most failures), the server will return an X-ResponseCode header.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2229");

                // If the code execute to here, that means the server returned an X-ResponseCode header if failure occurs.
                this.Site.CaptureRequirement(
                    2229,
                    @"[In Responding to All Request Type Requests] The server will include a X-ResponseCode header, as specified in section 2.2.3.3.3, if a failure occurs after the initial response is sent.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2229");

                // If the code execute to here, that means the server returned an X-ResponseCode header if failure occurs.
                this.Site.CaptureRequirement(
                    2229,
                    @"[In Responding to All Request Type Requests] The server will include a X-ResponseCode header, as specified in section 2.2.3.3.3, if a failure occurs after the initial response is sent.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R2230");

                // If the code execute to here, that means the server returned an X-ResponseCode header and the value is not 0 when failure occurs.
                this.Site.CaptureRequirement(
                    2230,
                    @"[In Responding to All Request Type Requests] More specifically, this gives the server the ability to later fail the request and return a different value in the X-ResponseCode header.");
            }
        }

        #endregion

        #region Verify Request Types For Mailbox Server Endpoint
        /// <summary>
        /// Verify the Request Types For Mailbox Server Endpoint related requirements.
        /// </summary>
        /// <param name="headers">The headers to be verified.</param>
        /// <param name="commonResponse">The CommonResponse to be verified.</param>
        private void VerifyRequestTypesForMailboxServerEndpoint(WebHeaderCollection headers, CommonResponse commonResponse)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R176, the value of X-RequestType header is {0}", headers["X-RequestType"]);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R176
            this.Site.CaptureRequirementIfIsTrue(
                this.requestTypeList.Contains(headers["X-RequestType"]),
                176,
                @"[In Request Types for Mailbox Server Endpoint] The X-RequestType header, specified in section 2.2.3.3.1, identifies which request type is being used.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1431");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1431
            // This structure of response body is parsed according to the request of this Specification, if the response body exists, R1431 is verified successfully.
            this.Site.CaptureRequirementIfIsNotNull(
                commonResponse,
                1431,
                @"[In Request Types for Mailbox Server Endpoint] The response body associated with the specific request types are each a raw binary data binary large object (BLOB) that follows the common response format, as specified in section 2.2.2.2.");

            // This structure of common response is parsed according to the request of this Specification, if the code run to here and passed, 
            // that means R1461 is invoked successfully.
            this.Site.CaptureRequirement(
                1461,
                @"[In Request Types for Mailbox Server Endpoint] Response body is separated from the common response by a blank line, as specified in [RFC2616].");
        }
        /// <summary>
        /// Verify the NotificationWait Request Types related requirements.
        /// </summary>
        /// <param name="headers">The headers to be verified.</param>
        private void VerifyNotificationWaitRequestType(WebHeaderCollection headers)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1261");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1261
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(headers["Set-Cookie"]),
                1261,
                @"[In Responding to a NotificationWait Request Type Request] The response headers include Set-Cookie headers as specified in section 2.2.3.2.4 for all cookies related to the Session Context.");

        }
        #endregion

        #region Verify each request type for mailbox server endpoint
        #region Verify Connect or Bind response
        /// <summary>
        /// Verify the Connect or Bind response related requirements.
        /// </summary>
        /// <param name="headers">The headers to be verified.</param>
        private void VerifyConnectOrBindResponse(WebHeaderCollection headers)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1146");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1146
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(headers["X-ClientInfo"]),
                1146,
                @"[In Creating a Session Context by Using the Connect or Bind Request Type] This information [X-ClientInfo header] is simply returned to the client in the response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1221");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1221
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(headers["Set-Cookie"]),
                1221,
                @"[In Responding to a Connect or Bind Request Type Request] The server MUST return the cookie that represents the Session Context as the value of the Set-Cookie header field, as specified in section 2.2.3.2.3.");
        }

        #endregion

        #region Verify Connect response
        /// <summary>
        /// Verify the Connect response related requirements.
        /// </summary>
        /// <param name="httpWebResponse">The HttpWebResponse to be verified.</param>
        private void VerifyConnectResponse(HttpWebResponse httpWebResponse)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1218");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1218
            // If the http web response is not null, that means the server issues a response to a connect request type.
            this.Site.CaptureRequirementIfIsNotNull(
                httpWebResponse,
                1218,
                @"[In Responding to a Connect or Bind Request Type Request] The server issues a response, as specified in section 2.2.2.2, to a Connect request type request, as specified in section 2.2.4.1.1.");
        }
        #endregion

        #region Verify Connect success response body
        /// <summary>
        /// Verify the Connect success response body related requirements.
        /// </summary>
        /// <param name="connectSuccessResponseBody">The Connect success response body to be verified.</param>
        private void VerifyConnectSuccessResponseBody(ConnectSuccessResponseBody connectSuccessResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R209");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R209
            this.Site.CaptureRequirementIfIsInstanceOfType(
                connectSuccessResponseBody.StatusCode,
                typeof(uint),
                209,
                @"[In Connect Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1349");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1349
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                connectSuccessResponseBody.StatusCode,
                1349,
                @"[In Connect Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R210");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R210
            this.Site.CaptureRequirementIfIsInstanceOfType(
                connectSuccessResponseBody.ErrorCode,
                typeof(uint),
                210,
                @"[In Connect Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R211");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R211
            this.Site.CaptureRequirementIfIsInstanceOfType(
                connectSuccessResponseBody.PollsMax,
                typeof(uint),
                211,
                @"[In Connect Request Type Success Response Body] PollsMax (4 bytes): An unsigned integer that specifies the number of milliseconds for the maximum polling interval.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R212");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R212
            this.Site.CaptureRequirementIfIsInstanceOfType(
                connectSuccessResponseBody.RetryCount,
                typeof(uint),
                212,
                @"[In Connect Request Type Success Response Body] RetryCount (4 bytes): An unsigned integer that specifies the number of times to retry request types.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R214");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R214
            this.Site.CaptureRequirementIfIsInstanceOfType(
                connectSuccessResponseBody.RetryDelay,
                typeof(uint),
                214,
                @"[In Connect Request Type Success Response Body] RetryDelay (4 bytes): An unsigned integer that specifies the number of milliseconds for the client to wait before retrying a failed request type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1351");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1351
            this.Site.CaptureRequirementIfIsInstanceOfType(
                connectSuccessResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                1351,
                @"[In Connect Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R218");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R218
            this.Site.CaptureRequirementIfIsInstanceOfType(
                connectSuccessResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                218,
                @"[In Connect Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug, 
                "Verify MS-OXCMAPIHTTP_R1352, the length of AuxiliaryBuffer is {0}, the value of AuxiliaryBufferSize is {1}.", 
                connectSuccessResponseBody.AuxiliaryBuffer.Length, 
                connectSuccessResponseBody.AuxiliaryBufferSize);
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1352
            this.Site.CaptureRequirementIfAreEqual<uint>(
                connectSuccessResponseBody.AuxiliaryBufferSize,
                (uint)connectSuccessResponseBody.AuxiliaryBuffer.Length,
                1352,
                @"[In Connect Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify Execute success response body
        /// <summary>
        /// Verify the Execute success response body related requirements.
        /// </summary>
        /// <param name="executeSuccessResponseBody">The Execute success response body to be verified.</param>
        private void VerifyExecuteSuccessResponseBody(ExecuteSuccessResponseBody executeSuccessResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R250");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R250
            this.Site.CaptureRequirementIfIsInstanceOfType(
                executeSuccessResponseBody.StatusCode,
                typeof(uint),
                250,
                @"[In Execute Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1356");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1356
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                executeSuccessResponseBody.StatusCode,
                1356,
                @"[In Execute Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R251");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R251
            this.Site.CaptureRequirementIfIsInstanceOfType(
                executeSuccessResponseBody.ErrorCode,
                typeof(uint),
                251,
                @"[In Execute Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R245");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R245
            this.Site.CaptureRequirementIfIsInstanceOfType(
                executeSuccessResponseBody.Flags,
                typeof(uint),
                245,
                @"[In Execute Request Type Success Response Body] The Data type of Field Flags is DWORD.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1358");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1358
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                executeSuccessResponseBody.Flags,
                1358,
                @"[In Execute Request Type Success Response Body] [Flags] The server MUST set this field to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R253");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R253
            this.Site.CaptureRequirementIfIsInstanceOfType(
                executeSuccessResponseBody.RopBufferSize,
                typeof(uint),
                253,
                @"[In Execute Request Type Success Response Body] RopBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the RopBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R254");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R254
            this.Site.CaptureRequirementIfIsInstanceOfType(
                executeSuccessResponseBody.RopBuffer,
                typeof(byte[]),
                254,
                @"[In Execute Request Type Success Response Body] RopBuffer (variable): An array of bytes that constitute the ROP responses payload.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R1360, the length of RopBuffer is {0}, the value of RopBufferSize is {1}.",
                executeSuccessResponseBody.RopBuffer.Length,
                executeSuccessResponseBody.RopBufferSize);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1360
            this.Site.CaptureRequirementIfAreEqual<uint>(
                executeSuccessResponseBody.RopBufferSize,
                (uint)executeSuccessResponseBody.RopBuffer.Length,
                1360,
                @"[In Execute Request Type Success Response Body] [RopBuffer] The size of this field, in bytes, is specified by the RopBufferSize field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R255");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R255
            this.Site.CaptureRequirementIfIsInstanceOfType(
                executeSuccessResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                255,
                @"[In Execute Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R256");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R256
            this.Site.CaptureRequirementIfIsInstanceOfType(
                executeSuccessResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                256,
                @"[In Execute Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R1361, the length of AuxiliaryBuffer is {0}, the value of AuxiliaryBufferSize is {1}.",
                executeSuccessResponseBody.AuxiliaryBuffer.Length,
                executeSuccessResponseBody.AuxiliaryBufferSize);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1361
            this.Site.CaptureRequirementIfAreEqual<uint>(
                executeSuccessResponseBody.AuxiliaryBufferSize,
                (uint)executeSuccessResponseBody.AuxiliaryBuffer.Length,
                1361,
                @"[In Execute Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify Disconnect response
        /// <summary>
        /// Verify the Disconnect response related requirements.
        /// </summary>
        /// <param name="httpWebResponse">The HttpWebResponse to be verified.</param>
        private void VerifyDisconnectResponse(HttpWebResponse httpWebResponse)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1255");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1255
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(httpWebResponse.Headers["Cookie"]),
                1255,
                @"[In Responding to a Disconnect Request Type Request] All relevant Set-Cookie headers, as specified in section 2.2.3.2.3,  are included in the response even though the session context cookie has been invalidated.");
        }
        #endregion

        #region Verify Disconnect success response body
        /// <summary>
        /// Verify the Disconnect success response body related requirements.
        /// </summary>
        /// <param name="disconnectSuccessResponseBody">The Disconnect success response body to be verified.</param>
        private void VerifyDisconnectSuccessResponseBody(DisconnectSuccessResponseBody disconnectSuccessResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R281");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R281
            this.Site.CaptureRequirementIfIsInstanceOfType(
                disconnectSuccessResponseBody.StatusCode,
                typeof(uint),
                281,
                @"[In Disconnect Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1363");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1363
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                disconnectSuccessResponseBody.StatusCode,
                1363,
                @"[In Disconnect Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R288");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R288
            this.Site.CaptureRequirementIfIsInstanceOfType(
                disconnectSuccessResponseBody.ErrorCode,
                typeof(uint),
                288,
                @"[In Disconnect Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R289");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R289
            this.Site.CaptureRequirementIfIsInstanceOfType(
                disconnectSuccessResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                289,
                @"[In Disconnect Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R290");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R290
            this.Site.CaptureRequirementIfIsInstanceOfType(
                disconnectSuccessResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                290,
                @"[In Disconnect Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMAPIHTTP_R1365, the length of AuxiliaryBuffer is {0}, the value of AuxiliaryBufferSize is {1}.",
                disconnectSuccessResponseBody.AuxiliaryBuffer.Length,
                disconnectSuccessResponseBody.AuxiliaryBufferSize);

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1365
            this.Site.CaptureRequirementIfAreEqual<uint>(
                disconnectSuccessResponseBody.AuxiliaryBufferSize,
                (uint)disconnectSuccessResponseBody.AuxiliaryBuffer.Length,
                1365,
                @"[In Disconnect Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify NotificationWait success response body
        /// <summary>
        /// Verify the NotificationWait success response body related requirements.
        /// </summary>
        /// <param name="notificationWaitSuccessResponseBody">The NotificationWaitSuccessResponseBody to be verified.</param>
        private void VerifyNotificationWaitSuccessResponseBody(NotificationWaitSuccessResponseBody notificationWaitSuccessResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R283");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R283
            this.Site.CaptureRequirementIfIsInstanceOfType(
                notificationWaitSuccessResponseBody.StatusCode,
                typeof(uint),
                283,
                @"[In NotificationWait Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1369");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1369
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                notificationWaitSuccessResponseBody.StatusCode,
                1369,
                @"[In NotificationWait Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R297");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R297
            this.Site.CaptureRequirementIfIsInstanceOfType(
                notificationWaitSuccessResponseBody.ErrorCode,
                typeof(uint),
                297,
                @"[In NotificationWait Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R285");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R285
            this.Site.CaptureRequirementIfIsInstanceOfType(
                notificationWaitSuccessResponseBody.EventPending,
                typeof(uint),
                285,
                @"[In NotificationWait Request Type Success Response Body] EventPending (4 bytes): An unsigned integer that indicates whether an event is pending.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1373");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1373
            this.Site.CaptureRequirementIfIsInstanceOfType(
                notificationWaitSuccessResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                1373,
                @"[In NotificationWait Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R298");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R298
            this.Site.CaptureRequirementIfIsInstanceOfType(
                notificationWaitSuccessResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                298,
                @"[In NotificationWait Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1374");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1374
            this.Site.CaptureRequirementIfAreEqual<uint>(
                notificationWaitSuccessResponseBody.AuxiliaryBufferSize,
                (uint)notificationWaitSuccessResponseBody.AuxiliaryBuffer.Length,
                1374,
                @"[In NotificationWait Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #endregion

        #region Verify request types for address book server endpoint
        /// <summary>
        /// Verify the request types for address book server endpoint related requirements.
        /// </summary>
        /// <param name="headers">The headers to be verified.</param>
        /// <param name="commonResponse">The commonResponse to be verified.</param>
        private void VerifyRequestTypesForAddressBookServerEndpoint(WebHeaderCollection headers, CommonResponse commonResponse)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R321, the value of X-RequestType header is {0}", headers["X-RequestType"]);
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R321
            this.Site.CaptureRequirementIfIsTrue(
                this.requestTypeList.Contains(headers["X-RequestType"]),
                321,
                @"[In Request Types for Address Book Server Endpoint] The X-RequestType header, specified in section 2.2.3.3.1, identifies which request type is being used.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1463");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1463
            // This structure of response body is parsed according to the request of this Specification, if the response body exists, R1463 is verified successfully.
            this.Site.CaptureRequirementIfIsNotNull(
                commonResponse,
                1463,
                @"[In Request Types for Address Book Server Endpoint] The response body associated with the specific request types are each a raw binary data binary large object (BLOB) that follows the common response format, as specified in section 2.2.2.2.");

            // This structure of common response is parsed according to the request of this Specification, if the code run to here and passed, 
            // that means R1462 is invoked successfully.
            this.Site.CaptureRequirement(
                1462,
                @"[In Request Types for Address Book Server Endpoint] Response body is separated from the common response by a blank line, as specified in [RFC2616].");
        }

        #endregion

        #region Verify each request type for address book server endpoint
        #region Verify QueryRows request type response body
        /// <summary>
        /// Verify the QueryRows request type response body related requirements.
        /// </summary>
        /// <param name="queryRowsResponseBody">The QueryRowsResponseBody to be verified.</param>
        /// <param name="queryRowsRequestBody">The QueryRowsRequestBody to be verified.</param>
        private void VerifyQueryRowsResponseBody(QueryRowsResponseBody queryRowsResponseBody, QueryRowsRequestBody queryRowsRequestBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R843");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R843
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryRowsResponseBody.StatusCode,
                typeof(uint),
                843,
                @"[In QueryRows Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R844");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R844
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                queryRowsResponseBody.StatusCode,
                844,
                @"[In QueryRows Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R845");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R845
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryRowsResponseBody.ErrorCode,
                typeof(uint),
                845,
                @"[In QueryRows Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R846");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R846
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryRowsResponseBody.HasState,
                typeof(bool),
                846,
                @"[In QueryRows Request Type Success Response Body] HasState (1 byte): A Boolean value that specifies whether the State field is present.");

            if (queryRowsResponseBody.HasState)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R850");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R850
                this.Site.CaptureRequirementIfIsNotNull(
                    queryRowsResponseBody.State,
                    850,
                    @"[In QueryRows Request Type Success Response Body] [State] This field is present when the HasState field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R847");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R847
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    queryRowsResponseBody.State,
                    typeof(STAT),
                    847,
                    @"[In QueryRows Request Type Success Response Body] State (optional) (36bytes): A STAT structure ([MS-OXNSPI] section 2.3.7) that specifies the state of a specific address book container.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R851.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R851
                this.Site.CaptureRequirementIfIsNull(
                    queryRowsResponseBody.State,
                    851,
                    @"[In QueryRows Request Type Success Response Body] [State] This field is not present when the HasState field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R852");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R852
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryRowsResponseBody.HasColumnsAndRows,
                typeof(bool),
                852,
                @"[In QueryRows Request Type Success Response Body] HasColumnsAndRows (1 byte): A Boolean value that specifies whether the Columns, RowCount, and RowData fields are present.");

            if (queryRowsResponseBody.HasColumnsAndRows)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R853");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R853
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    queryRowsResponseBody.Columns,
                    typeof(LargePropertyTagArray),
                    853,
                    @"[In QueryRows Request Type Success Response Body] Columns (optional) (variable): A LargePropTagArray structure (section 2.2.1.3) that specifies the columns for the rows returned.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R854.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R854
                this.Site.CaptureRequirementIfIsNotNull(
                    queryRowsResponseBody.Columns,
                    854,
                    @"[In QueryRows Request Type Success Response Body] [Columns] This field is present when the value of the HasColumnsAndRows field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R856.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R856
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    queryRowsResponseBody.RowCount,
                    typeof(uint),
                    856,
                    @"[In QueryRows Request Type Success Response Body] RowCount (optional) (4 bytes): An unsigned integer that specifies the number of structures in the RowData field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R857.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R857
                this.Site.CaptureRequirementIfIsNotNull(
                    queryRowsResponseBody.RowCount,
                    857,
                    @"[In QueryRows Request Type Success Response Body] [RowCount] This field is present when the value of the HasColumnsAndRows field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R859");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R859
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    queryRowsResponseBody.RowData,
                    typeof(AddressBookPropertyRow[]),
                    859,
                    @"[In QueryRows Request Type Success Response Body] RowData (optional) (variable): An array of AddressBookPropertyRow structures (section 2.2.1.2), each of which specifies the row data of the Explicit Table.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R860.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R860
                this.Site.CaptureRequirementIfIsNotNull(
                    queryRowsResponseBody.RowData,
                    860,
                    @"[In QueryRows Request Type Success Response Body] [RowData] This field is present when the HasColumnsAndRows field is nonzero.");

                //// Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R825, the value of RowCount for queryRowsResponseBody is {0}.", queryRowsResponseBody.RowCount.Value);

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R825
                // If the row count server maintained is bigger than the requested row number, the server will return the number of rows it maintained, 
                // otherwise, server will return the rows according to the requested row count.
                this.Site.CaptureRequirementIfIsTrue(
                    queryRowsRequestBody.RowCount >= queryRowsResponseBody.RowCount.Value,
                    825,
                    @"[In QueryRows Request Type Request Body] RowCount (4 bytes): An unsigned integer that specifies the number of rows the client is requesting.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R855.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R855
                this.Site.CaptureRequirementIfIsNull(
                    queryRowsResponseBody.Columns,
                    855,
                    @"[In QueryRows Request Type Success Response Body] [Columns] This field is not present when the value of the HasColumnsAndRows field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R858.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R858
                this.Site.CaptureRequirementIfIsNull(
                    queryRowsResponseBody.RowCount,
                    858,
                    @"[In QueryRows Request Type Success Response Body] [RowCount] This field is not present when the value of the HasColumnsAndRows field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R861.");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R861
                this.Site.CaptureRequirementIfIsNull(
                    queryRowsResponseBody.RowData,
                    861,
                    @"[In QueryRows Request Type Success Response Body] [RowData] This field is not present when the HasColumnsAndRows field is zero.");
            }
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R862");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R862
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryRowsResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                862,
                @"[In QueryRows Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R863");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R863
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryRowsResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                863,
                @"[In QueryRows Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R864");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R864
            this.Site.CaptureRequirementIfAreEqual<uint>(
                queryRowsResponseBody.AuxiliaryBufferSize,
                (uint)queryRowsResponseBody.AuxiliaryBuffer.Length,
                864,
                @"[In QueryRows Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }

        #endregion

        #region Verify Bind response body

        /// <summary>
        /// Verify the Bind response body related requirements.
        /// </summary>
        /// <param name="bindResponseBody">The Bind response body to be verified.</param>
        private void VerifyBindResponseBody(BindResponseBody bindResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R346");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R346
            this.Site.CaptureRequirementIfIsInstanceOfType(
                bindResponseBody.StatusCode,
                typeof(uint),
                346,
                @"[In Bind Request Type Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R347");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R347
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                bindResponseBody.StatusCode,
                347,
                @"[In Bind Request Type Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R348");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R348
            this.Site.CaptureRequirementIfIsInstanceOfType(
                bindResponseBody.ErrorCode,
                typeof(uint),
                348,
                @"[In Bind Request Type Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R349");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R349
            this.Site.CaptureRequirementIfIsInstanceOfType(
                bindResponseBody.ServerGuid,
                typeof(Guid),
                349,
                @"[In Bind Request Type Response Body] ServerGuid (16 bytes): A GUID that is associated with a specific address book server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R350");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R350
            this.Site.CaptureRequirementIfIsInstanceOfType(
                bindResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                350,
                @"[In Bind Request Type Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R351");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R351
            this.Site.CaptureRequirementIfIsInstanceOfType(
                bindResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                351,
                @"[In Bind Request Type Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R352");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R352
            this.Site.CaptureRequirementIfAreEqual<uint>(
                bindResponseBody.AuxiliaryBufferSize,
                (uint)bindResponseBody.AuxiliaryBuffer.Length,
                352,
                @"[In Bind Request Type Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1446");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1446
            // If the response body is not null, it indicates that the server issues a response to a Bind request type.
            this.Site.CaptureRequirementIfIsNotNull(
                bindResponseBody,
                1446,
                @"[In Responding to a Connect or Bind Request Type Request] The server issues a response, as specified in section 2.2.2.2, to a Bind request type request, as specified in section 2.2.5.1.1.");
        }
        #endregion

        #region Verify Unbind response body

        /// <summary>
        /// Verify the Unbind response body related requirements.
        /// </summary>
        /// <param name="unbindResponseBody">The Unbind response body to be verified.</param>
        private void VerifyUnbindResponseBody(UnbindResponseBody unbindResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R367");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R367
            this.Site.CaptureRequirementIfIsInstanceOfType(
                unbindResponseBody.StatusCode,
                typeof(uint),
                367,
                @"[In Unbind Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1291");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1291
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                unbindResponseBody.StatusCode,
                1291,
                @"[In Unbind Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R368");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R368
            this.Site.CaptureRequirementIfIsInstanceOfType(
                unbindResponseBody.ErrorCode,
                typeof(uint),
                368,
                @"[In Unbind Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R369");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R369
            this.Site.CaptureRequirementIfIsInstanceOfType(
                unbindResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                369,
                @"[In Unbind Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R370");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R370
            this.Site.CaptureRequirementIfIsInstanceOfType(
                unbindResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                370,
                @"[In Unbind Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R371");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R371
            this.Site.CaptureRequirementIfAreEqual<uint>(
                unbindResponseBody.AuxiliaryBufferSize,
                (uint)unbindResponseBody.AuxiliaryBuffer.Length,
                371,
                @"[In Unbind Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify CompareMinIds response body
        /// <summary>
        /// Verify the ComapreMinIds response body related requirements.
        /// </summary>
        /// <param name="compareMinIdsResponseBody">The CompareMinIds response body to be verified.</param>
        private void VerifyComapreMinIdsResponsebody(CompareMinIdsResponseBody compareMinIdsResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R399");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R399
            this.Site.CaptureRequirementIfIsInstanceOfType(
                compareMinIdsResponseBody.StatusCode,
                typeof(uint),
                399,
                @"[In CompareMinIds Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1290");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1290
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                compareMinIdsResponseBody.StatusCode,
                1290,
                @"[In CompareMinIds Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R400");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R400
            this.Site.CaptureRequirementIfIsInstanceOfType(
                compareMinIdsResponseBody.ErrorCode,
                typeof(uint),
                400,
                @"[In CompareMinIds Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R401");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R401
            this.Site.CaptureRequirementIfIsInstanceOfType(
                compareMinIdsResponseBody.Result,
                typeof(int),
                401,
                @"[In CompareMinIds Request Type Success Response Body] Result (4 bytes): A signed integer that specifies the result of the comparison.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R402");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R402
            this.Site.CaptureRequirementIfIsInstanceOfType(
                compareMinIdsResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                402,
                @"[In CompareMinIds Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R403");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R403
            this.Site.CaptureRequirementIfIsInstanceOfType(
                compareMinIdsResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                403,
                @"[In CompareMinIds Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R404");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R404
            this.Site.CaptureRequirementIfAreEqual<uint>(
                compareMinIdsResponseBody.AuxiliaryBufferSize,
                (uint)compareMinIdsResponseBody.AuxiliaryBuffer.Length,
                404,
                @"[In CompareMinIds Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify DnToMinId response body

        /// <summary>
        /// Verify the DnToMinId response body related requirements.
        /// </summary>
        /// <param name="responseBodyDnToMinId">The DnToMinId response body to be verified.</param>
        private void VerifyDnToMinIdResponseBody(DnToMinIdResponseBody responseBodyDnToMinId)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1286");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1286
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBodyDnToMinId.StatusCode,
                typeof(uint),
                1286,
                @"[In DnToMinId Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1287");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1287
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseBodyDnToMinId.StatusCode,
                1287,
                @"[In DnToMinId Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R430");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R430
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBodyDnToMinId.ErrorCode,
                typeof(uint),
                430,
                @"[In DnToMinId Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R431");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R431
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBodyDnToMinId.HasMinimalIds,
                typeof(bool),
                431,
                @"[In DnToMinId Request Type Success Response Body] HasMinimalIds (1 byte): A Boolean value that specifies whether the MinimalIdCount and MinimalIds fields are present.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R432");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R432
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBodyDnToMinId.MinimalIdCount,
                typeof(uint),
                432,
                @"[In DnToMinId Request Type Success Response Body] MinimalIdCount (optional) (4 bytes): An unsigned integer that specifies the number of structures in the MinimalIds field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R440");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R440
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBodyDnToMinId.AuxiliaryBufferSize,
                typeof(uint),
                440,
                @"[In DnToMinId Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R441");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R441
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBodyDnToMinId.AuxiliaryBuffer,
                typeof(byte[]),
                441,
                @"[In DnToMinId Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R442");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R442
            this.Site.CaptureRequirementIfAreEqual<uint>(
                responseBodyDnToMinId.AuxiliaryBufferSize,
                (uint)responseBodyDnToMinId.AuxiliaryBuffer.Length,
                442,
                @"[In DnToMinId Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");

            if (responseBodyDnToMinId.HasMinimalIds)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R433");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R433
                // Because the HasMinimalIds field is true, so if the MinimalIdCount field has value, then R433 will be verified. 
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBodyDnToMinId.MinimalIdCount,
                    433,
                    @"[In DnToMinId Request Type Success Response Body] [MinimalIdCount] This field is present when the value of the HasMinimalIds field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1288");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1288
                // Because the HasMinimalIds field is true, so if the MinimalIds field has value, then R1288 will be verified.  
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBodyDnToMinId.MinimalIds,
                    1288,
                    @"[In DnToMinId Request Type Success Response Body] [MinimalIds] This field is present when the value of the HasMinimalIds field is nonzero.");
            }
        }
        #endregion

        #region Verify GetSpecialTable response body

        /// <summary>
        /// Verify the GetSpecialTable response body related requirements.
        /// </summary>
        /// <param name="getSpecialTableResponseBody">The GetSpecialTable response body to be verified.</param>
        private void VerifyGetSpecialTableResponseBody(GetSpecialTableResponseBody getSpecialTableResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R664");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R664
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getSpecialTableResponseBody.StatusCode,
                typeof(uint),
                664,
                @"[In GetSpecialTable Request Type  Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R665");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R665
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getSpecialTableResponseBody.StatusCode,
                665,
                @"[In GetSpecialTable Request Type  Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R666");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R666
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getSpecialTableResponseBody.ErrorCode,
                typeof(uint),
                666,
                @"[In GetSpecialTable Request Type  Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R667");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R667
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getSpecialTableResponseBody.CodePage,
                typeof(uint),
                667,
                @"[In GetSpecialTable Request Type  Success Response Body] CodePage (4 bytes): An unsigned integer that specifies the code page the server used to express string properties.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R668");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R668
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getSpecialTableResponseBody.HasVersion,
                typeof(bool),
                668,
                @"[In GetSpecialTable Request Type  Success Response Body] HasVersion (1 byte): A Boolean value that specifies whether the Version field is present.");

            if (getSpecialTableResponseBody.HasVersion)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R670");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R670
                this.Site.CaptureRequirementIfIsNotNull(
                    getSpecialTableResponseBody.Version,
                    670,
                    @"[In GetSpecialTable Request Type  Success Response Body] [Version] This field is present when the value of the HasVersion field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R669");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R669
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    getSpecialTableResponseBody.Version,
                    typeof(uint),
                    669,
                    @"[In GetSpecialTable Request Type  Success Response Body] Version (optional) (4 bytes): An unsigned integer that specifies the version number of the address book hierarchy table that the server has.");
            }
                        
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R672");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R672
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getSpecialTableResponseBody.HasRows,
                typeof(bool),
                672,
                @"[In GetSpecialTable Request Type  Success Response Body] HasRows (1 byte): A Boolean value that specifies whether the RowCount and Rows fields are present.");

            if (getSpecialTableResponseBody.HasRows)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R675");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R675
                this.Site.CaptureRequirementIfIsNotNull(
                    getSpecialTableResponseBody.RowCount,
                    675,
                    @"[In GetSpecialTable Request Type  Success Response Body] [RowsCount] This field is present when the value of the HasRows field is nonzero.");
                
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R673");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R673
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    getSpecialTableResponseBody.RowCount,
                    typeof(uint),
                    673,
                    @"[In GetSpecialTable Request Type  Success Response Body] RowsCount (optional) (4 bytes): An unsigned integer that specifies the number of structures in the Rows field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R678");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R678
                this.Site.CaptureRequirementIfIsNotNull(
                    getSpecialTableResponseBody.Rows,
                    678,
                    @"[In GetSpecialTable Request Type  Success Response Body] [Rows] This field is present when the value of the HasRows field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1452");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1452
                this.Site.CaptureRequirementIfIsNull(
                    getSpecialTableResponseBody.RowCount,
                    1452,
                    @"[In GetSpecialTable Request Type  Success Response Body] [RowsCount] This field is not present when the value of the HasRows field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R679");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R679
                this.Site.CaptureRequirementIfIsNull(
                    getSpecialTableResponseBody.Rows,
                    679,
                    @"[In GetSpecialTable Request Type  Success Response Body] [Rows] This field is not present when the value of the HasRows field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R680");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R680
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getSpecialTableResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                680,
                @"[In GetSpecialTable Request Type  Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R681");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R681
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getSpecialTableResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                681,
                @"[In GetSpecialTable Request Type  Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R682");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R682
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getSpecialTableResponseBody.AuxiliaryBufferSize,
                (uint)getSpecialTableResponseBody.AuxiliaryBuffer.Length,
                682,
                @"[In GetSpecialTable Request Type  Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify GetTemplateInfo response body

        /// <summary>
        /// Verify the GetTemplateInfo response body related requirements.
        /// </summary>
        /// <param name="getTemplateInfoResponseBody">The GetTemplateInfo response body to be verified.</param>
        private void VerifyGetTemplateInfoResponseBody(GetTemplateInfoResponseBody getTemplateInfoResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R713");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R713
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getTemplateInfoResponseBody.StatusCode,
                typeof(uint),
                713,
                @"[In GetTemplateInfo Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R714");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R714
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getTemplateInfoResponseBody.StatusCode,
                714,
                @"[In GetTemplateInfo Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R715");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R715
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getTemplateInfoResponseBody.ErrorCode,
                typeof(uint),
                715,
                @"[In GetTemplateInfo Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R716");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R716
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getTemplateInfoResponseBody.CodePage,
                typeof(uint),
                716,
                @"[In GetTemplateInfo Request Type Success Response Body] CodePage (4 bytes): An unsigned integer that specifies the code page the server used to express string values of properties.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R717");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R717
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getTemplateInfoResponseBody.HasRow,
                typeof(bool),
                717,
                @"[In GetTemplateInfo Request Type Success Response Body] HasRow (1 byte): A Boolean value that specifies whether the Rows field is present.");

            if (getTemplateInfoResponseBody.HasRow)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R719");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R719
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    getTemplateInfoResponseBody.Row,
                    typeof(AddressBookPropertyValueList),
                    719,
                    @"[In GetTemplateInfo Request Type Success Response Body] [Row] This field is present when the value of the HasRow field is nonzero.");

                for (int i = 0; i < getTemplateInfoResponseBody.Row.Value.PropertyValueCount; i++)
                {
                    this.VerifyAddressBookTaggedPropertyValueStructure(getTemplateInfoResponseBody.Row.Value.PropertyValues[i]);
                }
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R720");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R720
                this.Site.CaptureRequirementIfIsNull(
                    getTemplateInfoResponseBody.Row,
                    720,
                    @"[In GetTemplateInfo Request Type Success Response Body] [Row] This field is not present when the value of the HasRow field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R721");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R721
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getTemplateInfoResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                721,
                @"[In GetTemplateInfo Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R722");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R722
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getTemplateInfoResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                722,
                @"[In GetTemplateInfo Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R723");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R723
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getTemplateInfoResponseBody.AuxiliaryBufferSize,
                (uint)getTemplateInfoResponseBody.AuxiliaryBuffer.Length,
                723,
                @"[In GetTemplateInfo Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify the ModLinkAttResponseBody response body

        /// <summary>
        ///  Verify the ModLinkAtt response body related requirements.
        /// </summary>
        /// <param name="modLinkAttResponseBody">The ModLinkAtt response body to be verified.</param>
        private void VerifyModLinkAttResponseBody(ModLinkAttResponseBody modLinkAttResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R757");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R757
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modLinkAttResponseBody.StatusCode,
                typeof(uint),
                757,
                @"[In ModLinkAtt Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R758");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R758
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                modLinkAttResponseBody.StatusCode,
                758,
                @"[In ModLinkAtt Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R759");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R759
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modLinkAttResponseBody.ErrorCode,
                typeof(uint),
                759,
                @"[In ModLinkAtt Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R760");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R760
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modLinkAttResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                760,
                @"[In ModLinkAtt Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R761");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R761
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modLinkAttResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                761,
                @"[In ModLinkAtt Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R762");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R762
            this.Site.CaptureRequirementIfAreEqual<uint>(
                modLinkAttResponseBody.AuxiliaryBufferSize,
                (uint)modLinkAttResponseBody.AuxiliaryBuffer.Length,
                762,
                @"[In ModLinkAtt Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify ResortRestriction response body
        /// <summary>
        /// Verify the ResortRestriction response body related requirements.
        /// </summary>
        /// <param name="resortRestrictionResponseBody">The ResortRestriction response body to be verified.</param>
        private void VerifyResortRestrictionResponseBody(ResortRestrictionResponseBody resortRestrictionResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R993");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R993
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resortRestrictionResponseBody.StatusCode,
                typeof(uint),
                993,
                @"[In ResortRestriction Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R994");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R994
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                resortRestrictionResponseBody.StatusCode,
                994,
                @"[In ResortRestriction Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R995");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R995
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resortRestrictionResponseBody.ErrorCode,
                typeof(uint),
                995,
                @"[In ResortRestriction Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R996");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R996
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resortRestrictionResponseBody.HasState,
                typeof(bool),
                996,
                @"[In ResortRestriction Request Type RSuccess esponse Body] HasState (1 byte): A Boolean value that specifies whether the State field is present.");

            if (resortRestrictionResponseBody.HasState)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R997");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R997
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resortRestrictionResponseBody.State,
                    typeof(STAT),
                    997,
                    @"[In ResortRestriction Request Type Success Response Body] State (optional) (36 butes): A STAT structure ([MS-OXNSPI] section 2.3.7) that specifies the state of a specific address book container.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R999");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R999
                this.Site.CaptureRequirementIfIsNotNull(
                    resortRestrictionResponseBody.State,
                    999,
                    @"[In ResortRestriction Request Type Success Response Body] [State] This field is present when the HasState field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1000");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1000
                this.Site.CaptureRequirementIfIsNull(
                    resortRestrictionResponseBody.State,
                    1000,
                    @"[In ResortRestriction Request Type Success Response Body] [State] This field is not present when the HasState field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1001");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1001
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resortRestrictionResponseBody.HasMinimalIds,
                typeof(bool),
                1001,
                @"[In ResortRestriction Request Type Success Response Body] HasMinimalIds (1 byte): A Boolean value that specifies whether the MinimalIdCount and MinimalIds fields are present.");

            if (resortRestrictionResponseBody.HasMinimalIds)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1002");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1002
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resortRestrictionResponseBody.MinimalIdCount,
                    typeof(uint),
                    1002,
                    @"[In ResortRestriction Request Type Success Response Body] MinimalIdCount (optional) (4 bytes): An unsigned integer that specifies the number of structures present in the Minimalids field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1005");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1005
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resortRestrictionResponseBody.MinimalIds,
                    typeof(uint[]),
                    1005,
                    @"[In ResortRestriction Request Type Success Response Body] MinimalIds (optional) (variable): An array of MinimalEntryID structures ([MS-OXNSPI] section 2.3.8.1) that compose a restricted address book container.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1006");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1006
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    resortRestrictionResponseBody.MinimalIdCount.Value,
                    (uint)resortRestrictionResponseBody.MinimalIds.Length,
                    1006,
                    @"[In ResortRestriction Request Type Success Response Body] [MinimalIds] The number of structures contained in this field is specified by the MinimalIdCount field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1003");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1003
                this.Site.CaptureRequirementIfIsNotNull(
                    resortRestrictionResponseBody.MinimalIdCount,
                    1003,
                    @"[In ResortRestriction Request Type Success Response Body] [MinimalIdCount] This field is present when the value of the HasMinimalIds field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1007");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1007
                this.Site.CaptureRequirementIfIsNotNull(
                    resortRestrictionResponseBody.MinimalIds,
                    1007,
                    @"[In ResortRestriction Request Type RSuccess esponse Body] [MinimalIds] This field is present when the value of the HasMinimalIds field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1004");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1004
                this.Site.CaptureRequirementIfIsNull(
                    resortRestrictionResponseBody.MinimalIdCount,
                    1004,
                    @"[In ResortRestriction Request Type Success Response Body] [MinimalIdCount] This field is not present when the value of the HasMinimalIds field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1008");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1008
                this.Site.CaptureRequirementIfIsNull(
                    resortRestrictionResponseBody.MinimalIds,
                    1008,
                    @"[In ResortRestriction Request Type RSuccess esponse Body] [MinimalIds] This field is not present when the value of the HasMinimalIds field is zero.");
            }
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1009");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1009
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resortRestrictionResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                1009,
                @"[In ResortRestriction Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1010");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1010
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resortRestrictionResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                1010,
                @"[In ResortRestriction Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1011");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1011
            this.Site.CaptureRequirementIfAreEqual<uint>(
                resortRestrictionResponseBody.AuxiliaryBufferSize,
                (uint)resortRestrictionResponseBody.AuxiliaryBuffer.Length,
                1011,
                @"[In ResortRestriction Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify UpdateStat response body

        /// <summary>
        ///  Verify the UpdateStat response body related requirements.
        /// </summary>
        /// <param name="updateStatResponseBody">The UpdateStat response body to be verified.</param>
        private void VerifyUpdateStatResponseBody(UpdateStatResponseBody updateStatResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1074");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1074
            this.Site.CaptureRequirementIfIsInstanceOfType(
                updateStatResponseBody.StatusCode,
                typeof(uint),
                1074,
                @"[In UpdateStat Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1075");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1075
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                updateStatResponseBody.StatusCode,
                1075,
                @"[In UpdateStat Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1076");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1076
            this.Site.CaptureRequirementIfIsInstanceOfType(
                updateStatResponseBody.ErrorCode,
                typeof(uint),
                1076,
                @"[In UpdateStat Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1077");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1077
            this.Site.CaptureRequirementIfIsInstanceOfType(
                updateStatResponseBody.HasState,
                typeof(bool),
                1077,
                @"[In UpdateStat Request Type Success Response Body] HasState (1 byte): A Boolean value that specifies whether the State field is present.");

            if (updateStatResponseBody.HasState)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1078");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1078
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    updateStatResponseBody.State,
                    typeof(STAT),
                    1078,
                    @"[In UpdateStat Request Type Success Response Body] State (optional) (36 bytes): A STAT structure ([MS-OXNSPI] section 2.3.7) that specifies the state of a specific address book container.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1080");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1080
                this.Site.CaptureRequirementIfIsNotNull(
                    updateStatResponseBody.State,
                    1080,
                    @"[In UpdateStat Request Type Success Response Body] [State] This field is present when the HasState field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1081");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1081
                this.Site.CaptureRequirementIfIsNull(
                    updateStatResponseBody.State,
                    1081,
                    @"[In UpdateStat Request Type Success Response Body] [State] This field is not present when the HasState field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1082");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1082
            this.Site.CaptureRequirementIfIsInstanceOfType(
                updateStatResponseBody.HasDelta,
                typeof(bool),
                1082,
                @"[In UpdateStat Request Type Success Response Body] HasDelta (1 byte): A Boolean value that specifies whether the Delta field is present.");

            if (updateStatResponseBody.HasDelta)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1084");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1084
                this.Site.CaptureRequirementIfIsNotNull(
                    updateStatResponseBody.Delta,
                    1084,
                    @"[In UpdateStat Request Type Success Response Body] [Delta] This field is present when the value of the HasDelta field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1085");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1085
                this.Site.CaptureRequirementIfIsNull(
                    updateStatResponseBody.Delta,
                    1085,
                    @"[In UpdateStat Request Type Success Response Body] [Delta] This field is not present when the value of the HasDelta field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1086");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1086
            this.Site.CaptureRequirementIfIsInstanceOfType(
                updateStatResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                1086,
                @"[In UpdateStat Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1087");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1087
            this.Site.CaptureRequirementIfIsInstanceOfType(
                updateStatResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                1087,
                @"[In UpdateStat Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1088");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1088
            this.Site.CaptureRequirementIfAreEqual<uint>(
                updateStatResponseBody.AuxiliaryBufferSize,
                (uint)updateStatResponseBody.AuxiliaryBuffer.Length,
                1088,
                @"[In UpdateStat Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }

        #endregion

        #region Verify GetMailboxUrl response body
        /// <summary>
        ///  Verify the GetMailboxUrl response body related requirements.
        /// </summary>
        /// <param name="getMailboxUrlResponseBody">The GetMailboxUrl response body to be verified.</param>
        private void VerifyGetMailboxUrlResponseBody(GetMailboxUrlResponseBody getMailboxUrlResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1099");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1099
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getMailboxUrlResponseBody.StatusCode,
                typeof(uint),
                1099,
                @"[In GetMailboxUrl Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1100");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1100
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getMailboxUrlResponseBody.StatusCode,
                1100,
                @"[In GetMailboxUrl Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1101");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1101
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getMailboxUrlResponseBody.ErrorCode,
                typeof(uint),
                1101,
                @"[In GetMailboxUrl Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1102");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1102
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getMailboxUrlResponseBody.ServerUrl,
                typeof(string),
                1102,
                @"[In GetMailboxUrl Request Type Success Response Body] ServerUrl (variable): A null-terminated Unicode string that specifies URL of the EMSMDB server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1103");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1103
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getMailboxUrlResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                1103,
                @"[In GetMailboxUrl Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1104");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1104
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getMailboxUrlResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                1104,
                @"[In GetMailboxUrl Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1105");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1105
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getMailboxUrlResponseBody.AuxiliaryBufferSize,
                (uint)getMailboxUrlResponseBody.AuxiliaryBuffer.Length,
                1105,
                @"[In GetMailboxUrl Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify GetAddressBookUrl response body
        /// <summary>
        /// Verify GetAddressBookUrl response body related requirements
        /// </summary>
        /// <param name="getAddressBookUrlResponseBody">The GetAddressBookUrl response body to be verified.</param>
        private void VerifyGetAddressBookUrlResponseBody(GetAddressBookUrlResponseBody getAddressBookUrlResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1117");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1117
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getAddressBookUrlResponseBody.StatusCode,
                typeof(uint),
                1117,
                @"[In GetAddressBookUrl Request Type  Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1118");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1118
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getAddressBookUrlResponseBody.StatusCode,
                1118,
                @"[In GetAddressBookUrl Request Type  Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1119");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1119
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getAddressBookUrlResponseBody.ErrorCode,
                typeof(uint),
                1119,
                @"[In GetAddressBookUrl Request Type  Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1120");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1120
            this.Site.CaptureRequirementIfAreEqual<string>(
                getAddressBookUrlResponseBody.ServerUrl,
                this.addressBookUrl + "\0",
                1120,
                @"[In GetAddressBookUrl Request Type  Success Response Body] ServerUrl (variable): A null-terminated Unicode string that specifies the URL of the NSPI server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1121");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1121
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getAddressBookUrlResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                1121,
                @"[In GetAddressBookUrl Request Type  Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1122");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1122
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getAddressBookUrlResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                1122,
                @"[In GetAddressBookUrl Request Type  Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1123");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1123
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getAddressBookUrlResponseBody.AuxiliaryBufferSize,
                (uint)getAddressBookUrlResponseBody.AuxiliaryBuffer.Length,
                1123,
                @"[In GetAddressBookUrl Request Type  Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }

        #endregion

        #region Verify GetMatches request type response body
        /// <summary>
        /// Verify the requirements related to GetMatches request type response body. 
        /// </summary>
        /// <param name="requesutBody">The request body of GetMatches request type.</param>
        /// <param name="responseBody">The response body of GetMatches request type.</param>
        private void VerifyGetMatchsResponseBody(GetMatchesRequestBody requesutBody, GetMatchesResponseBody responseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R524");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R524
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.StatusCode,
                typeof(uint),
                524,
                @"[In GetMatches Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R525");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R525
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                responseBody.StatusCode,
                525,
                @"[In GetMatches Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R526");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R526
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.ErrorCode,
                typeof(uint),
                526,
                @"[In GetMatches Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R527");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R527
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.HasState,
                typeof(bool),
                527,
                @"[In GetMatches Request Type Success Response Body] HasState (1 byte): A Boolean value that specifies whether the State field is present.");

            if (responseBody.HasState)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R528");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R528
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.State.Value,
                    typeof(STAT),
                    528,
                    @"[In GetMatches Request Type Success Response Body] State (optional) (36 bytes): A STAT structure ([MS-OXNSPI] section 2.3.7) that specifies the state of a specific address book container.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R530");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R530
                // Because the HasState field is true, so if the State has value, then R530 will be verified.
                this.Site.CaptureRequirementIfIsTrue(
                    responseBody.State.HasValue,
                    530,
                    @"[In GetMatches Request Type Success Response Body] [State] This field is present when the HasState field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1449");

                // Because the HasState field is false, so if the State does not have a value, then R1449 will be verified.
                this.Site.CaptureRequirementIfIsFalse(
                    responseBody.State.HasValue,
                    1449,
                    @"[In GetMatches Request Type Success Response Body] [State] This field is not present when the HasState field is zero.");
            }
           
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R533");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R533
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.HasMinimalIds,
                typeof(bool),
                533,
                @"[In GetMatches Request Type Success Response Body] HasMinimalIds (1 byte): A Boolean value that specifies whether the MinimalIdCount and MinimalIds fields are present.");

            if (responseBody.HasMinimalIds)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R516");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R516
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.MinimalIdCount,
                    typeof(uint),
                    516,
                    @"[In GetMatches Request Type Success Response Body] The Date type of Field MinimalIdCount is (4 bytes) unsigned integer (optional).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R534");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R534
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)responseBody.MinimalIds.Length,
                    responseBody.MinimalIdCount.Value,
                    534,
                    @"[In GetMatches Request Type Success Response Body] MinimalIdCount: An unsigned integer that specifies the number of structures present in the MinimalIds field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R535. The MinimalIdCount is {0}", responseBody.MinimalIdCount.Value);

                this.Site.CaptureRequirementIfIsTrue(
                    responseBody.MinimalIdCount.HasValue,
                    535,
                    @"[In GetMatches Request Type Success Response Body] [MinimalIdCount] This field is present when the value of the HasMinimalIds field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R538");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R538
                // According to [MS-OXNSPI], the MinimalEntryID is a single DWORD value.
                // So if MinimalIds is an array of unsigned integer, R538 will be verified.
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.MinimalIds,
                    typeof(uint[]),
                    538,
                    @"[In GetMatches Request Type Success Response Body] MinimalIds (optional) (variable): An array of MinimalEntryID structures ([MS-OXNSPI] section 2.3.8.1), each of which is the ID of an object found.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R539");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R539
                // Because the HasMinimalIds field is true, so if the MinimalIds has value, then R539 will be verified. 
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBody.MinimalIds,
                    539,
                    @"[In GetMatches Request Type Success Response Body] [MinimalIds] This field is present when the value of the HasMinimalIds field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1450");

                // Because the HasMinimalIds field is false, so if the MinimalIdCount does not have a value, then R1450 will be verified. 
                this.Site.CaptureRequirementIfIsFalse(
                    responseBody.MinimalIdCount.HasValue,
                    1450,
                    @"[In GetMatches Request Type Success Response Body] [MinimalIdCount] This field is not present when the value of the HasMinimalIds field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1294");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1294
                // Because the HasMinimalIds field is false, so if the MinimalIds does not have a value, then R1294 will be verified. 
                this.Site.CaptureRequirementIfIsNull(
                    responseBody.MinimalIds,
                    1294,
                    @"[In GetMatches Request Type Success Response Body] [MinimalIds] This field is not present when the value of the HasMinimalIds field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R542");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R542
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.HasColumnsAndRows,
                typeof(bool),
                542,
                @"[In GetMatches Request Type Success Response Body] HasColumnsAndRows (1 byte): A Boolean value that specifies whether the Columns, RowCount, and RowData fields are present.");

            if (responseBody.HasColumnsAndRows)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R543");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R543
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.Columns.Value,
                    typeof(LargePropertyTagArray),
                    543,
                    @"[In GetMatches Request Type Success Response Body] Columns (optional) (variable): A LargePropertyTagArray structure (section 2.2.1.3) that specifies the columns used for each row returned.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R544");

                // Because the HasColumnsAndRows field is true, so if the Columns field has value, then R544 will be verified. 
                this.Site.CaptureRequirementIfIsTrue(
                    responseBody.Columns.HasValue,
                    544,
                    @"[In GetMatches Request Type Success Response Body] [Columns] This field is present when the value of the HasColumnsAndRows field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R547");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R547
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.RowCount.Value,
                    typeof(uint),
                    547,
                    @"[In GetMatches Request Type Success Response Body] RowCount (optional) (4 bytes): An unsigned integer that specifies the number of structures in the RowData field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R545 : the RowCount field is {0}.", responseBody.RowCount.Value);

                // Because the HasColumnsAndRows field is true, so if the RowCount field has value, then R545 will be verified. 
                this.Site.CaptureRequirementIfIsTrue(
                    responseBody.RowCount.HasValue,
                    545,
                    @"[In GetMatches Request Type Success Response Body] [RowCount] This field is present when the value of the HasColumnsAndRows field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R551");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R551
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.RowData,
                    typeof(AddressBookPropertyRow[]),
                    551,
                    @"[In GetMatches Request Type Success Response Body] RowData (optional) (variable): An array of AddressBookPropertyRow structures (section 2.2.1.2), each of which specifies the row data for the entries requested.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R552");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R552
                // Because the HasColumnsAndRows field is true, so if the RowData field has value, then R552 will be verified. 
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBody.RowData,
                    552,
                    @"[In GetMatches Request Type Success Response Body] [RowData] This field is present when the HasColumnsAndRows field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R502 : the RowCount is {0}.", responseBody.RowCount);

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R502
                // Because the rows count exits on server can smaller than the RowCount value in request.
                // So if RowCount field in response is less than or equal to the RowCount field in request, R502 will be verified.
                bool isVerifiedR502 = responseBody.RowCount <= requesutBody.RowCount;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR502,
                    502,
                    @"[In GetMatches Request Type Request Body] RowCount (4 bytes): An unsigned integer that specifies the number of rows the client is requesting.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1451");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1451
                // Because the HasColumnsAndRows field is false, so if the Columns field does not have value, then R1451 will be verified. 
                this.Site.CaptureRequirementIfIsFalse(
                    responseBody.Columns.HasValue,
                    1451,
                    @"[In GetMatches Request Type Success Response Body] [Columns] This field is not present when the value of the HasColumnsAndRows field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R548");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R548
                // Because the HasColumnsAndRows field is false, so if the RowCount field does not have value, then R548 will be verified. 
                this.Site.CaptureRequirementIfIsFalse(
                    responseBody.RowCount.HasValue,
                    548,
                    @"[In GetMatches Request Type Success Response Body] [RowCount] This field is not present when the value of the HasColumnsAndRows field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1296");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1296
                // Because the HasColumnsAndRows field is false, so if the RowData field does not have value, then R1296 will be verified. 
                this.Site.CaptureRequirementIfIsNull(
                    responseBody.RowData,
                    1296,
                    @"[In GetMatches Request Type Success Response Body] [RowData] This field is not present when the HasColumnsAndRows field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R555");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R555
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.AuxiliaryBufferSize,
                typeof(uint),
                555,
                @"[In GetMatches Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R556");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R556
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.AuxiliaryBuffer,
                typeof(byte[]),
                556,
                @"[In GetMatches Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R557");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R557
            this.Site.CaptureRequirementIfAreEqual<uint>(
                responseBody.AuxiliaryBufferSize,
                (uint)responseBody.AuxiliaryBuffer.Length,
                557,
                @"[In GetMatches Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify SeekEntries request type response body
        /// <summary>
        ///  Verify the requirements related to SeekEntries request type response body.
        /// </summary>
        /// <param name="responseBody">The SeekEntries response body to be verified.</param>
        private void VerifySeekEntriesResponseBody(SeekEntriesResponseBody responseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1038");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1038
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.StatusCode,
                typeof(uint),
                1038,
                @"[In SeekEntries Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1039");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1039
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                responseBody.StatusCode,
                1039,
                @"[In SeekEntries Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1040");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1040
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.ErrorCode,
                typeof(uint),
                1040,
                @"[In SeekEntries Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1041");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1041
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.HasState,
                typeof(bool),
                1041,
                @"[In SeekEntries Request Type RSuccess esponse Body] HasState (1 byte): A Boolean value that specifies whether the State field is present.");

            if (responseBody.HasState)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1042");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1042
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.State.Value,
                    typeof(STAT),
                    1042,
                    @"[In SeekEntries Request Type Success Response Body] State (optional) (36 bytes): A STAT structure ([MS-OXNSPI] section 2.3.7) that specifies the state of a specific address book container.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1044");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1044
                // Because the HasState is true. So if the State has a value, then R1044 will be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBody.State,
                    1044,
                    @"[In SeekEntries Request Type Success Response Body] [State] This field is present when the HasState field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1045");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1045
                // Because the HasState is false. So if the State does not have a value, R1045 will be verified.
                this.Site.CaptureRequirementIfIsNull(
                    responseBody.State,
                    1045,
                    @"[In SeekEntries Request Type Success Response Body] [State] This field is not present when the HasState field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1046");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1046
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.HasColumnsAndRows,
                typeof(bool),
                1046,
                @"[In SeekEntries Request Type Success Response Body] HasColumnsAndRows (1 byte): A Boolean value that specifies whether the Columns, RowCount, and RowData fields are present.");

            if (responseBody.HasColumnsAndRows)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1048");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1048
                // Because the HasColumnsAndRows is true. So if the Columns field is not Null, R1048 will be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBody.Columns,
                    1048,
                    @"[In SeekEntries Request Type Success Response Body] [Columns] This field is present when the value of the HasColumnsAndRows field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1050");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1050
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    responseBody.RowCount.Value,
                    typeof(uint),
                    1050,
                    @"[In SeekEntries Request Type Success Response Body] RowCount (optional) (4 bytes): An unsigned integer that specifies the number of structures contained in the RowData field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1051");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1051
                // Because the HasColumnsAndRows is true. So if the RowCount field is not Null, R1051 will be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBody.RowCount,
                    1051,
                    @"[In SeekEntries Request Type Success Response Body] [RowCount] This field is present when the value of the HasColumnsAndRows field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1054");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1054
                // Because the HasColumnsAndRows is true. So if the RowCount field is not Null, R1051 will be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    responseBody.RowData,
                    1054,
                    @"[In SeekEntries Request Type Success Response Body] [RowData] This field is present when the HasColumnsAndRows field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1049");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1049
                // Because the HasColumnsAndRows is false. So if the Columns field is Null, R1049 will be verified.
                this.Site.CaptureRequirementIfIsNull(
                    responseBody.Columns,
                    1049,
                    @"[In SeekEntries Request Type Success Response Body] [Columns] This field is not present when the value of the HasColumnsAndRows field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1052");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1052
                // Because the HasColumnsAndRows is false. So if the RowCount field is Null, R1052 will be verified.
                this.Site.CaptureRequirementIfIsNull(
                    responseBody.RowCount,
                    1052,
                    @"[In SeekEntries Request Type Success Response Body] [RowCount] This field is not present when the value of the HasColumnsAndRows field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1055");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1055
                // Because the HasColumnsAndRows is false. So if the RowData field is Null, R1055 will be verified.
                this.Site.CaptureRequirementIfIsNull(
                    responseBody.RowData,
                    1055,
                    @"[In SeekEntries Request Type Success Response Body] [RowData] This field is not present when the HasColumnsAndRows field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1056");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1056
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.AuxiliaryBufferSize,
                typeof(uint),
                1056,
                @"[In SeekEntries Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1057");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1057
            this.Site.CaptureRequirementIfIsInstanceOfType(
                responseBody.AuxiliaryBuffer,
                typeof(byte[]),
                1057,
                @"[In SeekEntries Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R1058");

            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R1058
            this.Site.CaptureRequirementIfAreEqual<uint>(
                responseBody.AuxiliaryBufferSize,
                (uint)responseBody.AuxiliaryBuffer.Length,
                1058,
                @"[In SeekEntries Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify GetPropList request type response body
        /// <summary>
        /// Verify the requirements related with GetPropList request type response body. 
        /// </summary>
        /// <param name="getPropListResponseBody">The response body of GetPropList request type.</param>
        private void VerifyGetPropListResponseBody(GetPropListResponseBody getPropListResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R579");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R579
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropListResponseBody.StatusCode,
                typeof(uint),
                579,
                @"[In GetPropList Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R580");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R580
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getPropListResponseBody.StatusCode,
                580,
                @"[In GetPropList Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R581");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R581
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropListResponseBody.ErrorCode,
                typeof(uint),
                581,
                @"[In GetPropList Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R582");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R582
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropListResponseBody.HasPropertyTags,
                typeof(bool),
                582,
                @"[In GetPropList Request Type Success Response Body] HasPropertyTags (1 byte): A Boolean value that specifies whether the PropertyTags field is present.");

            if (getPropListResponseBody.HasPropertyTags)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R584");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R584
                // Because the HasPropertyTags field is true, so if the PropertyTags field has value, R584 will be verified. 
                this.Site.CaptureRequirementIfIsNotNull(
                    getPropListResponseBody.PropertyTags,
                    584,
                    @"[In GetPropList Request Type Success Response Body] [PropertyTags] This field is present when the value of the HasPropertyTags field is nonzero.");
            }
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R587");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R587
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropListResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                587,
                @"[In GetPropList Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R588");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R588
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropListResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                588,
                @"[In GetPropList Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R589");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R589
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getPropListResponseBody.AuxiliaryBufferSize,
                (uint)getPropListResponseBody.AuxiliaryBuffer.Length,
                589,
                @"[In GetPropList Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify GetProps request type response body
        /// <summary>
        /// Verify the requirements related to GetProps request type response body. 
        /// </summary>
        /// <param name="getPropsResponseBody">The response body of GetProps request type.</param>
        private void VerifyGetPropsResponseBody(GetPropsResponseBody getPropsResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R619");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R619
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBody.StatusCode,
                typeof(uint),
                619,
                @"[In GetProps Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R620");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R620
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getPropsResponseBody.StatusCode,
                620,
                @"[In GetProps Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R621");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R621
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBody.ErrorCode,
                typeof(uint),
                621,
                @"[In GetProps Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R622");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R622
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBody.CodePage,
                typeof(uint),
                622,
                @"[In GetProps Request Type Success Response Body] CodePage (4 bytes): An unsigned integer that specifies the code page that the server used to express string properties.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R623");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R623
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBody.HasPropertyValues,
                typeof(bool),
                623,
                @"[In GetProps Request Type Success Response Body] HasPropertyValues (1 byte): A Boolean value that specifies whether the PropertyValues field is present.");

            if (getPropsResponseBody.HasPropertyValues)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R626");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R626
                // Because the HasPropertyValues field is true, so if the PropertyValues field has value, R626 will be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    getPropsResponseBody.PropertyValues,
                    626,
                    @"[In GetProps Request Type Success Response Body] [PropertyValues] This field is present when the value of the HasPropertyValues field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R627");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R627
                // Because the HasPropertyValues field is false, so if the PropertyValues field has not value, R627 will be verified.
                this.Site.CaptureRequirementIfIsNull(
                    getPropsResponseBody.PropertyValues,
                    627,
                    @"[In GetProps Request Type Success Response Body] [PropertyValues] This field is not present when the value of the HasPropertyValues field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R628");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R628
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                628,
                @"[In GetProps Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R629");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R629
            this.Site.CaptureRequirementIfIsInstanceOfType(
                getPropsResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                629,
                @"[In GetProps Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R630");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R630
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getPropsResponseBody.AuxiliaryBufferSize,
                (uint)getPropsResponseBody.AuxiliaryBuffer.Length,
                630,
                @"[In GetProps Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify ModProps request type response body
        /// <summary>
        /// Verify the requirements related to ModProps request type response body. 
        /// </summary>
        /// <param name="modPropsResponseBody">The response body of ModProps request type.</param>
        private void VerifyModPropsResponseBody(ModPropsResponseBody modPropsResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R796");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R796
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modPropsResponseBody.StatusCode,
                typeof(uint),
                796,
                @"[In ModProps Request Type Success Response Body] StatusCode: An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R797");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R797
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                modPropsResponseBody.StatusCode,
                797,
                @"[In ModProps Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R798");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R798
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modPropsResponseBody.ErrorCode,
                typeof(uint),
                798,
                @"[In ModProps Request Type Success Response Body] ErrorCode: An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R799");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R799
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modPropsResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                799,
                @"[In ModProps Request Type Success Response Body] AuxiliaryBufferSize: An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R800");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R800
            this.Site.CaptureRequirementIfIsInstanceOfType(
                modPropsResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                800,
                @"[In ModProps Request Type Success Response Body] AuxiliaryBuffer: An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R801");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R801
            this.Site.CaptureRequirementIfAreEqual<uint>(
                modPropsResponseBody.AuxiliaryBufferSize,
                (uint)modPropsResponseBody.AuxiliaryBuffer.Length,
                801,
                @"[In ModProps Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify QueryColumns request type response body
        /// <summary>
        /// Verify the requirements related to QueryColumns request type response body. 
        /// </summary>
        /// <param name="queryColumnsResponseBody">The response body of QueryColumns request type.</param>
        private void VerifyQueryColumnsResponseBody(QueryColumnsResponseBody queryColumnsResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R885");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R885
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryColumnsResponseBody.StatusCode,
                typeof(uint),
                885,
                @"[In QueryColumns Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R886");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R886
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                queryColumnsResponseBody.StatusCode,
                886,
                @"[In QueryColumns Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R887");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R887
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryColumnsResponseBody.ErrorCode,
                typeof(uint),
                887,
                @"[In QueryColumns Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R888");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R888
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryColumnsResponseBody.HasColumns,
                typeof(bool),
                888,
                @"[In QueryColumns Request Type Success Response Body] HasColumns (1 byte): A Boolean value that specifies whether the Columns field is present.");

            if (queryColumnsResponseBody.HasColumns)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R889");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R889
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    queryColumnsResponseBody.Columns,
                    typeof(LargePropertyTagArray),
                    889,
                    @"[In QueryColumns Request Type Success Response Body] Columns (optional) (variable): A LargePropTagArray structure (section 2.2.1.3) that specifies the properties that exist on the address book.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R890");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R890
                this.Site.CaptureRequirementIfIsNotNull(
                    queryColumnsResponseBody.Columns,
                    890,
                    @"[In QueryColumns Request Type Success Response Body] [Columns] This field is present when the HasColumns field is nonzero.");
            }
           
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R892");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R892
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryColumnsResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                892,
                @"[In QueryColumns Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R893");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R893
            this.Site.CaptureRequirementIfIsInstanceOfType(
                queryColumnsResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                893,
                @"[In QueryColumns Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R894");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R894
            this.Site.CaptureRequirementIfAreEqual<uint>(
                queryColumnsResponseBody.AuxiliaryBufferSize,
                (uint)queryColumnsResponseBody.AuxiliaryBuffer.Length,
                894,
                @"[In QueryColumns Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion

        #region Verify ResolveNames request type response body
        /// <summary>
        /// Verify the requirements related to ResolveNames request type response body. 
        /// </summary>
        /// <param name="resolveNamesResponseBody">The response body of ResolveNames request type.</param>
        private void VerifyResolveNamesResponseBody(ResolveNamesResponseBody resolveNamesResponseBody)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R934");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R934
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resolveNamesResponseBody.StatusCode,
                typeof(uint),
                934,
                @"[In ResolveNames Request Type Success Response Body] StatusCode (4 bytes): An unsigned integer that specifies the status of the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R935");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R935
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                resolveNamesResponseBody.StatusCode,
                935,
                @"[In ResolveNames Request Type Success Response Body] [StatusCode] This field MUST be set to 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R936");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R936
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resolveNamesResponseBody.ErrorCode,
                typeof(uint),
                936,
                @"[In ResolveNames Request Type Success Response Body] ErrorCode (4 bytes): An unsigned integer that specifies the return status of the operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R937");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R937
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resolveNamesResponseBody.CodePage,
                typeof(uint),
                937,
                @"[In ResolveNames Request Type Success Response Body] CodePage (4 bytes): An unsigned integer that specifies the code page the server used to express string values of properties.");

            if (resolveNamesResponseBody.HasMinimalIds)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R938");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R938
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resolveNamesResponseBody.HasMinimalIds,
                    typeof(bool),
                    938,
                    @"[In ResolveNames Request Type Success Response Body] HasMinimalIds (1 byte): A Boolean value that specifies whether the MinimalIdCount and MinimalIds fields are present.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R939");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R939
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resolveNamesResponseBody.MinimalIdCount,
                    typeof(uint),
                    939,
                    @"[In ResolveNames Request Type Success Response Body] MinimalIdCount (optional) (4 bytes): An unsigned integer that specifies the number of structures in the MinimalIds field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R940");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R940
                this.Site.CaptureRequirementIfIsNotNull(
                    resolveNamesResponseBody.MinimalIdCount,
                    940,
                    @"[In ResolveNames Request Type Success Response Body] [MinimalIdCount] This field is present when the value of the HasMinimalIds field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R943");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R943
                this.Site.CaptureRequirementIfIsNotNull(
                    resolveNamesResponseBody.MinimalIds,
                    943,
                    @"[In ResolveNames Request Type Success Response Body] [MinimalIds] This field is present when the value of the HasMinimalIds field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R941");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R941
                this.Site.CaptureRequirementIfIsNull(
                    resolveNamesResponseBody.MinimalIdCount,
                    941,
                    @"[In ResolveNames Request Type Success Response Body] [MinimalIdCount] This field is not present when the value of the HasMinimalIds field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R944");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R944
                this.Site.CaptureRequirementIfIsNull(
                    resolveNamesResponseBody.MinimalIds,
                    944,
                    @"[In ResolveNames Request Type Success Response Body] [MinimalIds] This field is not present when the value of the HasMinimalIds field is zero.");
            }

            if (resolveNamesResponseBody.HasRowsAndPropertyTags)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R945");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R945
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resolveNamesResponseBody.HasRowsAndPropertyTags,
                    typeof(bool),
                    945,
                    @"[In ResolveNames Request Type Success Response Body] HasRowsAndCols (1 byte): A Boolean value that specifies whether the PropertyTags, RowCount, and RowData fields are present.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R949");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R949
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resolveNamesResponseBody.RowCount,
                    typeof(uint),
                    949,
                    @"[In ResolveNames Request Type Success Response Body] RowCount (4 bytes): An unsigned integer that specifies the number of structures in the RowData field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R952");

                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R952
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    resolveNamesResponseBody.RowData,
                    typeof(AddressBookPropertyRow[]),
                    952,
                    @"[In ResolveNames Request Type Success Response Body] RowData (optional) (variable): An array of AddressBookPropertyRow structures (section 2.2.1.2), each of which specifies the row data requested.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R947");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R947
                this.Site.CaptureRequirementIfIsNotNull(
                    resolveNamesResponseBody.PropertyTags,
                    947,
                    @"[In ResolveNames Request Type Success Response Body] [PropertyTags] This field is present when the value of the HasRowsAndColumns field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R950");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R950
                this.Site.CaptureRequirementIfIsNotNull(
                    resolveNamesResponseBody.RowCount,
                    950,
                    @"[In ResolveNames Request Type Success Response Body] [RowCount] This field is present when the value of the HasRowsAndCols field is nonzero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R953");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R953
                this.Site.CaptureRequirementIfIsNotNull(
                    resolveNamesResponseBody.RowData,
                    953,
                    @"[In ResolveNames Request Type Success Response Body] [RowData] This field is present when the value of the HasRowsAndCols field is nonzero.");
            }
            else
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R948");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R948
                this.Site.CaptureRequirementIfIsNull(
                    resolveNamesResponseBody.PropertyTags,
                    948,
                    @"[In ResolveNames Request Type Success Response Body] [PropertyTags] This field is not present when the value of the HasRowsAndColumns field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R951");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R951
                this.Site.CaptureRequirementIfIsNull(
                    resolveNamesResponseBody.RowCount,
                    951,
                    @"[In ResolveNames Request Type Success Response Body] [RowCount] This field is not present when the value of the HasRowsAndCols field is zero.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R954");
        
                // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R954
                this.Site.CaptureRequirementIfIsNull(
                    resolveNamesResponseBody.RowData,
                    954,
                    @"[In ResolveNames Request Type Success Response Body] [RowData] This field is not present when the value of the HasRowsAndCols field is zero.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R955");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R955
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resolveNamesResponseBody.AuxiliaryBufferSize,
                typeof(uint),
                955,
                @"[In ResolveNames Request Type Success Response Body] AuxiliaryBufferSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R956");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R956
            this.Site.CaptureRequirementIfIsInstanceOfType(
                resolveNamesResponseBody.AuxiliaryBuffer,
                typeof(byte[]),
                956,
                @"[In ResolveNames Request Type Success Response Body] AuxiliaryBuffer (variable): An array of bytes that constitute the auxiliary payload data returned from the server.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMAPIHTTP_R957");
        
            // Verify MS-OXCMAPIHTTP requirement: MS-OXCMAPIHTTP_R957
            this.Site.CaptureRequirementIfAreEqual<uint>(
                resolveNamesResponseBody.AuxiliaryBufferSize,
                (uint)resolveNamesResponseBody.AuxiliaryBuffer.Length,
                957,
                @"[In ResolveNames Request Type Success Response Body] [AuxiliaryBuffer] The size of this field, in bytes, is specified by the AuxiliaryBufferSize field.");
        }
        #endregion
        #endregion
    }
}