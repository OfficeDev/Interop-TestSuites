namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASHTTP.
    /// </summary>
    public partial class MS_ASHTTPAdapter
    {
        #region Verify transport related requirements.
        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            ProtocolTransportType transport = (ProtocolTransportType)Enum.Parse(typeof(ProtocolTransportType), Common.GetConfigurationPropertyValue("TransportType", Site), true);

            if (transport == ProtocolTransportType.HTTPS)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R518");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R518
                // When test suite is running on HTTPS and the command is executed successfully, this requirement will be captured.
                Site.CaptureRequirement(
                    518,
                    @"[In Transport] These commands [messages] are sent via [HTTP or] Hypertext Transfer Protocol over Secure Sockets Layer (HTTPS).");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R401");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R401
                // When test suite is running on HTTP and the command is executed successfully, this requirement will be captured.
                Site.CaptureRequirement(
                    401,
                    @"[In Transport] These commands [messages] are sent via HTTP [or Hypertext Transfer Protocol over Secure Sockets Layer (HTTPS)].");
            }
        }
        #endregion

        /// <summary>
        /// Verify requirements about HTTP POST response.
        /// </summary>
        /// <param name="postResponse">The HTTP POST response.</param>
        private void VerifyHTTPPOSTResponse(SendStringResponse postResponse)
        {
            // Verify MS-ASHTTP requirement: MS-ASHTTP_R175
            Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                HttpStatusCode.OK,
                postResponse.StatusCode,
                175,
                @"[In Status Line] [Status code] 200 OK [is described as] the command succeeded.");

            if (Common.IsRequirementEnabled(476, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R476");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R476
                // The response is returned successfully and the StatusCode is OK, so this requirement can be captured.
                Site.CaptureRequirement(
                    476,
                    @"[In Appendix A: Product Behavior] Implementation does format a response to the request as specified in section 2.2.2 with an appropriate HTTP status code as specified in section 2.2.2.1.1. (Exchange 2007 SP1 and above follow this behavior.)");
            }

            this.VerifyHTTPResponse();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R1");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R1
            // The HTTP POST command executes successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                1,
                @"[In Transport] Messages are transported by using HTTP POST, as specified in [RFC2616].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R167");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R167
            // The HTTP POST command executes successfully and the HTTP POST response is not null, so this requirement can be captured.
            Site.CaptureRequirement(
                167,
                @"[In HTTP POST Response] After receiving and interpreting a request, a server responds with an HTTP response that contains data returned from the server.");

            // GetAttachment command response doesn't contain xml format data, so GetAttachment command response is excluded from this capture.
            if (!string.IsNullOrEmpty(postResponse.ResponseDataXML) && !postResponse.ResponseDataXML.Contains("PNG"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R169");
                Site.Log.Add(LogEntryKind.Debug, "The ResponseDataXML of HTTP POST command is {0}.", postResponse.ResponseDataXML);

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R169
                // The HTTP POST command execute successfully, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    postResponse.ResponseDataXML.Contains("utf-8"),
                    169,
                    @"[In Response Format] Note that these [HTTP POST] responses are UTF-8 encoded.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R170");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R170
            // The HTTP POST command executes successfully and raw http response is converted successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                170,
                @"[In Response Format] As specified by [RFC2616], the format [of the command response] is the same as for the following requests.
                    Status-line
                    Response-headers
                    CR/LF
                    Message Body");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R172");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R172
            // The HTTP POST command executes successfully and raw http response is converted successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                172,
                @"[In Status Line] The status line consists of the HTTP version and a status code.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R228");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R228
            // The HTTP POST command executes successfully and raw http response is converted successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                228,
                @"[In Response Body] The response body contains data returned from the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R281");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R281
            // The HTTP POST command executes successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                281,
                @"[In Message Processing Events and Sequencing Rules] The server can receive HTTP POST requests (section 2.2.1) [or HTTP OPTIONS requests (section 2.2.3)] from the client.");

            this.VerifyHTTPPOSTHeaders(postResponse);
        }

        /// <summary>
        /// Verify requirements about HTTP POST response headers.
        /// </summary>
        /// <param name="postResponse">The HTTP POST response.</param>
        private void VerifyHTTPPOSTHeaders(SendStringResponse postResponse)
        {
            int contentLengthNumber;
            string responseHeaders = postResponse.Headers.ToString();
            Site.Log.Add(LogEntryKind.Debug, "The response headers are: {0}.", responseHeaders);
            bool isContentLengthExist = responseHeaders.Contains("Content-Length");
            bool isContentTypeExist = responseHeaders.Contains("Content-Type");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R199");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R199
            // The Content-Length header exists in the HTTP POST response and can be parsed to int, so this requirement can be captured.
            bool isVerifiedR199 = isContentLengthExist && int.TryParse(postResponse.Headers["Content-Length"], out contentLengthNumber);

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR199,
                199,
                @"[In Response Headers] Required [header] Content-Length specifies the size of the response body in bytes, [whose example value is] 56.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R521");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R521
            // The Content-Length header follows the specified format, so this requirement can be captured.
            bool isVerifiedR521 = responseHeaders.Contains("Content-Length: " + postResponse.Headers["Content-Length"]) && int.TryParse(postResponse.Headers["Content-Length"], out contentLengthNumber);

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR521,
                521,
                @"[In Content-Length] Content-Length = ""Content-Length"" "":"" 1*DIGIT");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R216");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R216
            // The Content-Length header exists in the HTTP POST response, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isContentLengthExist,
                216,
                @"[In Content-Length] This [Content-Length] header is required.");

            bool hasContent = int.Parse(postResponse.Headers["Content-Length"]) > 0;
            bool isVerifiedR212;
            if (hasContent)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R201");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R201
                // The Content-Type header exists in the HTTP POST response, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isContentTypeExist,
                    201,
                    @"[In Response Headers] Required [header] Content-Type specifies that the media-type of the response body is WBXML, [whose example value is] application/vnd.ms-sync.wbxml.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R218");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R218
                // The Content-Type header exists in the HTTP POST response, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isContentTypeExist,
                    218,
                    @"[In Content-Type] This [Content-Type] header is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R522");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R522
                // The Content-Type header follows the specified format, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    responseHeaders.Contains("Content-Type: " + postResponse.Headers["Content-Type"]),
                    522,
                    @"[In Content-Type] Content-Type = ""Content-Type"" "":"" media-type");

                isVerifiedR212 = isContentLengthExist && isContentTypeExist;
            }
            else
            {
                isVerifiedR212 = isContentLengthExist;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R212");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R212
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR212,
                212,
                @"[In Response Headers] When these two conditions [the response is to an HTTP POST request and the response has HTTP status 200] are met, only the following headers are necessary in the response: 
                  Content-Length
                  Content-Type, required only if Content-Length is greater than zero.");
        }

        /// <summary>
        /// Verify requirements about HTTP OPTIONS response.
        /// </summary>
        /// <param name="optionsResponse">The HTTP OPTIONS response.</param>
        private void VerifyHTTPOPTIONSResponse(OptionsResponse optionsResponse)
        {
            this.VerifyHTTPResponse();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R319");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R319
            // The HTTP OPTIONS command executes successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                319,
                @"[In Transport] Messages are transported by using HTTP OPTIONS, as specified in [RFC2616].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R239");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R239
            // The HTTP OPTIONS command executes successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                239,
                @"[In Response Format] Each response is sent from the server to the client as an HTTP OPTIONS response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R240");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R240
            // The HTTP OPTIONS command executes successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                240,
                @"[In Response Format] Note that these [sent from the server to the client as an HTTP OPTIONS] responses are UTF-8 encoded.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R241");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R241
            // The HTTP OPTIONS command executes successfully and raw http response is converted successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                241,
                @"[In Response Format] As specified by [RFC2616], the format is the same as for the following requests:
                    Status-line
                    Response-headers");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R244");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R244
            // The HTTP OPTIONS command executes successfully and raw http response is converted successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                244,
                @"[In Response Headers] The headers [MS-ASProtocolCommands,MS-ASProtocolVersions] follow the status line in the HTTP part of a response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R450");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R450
            // The HTTP OPTIONS command executes successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                450,
                @"[In Message Processing Events and Sequencing Rules] The server can receive [HTTP POST requests (section 2.2.1) or] HTTP OPTIONS requests (section 2.2.3) from the client.");

            this.VerifyHTTPOPTIONSHeaders(optionsResponse);
        }

        /// <summary>
        /// Verify requirements about HTTP OPTIONS response headers.
        /// </summary>
        /// <param name="optionsResponse">The HTTP OPTIONS response.</param>
        private void VerifyHTTPOPTIONSHeaders(OptionsResponse optionsResponse)
        {
            string commandHeaders = optionsResponse.Headers["MS-ASProtocolCommands"];
            string versionHeaders = optionsResponse.Headers["MS-ASProtocolVersions"];
            Site.Log.Add(LogEntryKind.Debug, "The MS-ASProtocolCommands header in response is: {0}.", commandHeaders);
            Site.Log.Add(LogEntryKind.Debug, "The MS-ASProtocolVersions header in response is: {0}.", versionHeaders);


            if (Common.IsRequirementEnabled(459, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R459");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R459
                // There is "12.1" in MS-ASProtocolVersions header, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    versionHeaders.Contains("12.1"),
                    459,
                    @"[In Appendix A: Product Behavior] Implementation does return the MS-ASProtocolVersions value of 12.1. (Exchange 2007 SP1 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(460, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R460");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R460
                // There is "14.0" in MS-ASProtocolVersions header, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    versionHeaders.Contains("14.0"),
                    460,
                    @"[In Appendix A: Product Behavior] Implementation does return the MS-ASProtocolVersions value of 14.0. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(461, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R461");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R461
                // There is "14.1" in MS-ASProtocolVersions header, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    versionHeaders.Contains("14.1"),
                    461,
                    @"[In Appendix A: Product Behavior] Implementation does return the MS-ASProtocolVersions value of 14.1. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1201, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R1201");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R1201
                // There is "16.0" in MS-ASProtocolVersions header, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    versionHeaders.Contains("16.0"),
                    1201,
                    @"[In Appendix A: Product Behavior] Implementation does return the MS-ASProtocolVersions value of 16.0. (Exchange 2016 Preview and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(12011, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R12011");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R12011
                // There is "16.1" in MS-ASProtocolVersions header, so this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    versionHeaders.Contains("16.1"),
                    12011,
                    @"[In Appendix A: Product Behavior] Implementation does return the MS-ASProtocolVersions value of 16.1. (Exchange 2016 Preview and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R249");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R249
            // R459, R460 and R461 are captured which means implementation returns MS-ASProtocolVersions value of 12.1, 14.0, 14.1,16.0 or 16.1, so this requirement can be captured directly.
            Site.CaptureRequirement(
                249,
                @"[In MS-ASProtocolVersions] The following values [MS-ASProtocolVersions] correspond to the ActiveSync protocol versions that are specified by [MS-ASCMD]:""16.1"", ""16.0"", ""14.1"", ""14.0"", ""12.1"", ""12.0"" and ""2.5"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R1107");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R1107
            Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(commandHeaders) && !string.IsNullOrEmpty(versionHeaders),
                1107,
                @"[In Handling HTTP OPTIONS Command] The server's response MUST contain both the MS-ASProtocolCommands header, as specified in section 2.2.4.1.2.1, and the MS-ASProtocolVersions header, as specified in section 2.2.4.1.2.2.");

            string[] splitCommand = commandHeaders.Split(',');

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R247");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R247
            // The MS-ASProtocolCommands header is split by ",", so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                splitCommand.Length > 1,
                247,
                @"[In MS-ASProtocolCommands] The MS-ASProtocolCommands header contains a comma-delimited list of the ActiveSync commands supported by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R238");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R238
            // R291 is captured, so this requirement can be captured.
            Site.CaptureRequirement(
                238,
                @"[In HTTP OPTIONS Response] After receiving an HTTP OPTIONS request, a server responds with an HTTP OPTIONS response that specifies the protocol versions it supports.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R296");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R296
            // After R459, R460 and R461 are captured which means implementation returns MS-ASProtocolVersions value of 12.1, 14.0 or 14.1, so this requirement can be captured.
            Site.CaptureRequirement(
                296,
                @"[In Handling HTTP OPTIONS Command] A protocol server can support multiple versions of the ActiveSync protocol.");

            if (Common.IsRequirementEnabled(298, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R298");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R298
                // When the server version is ExchangeServer2007, 14.0, 14.1,16.0 and 16.1 are not returned, this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    !versionHeaders.Contains("14.0") && !versionHeaders.Contains("14.1") && !versionHeaders.Contains("16.0") && !versionHeaders.Contains("16.1"),
                    298,
                    @"[In Appendix A: Product Behavior] Implementation does not return MS-ASProtocolVersions values of 16.1,16.0, 14.1 or 14.0. (<11> Section 3.2.5.2: Exchange 2007 SP1 does not return the value ""16.1"",""16.0"", ""14.1"", or ""14.0"" in the MS-ASProtocolVersions header.)");
            }
        }

        /// <summary>
        /// Verify requirements about HTTP response.
        /// </summary>
        private void VerifyHTTPResponse()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R261");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R261
            // The command executed successfully, so this requirement can be captured.
            Site.CaptureRequirement(
                261,
                @"[In Message Processing Events and Sequencing Rules] Clients receive HTTP responses from the server only in response to HTTP requests sent by the client.");
        }
    }
}