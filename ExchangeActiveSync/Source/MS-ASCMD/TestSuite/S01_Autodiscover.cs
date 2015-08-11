namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the Autodiscover command.
    /// </summary>
    [TestClass]
    public class S01_Autodiscover : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is used to verify if the Type element value is set to 'MobileSync', the Name element should be returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S01_TC01_Autodiscover_TypeIsMobileSync()
        {
            Site.Assume.IsFalse(Common.GetSutVersion(this.Site) == SutVersion.ExchangeServer2007 && string.Equals(Common.GetConfigurationPropertyValue("TransportType", this.Site).ToUpper(CultureInfo.InvariantCulture), "HTTP"), "Autodiscover request should be passed only through HTTPS to Exchange Server 2007.");
            AutodiscoverRequest request = new AutodiscoverRequest
            {
                RequestData = new Request.Autodiscover
                {
                    Request = new Request.RequestType
                    {
                        AcceptableResponseSchema = Common.GetConfigurationPropertyValue("AcceptableResponseSchema", Site),
                        EMailAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain)
                    }
                }
            };

            AutodiscoverResponse response = CMDAdapter.Autodiscover(request, ContentTypeEnum.Xml);
            Site.Assert.AreEqual<string>("MobileSync", response.ResponseData.Response.Action.Settings[0].Type, "The type of Action in Autodiscover command response should be MobileSync.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3482");

            // If the Type element value is "MobileSync", then the Name element specifies the URL that conveys the protocol which specified by Url element.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3482
            Site.CaptureRequirementIfAreEqual<string>(
                response.ResponseData.Response.Action.Settings[0].Url,
                response.ResponseData.Response.Action.Settings[0].Name,
                3482,
                @"[In Name(Autodiscover)] The Name element is an optional child element of the Server element in Autodiscover command responses that specifies a URL if the Type element (section 2.2.3.170.1) value is set to ""MobileSync"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3484");

            // If the Type element value is "MobileSync", then the Name element specifies the URL that conveys the protocol which specified by Url element.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3484
            Site.CaptureRequirementIfAreEqual<string>(
                response.ResponseData.Response.Action.Settings[0].Url,
                response.ResponseData.Response.Action.Settings[0].Name,
                3484,
                @"[In Name(Autodiscover)] If the Type element value is ""MobileSync"", then the Name element specifies the URL that conveys the protocol.");

            if (Common.IsRequirementEnabled(5160, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5160");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5160
                Site.CaptureRequirementIfAreEqual<string>(
                    "en:en",
                    response.ResponseData.Response.Culture,
                    5160,
                    "[In Appendix A: Product Behavior] Implementation does return the form \"en:en\" of Culture element, regardless of the culture that is sent by the client. (<26> Section 2.2.3.38: In Exchange 2007, the Culture element always returns \"en:en\", regardless of the culture that is sent by the client.)");
            }

            if (Common.IsRequirementEnabled(5823, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5823");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5823
                Site.CaptureRequirementIfAreEqual<string>(
                    "en:us",
                    response.ResponseData.Response.Culture,
                    5823,
                    "[In Appendix A: Product Behavior] Implementation does return the form \"en:us\" of Culture element. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(5718, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5718");

                // If the Response element do not have an Error child element when set the Content-Type header to "text/xml", it indicates an error does not occur in the Autodiscover command framework that hosts the Autodiscovery implementation.
                // Verify MS-ASCMD requirement: MS-ASCMD_R5718
                Site.CaptureRequirementIfIsNull(
                    response.ResponseData.Response.Error,
                    5718,
                    "[In Appendix A: Product Behavior] When sending an Autodiscover command request to implementation, the Content-Type header does accept the following values: \"text/xml\". (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(5123, this.Site))
            {
                response = CMDAdapter.Autodiscover(request, ContentTypeEnum.Html);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5123");

                // If the Response element do not have an Error child element when set the Content-Type header to "text/html", it indicates an error does not occur in the Autodiscover command framework that hosts the Autodiscovery implementation.
                // Verify MS-ASCMD requirement: MS-ASCMD_R5123
                Site.CaptureRequirementIfIsNull(
                    response.ResponseData.Response.Error,
                    5123,
                    "[In Appendix A: Product Behavior] When sending an Autodiscover command request to implementation, the Content-Type header does accept the following values: \"text/html\" [or \"text/xml\"]. (<1> Section 2.2.2.1: When sending an Autodiscover command request to Exchange 2007, the Content-Type header accepts the following values: \"text/html\" or \"text/xml\".)");
            }
        }

        /// <summary>
        /// This test case is used to verify if Autodiscover failed, the server should return an error child element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S01_TC02_Autodiscover_Fail()
        {
            Site.Assume.IsFalse(Common.GetSutVersion(Site) == SutVersion.ExchangeServer2007 && string.Equals(Common.GetConfigurationPropertyValue("TransportType", this.Site).ToUpper(CultureInfo.InvariantCulture), "HTTP"), "Autodiscover request should be passed only through HTTPS to Exchange Server 2007.");
            AutodiscoverRequest request = new AutodiscoverRequest
            {
                RequestData = new Request.Autodiscover
                {
                    Request = new Request.RequestType
                    {
                        AcceptableResponseSchema = Common.GetConfigurationPropertyValue("AcceptableResponseSchema", Site),
                        EMailAddress = Common.GetMailAddress("InvallidEmailAddress", this.User1Information.UserDomain)
                    }
                }
            };

            AutodiscoverResponse response = CMDAdapter.Autodiscover(request, ContentTypeEnum.Xml);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3818");

            // An Error child element returned in Response element indicate an error occurs in the Autodiscover command framework that hosts the Autodiscovery implementation.
            // Verify MS-ASCMD requirement: MS-ASCMD_R3818
            Site.CaptureRequirementIfIsNotNull(
                response.ResponseData.Response.Action.Error,
                3818,
                @"[In Response(Autodiscover)] If an error occurs in the Autodiscover command framework that hosts the Autodiscovery implementation, then the Response element MUST have an Error child element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4001");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4001
            Site.CaptureRequirementIfAreNotEqual<string>(
                "1",
                response.ResponseData.Response.Action.Error.Status,
                4001,
                @"[In Status(Autodiscover)] Because the Status element is only returned when the command encounters an error, the success status code is never included in a response message.");
        }
        #endregion
    }
}