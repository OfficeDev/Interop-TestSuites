//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS. Protocol client tries to perform operations associated with the customization of the specified cascading style sheet.
    /// </summary>
    [TestClass]
    public class S08_OperationsOnCSS : TestSuiteBase
    {
        #region Additional test attributes, initialization and clean up

        /// <summary>
        /// Class initialization.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }
        #endregion

        /// <summary>
        /// This test case aims to verify the CustomizeCss operation with a valid cssFile without authorization.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC01_CustomizeCss_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);
            try
            {
                Adapter.CustomizeCss(Common.GetConfigurationPropertyValue("CssFile_Valid", this.Site));
                Site.Assert.Fail("The expected http status code is not returned for the CustomizeCss operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1070
                // COMMENT: When the CustomizeCss operation is invoked by unauthenticated user, if 
                // the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1070,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[CustomizeCss], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertCss operation when cssFile doesn’t exist on the protocol server.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC02_RevertCss_CssFileNotExist()
        {
            Adapter.RevertCss(string.Format("{0}.css", this.GenerateRandomString(5)));
            
            // Verify MS-WEBSS requirement: MS-WEBSS_R844
            // When the RevertCss operation is invoked when the file specified by the cssFile 
            // element does not exist on the server, if there is no SOAP fault returned, then the requirement 
            // can be captured.
            Site.CaptureRequirement(
            844,
            @"[In RevertCss] If this value[CSS file name] does not match one of the CSS files on the protocol server, the protocol server [MUST NOT carry out the operation, and it] MUST NOT return SOAP fault message.");
        }

        /// <summary>
        /// This test case aims to verify the CustomizeCss operation when the cssFile element is blank or missing.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC03_CustomizeCss_CssFileBlankOrMissing()
        {
            try
            {
                Adapter.CustomizeCss(string.Empty);
                Site.Assert.Fail("The error code is not returned for the CustomizeCss operation.");
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1042
                // COMMENT: If the server return the server faults as SOAP faults, then the requirement 
                // can be captured.
                Site.CaptureRequirement(
                    1042,
                    @"[In Transport] Protocol server faults MUST be returned SOAP faults as specified either in [SOAP1.1] section 4.4, SOAP Fault, or in [SOAP1.2/1] section 5.4, SOAP Fault.");

                if (Common.IsRequirementEnabled(99, this.Site))
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R99
                    // COMMENT: When the CustomizeCss operation is invoked when the cssFile element is blank 
                    // or missing, if the returned error code is 0x82000001, then the requirement can be captured.
                    Site.CaptureRequirementIfAreEqual<string>(
                        SoapErrorCode.ErrorCode0x82000001,
                        exp.Detail.LastChild.InnerText,
                       99,
                    @"[In CustomizeCssResponse] [If implementation does encounter one of the error conditions described in the following table when running this operation, a SOAP fault MUST be returned that contain one of the error codes for the specified error condition.] Implementation does the Error Code ""0x82000001"" meaning is Blank cssFile specified, or cssFile element missing. (Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the CustomizeCss operation when the cssFile element has no file name extension.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC04_CustomizeCss_NoFileExtension()
        {
            string soapErrorCode = string.Empty;
            try
            {
                Adapter.CustomizeCss(this.GenerateRandomString(5));
                Site.Assert.Fail("The error code is not returned for the CustomizeCss operation.");
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1042
                // COMMENT: If the server return the server faults as SOAP faults, then the requirement 
                // can be captured.
                Site.CaptureRequirement(
                    1042,
                    @"[In Transport] Protocol server faults MUST be returned SOAP faults as specified either in [SOAP1.1] section 4.4, SOAP Fault, or in [SOAP1.2/1] section 5.4, SOAP Fault.");

                if (Common.IsRequirementEnabled(100, this.Site))
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R100
                    // COMMENT: When the CustomizeCss operation is invoked when the cssFile element has no file 
                    // extension, if the returned error code is 0x80131600, then the requirement can be captured.
                    Site.CaptureRequirementIfAreEqual<string>(
                        SoapErrorCode.ErrorCode0x80131600,
                        soapErrorCode = exp.Detail.LastChild.InnerText,
                         100,
                    @"[In CustomizeCssResponse] [If implementation does encounter one of the error conditions described in the following table when running this operation, a SOAP fault MUST be returned that contain one of the error codes for the specified error condition.]  Implementation does the Error Code equal ""0x80131600"" meaning is Specified cssFile has no file extension.(Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the CustomizeCss operation when the cssFile does not exist on the protocol server.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC05_CustomizeCss_CssFileNotExist()
        {
            try
            {
                Adapter.CustomizeCss(string.Format("{0}.css", this.GenerateRandomString(5)));
                Site.Assert.Fail("The expected error code '0x80070002' is not returned for the CustomizeCss operation.");
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1042
                // COMMENT: If the server return the server faults as SOAP faults, then the requirement 
                // can be captured.
                Site.CaptureRequirement(
                    1042,
                    @"[In Transport] Protocol server faults MUST be returned SOAP faults as specified either in [SOAP1.1] section 4.4, SOAP Fault, or in [SOAP1.2/1] section 5.4, SOAP Fault.");

                if (Common.IsRequirementEnabled(101, this.Site))
                {
                    // COMMENT: When the CustomizeCss operation is invoked when the file specified by the 
                    // cssFile element does not exist on the server, if the returned error code is 0x80070002, then 
                    // the requirement can be captured.
                    // Verify MS-WEBSS requirement: MS-WEBSS_R101
                    Site.CaptureRequirementIfAreEqual(
                        SoapErrorCode.ErrorCode0x80070002,
                        this.GetErrorCode(exp.Detail.LastChild.InnerText),
                         101,
                    @"[In CustomizeCssResponse] [If implementation does encounter one of the error conditions described in the following table when running this operation, a SOAP fault MUST be returned that contain one of the error codes for the specified error condition.] Implementation does the Error Code equal ""0x80070002"" Specified cssFile does not exist on the protocol server.(Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the CustomizeCss operation when the cssFile contains an asterisk.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC06_CustomizeCss_ContainAsterisk()
        {
            try
            {
                Adapter.CustomizeCss(string.Format("{0}*.css", this.GenerateRandomString(5)));
                Site.Assert.Fail("The expected error code '0x81020073' is not returned for the CustomizeCss operation.");
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1042
                // COMMENT: If the server return the server faults as SOAP faults, then the requirement 
                // can be captured.
                Site.CaptureRequirement(
                    1042,
                    @"[In Transport] Protocol server faults MUST be returned SOAP faults as specified either in [SOAP1.1] section 4.4, SOAP Fault, or in [SOAP1.2/1] section 5.4, SOAP Fault.");

                if (Common.IsRequirementEnabled(102, this.Site))
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R102
                    // COMMENT: When the CustomizeCss operation is invoked when the cssFile element containing
                    // an asterisk, if the returned error code is 0x80131600, then the requirement can be captured.
                    Site.CaptureRequirementIfAreEqual<string>(
                        SoapErrorCode.ErrorCode0x81020073,
                        exp.Detail.LastChild.InnerText,
                       102,
                    @"[In CustomizeCssResponse] [If implementation does encounter one of the error conditions described in the following table when running this operation, a SOAP fault MUST be returned that contain one of the error codes for the specified error condition.] Implementation does the Error Code equal ""0x81020073"" meaning is Client specified a cssFile value containing an asterisk (*). (Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertCss operation when the user is not authorized. 
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC07_RevertCss_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);
            try
            {
                Adapter.RevertCss(Common.GetConfigurationPropertyValue("CssFile_Valid", this.Site));
                Site.Assert.Fail("The expected http status code is not returned for the RevertCss operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1084
                // COMMENT: When the RevertCss operation is invoked by unauthenticated user, if the 
                // server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1084,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[RevertCss], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify RevertCss operation to let the customizations of the context site defined by the given CSS file return to the default style.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC08_RevertCss_Succeed()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R1035
            Site.Assert.IsFalse(!Common.IsRequirementEnabled(1035, this.Site), "This operation RevertCss failed.");
            if (Common.IsRequirementEnabled(1035, this.Site))
            {
                Adapter.RevertCss(Common.GetConfigurationPropertyValue("CssFile_Valid", this.Site));
                
                // When the System Under Test product name is Windows SharePoint Services 3.0 and above, if the server returns an
                // exception when invoke RevertCss operation, then the requirement can be captured.
                Site.CaptureRequirement(
                    1035,
                    @"[In Appendix B: Product Behavior]Implementation does support this[RevertCss] operation.(<19> Windows SharePoint Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertCss operation when cssFile is blank or null.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC09_RevertCss_CssFileBlankOrNull()
        {
            try
            {
                Adapter.RevertCss(string.Empty);
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertCss operation.");
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R431
                // COMMENT: When the RevertCss operation is invoked when the cssFile element is blank or
                // null, if the returned error code is 0x82000001, then the requirement can be captured.
                Site.CaptureRequirementIfAreEqual<string>(
                    SoapErrorCode.ErrorCode0x82000001,
                    exp.Detail.LastChild.InnerText,
                    431,
                    @"[In RevertCssResponse] In case the server encounters one of the following error conditions[Blank or null cssFile specified] during the execution of this operation[RevertCss], a SOAP fault MUST be returned[Error Code: 0x82000001] .");

                // Verify MS-WEBSS requirement: MS-WEBSS_R811
                if (Common.IsRequirementEnabled(811, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<string>(
                    SoapErrorCode.ErrorCode0x82000001,
                    exp.Detail.LastChild.InnerText,
                    811,
                    @"[In RevertCssResponse] If implementation does encounter one of the following error conditions while running this operation[RevertCss], a SOAP fault MUST be returned that contain one of the error codes[0x82000001] in the following table for the specified error condition[Blank or null cssFile specified] in products[
  The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013].");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertCss operation when cssFile does not exist on server. 
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S08_TC10_CustomizeCssFileMatchName()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R90    
            string soapErrorCode = string.Empty;
            try
            {
                Adapter.CustomizeCss(this.GenerateRandomString(5));
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R90
                // COMMENT: When the CustomizeCss operation is invoked when the cssFile element has no file 
                // extension, if the returned error code is 0x80131600, then the requirement can be captured.
                Site.CaptureRequirementIfAreEqual<string>(
                    SoapErrorCode.ErrorCode0x80131600,
                    soapErrorCode = exp.Detail.LastChild.InnerText,
                    90,
                    @"[In CustomizeCss] cssFile: The cssFile input parameter MUST specify the name of one of the CSS files that resides in the default central location on the protocol server.");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R92
            bool isVerifiedR92 = false;
            try
            {
                Adapter.CustomizeCss(string.Format("{0}.doc", this.GenerateRandomString(5)));
                Site.Assert.Fail("The expected including the file extension '.css' is not returned for the CustomizeCss operation.");
            }
            catch (SoapException exp)
            {
                if (exp.Detail.InnerText.Contains(".css"))
                {
                    isVerifiedR92 = true;
                }
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R92
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR92,
                92,
                @"[In CustomizeCss] This file[cssFile] name MUST match the file name on the protocol server, including the file extension "".css"", for example: ""core.css"".");
        }
    }
}