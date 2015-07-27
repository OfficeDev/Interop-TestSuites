//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This test class is used to test GetServerInfo operation.
    /// </summary>
    [TestClass]
    public class S03_GetServerInfo : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear up the class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion

        /// <summary>
        /// This test case is used to test GetServerInfo on a repository that is configured for routing content.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S03_TC01_GetServerInfo()
        {
            // Initial parameters to use a repository that is configured for routing content
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Retrieves data about the type, version of the repository 
            // we just add one hold in server for test
            ServerInfo serverInfo = this.Adapter.GetServerInfo();

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R186
            // Not null means adapter has checked structure of message, so this can be captured.
            Site.CaptureRequirementIfIsNotNull(
                     serverInfo,
                     "MS-OFFICIALFILE",
                     186,
                     @"[In GetServerInfo] The protocol client sends a GetServerInfoSoapIn request WSDL message, and the protocol server MUST respond with a GetServerInfoSoapOut response WSDL message.");

            if (Common.Common.IsRequirementEnabled(3702, this.Site))
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R3702
                // Not null means adapter has checked structure of message, so this can be captured.
                Site.CaptureRequirementIfIsNull(
                         serverInfo.RoutingWeb,
                         "MS-OFFICIALFILE",
                         3702,
                         @"[In Appendix C: Product Behavior] Implementation does not include RoutingWeb element. <4> Section 2.2.4.4:  Office SharePoint Server 2007 does not use this element [RoutingWeb].");
            }

            if (Common.Common.IsRequirementEnabled(1031, this.Site))
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1031
                Site.CaptureRequirementIfIsNotNull(
                         serverInfo.RoutingWeb,
                         "MS-OFFICIALFILE",
                         1031,
                         @"[In Appendix C: Product Behavior] Implementation does include RoutingWeb element. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
            }

            if (Common.Common.IsRequirementEnabled(268, this.Site))
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R268
                bool isServerTypeVerified = string.Compare("Microsoft.Office.Server", serverInfo.ServerType, true) == 0;
                bool isServerVersionVerified = string.Compare("Microsoft.Office.OfficialFileSoap, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c", serverInfo.ServerVersion, true) == 0;

                Site.Assert.IsTrue(
                    isServerTypeVerified,
                    "When MS-OFFICIALFILE_R268 is enabled, the ServerType should return Microsoft.Office.Server, actual result is {0}",
                    serverInfo.ServerType);

                Site.Assert.IsTrue(
                    isServerVersionVerified,
                    "When MS-OFFICIALFILE_R268 is enabled, the ServerVersion should return Microsoft.Office.OfficialFileSoap, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c, actual result is {0}",
                    serverInfo.ServerVersion);

                Site.CaptureRequirementIfIsTrue(
                         isServerTypeVerified && isServerVersionVerified,
                         "MS-OFFICIALFILE",
                         268,
                         @"[In Appendix C: Product Behavior] Implementation does return implementation-specific information in the ServerInfo element. <9> Section 3.1.4.5:  Office SharePoint Server 2007 returns ""Microsoft.Office.Server"" as the ServerType and ""Microsoft.Office.OfficialFileSoap, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" as the ServerVersion.");
            }

            if (Common.Common.IsRequirementEnabled(269, this.Site))
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R269
                bool isServerTypeVerified = string.Compare("Microsoft.Office.Server v4", serverInfo.ServerType, true) == 0;
                bool isServerVersionVerified = string.Compare("Microsoft.Office.OfficialFileSoap, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c", serverInfo.ServerVersion, true) == 0;

                Site.Assert.IsTrue(
                    isServerTypeVerified,
                    "When MS-OFFICIALFILE_R269 is enabled, the ServerType should return Microsoft.Office.Server v4, actual result is {0}",
                    serverInfo.ServerType);

                Site.Assert.IsTrue(
                    isServerVersionVerified,
                    "When MS-OFFICIALFILE_R269 is enabled, the ServerVersion should return Microsoft.Office.OfficialFileSoap, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c, actual result is {0}",
                    serverInfo.ServerVersion);

                Site.CaptureRequirementIfIsTrue(
                         isServerTypeVerified && isServerVersionVerified,
                         "MS-OFFICIALFILE",
                         269,
                         @"[In Appendix C: Product Behavior] Implementation does return implementation-specific information in the ServerInfo element. [<9> Section 3.1.4.5:] SharePoint Server 2010 returns ""Microsoft.Office.Server v4"" as the ServerType and "" Microsoft.Office.OfficialFileSoap, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" as the ServerVersion.");
            }

            if (Common.Common.IsRequirementEnabled(357, this.Site))
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R357
                bool isServerTypeVerified = string.Compare("Microsoft.Office.Server v4", serverInfo.ServerType, true) == 0;
                bool isServerVersionVerified = string.Compare("Microsoft.Office.OfficialFileSoap, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c", serverInfo.ServerVersion, true) == 0;

                Site.Assert.IsTrue(
                    isServerTypeVerified,
                    "When MS-OFFICIALFILE_R357 is enabled, the ServerType should return Microsoft.Office.Server v4, actual result is {0}",
                    serverInfo.ServerType);

                Site.Assert.IsTrue(
                    isServerVersionVerified,
                    "When MS-OFFICIALFILE_R357 is enabled, the ServerVersion should return Microsoft.Office.OfficialFileSoap, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c, actual result is {0}",
                    serverInfo.ServerVersion);

                Site.CaptureRequirementIfIsTrue(
                         isServerTypeVerified && isServerVersionVerified,
                         "MS-OFFICIALFILE",
                         357,
                         @"[In Appendix C: Product Behavior] Implementation does return implementation-specific information in the ServerInfo element. [<9> Section 3.1.4.5:] SharePoint Server 2013 returns ""Microsoft.Office.Server v4"" as the ServerType and "" Microsoft.Office.OfficialFileSoap, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" as the ServerVersion.");
            }
        }
    }
}
