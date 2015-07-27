//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test HTTP OPTIONS.
    /// </summary>
    [TestClass]
    public class S04_HTTPOPTIONSMessage : TestSuiteBase
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
        /// This test case is intended to validate the HTTP OPTIONS command.
        /// </summary>
        [TestCategory("MSASHTTP"), TestMethod()]
        public void MSASHTTP_S04_TC01_HTTPOPTIONS()
        {
            #region Call HTTP OPTIONS command.
            OptionsResponse optionsResponse = this.HTTPAdapter.HTTPOPTIONS();
            Site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, optionsResponse.StatusCode, "The StatusCode of HTTP OPTIONS command response should be OK, actual {0}.", optionsResponse.StatusCode);
            #endregion
        }
        #endregion
    }
}