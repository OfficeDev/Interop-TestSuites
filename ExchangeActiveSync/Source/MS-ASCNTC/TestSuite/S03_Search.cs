//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCNTC
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to use Search command to search Contact class data on the server.
    /// </summary>
    [TestClass]
    public class S03_Search : TestSuiteBase
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

        #region MSASCNTC_S03_TC01_Search_RetrieveContact
        /// <summary>
        /// This case is designed to retrieve contact using Search command.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S03_TC01_Search_RetrieveContact()
        {
            #region Call Sync command with Add element to add a contact with all Contact class elements to the server
            string picture = "SmallPhoto.jpg";

            Dictionary<Request.ItemsChoiceType8, object> contactProperties = this.SetContactProperties(picture);
            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());
            #endregion

            #region Call Search command to retrieve the contact item that added in previous step
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 1 };

            Search searchItem = this.GetSearchResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, bodyPreference);
            #endregion

            #region Verify requirements
            this.VerifyContactClassElements(contactProperties, searchItem.Contact);
            #endregion
        }
        #endregion
    }
}