namespace Microsoft.Protocols.TestSuites.MS_ASCNTC
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to use the Sync command to synchronize the Contact class data between client and server.
    /// </summary>
    [TestClass]
    public class S01_Sync : TestSuiteBase
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

        #region MSASCNTC_S01_TC01_Sync_AddContact
        /// <summary>
        /// This case is designed to use Sync Add operation to add a contact.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC01_Sync_AddContact()
        {
            #region Call Sync command with Add element to add a contact with all Contact class elements to the server.
            string picture = "SmallPhoto.jpg";

            Dictionary<Request.ItemsChoiceType8, object> contactProperties = this.SetContactProperties(picture);
            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());
            #endregion

            #region Call Sync command to synchronize the contact item that added in previous step
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 1 };

            // Get the new added contact
            Sync newAddedItem = this.GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, bodyPreference, null);

            Site.Assert.IsNull(
                newAddedItem.Contact.WeightedRank,
                "The Sync response should not contain WeightedRank since it is only returned in a recipient information cache response.");
            #endregion

            #region Verify requirements
            this.VerifyContactClassElements(contactProperties, newAddedItem.Contact);
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC02_Sync_GhostedElements_ExceptAssistantName
        /// <summary>
        /// This case is designed to test ghosted elements except AssistantName element.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC02_Sync_GhostedElements_ExceptAssistantName()
        {
            #region Call Sync command with Add element to add a contact to the server
            Dictionary<Request.ItemsChoiceType8, object> contactProperties = this.SetContactProperties(null);
            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());
            #endregion

            #region Call Sync command with Supported element in initial Sync request to synchronize the contact item that added in previous step
            // Put the AssistantName into Supported element
            Request.Supported supportedElements = new Request.Supported
            {
                Items = new string[] { string.Empty },
                ItemsElementName = new Request.ItemsChoiceType[] { Request.ItemsChoiceType.AssistantName }
            };

            // Get the new added contact
            Sync newAddedItem = this.GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, null, supportedElements);
            #endregion

            #region Call Sync command with Change element to change the AssistantName value of the contact
            Request.SyncCollectionChange changeData = new Request.SyncCollectionChange
            {
                ApplicationData = new Request.SyncCollectionChangeApplicationData
                {
                    Items = new object[] { "EditedAssistantName" },
                    ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.AssistantName }
                },
                ServerId = newAddedItem.ServerId
            };

            this.UpdateContact(this.SyncKey, this.User1Information.ContactsCollectionId, changeData);
            #endregion

            #region Call Sync command to synchronize the changed contact on the server
            // Get the updated contact
            Sync changedItem = this.GetSyncChangeResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, this.SyncKey, null);

            Site.Assert.AreEqual<string>(
                "EditedAssistantName",
                changedItem.Contact.AssistantName,
                "The value of AssistantName should be changed.");
            #endregion

            #region Verify requirements
            // If the value of the ghosted elements except the one added into Supported element in Sync change response equals the value in Sync.Add response, it means the existing value for these elements are preserved, then the ghosted elements can be verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R100");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R100
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.AccountName,
                changedItem.Contact.AccountName,
                100,
                @"[In AccountName] This element[contacts2:AccountName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R110");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R110
            Site.CaptureRequirementIfAreEqual<DateTime?>(
                newAddedItem.Contact.Anniversary,
                changedItem.Contact.Anniversary,
                110,
                @"[In Anniversary] This element[Anniversary] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R120");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R120
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.AssistantPhoneNumber,
                changedItem.Contact.AssistantPhoneNumber,
                120,
                @"[In AssistantPhoneNumber] This element[AssistantPhoneNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R125");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R125
            Site.CaptureRequirementIfAreEqual<DateTime?>(
                newAddedItem.Contact.Birthday,
                changedItem.Contact.Birthday,
                125,
                @"[In Birthday] This element[Birthday] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R133");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R133
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.BusinessAddressCity,
                changedItem.Contact.BusinessAddressCity,
                133,
                @"[In BusinessAddressCity] This element[BusinessAddressCity] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R138");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R138
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.BusinessAddressCountry,
                changedItem.Contact.BusinessAddressCountry,
                138,
                @"[In BusinessAddressCountry] This element [BusinessAddressCountry]can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R700");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R700
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.BusinessAddressPostalCode,
                changedItem.Contact.BusinessAddressPostalCode,
                700,
                @"[In BusinessAddressPostalCode] This element[BusinessAddressPostalCode] can be ghosted. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R146");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R146
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.BusinessAddressState,
                changedItem.Contact.BusinessAddressState,
                146,
                @"[In BusinessAddressState] This element[BusinessAddressState] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R151");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R151
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.BusinessAddressStreet,
                changedItem.Contact.BusinessAddressStreet,
                151,
                @"[In BusinessAddressStreet] This element[BusinessAddressStreet] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R156");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R156
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.BusinessFaxNumber,
                changedItem.Contact.BusinessFaxNumber,
                156,
                @"[In BusinessFaxNumber] This element[BusinessFaxNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R161");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R161
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.BusinessPhoneNumber,
                changedItem.Contact.BusinessPhoneNumber,
                161,
                @"[In BusinessPhoneNumber] This element[BusinessPhoneNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R166");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R166
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Business2PhoneNumber,
                changedItem.Contact.Business2PhoneNumber,
                166,
                @"[In Business2PhoneNumber] This element[Business2PhoneNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R171");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R171
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.CarPhoneNumber,
                changedItem.Contact.CarPhoneNumber,
                171,
                @"[In CarPhoneNumber] This element[CarPhoneNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R177");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R177
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Categories.Category[0],
                changedItem.Contact.Categories.Category[0],
                177,
                @"[In Categories] This element[Categories] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R188");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R188
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Children.Child[0],
                changedItem.Contact.Children.Child[0],
                188,
                @"[In Children] This element[Children] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R198");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R198
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.CompanyMainPhone,
                changedItem.Contact.CompanyMainPhone,
                198,
                @"[In CompanyMainPhone] This element[contacts2:CompanyMainPhone] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R203");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R203
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.CompanyName,
                changedItem.Contact.CompanyName,
                203,
                @"[In CompanyName] This element[CompanyName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R208");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R208
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.CustomerId,
                changedItem.Contact.CustomerId,
                208,
                @"[In CustomerId] This element[contacts2:CustomerId] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R212");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R212
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Department,
                changedItem.Contact.Department,
                212,
                @"[In Department] This element[Department] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R217");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R217
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Email1Address,
                changedItem.Contact.Email1Address,
                217,
                @"[In Email1Address] This element[Email1Address] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R224");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R224
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Email2Address,
                changedItem.Contact.Email2Address,
                224,
                @"[In Email2Address] This element[Email2Address] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R229");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R229
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Email3Address,
                changedItem.Contact.Email3Address,
                229,
                @"[In Email3Address] This element[Email3Address] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R234");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R234
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.FileAs,
                changedItem.Contact.FileAs,
                234,
                @"[In FileAs] This element[FileAs] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R241");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R241
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.FirstName,
                changedItem.Contact.FirstName,
                241,
                @"[In FirstName] This element[FirstName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R246");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R246
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.GovernmentId,
                changedItem.Contact.GovernmentId,
                246,
                @"[In GovernmentId] This element[contacts2:GovernmentId] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R251");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R251
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.HomeAddressCity,
                changedItem.Contact.HomeAddressCity,
                251,
                @"[In HomeAddressCity] This element[HomeAddressCity] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R256");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R256
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.HomeAddressCountry,
                changedItem.Contact.HomeAddressCountry,
                256,
                @"[In HomeAddressCountry] This element[HomeAddressCountry] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R261");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R261
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.HomeAddressPostalCode,
                changedItem.Contact.HomeAddressPostalCode,
                261,
                @"[In HomeAddressPostalCode] This element[HomeAddressPostalCode] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R266");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R266
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.HomeAddressState,
                changedItem.Contact.HomeAddressState,
                266,
                @"[In HomeAddressState] This element[HomeAddressState] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R271");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R271
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.HomeAddressStreet,
                changedItem.Contact.HomeAddressStreet,
                271,
                @"[In HomeAddressStreet] This element[HomeAddressStreet] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R276");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R276
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.HomeFaxNumber,
                changedItem.Contact.HomeFaxNumber,
                276,
                @"[In HomeFaxNumber] This element[HomeFaxNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R281");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R281
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.HomePhoneNumber,
                changedItem.Contact.HomePhoneNumber,
                281,
                @"[In HomePhoneNumber] This element[HomePhoneNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R286");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R286
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Home2PhoneNumber,
                changedItem.Contact.Home2PhoneNumber,
                286,
                @"[In Home2PhoneNumber] This element[Home2PhoneNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R291");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R291
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.IMAddress,
                changedItem.Contact.IMAddress,
                291,
                @"[In IMAddress] This element[contacts2:IMAddress] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R296");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R296
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.IMAddress2,
                changedItem.Contact.IMAddress2,
                296,
                @"[In IMAddress2] This element[contacts2:IMAddress2] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R301");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R301
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.IMAddress3,
                changedItem.Contact.IMAddress3,
                301,
                @"[In IMAddress3] This element[contacts2:IMAddress3] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R306");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R306
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.JobTitle,
                changedItem.Contact.JobTitle,
                306,
                @"[In JobTitle] This element[JobTitle] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R311");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R311
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.LastName,
                changedItem.Contact.LastName,
                311,
                @"[In LastName] This element[LastName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R316");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R316
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.ManagerName,
                changedItem.Contact.ManagerName,
                316,
                @"[In ManagerName] This element[contacts2:ManagerName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R321");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R321
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.MiddleName,
                changedItem.Contact.MiddleName,
                321,
                @"[In MiddleName] This element[MiddleName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R326");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R326
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.MMS,
                changedItem.Contact.MMS,
                326,
                @"[In MMS] This element[contacts2:MMS] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R331");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R331
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.MobilePhoneNumber,
                changedItem.Contact.MobilePhoneNumber,
                331,
                @"[In MobilePhoneNumber] This element[MobilePhoneNumber] can be ghosted.");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R336
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.NickName,
                changedItem.Contact.NickName,
                336,
                @"[In NickName] This element[contacts2:NickName] can be ghosted.");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R341
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.OfficeLocation,
                changedItem.Contact.OfficeLocation,
                341,
                @"[In OfficeLocation] This element[OfficeLocation] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R346");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R346
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.OtherAddressCity,
                changedItem.Contact.OtherAddressCity,
                346,
                @"[In OtherAddressCity] This element[OtherAddressCity] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R351");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R351
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.OtherAddressCountry,
                changedItem.Contact.OtherAddressCountry,
                351,
                @"[In OtherAddressCountry] This element[OtherAddressCountry] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R356");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R356
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.OtherAddressPostalCode,
                changedItem.Contact.OtherAddressPostalCode,
                356,
                @"[In OtherAddressPostalCode] This element[OtherAddressPostalCode] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R361");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R361
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.OtherAddressState,
                changedItem.Contact.OtherAddressState,
                361,
                @"[In OtherAddressState] This element[OtherAddressState] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R366");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R366
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.OtherAddressStreet,
                changedItem.Contact.OtherAddressStreet,
                366,
                @"[In OtherAddressStreet] This element[OtherAddressStreet] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R371");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R371
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.PagerNumber,
                changedItem.Contact.PagerNumber,
                371,
                @"[In PagerNumber] This element[PagerNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R388");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R388
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.RadioPhoneNumber,
                changedItem.Contact.RadioPhoneNumber,
                388,
                @"[In RadioPhoneNumber] This element[RadioPhoneNumber] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R393");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R393
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Spouse,
                changedItem.Contact.Spouse,
                393,
                @"[In Spouse] This element[Spouse] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R398");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R398
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Suffix,
                changedItem.Contact.Suffix,
                398,
                @"[In Suffix] This element[Suffix] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R403");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R403
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Title,
                changedItem.Contact.Title,
                403,
                @"[In Title] This element[Title] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R408");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R408
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.WebPage,
                changedItem.Contact.WebPage,
                408,
                @"[In WebPage] This element[WebPage] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R420");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R420
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.YomiCompanyName,
                changedItem.Contact.YomiCompanyName,
                420,
                @"[In YomiCompanyName] This element[YomiCompanyName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R425");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R425
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.YomiFirstName,
                changedItem.Contact.YomiFirstName,
                425,
                @"[In YomiFirstName] This element[YomiFirstName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R430");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R430
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.YomiLastName,
                changedItem.Contact.YomiLastName,
                430,
                @"[In YomiLastName] This element[YomiLastName] can be ghosted.");

            // If all above requirements can be captured successfully, then requirement MS-ASCNTC_R499 can be captured directly.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R499");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R499
            Site.CaptureRequirement(
                499,
                @"[In Omitting Ghosted Properties from a Sync Change Request] Instead of deleting these excluded properties [Ghosted elements], the server preserves their previous value.");
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC03_Sync_GhostedElement_AssistantName
        /// <summary>
        /// This case is designed to test the AssistantName element which can be ghosted.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC03_Sync_GhostedElement_AssistantName()
        {
            #region Call Sync command with Add element to add a contact to the server
            Dictionary<Request.ItemsChoiceType8, object> contactProperties = this.SetContactProperties(null);
            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());
            #endregion

            #region Call Sync command with Supported element in initial Sync request to synchronize the contact item that added in previous step
            // Put JobTitle into the Supported element
            Request.Supported supportedElements = new Request.Supported
            {
                Items = new string[] { string.Empty },
                ItemsElementName = new Request.ItemsChoiceType[] { Request.ItemsChoiceType.JobTitle }
            };

            // Set the BodyPreference element in Sync command request
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 1 };

            // Get the new added contact
            Sync newAddedItem = this.GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, bodyPreference, supportedElements);
            #endregion

            #region Call Sync command with Change element to change the JobTitle value of the contact
            Request.SyncCollectionChange changeData = new Request.SyncCollectionChange
            {
                ApplicationData = new Request.SyncCollectionChangeApplicationData
                {
                    Items = new object[] { "EditedJobTitle" },
                    ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.JobTitle }
                },
                ServerId = newAddedItem.ServerId
            };

            this.UpdateContact(this.SyncKey, this.User1Information.ContactsCollectionId, changeData);
            #endregion

            #region Call Sync command with Supported element in initial Sync request to synchronize the changed contact on the server
            // Get the updated contact
            Sync changedItem = this.GetSyncChangeResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, this.SyncKey, bodyPreference);

            Site.Assert.AreEqual<string>(
                "EditedJobTitle",
                changedItem.Contact.JobTitle,
                "The value of JobTitle should be changed.");
            #endregion

            #region Verify requirements
            // If the value of the AssistantName element in the Sync Change response equals the value in the Sync Add response, it means the existing value for the AssistantName element is preserved, then requirements MS-ASCNTC_R461 and MS-ASCNTC_R115 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R115");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R115
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.AssistantName,
                changedItem.Contact.AssistantName,
                115,
                @"[In AssistantName] This element[AssistantName] can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R461");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R461
            Site.CaptureRequirementIfAreEqual<string>(
                newAddedItem.Contact.Body.Data,
                changedItem.Contact.Body.Data,
                461,
                @"[In Truncating the Contact Notes Field]If an airsyncbase:Body element is not included in the request that is sent from the client to the server, the server MUST NOT delete the stored Notes for the contact.");
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC04_Sync_RefreshRecipientInformationCache
        /// <summary>
        /// This case is designed to retrieve a minimal set of Contact class data from the server by issuing a Sync command request against the recipient information cache.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC04_Sync_RefreshRecipientInformationCache()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recipient information cache is not supported with the value of the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call SendMail command to send a mail to recipient
            string subject = Common.GenerateResourceName(Site, "Subject");
            string emailBody = Common.GenerateResourceName(Site, "Body");

            this.SendEmail(subject, emailBody);

            // Make sure the email has reached the recipient's inbox folder
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.GetSyncAddResult(subject, this.User2Information.InboxCollectionId, null, null);
            #endregion

            #region Call Sync command to synchronize the recipient information cache folder of the sender
            this.SwitchUser(this.User1Information, false);

            SyncRequest request = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User1Information.RecipientInformationCacheCollectionId, null), this.User1Information.RecipientInformationCacheCollectionId, null);

            SyncStore store = this.ASCNTCAdapter.Sync(request);

            Site.Assert.AreEqual<byte>(
                1,
                store.CollectionStatus,
                "The recipient information cache folder should be successfully synchronized.");
            #endregion

            #region Verify requirements
            foreach (Sync recipientInformationCacheItem in store.AddElements)
            {
                // If the returned WeightedRank element is not null, it means this element is returned in response, then requirement MS-ASCNTC_R415 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R415.");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R415
                Site.CaptureRequirementIfIsNotNull(
                    recipientInformationCacheItem.Contact.WeightedRank,
                    415,
                    @"[In WeightedRank] The WeightedRank element is only returned in a recipient information cache response.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R410.");

                // If MS-ASCNTC_R415 can be captured successfully, then requirement MS-ASCNTC_R410 can be captured directly.
                // Verify MS-ASCNTC requirement: MS-ASCNTC_R410
                Site.CaptureRequirement(
                    410,
                    @"[In WeightedRank] The WeightedRank element<2> specifies the rank of this contact entry in the recipient information cache.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R236.");

                // If the returned FileAs element is not null, it means this element is returned in response, then requirement MS-ASCNTC_R236 can be captured.
                // Verify MS-ASCNTC requirement: MS-ASCNTC_R236
                Site.CaptureRequirementIfIsNotNull(
                    recipientInformationCacheItem.Contact.FileAs,
                    236,
                    @"[In FileAs] The FileAs element is one of the Contact class elements that is returned in a recipient information cache response.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R219.");

                // If the returned Email1Address element is not null, it means this element is returned in response, then requirement MS-ASCNTC_R219 can be captured.
                // Verify MS-ASCNTC requirement: MS-ASCNTC_R219
                Site.CaptureRequirementIfIsNotNull(
                    recipientInformationCacheItem.Contact.Email1Address,
                    219,
                    @"[In Email1Address] The Email1Address element is one of the Contact class elements that is returned in a recipient information cache response.");

                if (Common.IsRequirementEnabled(1010, this.Site))
                {
                    // If the returned Alias element is null, it means this element isn't returned in response, then requirement MS-ASCNTC_R1010 can be captured.
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R1010.");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R1010
                    Site.CaptureRequirementIfIsNull(
                        recipientInformationCacheItem.Contact.Alias,
                        1010,
                        @"[In Appendix B: Product Behavior] Implementation does not return Alias element in a recipient information cache response. (Exchange Server 2007 and above follow this behavior.)");
                }
            }
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC05_Sync_Status6_PictureExceeds48KB
        /// <summary>
        /// This case is designed to test server must return a status error of 6 if the value of the Picture element exceeds 48 KB.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC05_Sync_Status6_PictureExceeds48KB()
        {
            #region Call Sync command with Add element to add a contact with a picture whose value exceeds 48 KB
            string picture = "BigPhoto.jpg";

            Dictionary<Request.ItemsChoiceType8, object> contactProperties = this.SetContactProperties(picture);

            SyncStore syncAddStore = this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R384.");

            // If the Status in the response is 6, then requirement MS-ASCNTC_R384 can be captured.
            // Verify MS-ASCNTC requirement: MS-ASCNTC_R384
            Site.CaptureRequirementIfAreEqual<int>(
                6,
                int.Parse(syncAddStore.AddResponses[0].Status),
                384,
                @"[In Picture] If the value of the Picture element exceeds 48 KB of content with base64 encoding, the server MUST return a status error of 6.");
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC06_Sync_AddContact_300CategoryElements
        /// <summary>
        /// This case is designed to test the Categories element can have up to 300 Category elements.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC06_Sync_AddContact_300CategoryElements()
        {
            #region Call Sync command with 300 Category elements per Categories element to add a contact to the server
            string fileAs = Common.GenerateResourceName(Site, "Contact");
            Dictionary<Request.ItemsChoiceType8, object> contactProperties = new Dictionary<Request.ItemsChoiceType8, object>();
            Collection<string> categories = new Collection<string>();

            for (int i = 1; i <= 300; i++)
            {
                string category = "Category" + i;
                categories.Add(category);
            }

            Request.Categories1 contactCategories = new Request.Categories1 { Category = new string[categories.Count] };
            categories.CopyTo(contactCategories.Category, 0);
            contactProperties.Add(Request.ItemsChoiceType8.Categories1, contactCategories);
            contactProperties.Add(Request.ItemsChoiceType8.FileAs, fileAs);

            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, fileAs);
            #endregion

            #region Call Sync command to synchronize the contact added in previous step
            // Get the new added contact
            Sync newAddedItem = this.GetSyncAddResult(fileAs, this.User1Information.ContactsCollectionId, null, null);
            #endregion

            #region Verify requirement
            // If the Sync Add operation is successful and the count of Category is 300, it means the contact can have 300 Category elements, then requirement MS-ASCNTC_R183 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R183.");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R183
            Site.CaptureRequirementIfAreEqual<int>(
                300,
                newAddedItem.Contact.Categories.Category.Length,
                183,
                @"[In Category] It[Categories] can have up to 300 elements[Category] per Categories element.");
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC07_Sync_AddContact_300ChildElements
        /// <summary>
        /// This case is designed to test the Children element can have up to 300 Child elements.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC07_Sync_AddContact_300ChildElements()
        {
            #region Call Sync command with 300 Child elements per Children element to add a contact to the server
            string fileAs = Common.GenerateResourceName(Site, "Contact");
            Dictionary<Request.ItemsChoiceType8, object> contactProperties = new Dictionary<Request.ItemsChoiceType8, object>();

            List<string> children = new List<string>();
            for (int i = 1; i <= 300; i++)
            {
                string child = "Child" + i;
                children.Add(child);
            }

            Request.Children contactChildren = new Request.Children { Child = children.ToArray() };
            contactProperties.Add(Request.ItemsChoiceType8.Children, contactChildren);
            contactProperties.Add(Request.ItemsChoiceType8.FileAs, fileAs);

            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, fileAs);
            #endregion

            #region Call Sync command to synchronize the contact added in previous step
            // Get the new added contact
            Sync newAddedItem = this.GetSyncAddResult(fileAs, this.User1Information.ContactsCollectionId, null, null);
            #endregion

            #region Verify requirement
            // If the Sync Add operation is successful and the count of Child is 300, it means the contact can have 300 Child elements, then requirement MS-ASCNTC_R194 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R194.");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R194
            Site.CaptureRequirementIfAreEqual<int>(
                300,
                newAddedItem.Contact.Children.Child.Length,
                194,
                @"[In Child] It[Children] can have up to 300 elements[Child] per Children element.");
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC08_Sync_TruncatedBody
        /// <summary>
        /// This case is designed to test server truncates the contents of the airsyncbase:Body element in the Sync command response if the client requests truncation.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC08_Sync_TruncatedBody()
        {
            #region Call Sync command with Add element to add a contact to the server
            Dictionary<Request.ItemsChoiceType8, object> contactProperties = this.SetContactProperties(null);
            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());
            #endregion

            #region Call Sync command with TruncationSize element smaller than the available data of the body
            Request.BodyPreference bodyPreference = new Request.BodyPreference
            {
                Type = 1,
                TruncationSize = 8,
                TruncationSizeSpecified = true
            };

            Sync newAddedItem = this.GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, bodyPreference, null);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R512.");

            // If the body content is truncated, then requirement MS-ASCNTC_R512 can be captured.
            // Verify MS-ASCNTC requirement: MS-ASCNTC_R512
            Site.CaptureRequirementIfAreEqual<string>(
                ((Request.Body)contactProperties[Request.ItemsChoiceType8.Body]).Data.Substring(0, (int)bodyPreference.TruncationSize),
                newAddedItem.Contact.Body.Data,
                512,
                @"[In Truncating the Contact Notes Field] Once a client requests truncation, the server truncates the contents of the airsyncbase:Body element in the subsequent Sync command response.");
            #endregion
        }
        #endregion

        #region MSASCNTC_S01_TC09_Sync_NonTruncatedBody
        /// <summary>
        /// This case is designed to test server will no longer truncate the contents of the airsyncbase:Body element if the client sends an airsyncbase:BodyPreference element in the request that contains a Type element to specify the desired format, but does not include the airsyncbase:TruncationSize element.
        /// </summary>
        [TestCategory("MSASCNTC"), TestMethod()]
        public void MSASCNTC_S01_TC09_Sync_NonTruncatedBody()
        {
            #region Call Sync command with Add element to add a contact to the server
            Dictionary<Request.ItemsChoiceType8, object> contactProperties = this.SetContactProperties(null);
            this.AddContact(this.User1Information.ContactsCollectionId, contactProperties);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());
            #endregion

            #region Call Sync command without TruncationSize element to synchronize the new added contact
            // Set Type to 1 to get the plain text format content
            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 1 };

            Sync newAddedItem = this.GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), this.User1Information.ContactsCollectionId, bodyPreference, null);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R460.");

            // If the body content is not truncated, then requirement MS-ASCNTC_R460 can be captured.
            // Verify MS-ASCNTC requirement: MS-ASCNTC_R460
            Site.CaptureRequirementIfAreEqual<string>(
                ((Request.Body)contactProperties[Request.ItemsChoiceType8.Body]).Data,
                newAddedItem.Contact.Body.Data,
                460,
                @"[In Truncating the Contact Notes Field] A client can request that the server no longer truncate the contents of the airsyncbase:Body element by sending an airsyncbase:BodyPreference element ([MS-ASAIRS] section 2.2.2.7) in the request that contains a Type element ([MS-ASAIRS] section 2.2.2.22.4) to specify the desired format, but does not include the airsyncbase:TruncationSize element.");
            #endregion
        }
        #endregion
    }
}