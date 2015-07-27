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
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Gets or sets the information of User1.
        /// </summary>
        protected UserInformation User1Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User2.
        /// </summary>
        protected UserInformation User2Information { get; set; }

        /// <summary>
        /// Gets the instance of the ASCNTC adapter interface. 
        /// </summary>
        protected IMS_ASCNTCAdapter ASCNTCAdapter { get; private set; }

        /// <summary>
        /// Gets the value of syncKey.
        /// </summary>
        protected string SyncKey { get; private set; }

        #endregion

        #region Test case initialize and cleanup
        /// <summary>
        /// Initialize the test case
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            if (this.ASCNTCAdapter == null)
            {
                this.ASCNTCAdapter = this.Site.GetAdapter<IMS_ASCNTCAdapter>();
            }

            // Get the information of User1.
            this.User1Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User1Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            // Get the information of User2.
            this.User2Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User2Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User2Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "12.1"))
            {
                this.SwitchUser(this.User1Information, true);
            }
        }

        /// <summary>
        /// Clean up the test case
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();

            if (this.User1Information.UserCreatedItems.Count != 0)
            {
                // Switch to User1
                this.SwitchUser(this.User1Information, false);
                this.DeleteItemsInFolder(this.User1Information.UserCreatedItems);
            }

            if (this.User2Information.UserCreatedItems.Count != 0)
            {
                // Switch to User2
                this.SwitchUser(this.User2Information, false);
                this.DeleteItemsInFolder(this.User2Information.UserCreatedItems);
            }
        }
        #endregion

        #region Switch user
        /// <summary>
        /// This method is used to change user to call ActiveSync commands and resynchronize the folder collection hierarchy.
        /// </summary>
        /// <param name="userInformation">The information of the user that will switch to.</param>
        /// <param name="isFolderSyncNeeded">A boolean value indicates whether needs to synchronize the folder hierarchy.</param>
        protected void SwitchUser(UserInformation userInformation, bool isFolderSyncNeeded)
        {
            this.ASCNTCAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

            if (isFolderSyncNeeded)
            {
                FolderSyncResponse folderSyncResponse = this.FolderSync();

                // Get the folder collectionId of User1
                if (userInformation.UserName == this.User1Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User1Information.InboxCollectionId))
                    {
                        this.User1Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.ContactsCollectionId))
                    {
                        this.User1Information.ContactsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Contacts, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.RecipientInformationCacheCollectionId))
                    {
                        this.User1Information.RecipientInformationCacheCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.RecipientInformationCache, this.Site);
                    }
                }

                // Get the folder collectionId of User2
                if (userInformation.UserName == this.User2Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User2Information.InboxCollectionId))
                    {
                        this.User2Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }
                }
            }
        }
        #endregion

        #region Send mail
        /// <summary>
        /// The method is used to send a mail
        /// </summary>
        /// <param name="subject">The subject of the mail.</param>
        /// <param name="body">The body of the item.</param>
        protected void SendEmail(string subject, string body)
        {
            SendMailRequest request = new SendMailRequest
            {
                RequestData =
                {
                    ClientId = Guid.NewGuid().ToString("N"),
                    Mime = TestSuiteHelper.CreateMIME(
                           Common.GetMailAddress(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site)),
                           Common.GetMailAddress(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site)),
                           subject,
                           body)
                }
            };

            SendMailResponse sendMailResponse = this.ASCNTCAdapter.SendMail(request);
            this.Site.Assert.AreEqual<string>(
                 string.Empty,
                 sendMailResponse.ResponseDataXML,
                 "The server should return an empty XML body to indicate SendMail command executes successfully.");
        }
        #endregion

        #region Get initial SyncKey
        /// <summary>
        /// This method is used to get the initial syncKey of the specified folder.
        /// </summary>
        /// <param name="collectionId">The collection id of the specified folder.</param>
        /// <param name="supportedElements">The elements in Supported element.</param>
        /// <returns>The initial syncKey of the specified folder.</returns>
        protected string GetInitialSyncKey(string collectionId, Request.Supported supportedElements)
        {
            // Obtains the key by sending an initial Sync request with a SyncKey element value of zero and the CollectionId element
            SyncRequest syncRequest = TestSuiteHelper.CreateInitialSyncRequest(collectionId, supportedElements);
            DataStructures.SyncStore syncResult = this.ASCNTCAdapter.Sync(syncRequest);

            // Status code '12' means the folder hierarchy has changed
            while (syncResult.Status == 12)
            {
                // Resynchronize the folder hierarchy
                this.FolderSync();

                syncResult = this.ASCNTCAdapter.Sync(syncRequest);
            }

            this.Site.Assert.IsNotNull(
                syncResult,
                "The result for an initial synchronize should not be null.");

            // Verify Sync result
            this.Site.Assert.AreEqual<byte>(
                1,
                syncResult.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            return syncResult.SyncKey;
        }
        #endregion

        #region Get commands execution result
        /// <summary>
        /// This method is used to synchronize item with specified subject.
        /// </summary>
        /// <param name="subject">The subject of the item.</param>
        /// <param name="collectionId">The collection id which to sync with.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <param name="supported">The elements in Supported element.</param>
        /// <returns>The item with specified subject.</returns>
        protected DataStructures.Sync GetSyncAddResult(string subject, string collectionId, Request.BodyPreference bodyPreference, Request.Supported supported)
        {
            DataStructures.SyncStore syncStore;
            DataStructures.Sync item = null;
            SyncRequest request = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(collectionId, supported), collectionId, bodyPreference);

            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                Thread.Sleep(waitTime);
                syncStore = this.ASCNTCAdapter.Sync(request);
                if (syncStore != null && syncStore.CollectionStatus == 1)
                {
                    item = TestSuiteHelper.GetSyncAddItem(syncStore, subject);
                }

                counter++;
            }
            while ((syncStore == null || item == null) && counter < retryCount);

            this.Site.Assert.IsNotNull(item, "The email item with subject {0} should be found, retry count: {1}.", subject, counter);

            this.SyncKey = syncStore.SyncKey;

            return item;
        }

        /// <summary>
        /// This method is used to get the changed item from server.
        /// </summary>
        /// <param name="fileAs">The FileAs of the contact.</param>
        /// <param name="collectionId">The collectionId of the folder that store the contact.</param>
        /// <param name="syncKey">The syncKey of the last Sync response.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <returns>The changed item.</returns>
        protected DataStructures.Sync GetSyncChangeResult(string fileAs, string collectionId, string syncKey, Request.BodyPreference bodyPreference)
        {
            DataStructures.SyncStore syncStore;
            DataStructures.Sync item = null;
            SyncRequest request = TestSuiteHelper.CreateSyncRequest(syncKey, collectionId, bodyPreference);

            syncStore = this.ASCNTCAdapter.Sync(request);
            if (syncStore != null && syncStore.CollectionStatus == 1)
            {
                item = TestSuiteHelper.GetSyncChangeItem(syncStore, fileAs);
            }

            this.Site.Assert.IsNotNull(item, "The email item with FileAs: {0} should be found.", fileAs);

            return item;
        }

        /// <summary>
        /// This method is used to fetch item with specified serverId on the server.
        /// </summary>
        /// <param name="collectionId">The collection id.</param>
        /// <param name="serverId">The server id of the mail.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <param name="schema">The Schema of the item to be fetched.</param>
        /// <returns>The item retrieved from server.</returns>
        protected DataStructures.ItemOperations GetItemOperationsResult(string collectionId, string serverId, Request.BodyPreference bodyPreference, Request.Schema schema)
        {
            DataStructures.ItemOperationsStore itemOperationsStore;
            DataStructures.ItemOperations item = null;
            ItemOperationsRequest request = TestSuiteHelper.CreateItemOperationsRequest(collectionId, serverId, bodyPreference, schema);

            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                Thread.Sleep(waitTime);
                itemOperationsStore = this.ASCNTCAdapter.ItemOperations(request, DeliveryMethodForFetch.Inline);

                // Since the item serverId is unique, there should be only one item in response
                this.Site.Assert.AreEqual<int>(
                    1,
                    itemOperationsStore.Items.Count,
                    "The count of Items in ItemOperations command response should be 1.");

                if (itemOperationsStore.Items[0].Contact != null)
                {
                    item = itemOperationsStore.Items[0];
                }
            }
            while (item == null && counter < retryCount);

            this.Site.Assert.IsNotNull(item, "The item should be found, retry count: {0}.", counter);

            return item;
        }

        /// <summary>
        /// Search item with specified criteria on the server.
        /// </summary>
        /// <param name="fileAs">The FileAs of the item.</param>
        /// <param name="collectionId">The collection id.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <returns>The server response.</returns>
        protected DataStructures.Search GetSearchResult(string fileAs, string collectionId, Request.BodyPreference bodyPreference)
        {
            SearchRequest request = TestSuiteHelper.CreateSearchRequest(fileAs, collectionId, bodyPreference);

            DataStructures.SearchStore searchStore = this.ASCNTCAdapter.Search(request);
            DataStructures.Search searchItem = null;
            if (searchStore.Results.Count != 0)
            {
                searchItem = TestSuiteHelper.GetSearchItem(searchStore, fileAs);
            }

            this.Site.Assert.IsNotNull(searchItem, "The contact item with FileAs: {0} should be found.", fileAs);

            return searchItem;
        }
        #endregion

        #region Set contact properties
        /// <summary>
        /// This method is used to set value for contact properties.
        /// </summary>
        /// <param name="picture">The picture of the contact item.</param>
        /// <returns>The value for contact properties.</returns>
        protected Dictionary<Request.ItemsChoiceType8, object> SetContactProperties(string picture)
        {
            Request.Body contactBody = new Request.Body { Type = 1, Data = Common.GenerateResourceName(this.Site, "Body") };
            List<string> categories = new List<string> { this.Site.DefaultProtocolDocShortName + "Category" };
            Request.Categories1 contactCategories = new Request.Categories1 { Category = categories.ToArray() };
            List<string> children = new List<string> { Common.GenerateResourceName(this.Site, "Child") };
            Request.Children contactChildren = new Request.Children { Child = children.ToArray() };

            Dictionary<Request.ItemsChoiceType8, object> contactProperties = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.FileAs, Common.GenerateResourceName(this.Site, "Contact")
                },
                {
                    Request.ItemsChoiceType8.AccountName, "Contact1"
                },
                {
                    Request.ItemsChoiceType8.Anniversary, DateTime.Parse("2012-10-01")
                },
                {
                    Request.ItemsChoiceType8.AssistantName, "Assistant"
                },
                {
                    Request.ItemsChoiceType8.AssistantPhoneNumber, "54689733"
                },
                {
                    Request.ItemsChoiceType8.Birthday, DateTime.Parse("1980-04-12")
                },
                {
                    Request.ItemsChoiceType8.Body, contactBody
                },
                {
                    Request.ItemsChoiceType8.BusinessPhoneNumber, "987654321"
                },
                {
                    Request.ItemsChoiceType8.Business2PhoneNumber, "123456789"
                },
                {
                    Request.ItemsChoiceType8.BusinessAddressCity, "Beijing"
                },
                {
                    Request.ItemsChoiceType8.BusinessAddressCountry, "China"
                },
                {
                    Request.ItemsChoiceType8.BusinessAddressPostalCode, "100080"
                },
                {
                    Request.ItemsChoiceType8.BusinessAddressState, "Haidian"
                },
                {
                    Request.ItemsChoiceType8.BusinessAddressStreet, "Danling"
                },
                {
                    Request.ItemsChoiceType8.BusinessFaxNumber, "0123456"
                },
                {
                    Request.ItemsChoiceType8.CarPhoneNumber, "123456789"
                },
                {
                    Request.ItemsChoiceType8.Categories1, contactCategories
                },
                {
                    Request.ItemsChoiceType8.Children, contactChildren
                },
                {
                    Request.ItemsChoiceType8.CompanyMainPhone, "23456987"
                },
                {
                    Request.ItemsChoiceType8.CompanyName, "Microsoft"
                },
                {
                    Request.ItemsChoiceType8.CustomerId, "Customer1"
                },
                {
                    Request.ItemsChoiceType8.Department, "Testing"
                },

                // Domain name "contoso.com" is just a sample domain name. It does not have any relationship to the domain name configured in common ptfconfig file.
                {
                    Request.ItemsChoiceType8.Email1Address, "Email1@contoso.com"
                },
                {
                    Request.ItemsChoiceType8.Email2Address, "Email2@contoso.com"
                },
                {
                    Request.ItemsChoiceType8.Email3Address, "Email3@contoso.com"
                },
                {
                    Request.ItemsChoiceType8.FirstName, "Aaron"
                },
                {
                    Request.ItemsChoiceType8.GovernmentId, "99910819810809502X"
                },
                {
                    Request.ItemsChoiceType8.Home2PhoneNumber, "1234568742"
                },
                {
                    Request.ItemsChoiceType8.HomeAddressCity, "Princeton"
                },
                {
                    Request.ItemsChoiceType8.HomeAddressCountry, "United States"
                },
                {
                    Request.ItemsChoiceType8.HomeAddressPostalCode, "08989"
                },
                {
                    Request.ItemsChoiceType8.HomeAddressState, "New Jersey"
                },
                {
                    Request.ItemsChoiceType8.HomeAddressStreet, "Nassauy"
                },
                {
                    Request.ItemsChoiceType8.HomeFaxNumber, "012345678"
                },
                {
                    Request.ItemsChoiceType8.HomePhoneNumber, "014578932"
                },
                {
                    Request.ItemsChoiceType8.IMAddress, "InstantIMAddress"
                },
                {
                    Request.ItemsChoiceType8.IMAddress2, "AlternativeIMAddress"
                },
                {
                    Request.ItemsChoiceType8.IMAddress3, "TertiaryIMAddress"
                },
                {
                    Request.ItemsChoiceType8.JobTitle, "Tester"
                },
                {
                    Request.ItemsChoiceType8.LastName, "Clinton"
                },
                {
                    Request.ItemsChoiceType8.ManagerName, "Jodge"
                },
                {
                    Request.ItemsChoiceType8.MiddleName, "Jafferson"
                },
                {
                    Request.ItemsChoiceType8.MMS, "mms://enmms.chinabroadcast.cn"
                },
                {
                    Request.ItemsChoiceType8.MobilePhoneNumber, "12345678932"
                },
                {
                    Request.ItemsChoiceType8.NickName, "Ara"
                },
                {
                    Request.ItemsChoiceType8.OfficeLocation, "Beijing"
                },
                {
                    Request.ItemsChoiceType8.OtherAddressCity, "Toronto"
                },
                {
                    Request.ItemsChoiceType8.OtherAddressCountry, "Canada"
                },
                {
                    Request.ItemsChoiceType8.OtherAddressPostalCode, "K1N5H3"
                },
                {
                    Request.ItemsChoiceType8.OtherAddressState, "Ontario"
                },
                {
                    Request.ItemsChoiceType8.OtherAddressStreet, "Avenue"
                },
                {
                    Request.ItemsChoiceType8.PagerNumber, "1"
                },
                {
                    Request.ItemsChoiceType8.RadioPhoneNumber, "1234562589"
                },
                {
                    Request.ItemsChoiceType8.Spouse, "Emily"
                },
                {
                    Request.ItemsChoiceType8.Suffix, "CL."
                },
                {
                    Request.ItemsChoiceType8.Title, "Test"
                },
                {
                    Request.ItemsChoiceType8.WebPage, "http://Test.html"
                },
                {
                    Request.ItemsChoiceType8.YomiCompanyName, "Sony"
                },
                {
                    Request.ItemsChoiceType8.YomiFirstName, "Hanasaki"
                },
                {
                    Request.ItemsChoiceType8.YomiLastName, "Tsuki"
                }
            };

            if (picture != null)
            {
                contactProperties.Add(Request.ItemsChoiceType8.Picture, Convert.ToBase64String(File.ReadAllBytes(picture)));
            }

            return contactProperties;
        }
        #endregion

        #region Add contact
        /// <summary>
        /// This method is used to add a contact to server.
        /// </summary>
        /// <param name="collectionId">The collectionId of the folder which the contact is added to.</param>
        /// <param name="contactProperties">The dictionary for contact properties and its values.</param>
        /// <returns>The result returned from server for adding a contact.</returns>
        protected DataStructures.SyncStore AddContact(string collectionId, Dictionary<Request.ItemsChoiceType8, object> contactProperties)
        {
            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData
            {
                ItemsElementName = new Request.ItemsChoiceType8[contactProperties.Count],
                Items = new object[contactProperties.Count]
            };

            contactProperties.Keys.CopyTo(applicationData.ItemsElementName, 0);
            contactProperties.Values.CopyTo(applicationData.Items, 0);

            string syncKey = this.GetInitialSyncKey(collectionId, null);
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncAddRequest(syncKey, collectionId, applicationData);

            DataStructures.SyncStore syncAddStore = this.ASCNTCAdapter.Sync(syncAddRequest);
            this.Site.Assert.AreEqual<byte>(
                1,
                syncAddStore.CollectionStatus,
                "The Sync Add operation to add a contact should be successful.");

            // Since only one contact is added, the count of AddResponses should be 1.
            this.Site.Assert.AreEqual<int>(
                1,
                syncAddStore.AddResponses.Count,
                "Only one contact is added, the count of AddResponses should be 1.");

            return syncAddStore;
        }
        #endregion

        #region Update contact
        /// <summary>
        /// This method is used to update the contact properties.
        /// </summary>
        /// <param name="syncKey">The SyncKey returned form last Sync response.</param>
        /// <param name="collectionId">The collectionId of folder that the item belongs to.</param>
        /// <param name="changeData">The change data.</param>
        protected void UpdateContact(string syncKey, string collectionId, Request.SyncCollectionChange changeData)
        {
            SyncRequest syncChangeRequest = TestSuiteHelper.CreateSyncChangeRequest(syncKey, collectionId, changeData);
            DataStructures.SyncStore changeStore = this.ASCNTCAdapter.Sync(syncChangeRequest);

            this.Site.Assert.AreEqual<byte>(
                1,
                changeStore.CollectionStatus,
                "The Sync Change operation should be successful.");
        }
        #endregion

        #region Record the userName, folder collectionId and item subject
        /// <summary>
        /// Record the user name, folder collectionId and subjects the current test case impacts.
        /// </summary>
        /// <param name="userName">The user that current test case used.</param>
        /// <param name="folderCollectionId">The collectionId of folder that the current test case impact.</param>
        /// <param name="itemSubjects">The subject of items that the current test case impact.</param>
        protected void RecordCaseRelativeItems(string userName, string folderCollectionId, params string[] itemSubjects)
        {
            // Record the item in the specified folder.
            CreatedItems createdItems = new CreatedItems { CollectionId = folderCollectionId };

            foreach (string subject in itemSubjects)
            {
                createdItems.ItemSubject.Add(subject);
            }

            // Record the created items of User1.
            if (userName == this.User1Information.UserName)
            {
                this.User1Information.UserCreatedItems.Add(createdItems);
            }

            // Record the created items of User2.
            if (userName == this.User2Information.UserName)
            {
                this.User2Information.UserCreatedItems.Add(createdItems);
            }
        }
        #endregion

        #region Verify Contact class elements
        /// <summary>
        /// This method is used to verify the Contact class elements that returned from Sync, ItemOperations or Search command response.
        /// </summary>
        /// <param name="contactProperties">The contact properties.</param>
        /// <param name="contact">The contact item returned from server.</param>
        protected void VerifyContactClassElements(Dictionary<Request.ItemsChoiceType8, object> contactProperties, DataStructures.Contact contact)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R97");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R97
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.AccountName].ToString(),
                contact.AccountName,
                97,
                @"[In AccountName] The contacts2:AccountName element specifies the account name and/or number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R107");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R107
            this.Site.CaptureRequirementIfAreEqual<DateTime?>(
                DateTime.Parse(contactProperties[Request.ItemsChoiceType8.Anniversary].ToString()),
                contact.Anniversary,
                107,
                @"[In Anniversary] The Anniversary element specifies the wedding anniversary date for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R112");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R112
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.AssistantName].ToString(),
                contact.AssistantName,
                112,
                @"[In AssistantName] The AssistantName element specifies the name of the contact's assistant.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R117");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R117
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.AssistantPhoneNumber].ToString(),
                contact.AssistantPhoneNumber,
                117,
                @"[In AssistantPhoneNumber] The AssistantPhoneNumber element specifies the phone number of the contact's assistant.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R122");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R122
            this.Site.CaptureRequirementIfAreEqual<DateTime?>(
                DateTime.Parse(contactProperties[Request.ItemsChoiceType8.Birthday].ToString()),
                contact.Birthday,
                122,
                @"[In Birthday] The Birthday element specifies the birth date of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R1120");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R1120
            this.Site.CaptureRequirementIfAreEqual<string>(
                ((Request.Body)contactProperties[Request.ItemsChoiceType8.Body]).Data,
                contact.Body.Data,
                1120,
                @"[In Body] The airsyncbase:Body element [is a container ([MS-ASDTYPE] section 2.2) element that] specifies the notes for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R130");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R130
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.BusinessAddressCity].ToString(),
                contact.BusinessAddressCity,
                130,
                @"[In BusinessAddressCity] The BusinessAddressCity element specifies the business city of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R135");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R135
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.BusinessAddressCountry].ToString(),
                contact.BusinessAddressCountry,
                135,
                @"[In BusinessAddressCountry] The BusinessAddressCountry element specifies the business country/region of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R140");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R140
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.BusinessAddressPostalCode].ToString(),
                contact.BusinessAddressPostalCode,
                140,
                @"[In BusinessAddressPostalCode] The BusinessAddressPostalCode element specifies the business postal code for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R143");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R143
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.BusinessAddressState].ToString(),
                contact.BusinessAddressState,
                143,
                @"[In BusinessAddressState] The BusinessAddressState element specifies the business state for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R148");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R148
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.BusinessAddressStreet].ToString(),
                contact.BusinessAddressStreet,
                148,
                @"[In BusinessAddressStreet] The BusinessAddressStreet element specifies the business street address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R153");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R153
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.BusinessFaxNumber].ToString(),
                contact.BusinessFaxNumber,
                153,
                @"[In BusinessFaxNumber] The BusinessFaxNumber element specifies the business fax number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R158");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R158
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.BusinessPhoneNumber].ToString(),
                contact.BusinessPhoneNumber,
                158,
                @"[In BusinessPhoneNumber] The BusinessPhoneNumber element specifies the primary business phone number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R163");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R163
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.Business2PhoneNumber].ToString(),
                contact.Business2PhoneNumber,
                163,
                @"[In Business2PhoneNumber] The Business2PhoneNumber element specifies the secondary business telephone number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R168");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R168
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.CarPhoneNumber].ToString(),
                contact.CarPhoneNumber,
                168,
                @"[In CarPhoneNumber] The CarPhoneNumber element specifies the car telephone number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R1030");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R1030
            this.Site.CaptureRequirementIfIsNotNull(
                contact.Categories,
                1030,
                @"[In Categories] The Categories element [is a container ([MS-ASDTYPE] section 2.2) element that] specifies a collection of user labels assigned to the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R10040");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R10040
            this.Site.CaptureRequirementIfIsNotNull(
                contact.Children,
                10040,
                @"[In Children] The Children element [is a container ([MS-ASDTYPE] section 2.2) element that] specifies a collection of the contact's children.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R190");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R190
            this.Site.CaptureRequirementIfAreEqual<string>(
                ((Request.Children)contactProperties[Request.ItemsChoiceType8.Children]).Child[0],
                contact.Children.Child[0],
                190,
                @"[In Child] The Child element is an optional child element of the Children element that specifies a child of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R195");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R195
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.CompanyMainPhone].ToString(),
                contact.CompanyMainPhone,
                195,
                @"[In CompanyMainPhone] The contacts2:CompanyMainPhone element specifies the main telephone number for the contact's company.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R200");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R200
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.CompanyName].ToString(),
                contact.CompanyName,
                200,
                @"[In CompanyName] The CompanyName element specifies the company name for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R205");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R205
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.CustomerId].ToString(),
                contact.CustomerId,
                205,
                @"[In CustomerId] The contacts2:CustomerId element specifies the customer identifier (ID) for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R209");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R209
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.Department].ToString(),
                contact.Department,
                209,
                @"[In Department] The Department element specifies the department name for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R214");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R214
            this.Site.CaptureRequirementIfIsTrue(
                contact.Email1Address.Contains(contactProperties[Request.ItemsChoiceType8.Email1Address].ToString()),
                214,
                @"[In Email1Address] The Email1Address element specifies the first e-mail address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R221");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R221
            this.Site.CaptureRequirementIfIsTrue(
                contact.Email2Address.Contains(contactProperties[Request.ItemsChoiceType8.Email2Address].ToString()),
                221,
                @"[In Email2Address] The Email2Address element specifies the second e-mail address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R226");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R226
            this.Site.CaptureRequirementIfIsTrue(
                contact.Email3Address.Contains(contactProperties[Request.ItemsChoiceType8.Email3Address].ToString()),
                226,
                @"[In Email3Address] The Email3Address element specifies the third e-mail address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R231");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R231
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.FileAs].ToString(),
                contact.FileAs,
                231,
                @"[In FileAs] The FileAs element specifies how a contact is filed in the Contacts folder or the recipient information cache folder. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R238");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R238
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.FirstName].ToString(),
                contact.FirstName,
                238,
                @"[In FirstName] The FirstName element specifies the first name of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R243");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R243
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.GovernmentId].ToString(),
                contact.GovernmentId,
                243,
                @"[In GovernmentId] The contacts2:GovernmentId element specifies the government-assigned identifier (ID) for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R248");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R248
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.HomeAddressCity].ToString(),
                contact.HomeAddressCity,
                248,
                @"[In HomeAddressCity] The HomeAddressCity element specifies the home city for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R253");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R253
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.HomeAddressCountry].ToString(),
                contact.HomeAddressCountry,
                253,
                @"[In HomeAddressCountry] The HomeAddressCountry element specifies the home country/region for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R258");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R258
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.HomeAddressPostalCode].ToString(),
                contact.HomeAddressPostalCode,
                258,
                @"[In HomeAddressPostalCode] The HomeAddressPostalCode element specifies the home postal code for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R263");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R263
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.HomeAddressState].ToString(),
                contact.HomeAddressState,
                263,
                @"[In HomeAddressState] The HomeAddressState element specifies the home state for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R268");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R268
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.HomeAddressStreet].ToString(),
                contact.HomeAddressStreet,
                268,
                @"[In HomeAddressStreet] The HomeAddressStreet element specifies the home street address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R273");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R273
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.HomeFaxNumber].ToString(),
                contact.HomeFaxNumber,
                273,
                @"[In HomeFaxNumber] The HomeFaxNumber element specifies the home fax number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R278");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R278
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.HomePhoneNumber].ToString(),
                contact.HomePhoneNumber,
                278,
                @"[In HomePhoneNumber] The HomePhoneNumber element specifies the home phone number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R283");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R283
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.Home2PhoneNumber].ToString(),
                contact.Home2PhoneNumber,
                283,
                @"[In Home2PhoneNumber] The Home2PhoneNumber element specifies the alternative home phone number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R288");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R288
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.IMAddress].ToString(),
                contact.IMAddress,
                288,
                @"[In IMAddress] The contacts2:IMAddress element specifies the instant messaging address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R293");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R293
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.IMAddress2].ToString(),
                contact.IMAddress2,
                293,
                @"[In IMAddress2] The contacts2:IMAddress2 element specifies the alternative instant messaging address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R298");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R298
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.IMAddress3].ToString(),
                contact.IMAddress3,
                298,
                @"[In IMAddress3] The contacts2:IMAddress3 element specifies the tertiary instant messaging address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R303");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R303
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.JobTitle].ToString(),
                contact.JobTitle,
                303,
                @"[In JobTitle] The JobTitle element specifies the contact's job title.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R308");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R308
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.LastName].ToString(),
                contact.LastName,
                308,
                @"[In LastName] The LastName element specifies the contact's last name.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R313");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R313
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.ManagerName].ToString(),
                contact.ManagerName,
                313,
                @"[In ManagerName] The contacts2:ManagerName element specifies the distinguished name (DN) (1) of the contact's manager.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R318");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R318
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.MiddleName].ToString(),
                contact.MiddleName,
                318,
                @"[In MiddleName] The MiddleName element specifies the middle name of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R323");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R323
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.MMS].ToString(),
                contact.MMS,
                323,
                @"[In MMS] The contacts2:MMS element specifies the Multimedia Messaging Service (MMS) address for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R328");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R328
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.MobilePhoneNumber].ToString(),
                contact.MobilePhoneNumber,
                328,
                @"[In MobilePhoneNumber] The MobilePhoneNumber element specifies the mobile phone number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R333");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R333
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.NickName].ToString(),
                contact.NickName,
                333,
                @"[In NickName] The contacts2:NickName element specifies the nickname for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R338");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R338
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.OfficeLocation].ToString(),
                contact.OfficeLocation,
                338,
                @"[In OfficeLocation] The OfficeLocation element specifies the office location for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R343");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R343
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.OtherAddressCity].ToString(),
                contact.OtherAddressCity,
                343,
                @"[In OtherAddressCity] The OtherAddressCity element specifies the city for the contact's alternate address.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R348");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R348
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.OtherAddressCountry].ToString(),
                contact.OtherAddressCountry,
                348,
                @"[In OtherAddressCountry] The OtherAddressCountry element specifies the country/region of the contact's alternate address.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R353");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R353
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.OtherAddressPostalCode].ToString(),
                contact.OtherAddressPostalCode,
                353,
                @"[In OtherAddressPostalCode] The OtherAddressPostalCode element specifies the postal code of the contact's alternate address.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R358");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R358
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.OtherAddressState].ToString(),
                contact.OtherAddressState,
                358,
                @"[In OtherAddressState] The OtherAddressState element specifies the state of the contact's alternate address.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R363");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R363
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.OtherAddressStreet].ToString(),
                contact.OtherAddressStreet,
                363,
                @"[In OtherAddressStreet] The OtherAddressStreet element specifies the street address of the contact's alternate address.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R368");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R368
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.PagerNumber].ToString(),
                contact.PagerNumber,
                368,
                @"[In PagerNumber] The PagerNumber element specifies the pager number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R373");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R373
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.Picture].ToString(),
                contact.Picture,
                373,
                @"[In Picture] The Picture element specifies the file that contains the picture of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R1021");

            // If the contact with picture is successfully added to the server and the picture returned from server is not null, it means the picture satisfies the specified image size.
            // Verify MS-ASCNTC requirement: MS-ASCNTC_R1021
            this.Site.CaptureRequirement(
                1021,
                @"[In Picture] The value of the Picture element MUST be limited to an image size of 36 KB.");

            try
            {
                byte[] pictureByte = Convert.FromBase64String(contact.Picture);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R381");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R381
                this.Site.CaptureRequirementIfIsTrue(
                    pictureByte.Length <= 48 * 1024,
                    381,
                    @"[In Picture] The value of the Picture element MUST be limited to 48 KB of binary content that is encoded with base64 encoding");
            }
            catch (FormatException)
            {
                this.Site.Assert.Fail("The Picture element should be encoded with base64 encoding.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R385");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R385
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.RadioPhoneNumber].ToString(),
                contact.RadioPhoneNumber,
                385,
                @"[In RadioPhoneNumber] The RadioPhoneNumber element specifies the radio phone number for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R390");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R390
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.Spouse].ToString(),
                contact.Spouse,
                390,
                @"[In Spouse] The Spouse element specifies the name of the contact's spouse/partner.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R395");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R395
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.Suffix].ToString(),
                contact.Suffix,
                395,
                @"[In Suffix] The Suffix element specifies the suffix for the contact's name.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R400");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R400
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.Title].ToString(),
                contact.Title,
                400,
                @"[In Title] The Title element specifies the contact's business title.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R405");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R405
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.WebPage].ToString(),
                contact.WebPage,
                405,
                @"[In WebPage] The WebPage element specifies the Web site or personal Web page for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R417");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R417
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.YomiCompanyName].ToString(),
                contact.YomiCompanyName,
                417,
                @"[In YomiCompanyName] The YomiCompanyName element specifies the Japanese phonetic rendering of the company name for the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R422");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R422
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.YomiFirstName].ToString(),
                contact.YomiFirstName,
                422,
                @"[In YomiFirstName] The YomiFirstName element specifies the Japanese phonetic rendering of the first name of the contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R427");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R427
            this.Site.CaptureRequirementIfAreEqual<string>(
                contactProperties[Request.ItemsChoiceType8.YomiLastName].ToString(),
                contact.YomiLastName,
                427,
                @"[In YomiLastName] The YomiLastName element specifies the Japanese phonetic rendering of the last name of the contact.");
        }
        #endregion

        #region Private methods
        /// <summary>
        /// This method is used to synchronize the folder collection hierarchy.
        /// </summary>
        /// <returns>The response of the FolderSync command.</returns>
        private FolderSyncResponse FolderSync()
        {
            FolderSyncResponse folderSyncResponse = this.ASCNTCAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));

            this.Site.Assert.AreEqual<byte>(
                1,
                folderSyncResponse.ResponseData.Status,
                "The Status value should be 1 to indicate the FolderSync command executes successfully.");

            return folderSyncResponse;
        }

        /// <summary>
        /// Delete all the items in a folder.
        /// </summary>
        /// <param name="createdItemsCollection">The created items collection which should be deleted.</param>
        private void DeleteItemsInFolder(Collection<CreatedItems> createdItemsCollection)
        {
            foreach (CreatedItems createdItems in createdItemsCollection)
            {
                string syncKey = this.GetInitialSyncKey(createdItems.CollectionId, null);
                SyncRequest request = TestSuiteHelper.CreateSyncRequest(syncKey, createdItems.CollectionId, null);
                DataStructures.SyncStore result = this.ASCNTCAdapter.Sync(request);

                List<Request.SyncCollectionDelete> deleteData = new List<Request.SyncCollectionDelete>();
                foreach (string subject in createdItems.ItemSubject)
                {
                    string serverId = null;
                    if (result != null)
                    {
                        foreach (DataStructures.Sync item in result.AddElements)
                        {
                            if (item.Email.Subject != null && item.Email.Subject.Equals(subject, StringComparison.CurrentCulture))
                            {
                                serverId = item.ServerId;
                                break;
                            }

                            if (item.Contact.FileAs != null && item.Contact.FileAs.Equals(subject, StringComparison.CurrentCulture))
                            {
                                serverId = item.ServerId;
                                break;
                            }
                        }
                    }

                    this.Site.Assert.IsNotNull(serverId, "The item with subject '{0}' should be found!", subject);
                    deleteData.Add(new Request.SyncCollectionDelete() { ServerId = serverId });
                }

                Request.SyncCollection syncCollection = TestSuiteHelper.CreateSyncCollection(result.SyncKey, createdItems.CollectionId);
                syncCollection.Commands = deleteData.ToArray();
                syncCollection.DeletesAsMoves = false;
                syncCollection.DeletesAsMovesSpecified = true;

                SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
                DataStructures.SyncStore deleteResult = this.ASCNTCAdapter.Sync(syncRequest);
                this.Site.Assert.AreEqual<byte>(
                    1,
                    deleteResult.CollectionStatus,
                    "The value of Status should be 1 to indicate the Sync command executes successfully.");
            }
        }
        #endregion
    }
}