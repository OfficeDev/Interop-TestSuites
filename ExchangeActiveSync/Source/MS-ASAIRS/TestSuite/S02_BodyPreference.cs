namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the BodyPreference element and Body element in the AirSyncBase namespace, which is used by the Sync command, Search command and ItemOperations command to identify the data sent by and returned to client.
    /// </summary>
    [TestClass]
    public class S02_BodyPreference : TestSuiteBase
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
        public static void ClassCleanUp()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region MSASAIRS_S02_TC01_BodyPreference_AllOrNoneTrue_AllContentReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPreference) element is set to 1 (TRUE) and the content has not been truncated, all of the content is synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC01_BodyPreference_AllOrNoneTrue_AllContentReturned()
        {
            #region Send a plain text email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyBodyElements(syncItem.Email.Body, true, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R373");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R373
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.Body.Data,
                syncItem.Email.Body.Data,
                373,
                @"[In AllOrNone] When the value [of the AllOrNone element] is set to 1 (TRUE) and the content has not been truncated, all of the content is synchronized.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R304");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R304
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                syncItem.Email.Body.Type,
                304,
                @"[In Type] [The value] 1 [of Type element] means Plain text.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R237");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R237
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                syncItem.Email.NativeBodyType.ToString(),
                237,
                @"[In NativeBodyType] [The value] 1 represents Plain text.");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyBodyElements(itemOperationsItem.Email.Body, true, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R54");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R54
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.Body.Data,
                itemOperationsItem.Email.Body.Data,
                54,
                @"[In AllOrNone] When the value [of the AllOrNone element] is set to 1 (TRUE) and the content has not been truncated, all of the content is retrieved.");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyBodyElements(searchItem.Email.Body, true, false, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R53");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R53
                Site.CaptureRequirementIfAreEqual<string>(
                    allContentItem.Email.Body.Data,
                    searchItem.Email.Body.Data,
                    53,
                    @"[In AllOrNone] When the value [of the AllOrNone element] is set to 1 (TRUE) and the content has not been truncated, all of the content is searched.");
                #endregion
            }
        }
        #endregion

        #region MSASAIRS_S02_TC02_BodyPreference_AllOrNoneTrue_AllContentNotReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPreference) element is set to 1 (TRUE) and the content has been truncated, the content is not synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC02_BodyPreference_AllOrNoneTrue_AllContentNotReturned()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 2,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyBodyElements(syncItem.Email.Body, true, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R376");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R376
            Site.CaptureRequirementIfIsNull(
                syncItem.Email.Body.Data,
                376,
                @"[In AllOrNone] When the value is set to 1 (TRUE) and the content has been truncated, the content is not synchronized. ");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyBodyElements(itemOperationsItem.Email.Body, true, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R377");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R377
            Site.CaptureRequirementIfIsNull(
                itemOperationsItem.Email.Body.Data,
                377,
                @"[In AllOrNone] When the value is set to 1 (TRUE) and the content has been truncated, the content is not retrieved. ");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyBodyElements(searchItem.Email.Body, true, true, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R375");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R375
                Site.CaptureRequirementIfIsNull(
                    searchItem.Email.Body.Data,
                    375,
                    @"[In AllOrNone] When the value is set to 1 (TRUE) and the content has been truncated, the content is not searched. ");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirement MS-ASAIRS_R78 can be covered directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R78");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R78
            Site.CaptureRequirement(
                78,
                @"[In AllOrNone (BodyPreference)] But, if the client also includes the AllOrNone element with a value of 1 (TRUE) along with the TruncationSize element, it is instructing the server not to return a truncated response for that type when the size (in bytes) of the available data exceeds the value of the TruncationSize element.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC03_BodyPreference_AllOrNoneFalse_TruncatedContentReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPreference) element is set to 0 (FALSE) and the available data exceeds the value of the TruncationSize element, the truncated content is synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC03_BodyPreference_AllOrNoneFalse_TruncatedContentReturned()
        {
            #region Send a plain text email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 8,
                    TruncationSizeSpecified = true,
                    AllOrNone = false,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyBodyElements(syncItem.Email.Body, false, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R378");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R378
            Site.CaptureRequirementIfAreEqual<string>(
                body.Substring(0, (int)bodyPreference[0].TruncationSize),
                syncItem.Email.Body.Data,
                378,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the truncated is synchronized. ");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyBodyElements(itemOperationsItem.Email.Body, false, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R379");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R379
            Site.CaptureRequirementIfAreEqual<string>(
                body.Substring(0, (int)bodyPreference[0].TruncationSize),
                itemOperationsItem.Email.Body.Data,
                379,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the truncated is retrieved. ");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyBodyElements(searchItem.Email.Body, false, true, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R55");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R55
                Site.CaptureRequirementIfAreEqual<string>(
                    body.Substring(0, (int)bodyPreference[0].TruncationSize),
                    searchItem.Email.Body.Data,
                    55,
                    @"[In AllOrNone] When the value is set to 0 (FALSE), the truncated is searched. ");
                #endregion
            }

            #region Verify requirement
            // According to above steps, requirement MS-ASAIRS_R180 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R180");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R180
            Site.CaptureRequirement(
                180,
                @"[In Data (Body)] If the Truncated element (section 2.2.2.39.1) is included in the response, the data in the Data element is truncated.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC04_BodyPreference_AllOrNoneFalse_NonTruncatedContentReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPreference) element is set to 0 (FALSE) and the available data doesn't exceed the value of the TruncationSize element, the non-truncated content will be synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC04_BodyPreference_AllOrNoneFalse_NonTruncatedContentReturned()
        {
            #region Send a plain text email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = false,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyBodyElements(syncItem.Email.Body, false, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R381");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R381
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.Body.Data,
                syncItem.Email.Body.Data,
                381,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the nontruncated content is synchronized. ");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyBodyElements(itemOperationsItem.Email.Body, false, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R382");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R382
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.Body.Data,
                itemOperationsItem.Email.Body.Data,
                382,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the nontruncated content is retrieved. ");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyBodyElements(searchItem.Email.Body, false, false, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R380");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R380
                Site.CaptureRequirementIfAreEqual<string>(
                    allContentItem.Email.Body.Data,
                    searchItem.Email.Body.Data,
                    380,
                    @"[In AllOrNone] When the value is set to 0 (FALSE), the nontruncated content is searched. ");
                #endregion
            }
        }
        #endregion

        #region MSASAIRS_S02_TC05_BodyPreference_NoAllOrNone_TruncatedContentReturned
        /// <summary>
        /// This case is designed to test if the AllOrNone (BodyPreference) element is not included in the request and the available data exceeds the value of the TruncationSize element, the truncated content will be synchronized, searched or retrieved as if the value was set to 0 (FALSE).
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC05_BodyPreference_NoAllOrNone_TruncatedContentReturned()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 8,
                    TruncationSizeSpecified = true,
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyBodyElements(syncItem.Email.Body, null, true, true);

            Site.Assert.AreEqual<string>(
                body.Substring(0, (int)bodyPreference[0].TruncationSize),
                syncItem.Email.Body.Data,
                "The server should return the data truncated to the size requested by TruncationSize when the AllOrNone element is not included in the request and the available data exceeds the truncation size.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R420");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R420
            Site.CaptureRequirement(
                420,
                @"[In AllOrNone (BodyPreference)] If the AllOrNone element is not included in the request, then the truncated content is synchronized as if the value was set to 0 (FALSE).");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyBodyElements(itemOperationsItem.Email.Body, null, true, true);

            Site.Assert.AreEqual<string>(
                body.Substring(0, (int)bodyPreference[0].TruncationSize),
                itemOperationsItem.Email.Body.Data,
                "The server should return the data truncated to the size requested by TruncationSize when the AllOrNone element is not included in the request and the available data exceeds the truncation size.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R421");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R421
            Site.CaptureRequirement(
                421,
                @"[In AllOrNone (BodyPreference)] If the AllOrNone element is not included in the request, then the truncated content is retrieved as if the value was set to 0 (FALSE).");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyBodyElements(searchItem.Email.Body, null, true, true);

                Site.Assert.AreEqual<string>(
                    body.Substring(0, (int)bodyPreference[0].TruncationSize),
                    searchItem.Email.Body.Data,
                    "The server should return the data truncated to the size requested by TruncationSize when the AllOrNone element is not included in the request and the available data exceeds the truncation size.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R73");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R73
                Site.CaptureRequirement(
                    73,
                    @"[In AllOrNone (BodyPreference)] If the AllOrNone element is not included in the request, then the truncated content is searched as if the value was set to 0 (FALSE).");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirements MS-ASAIRS_R276 and MS-ASAIRS_R77 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R276");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R276
            Site.CaptureRequirement(
                276,
                @"[In Truncated (Body)] If the value [of the Truncated element] is TRUE, then the body of the item has been truncated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R77");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R77
            Site.CaptureRequirement(
                77,
                @"[In AllOrNone (BodyPreference)] [A client can include multiple BodyPreference elements in a command request with different values for the Type element] By default, the server returns the data truncated to the size requested by TruncationSize for the Type element that matches the native storage format of the item's Body element (section 2.2.2.9).");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC06_BodyPreference_NoAllOrNone_NonTruncatedContentReturned
        /// <summary>
        /// This case is designed to test if the AllOrNone (BodyPreference) element is not included in the request and the available data doesn't exceed the value of the TruncationSize element, the non-truncated content will be synchronized, searched or retrieved as if the value was set to 0 (FALSE).
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC06_BodyPreference_NoAllOrNone_NonTruncatedContentReturned()
        {
            #region Send a plain text email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyBodyElements(syncItem.Email.Body, null, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R423");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R423
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.Body.Data,
                syncItem.Email.Body.Data,
                423,
                @"[In AllOrNone (BodyPreference)] If the AllOrNone element is not included in the request, then the non-truncated content is synchronized as if the value was set to 0 (FALSE).");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyBodyElements(itemOperationsItem.Email.Body, null, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R424");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R424
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.Body.Data,
                itemOperationsItem.Email.Body.Data,
                424,
                @"[In AllOrNone (BodyPreference)] If the AllOrNone element is not included in the request, then the non-truncated content is retrieved as if the value was set to 0 (FALSE).");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyBodyElements(searchItem.Email.Body, null, false, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R422");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R422
                Site.CaptureRequirementIfAreEqual<string>(
                    allContentItem.Email.Body.Data,
                    searchItem.Email.Body.Data,
                    422,
                    @"[In AllOrNone (BodyPreference)] If the AllOrNone element is not included in the request, then the non-truncated content is searched as if the value was set to 0 (FALSE).");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirement MS-ASAIRS_R277 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R277");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R277
            Site.CaptureRequirement(
                277,
                @"[In Truncated (Body)] If the value [of the Truncated element] is FALSE, or there is no Truncated element, then the body of the item has not been truncated.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC07_BodyPreference_NotIncludedTruncationSize
        /// <summary>
        /// This case is designed to test if the TruncationSize (BodyPreference) element is not included, the server will return the same response no matter whether AllOrNone is true or false.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC07_BodyPreference_NotIncludedTruncationSize()
        {
            #region Send a plain text email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreferenceAllOrNoneTrue = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };

            Request.BodyPreference[] bodyPreferenceAllOrNoneFalse = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    AllOrNone = false,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            // Call Sync command with AllOrNone setting to TRUE
            DataStructures.Sync syncItemAllOrNoneTrue = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreferenceAllOrNoneTrue, null);

            this.VerifyBodyElements(syncItemAllOrNoneTrue.Email.Body, true, false, false);

            // Call Sync command with AllOrNone setting to FALSE
            DataStructures.Sync syncItemAllOrNoneFalse = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreferenceAllOrNoneFalse, null);

            this.VerifyBodyElements(syncItemAllOrNoneFalse.Email.Body, false, false, false);

            Site.Log.Add(
                LogEntryKind.Debug,
                "Entire content: {0}, content for AllOrNone TRUE: {1}, content for AllOrNone FALSE: {2}.",
                allContentItem.Email.Body.Data,
                syncItemAllOrNoneTrue.Email.Body.Data,
                syncItemAllOrNoneFalse.Email.Body.Data);

            Site.Assert.IsTrue(
                allContentItem.Email.Body.Data == syncItemAllOrNoneTrue.Email.Body.Data && syncItemAllOrNoneTrue.Email.Body.Data == syncItemAllOrNoneFalse.Email.Body.Data,
                "Server should return the entire content for the request and same response no matter AllOrNone is true or false if the TruncationSize element is absent in Sync command request.");
            #endregion

            #region Verify ItemOperations command related elements
            // Call ItemOperations command with AllOrNone setting to true
            DataStructures.ItemOperations itemOperationsItemAllOrNoneTrue = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItemAllOrNoneTrue.ServerId, null, bodyPreferenceAllOrNoneTrue, null, null);

            this.VerifyBodyElements(itemOperationsItemAllOrNoneTrue.Email.Body, true, false, false);

            // Call ItemOperations command with AllOrNone setting to false
            DataStructures.ItemOperations itemOperationsItemAllOrNoneFalse = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItemAllOrNoneTrue.ServerId, null, bodyPreferenceAllOrNoneFalse, null, null);

            this.VerifyBodyElements(itemOperationsItemAllOrNoneFalse.Email.Body, false, false, false);

            Site.Log.Add(
                LogEntryKind.Debug,
                "Entire content: {0}, content for AllOrNone TRUE: {1}, content for AllOrNone FALSE: {2}.",
                allContentItem.Email.Body.Data,
                itemOperationsItemAllOrNoneTrue.Email.Body.Data,
                itemOperationsItemAllOrNoneFalse.Email.Body.Data);

            Site.Assert.IsTrue(
                allContentItem.Email.Body.Data == itemOperationsItemAllOrNoneTrue.Email.Body.Data && itemOperationsItemAllOrNoneTrue.Email.Body.Data == itemOperationsItemAllOrNoneFalse.Email.Body.Data,
                "Server should return the entire content for the request and same response no matter AllOrNone is true or false if the TruncationSize element is absent in ItemOperations command request.");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                // Call Search command with AllOrNone setting to true
                DataStructures.Search searchItemAllNoneTrue = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItemAllOrNoneTrue.Email.ConversationId, bodyPreferenceAllOrNoneTrue, null);

                this.VerifyBodyElements(searchItemAllNoneTrue.Email.Body, true, false, false);

                // Call Search command with AllOrNone setting to false
                DataStructures.Search searchItemAllNoneFalse = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItemAllOrNoneTrue.Email.ConversationId, bodyPreferenceAllOrNoneFalse, null);

                this.VerifyBodyElements(searchItemAllNoneFalse.Email.Body, false, false, false);

                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Entire content: {0}, content for AllOrNone TRUE when TruncationSize element is absent: {1}, content for AllOrNone FALSE when TruncationSize element is absent: {2}.",
                    allContentItem.Email.Body.Data,
                    searchItemAllNoneTrue.Email.Body.Data,
                    searchItemAllNoneFalse.Email.Body.Data);

                Site.Assert.IsTrue(
                    allContentItem.Email.Body.Data == searchItemAllNoneTrue.Email.Body.Data && searchItemAllNoneTrue.Email.Body.Data == searchItemAllNoneFalse.Email.Body.Data,
                    "Server should return the entire content for the request and same response no matter AllOrNone is true or false if the TruncationSize element is absent in Search command request.");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirements MS-ASAIRS_R300 and MS-ASAIRS_R401 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R300");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R300
            Site.CaptureRequirement(
                300,
                @"[In TruncationSize (BodyPreference)] If the TruncationSize element is absent, the entire content is used for the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R401");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R401
            Site.CaptureRequirement(
                401,
                @"[In AllOrNone (BodyPreference)]If the TruncationSize element is not included, the server will return the same response no matter whether AllOrNone is true or false.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC08_MultipleBodyPreference
        /// <summary>
        /// This case is designed to test if the client has specified multiple BodyPreference elements, the server will select the next BodyPreference element and return the maximum amount of body text to the client.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC08_MultipleBodyPreference()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreferences = new Request.BodyPreference[2]
            {
                new Request.BodyPreference()
                {
                    Type = 2,
                    TruncationSize = 8,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                },
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 18,
                    TruncationSizeSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreferences, null);

            this.VerifyMultipleBodyPreference(syncItem.Email, body, bodyPreferences);
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreferences, null, null);

            this.VerifyMultipleBodyPreference(itemOperationsItem.Email, body, bodyPreferences);
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreferences, null);

                this.VerifyMultipleBodyPreference(searchItem.Email, body, bodyPreferences);
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirement MS-ASAIRS_R80 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R80");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R80
            Site.CaptureRequirement(
                80,
                @"[In AllOrNone (BodyPreference)] In this case [if the client also includes the AllOrNone element with a value of 1 (TRUE) along with the TruncationSize element, it is instructing the server not to return a truncated response for that type when the size (in bytes) of the available data exceeds the value of the TruncationSize element], if the client has specified multiple BodyPreference elements, the server selects the next BodyPreference element that will return the maximum amount of body text to the client.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC09_Body_Preview
        /// <summary>
        /// This case is designed to test the Preview (Body) element which contains the Unicode plain text message or message part preview returned to the client.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC09_Body_Preview()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Preview element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send a plain text email and get the none truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 2,
                    Preview = 18,
                    PreviewSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyBodyPreview(syncItem.Email, allContentItem.Email, bodyPreference);
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyBodyPreview(itemOperationsItem.Email, allContentItem.Email, bodyPreference);
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyBodyPreview(searchItem.Email, allContentItem.Email, bodyPreference);
                #endregion
            }

            #region Verify common requirements
            // According to above steps, the following requirements can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R248");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R248
            Site.CaptureRequirement(
                248,
                @"[In Preview (Body)] The Preview element is an optional child element of the Body element (section 2.2.2.9) that contains the Unicode plain text message or message part preview returned to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R250");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R250
            Site.CaptureRequirement(
                250,
                @"[In Preview (Body)] The Preview element in a response MUST contain no more than the number of characters specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R2644");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R2644
            Site.CaptureRequirement(
                2644,
                @"[In Preview (BodyPreference)] [The Preview element] specifies the maximum length of the Unicode plain text message or message part preview to be returned to the client.");
            #endregion

            #region Set BodyPreference element
            bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type =1,
                    Preview = 18,
                    PreviewSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);
            this.VerifyBodyPreview(syncItem.Email, allContentItem.Email, bodyPreference);
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC10_IncludedBodyInResponse
        /// <summary>
        /// This case is designed to test the Body element must be included in a response message whenever new items are created or an item has changes.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC10_IncludedBodyInResponse()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);

            Site.Assert.IsNotNull(
                syncItem.Email.Body,
                "The Body element should be included in a response message whenever new items are created.");
            #endregion

            #region Verify ItemOperations related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, null, null);

            Site.Assert.IsNotNull(
                itemOperationsItem.Email.Body,
                "The Body element should be included in a response message whenever new items are created.");
            #endregion

            #region Verify requirements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R105");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R105
            Site.CaptureRequirement(
                105,
                @"[In Body] The Body element MUST be included in a response message whenever [an item has changes or] new items are created.");
            #endregion

            #region Update Read property of the item
            Request.SyncCollectionChange changeData = new Request.SyncCollectionChange
            {
                ServerId = syncItem.ServerId,
                ApplicationData =
                    new Request.SyncCollectionChangeApplicationData
                    {
                        Items = new object[] { !syncItem.Email.Read },
                        ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Read }
                    }
            };

            Request.SyncCollection syncCollection = TestSuiteHelper.CreateSyncCollection(this.SyncKey, this.User2Information.InboxCollectionId);
            syncCollection.Commands = new object[] { changeData };

            SyncRequest request = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });

            // Call Sync command to update the item
            DataStructures.SyncStore syncStore = this.ASAIRSAdapter.Sync(request);

            Site.Assert.AreEqual<byte>(
                1,
                syncStore.CollectionStatus,
                "The server should return status 1 in the Sync command response to indicate sync command executes successfully.");
            #endregion

            #region Verify Sync command after update item
            DataStructures.Sync updatedSyncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);

            // Assert the Read has been changed to the new value
            Site.Assert.AreEqual<bool?>(
                !syncItem.Email.Read,
                updatedSyncItem.Email.Read,
                "The Read property of the item should be updated.");

            // Assert the body is not null when the item property is changed
            Site.Assert.IsNotNull(
                updatedSyncItem.Email.Body,
                "The Body element should be included in a response message whenever an item has changes.");
            #endregion

            #region Verify ItemOperations command after update item
            DataStructures.ItemOperations updatedItemOperationItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, updatedSyncItem.ServerId, null, null, null, null);

            // Assert the Read has been changed to the new value
            Site.Assert.AreEqual<bool?>(
                !itemOperationsItem.Email.Read,
                updatedItemOperationItem.Email.Read,
                "The Read property of the item should be updated.");

            // Assert the body is not null when the item property is changed
            Site.Assert.IsNotNull(
                updatedItemOperationItem.Email.Body,
                "The Body element should be included in a response message whenever an item has changes.");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, null);

                Site.Assert.IsNotNull(
                    searchItem.Email.Body,
                    "The Body element should be included in a response message whenever new items are created.");
                #endregion

                #region Update Read property of the item again
                changeData = new Request.SyncCollectionChange
                {
                    ServerId = syncItem.ServerId,
                    ApplicationData =
                        new Request.SyncCollectionChangeApplicationData
                        {
                            Items = new object[] { !updatedItemOperationItem.Email.Read },
                            ItemsElementName = new Request.ItemsChoiceType7[] { Request.ItemsChoiceType7.Read }
                        }
                };

                syncCollection = TestSuiteHelper.CreateSyncCollection(this.SyncKey, this.User2Information.InboxCollectionId);
                syncCollection.Commands = new object[] { changeData };

                request = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });

                // Call Sync command to update the item
                syncStore = this.ASAIRSAdapter.Sync(request);

                Site.Assert.AreEqual<byte>(
                    1,
                    syncStore.CollectionStatus,
                    "The server should return status 1 in the Sync command response to indicate sync command executes successfully.");
                #endregion

                #region Verify Search command after update item
                DataStructures.Search updateSearchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, updatedItemOperationItem.Email.ConversationId, null, null);

                // Assert the Read has been changed to the new value
                Site.Assert.AreEqual<bool?>(
                   !searchItem.Email.Read,
                   updateSearchItem.Email.Read,
                   "The Read property of the item should be updated.");

                // Assert the body is not null when the item property is changed
                Site.Assert.IsNotNull(
                    updateSearchItem.Email.Body,
                    "The Body element should be included in a response message whenever an item has changes.");
                #endregion
            }

            #region Verify requirements
            // According to above steps, requirement MS-ASAIRS_R386 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R386");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R386
            Site.CaptureRequirement(
                386,
                @"[In Body] The Body element MUST be included in a response message whenever an item has changes [or new items are created].");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC11_TruncatedPresentOrNotInRequest
        /// <summary>
        /// This case is designed to test both the Truncated element and the value of Truncated element has no effect in a command request.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC11_TruncatedPresentOrNotInRequest()
        {
            bool isTruncatedSupported = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0")
                || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1");
            Site.Assume.IsTrue(isTruncatedSupported, "The Truncated element is only supported when the MS-ASProtocolVersion header is set to 12.1, 14.0 and 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Add three contacts with or without the Truncated element
            List<object> commandList = new List<object>();
            string data = Common.GenerateResourceName(Site, "ContactData");

            string fileAsWithoutTruncated = Common.GenerateResourceName(Site, "ContactWithoutTruncated");
            Request.SyncCollectionAdd syncAdd = CreateSyncAddContact(fileAsWithoutTruncated, data, null);
            commandList.Add(syncAdd);

            string fileAsWithTruncatedTrue = Common.GenerateResourceName(Site, "ContactWithTruncatedTrue");
            syncAdd = CreateSyncAddContact(fileAsWithTruncatedTrue, data, true);
            commandList.Add(syncAdd);

            string fileAsWithTruncatedFalse = Common.GenerateResourceName(Site, "ContactWithTruncatedFalse");
            syncAdd = CreateSyncAddContact(fileAsWithTruncatedFalse, data, false);
            commandList.Add(syncAdd);

            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User1Information.ContactsCollectionId), this.User1Information.ContactsCollectionId, commandList.ToArray(), null, null);

            DataStructures.SyncStore syncAddResponse = this.ASAIRSAdapter.Sync(syncAddRequest);
            Site.Assert.AreEqual<byte>(
                1,
                syncAddResponse.CollectionStatus,
                "The server should return a status 1 to indicate the Sync Add operation is successful.");

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.ContactsCollectionId, fileAsWithoutTruncated, fileAsWithTruncatedTrue, fileAsWithTruncatedFalse);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1
                }
            };
            #endregion

            #region Sychronize the three contacts
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User1Information.ContactsCollectionId), this.User1Information.ContactsCollectionId, null, bodyPreference, null);

            DataStructures.SyncStore contacts = this.ASAIRSAdapter.Sync(syncRequest);

            DataStructures.Sync contactWithoutTruncated = TestSuiteHelper.GetSyncAddItem(contacts, fileAsWithoutTruncated);
            this.VerifySyncItem(contactWithoutTruncated);

            DataStructures.Sync contactWithTruncatedTrue = TestSuiteHelper.GetSyncAddItem(contacts, fileAsWithTruncatedTrue);
            this.VerifySyncItem(contactWithTruncatedTrue);

            DataStructures.Sync contactWithTruncatedFalse = TestSuiteHelper.GetSyncAddItem(contacts, fileAsWithTruncatedFalse);
            this.VerifySyncItem(contactWithTruncatedFalse);

            Site.Log.Add(
                LogEntryKind.Debug,
                "Content without including Truncated in request: {0}; content with Truncated TRUE: {1}; content with Truncated FALSE: {2}.",
                contactWithoutTruncated.Contact.Body.Data,
                contactWithTruncatedTrue.Contact.Body.Data,
                contactWithTruncatedFalse.Contact.Body.Data);

            Site.Assert.IsTrue(
                contactWithoutTruncated.Contact.Body.Data == contactWithTruncatedTrue.Contact.Body.Data &&
                contactWithTruncatedTrue.Contact.Body.Data == contactWithTruncatedFalse.Contact.Body.Data,
                "The data should be same whenever the Truncated present or not in a request.");
            #endregion

            #region Verify requirements
            // According to above steps, the following requirements can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R11500");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R11500
            Site.CaptureRequirement(
                11500,
                @"[In Body] Reply is the same whether this element[Truncated] is used in a command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R406");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R406
            Site.CaptureRequirement(
                406,
                @"[In Truncated (Body)] The server will return the same response no matter what the value of Truncated element is.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC12_NativeBodyTypeAndType_RTF
        /// <summary>
        /// This case is designed to test if the value of the Type element is 3 (RTF), the value of the Data element is encoded using base64 encoding and the value of NativeBodyType is also 3.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC12_NativeBodyTypeAndType_RTF()
        {
            #region Send a mail with an embedded OLE object
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.AttachOLE, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 3
                }
            };
            #endregion

            #region Verify Sync command
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyType(syncItem.Email, bodyPreference[0].Type);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R239");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R239
            Site.CaptureRequirementIfAreEqual<byte?>(
                3,
                syncItem.Email.NativeBodyType,
                239,
                @"[In NativeBodyType] [The value] 3 represents RTF.");
            #endregion

            #region Verify ItemOperations command
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, null);

            this.VerifyType(itemOperationsItem.Email, bodyPreference[0].Type);
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, bodyPreference, null);

                this.VerifyType(searchItem.Email, bodyPreference[0].Type);

                try
                {
                    Convert.FromBase64String(searchItem.Email.Body.Data);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R179");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R179
                    Site.CaptureRequirement(
                        179,
                        @"[In Data (Body)] If the value of the Type element is 3 (RTF), the value of the Data element is encoded using base64 encoding.");
                }
                catch (FormatException)
                {
                    Site.Assert.Fail("The value of Data element should be encoded using base64 encoding.");
                }
                #endregion
            }

            #region Verify requirements
            // According to above steps, requirement MS-ASAIRS_R306 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R306");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R306
            Site.CaptureRequirement(
                306,
                @"[In Type] [The value] 3 [of Type element] means RTF.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC13_NativeBodyTypeAndType_HTML
        /// <summary>
        /// This case is designed to test both the NativeBodyType and Type has the same value unless the server has modified the format of the body to match the client's request.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC13_NativeBodyTypeAndType_HTML()
        {
            #region Send an html email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 2
                }
            };

            Request.BodyPreference[] bodyPreferenceWithType4 = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 4
                }
            };
            #endregion

            #region Verify Sync command
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            this.VerifyType(syncItem.Email, bodyPreference[0].Type);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R238");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R238
            Site.CaptureRequirementIfAreEqual<byte?>(
                2,
                syncItem.Email.NativeBodyType,
                238,
                @"[In NativeBodyType] [The value] 2 represents HTML.");

            // According to above steps, requirement MS-ASAIRS_R2411 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R2411");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R2411
            Site.CaptureRequirement(
                2411,
                @"[In NativeBodyType]  The NativeBodyType and Type elements have the same value if the server has not modified the format of the body to match the client's request.");
            #endregion

            #region Verify ItemOperations command
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null, DeliveryMethodForFetch.Inline);

            this.VerifyType(itemOperationsItem.Email, bodyPreference[0].Type);
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

                this.VerifyType(searchItem.Email, bodyPreference[0].Type);
                #endregion
            }

            #region Verify requirements
            // According to above steps, requirement MS-ASAIRS_R305 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R305");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R305
            Site.CaptureRequirement(
                305,
                @"[In Type] [The value] 2 [of Type element] means HTML.");
            #endregion

            #region Verify Sync command when setting Type to 4
            DataStructures.Sync syncItemWithType4 = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreferenceWithType4, null);

            Site.Assert.IsNotNull(
                syncItemWithType4.Email.Body,
                "The Body element should be included in Sync command response when the BodyPreference element is specified in Sync command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R307");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R307
            Site.CaptureRequirementIfAreEqual<byte>(
                4,
                syncItemWithType4.Email.Body.Type,
                307,
                @"[In Type] [The value] 4 [of Type element] means MIME.");

            Site.Assert.AreEqual<byte?>(
                2,
                syncItemWithType4.Email.NativeBodyType,
                "The NativeBodyType value should be 2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R2412");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R2412
            Site.CaptureRequirementIfAreNotEqual<byte?>(
                syncItemWithType4.Email.Body.Type,
                syncItemWithType4.Email.NativeBodyType,
                2412,
                @"[In NativeBodyType]  The NativeBodyType and Type elements have different values if the server has modified the format of the body to match the client's request.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S02_TC14_ItemOperations_Part
        /// <summary>
        /// This case is designed to test the Part element must not be present in non-multipart response and must be present in multipart response.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S02_TC14_ItemOperations_Part()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPReference element
            Request.BodyPreference[] bodyPreferences = new Request.BodyPreference[]
            {
                new Request.BodyPreference
                {
                    Type = 4
                }
            };
            #endregion

            #region Synchronize the email item
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);
            #endregion

            #region Call ItemOperations command with the DeliveryMethodForFetch setting to inline
            DataStructures.ItemOperations itemOperationsItemInline = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreferences, null, DeliveryMethodForFetch.Inline);

            Site.Assert.IsNotNull(
                itemOperationsItemInline.Email.Body,
                "The Body element should be included in response when new item is created.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R405");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R405
            Site.CaptureRequirementIfIsNull(
                itemOperationsItemInline.Email.Body.Part,
                405,
                @"[In Part] This element MUST NOT be present in non-multipart responses.");
            #endregion

            #region Call ItemOperations command with the DeliveryMethodForFetch setting to MultiPart
            DataStructures.ItemOperations itemOperationsItemMultiPart = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreferences, null, DeliveryMethodForFetch.MultiPart);

            Site.Assert.IsNotNull(
                itemOperationsItemMultiPart.Email.Body,
                "The Body element should be included in response when new item is created.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R244");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R244
            Site.CaptureRequirementIfIsNotNull(
                itemOperationsItemMultiPart.Email.Body.Part,
                244,
                @"[In Part] This element [the Part element] MUST be present in multipart responses, as specified in [MS-ASCMD] section 2.2.2.9.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R176");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R176
            Site.CaptureRequirementIfIsNull(
                itemOperationsItemMultiPart.Email.Body.Data,
                176,
                @"[In Data (Body)] This element [the Data (Body) element] MUST NOT be present in multipart responses, as specified in [MS-ASCMD] section 2.2.2.10.1.");
            #endregion
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Create a SyncCollectionAdd instance for adding a contact.
        /// </summary>
        /// <param name="fileAs">The FileAs element for the contact.</param>
        /// <param name="data">The body data of contact.</param>
        /// <param name="truncated">The Truncated element for the contact.</param>
        /// <returns>Return a SyncCollectionAdd instance.</returns>
        private static Request.SyncCollectionAdd CreateSyncAddContact(string fileAs, string data, bool? truncated)
        {
            Request.SyncCollectionAdd syncAdd = new Request.SyncCollectionAdd { ClientId = Guid.NewGuid().ToString("N") };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType8> itemsElementName = new List<Request.ItemsChoiceType8>();

            items.Add(fileAs);
            itemsElementName.Add(Request.ItemsChoiceType8.FileAs);

            Request.Body addBody = new Request.Body { Type = 1, Data = data };

            if (truncated != null)
            {
                addBody.TruncatedSpecified = true;
                addBody.Truncated = (bool)truncated;
            }

            items.Add(addBody);
            itemsElementName.Add(Request.ItemsChoiceType8.Body);

            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData
            {
                Items = items.ToArray(),
                ItemsElementName = itemsElementName.ToArray()
            };
            syncAdd.ApplicationData = applicationData;

            return syncAdd;
        }

        /// <summary>
        /// Verify elements in Body element.
        /// </summary>
        /// <param name="body">The body of item.</param>
        /// <param name="allOrNone">The value of AllOrNone element, "null" means the AllOrNone element is not present in request.</param>
        /// <param name="truncated">Whether the content is truncated.</param>
        /// <param name="includedTruncationSize">Whether includes the TruncationSize element.</param>
        private void VerifyBodyElements(Body body, bool? allOrNone, bool truncated, bool includedTruncationSize)
        {
            Site.Assert.IsNotNull(
                body,
                "The Body element should be included in response when the BodyPreference element is specified in request.");

            // Verify elements when TruncationSize element is absent
            if (!includedTruncationSize)
            {
                Site.Assert.IsTrue(
                    !body.TruncatedSpecified || (body.TruncatedSpecified && !body.Truncated),
                    "The data should not be truncated when the TruncationSize element is absent.");

                return;
            }

            if (truncated)
            {
                Site.Assert.IsTrue(
                    body.Truncated,
                    "The data should be truncated when the AllOrNone element value is {0} in the request and the available data exceeds the truncation size.",
                    allOrNone);
            }
            else
            {
                Site.Assert.IsTrue(
                    !body.TruncatedSpecified || (body.TruncatedSpecified && !body.Truncated),
                    "The data should not be truncated when the AllOrNone element value is {0} in the request and the truncation size exceeds the available data.",
                    allOrNone);
            }

            if (body.Truncated)
            {
                if (Common.IsRequirementEnabled(402, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R402");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R402
                    Site.CaptureRequirementIfIsTrue(
                        body.EstimatedDataSizeSpecified,
                        402,
                        @"[In Appendix B: Product Behavior] Implementation does include the EstimatedDataSize (Body) element in a response message whenever the Truncated element is set to TRUE. (Exchange Server 2007 SP1 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify elements in command response when the request includes multiple BodyPreference elements.
        /// </summary>
        /// <param name="email">The email item got from server.</param>
        /// <param name="body">The body of the email.</param>
        /// <param name="bodyPreferences">A BodyPreference object.</param>
        private void VerifyMultipleBodyPreference(DataStructures.Email email, string body, Request.BodyPreference[] bodyPreferences)
        {
            Site.Assert.IsNotNull(
                email.Body,
                "The Body element should be included in command response when the BodyPreference element is specified in command request.");

            Site.Assert.IsTrue(
                email.Body.Truncated,
                "The data should be truncated since the available data exceeds the TruncationSize value in the second BodyPreference and the first BodyPreference includes the AllOrNone element with a value of 1 (TRUE) along with the TruncationSize element which is smaller than the available data.");

            Site.Assert.AreEqual<string>(
                body.Substring(0, (int)bodyPreferences[1].TruncationSize),
                email.Body.Data,
                "The data should be truncated to the size that requested by TruncationSize element in the second BodyPreference if the request includes the AllOrNone element with a value of 1 (TRUE) along with the TruncationSize element which is smaller than the available data in the first BodyPreference.");
        }

        /// <summary>
        /// Verify the Preview element in Body element.
        /// </summary>
        /// <param name="email">The email item got from server.</param>
        /// <param name="allContentEmail">The email item which has full content.</param>
        /// <param name="bodyPreference">A BodyPreference object.</param>
        private void VerifyBodyPreview(DataStructures.Email email, DataStructures.Email allContentEmail, Request.BodyPreference[] bodyPreference)
        {
            if (bodyPreference[0].Type == 2)
            {
                Site.Assert.IsNotNull(
                email.Body,
                "The Body element should be included in command response when the BodyPreference element is specified in command request.");

                Site.Assert.IsNotNull(
                    email.Body.Preview,
                    "The Preview element should be included in command response when the Preview element is specified in command request.");

                Site.Assert.IsTrue(
                    email.Body.Preview.Length <= bodyPreference[0].Preview,
                    "The Preview element in a response should contain no more than the number of characters specified in the request. The length of Preview element in response is: {0}.",
                    email.Body.Preview.Length);

                Site.Assert.IsTrue(
                    allContentEmail.Body.Data.Contains(email.Body.Preview),
                    "The Preview element in a response should contain the message part preview returned to the client.");
            }

            if (bodyPreference[0].Type == 1)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R1001036");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R1001036
                Site.CaptureRequirementIfIsTrue(
                    allContentEmail.Body.Data == email.Body.Data && email.Body.Preview == null,
                    1001036,
                    @"[In Preview (Body)] If the Body element in the request contains a Type element of value 1 (Plain text) and there is valid data returned in the Data element (section 2.2.2.20.1), then the Preview element will not be returned in the same response.");
            }
        }

        /// <summary>
        /// Verify the item synchronized from server.
        /// </summary>
        /// <param name="item">The item synchronized from server.</param>
        private void VerifySyncItem(DataStructures.Sync item)
        {
            Site.Assert.IsNotNull(
                item,
                "The item should not be null.");

            Site.Assert.IsNotNull(
                item.Contact.Body,
                "The Body element should be included in response when the BodyPreference element is specified in request.");
        }

        /// <summary>
        /// Verify the Type element of an item.
        /// </summary>
        /// <param name="email">The email item got from server.</param>
        /// <param name="typeValue">The value of the Type element.</param>
        private void VerifyType(DataStructures.Email email, byte typeValue)
        {
            Site.Assert.IsNotNull(
                email.Body,
                "The Body element should be included in command response when the BodyPreference element is specified in command request.");

            Site.Assert.AreEqual<byte>(
                typeValue,
                email.Body.Type,
                "The Type value in command response should be {0}.",
                typeValue);
        }
        #endregion
    }
}