namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the BodyPartPreference element and BodyPart element in the AirSyncBase namespace, which is used by the Sync command, Search command and ItemOperations command to identify the data sent by and returned to client.
    /// </summary>
    [TestClass]
    public class S01_BodyPartPreference : TestSuiteBase
    {
        #region Class initialize and cleanup
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

        #region MSASAIRS_S01_TC01_BodyPartPreference_AllOrNoneTrue_AllContentReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPartPreference) element is set to 1 (TRUE) and the content has not been truncated, all of the content is synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC01_BodyPartPreference_AllOrNoneTrue_AllContentReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPartPreference element
            Request.BodyPartPreference[] bodyPartPreference = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreference);

            this.VerifyBodyPartElements(syncItem.Email.BodyPart, true, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R373");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R373
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.BodyPart.Data,
                syncItem.Email.BodyPart.Data,
                373,
                @"[In AllOrNone] When the value [of the AllOrNone element] is set to 1 (TRUE) and the content has not been truncated, all of the content is synchronized.");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, bodyPartPreference, null);

            this.VerifyBodyPartElements(itemOperationsItem.Email.BodyPart, true, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R54");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R54
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.BodyPart.Data,
                itemOperationsItem.Email.BodyPart.Data,
                54,
                @"[In AllOrNone] When the value [of the AllOrNone element] is set to 1 (TRUE) and the content has not been truncated, all of the content is retrieved.");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, bodyPartPreference);

                this.VerifyBodyPartElements(searchItem.Email.BodyPart, true, false, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R53");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R53
                Site.CaptureRequirementIfAreEqual<string>(
                    allContentItem.Email.BodyPart.Data,
                    searchItem.Email.BodyPart.Data,
                    53,
                    @"[In AllOrNone] When the value [of the AllOrNone element] is set to 1 (TRUE) and the content has not been truncated, all of the content is searched.");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirements MS-ASAIRS_R120 and MS-ASAIRS_R271 can be covered directly.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R120");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R120
            Site.CaptureRequirement(
                120,
                @"[In BodyPart] The BodyPart element MUST be included in a command response when the BodyPartPreference element (section 2.2.2.11) is specified in a request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R271");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R271
            Site.CaptureRequirement(
                271,
                @"[In Status] [The value] 1 [of Status element] means Success.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S01_TC02_BodyPartPreference_AllOrNoneTrue_AllContentNotReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPartPreference) element is set to 1 (TRUE) and the content has been truncated, the content is not synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC02_BodyPartPreference_AllOrNoneTrue_AllContentNotReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);
            #endregion

            #region Set BodyPartPreference element
            Request.BodyPartPreference[] bodyPartPreference = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    TruncationSize = 2,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreference);

            this.VerifyBodyPartElements(syncItem.Email.BodyPart, true, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R376");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R376
            Site.CaptureRequirementIfIsNull(
                syncItem.Email.BodyPart.Data,
                376,
                @"[In AllOrNone] When the value is set to 1 (TRUE) and the content has been truncated, the content is not synchronized. ");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, bodyPartPreference, null);

            this.VerifyBodyPartElements(itemOperationsItem.Email.BodyPart, true, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R377");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R377
            Site.CaptureRequirementIfIsNull(
                itemOperationsItem.Email.BodyPart.Data,
                377,
                @"[In AllOrNone] When the value is set to 1 (TRUE) and the content has been truncated, the content is not retrieved. ");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            { 
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, bodyPartPreference);

                this.VerifyBodyPartElements(searchItem.Email.BodyPart, true, true, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R375");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R375
                Site.CaptureRequirementIfIsNull(
                    searchItem.Email.BodyPart.Data,
                    375,
                    @"[In AllOrNone] When the value is set to 1 (TRUE) and the content has been truncated, the content is not searched. ");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirement MS-ASAIRS_R63 can be covered directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R63");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R63
            Site.CaptureRequirement(
                63,
                @"[In AllOrNone (BodyPartPreference)] But, if the client also includes the AllOrNone element with a value of 1 (TRUE) along with the TruncationSize element, it is instructing the server not to return a truncated response for that type when the size (in bytes) of the available data exceeds the value of the TruncationSize element.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S01_TC03_BodyPartPreference_AllOrNoneFalse_TruncatedContentReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPartPreference) element is set to 0 (FALSE) and the available data exceeds the value of the TruncationSize element, the truncated content is synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC03_BodyPartPreference_AllOrNoneFalse_TruncatedContentReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);

            this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            XmlElement lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;
            string allData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);
            #endregion

            #region Set BodyPartPreference element
            Request.BodyPartPreference[] bodyPartPreference = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    TruncationSize = 8,
                    TruncationSizeSpecified = true,
                    AllOrNone = false,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreference);
            lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;

            this.VerifyBodyPartElements(syncItem.Email.BodyPart, false, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R378");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R378
            Site.CaptureRequirementIfAreEqual<string>(
                TestSuiteHelper.TruncateData(allData, (int)bodyPartPreference[0].TruncationSize),
                TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject),
                378,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the truncated is synchronized. ");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, bodyPartPreference, null);
            lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;

            this.VerifyBodyPartElements(itemOperationsItem.Email.BodyPart, false, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R379");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R379
            Site.CaptureRequirementIfAreEqual<string>(
                TestSuiteHelper.TruncateData(allData, (int)bodyPartPreference[0].TruncationSize),
                TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject),
                379,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the truncated is retrieved. ");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, bodyPartPreference);
                lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;

                this.VerifyBodyPartElements(searchItem.Email.BodyPart, false, true, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R55");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R55
                Site.CaptureRequirementIfAreEqual<string>(
                    TestSuiteHelper.TruncateData(allData, (int)bodyPartPreference[0].TruncationSize),
                    TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject),
                    55,
                    @"[In AllOrNone] When the value is set to 0 (FALSE), the truncated is searched. ");
                #endregion
            }

            #region Verify requirement
            // According to above steps, requirement MS-ASAIRS_R188 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R188");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R188
            Site.CaptureRequirement(
                188,
                @"[In Data (BodyPart)] If the Truncated element (section 2.2.2.39.2) is included in the response, then the data in the Data element is truncated.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S01_TC04_BodyPartPreference_AllOrNoneFalse_NonTruncatedContentReturned
        /// <summary>
        /// This case is designed to test when the value of the AllOrNone (BodyPartPreference) element is set to 0 (FALSE) and the available data doesn't exceed the value of the TruncationSize element, the non-truncated content will be synchronized, searched or retrieved.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC04_BodyPartPreference_AllOrNoneFalse_NonTruncatedContentReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPartPreference element
            Request.BodyPartPreference[] bodyPartPreference = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = false,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreference);

            this.VerifyBodyPartElements(syncItem.Email.BodyPart, false, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R381");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R381
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.BodyPart.Data,
                syncItem.Email.BodyPart.Data,
                381,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the nontruncated content is synchronized. ");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, bodyPartPreference, null);

            this.VerifyBodyPartElements(itemOperationsItem.Email.BodyPart, false, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R382");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R382
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.BodyPart.Data,
                itemOperationsItem.Email.BodyPart.Data,
                382,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the nontruncated content is retrieved. ");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {          
                 #region Verify Search command related elements
            DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, bodyPartPreference);

            this.VerifyBodyPartElements(searchItem.Email.BodyPart, false, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R380");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R380
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.BodyPart.Data,
                searchItem.Email.BodyPart.Data,
                380,
                @"[In AllOrNone] When the value is set to 0 (FALSE), the nontruncated content is searched. ");
                #endregion
            }
        }
        #endregion

        #region MSASAIRS_S01_TC05_BodyPartPreference_NoAllOrNone_TruncatedContentReturned
        /// <summary>
        /// This case is designed to test if the AllOrNone (BodyPartPreference) element is not included in the request and the available data exceeds the value of the TruncationSize element, the truncated content synchronized, searched or retrieved as if the value was set to 0 (FALSE).
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC05_BodyPartPreference_NoAllOrNone_TruncatedContentReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);

            this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            XmlElement lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;
            string allData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);
            #endregion

            #region Set BodyPreference element
            Request.BodyPartPreference[] bodyPartPreference = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    TruncationSize = 8,
                    TruncationSizeSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreference);
            lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;

            this.VerifyBodyPartElements(syncItem.Email.BodyPart, null, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R407");

            // Verify MS-ASAIRS requirement: Verify MS-ASAIRS_R407
            Site.CaptureRequirementIfAreEqual<string>(
                TestSuiteHelper.TruncateData(allData, (int)bodyPartPreference[0].TruncationSize),
                TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject),
                407,
                @"[In AllOrNone (BodyPartPreference)] If the AllOrNone element is not included in the request, the truncated synchronized as if the value was set to 0 (FALSE).");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, bodyPartPreference, null);
            lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;

            this.VerifyBodyPartElements(itemOperationsItem.Email.BodyPart, null, true, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R408");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R408
            Site.CaptureRequirementIfAreEqual<string>(
                TestSuiteHelper.TruncateData(allData, (int)bodyPartPreference[0].TruncationSize),
                TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject),
                408,
                @"[In AllOrNone (BodyPartPreference)] If the AllOrNone element is not included in the request, the truncated retrieved as if the value was set to 0 (FALSE).");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, bodyPartPreference);
                lastRawResponse = (XmlElement)this.ASAIRSAdapter.LastRawResponseXml;

                this.VerifyBodyPartElements(searchItem.Email.BodyPart, null, true, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R392");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R392
                Site.CaptureRequirementIfAreEqual<string>(
                    TestSuiteHelper.TruncateData(allData, (int)bodyPartPreference[0].TruncationSize),
                    TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject),
                    392,
                    @"[In AllOrNone (BodyPartPreference)] If the AllOrNone element is not included in the request, the truncated searched as if the value was set to 0 (FALSE).");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirements MS-ASAIRS_R62 and MS-ASAIRS_R282 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R62");

            // Verify MS-ASAIRS requirement: Verify MS-ASAIRS_R62
            Site.CaptureRequirement(
                62,
                @"[In AllOrNone (BodyPartPreference)] [A client can include multiple BodyPartPreference elements in a command request with different values for the Type element] By default, the server returns the data truncated to the size requested by TruncationSize for the Type element that matches the native storage format of the item's Body element (section 2.2.2.9).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R282");

            // Verify MS-ASAIRS requirement: Verify MS-ASAIRS_R282
            Site.CaptureRequirement(
                282,
                @"[In Truncated (BodyPart)] If the value [of the Truncated element] is TRUE, then the body of the item has been truncated.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S01_TC06_BodyPartPreference_NoAllOrNone_NonTruncatedContentReturned
        /// <summary>
        /// This case is designed to test if the AllOrNone (BodyPartPreference) element is not included in the request and the available data doesn't exceed the value of the TruncationSize element, the non-truncated content will be synchronized, searched or retrieved as if the value was set to 0 (FALSE).
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC06_BodyPartPreference_NoAllOrNone_NonTruncatedContentReturned()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPartPreference element
            Request.BodyPartPreference[] bodyPartPreference = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreference);

            this.VerifyBodyPartElements(syncItem.Email.BodyPart, null, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R410");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R410
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.BodyPart.Data,
                syncItem.Email.BodyPart.Data,
                410,
                @"[In AllOrNone (BodyPartPreference)] If the AllOrNone element is not included in the request, the nontruncated content is synchronized as if the value was set to 0 (FALSE).");
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, bodyPartPreference, null);

            this.VerifyBodyPartElements(itemOperationsItem.Email.BodyPart, null, false, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R411");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R411
            Site.CaptureRequirementIfAreEqual<string>(
                allContentItem.Email.BodyPart.Data,
                itemOperationsItem.Email.BodyPart.Data,
                411,
                @"[In AllOrNone (BodyPartPreference)] If the AllOrNone element is not included in the request, the nontruncated content is retrieved as if the value was set to 0 (FALSE).");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, bodyPartPreference);

                this.VerifyBodyPartElements(searchItem.Email.BodyPart, null, false, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R409");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R409
                Site.CaptureRequirementIfAreEqual<string>(
                    allContentItem.Email.BodyPart.Data,
                    searchItem.Email.BodyPart.Data,
                    409,
                    @"[In AllOrNone (BodyPartPreference)] If the AllOrNone element is not included in the request, the nontruncated content is searched as if the value was set to 0 (FALSE).");
                #endregion
            }

            #region Verify common requirements
            // According to above steps, requirement MS-ASAIRS_R283 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R283");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R283
            Site.CaptureRequirement(
                283,
                @"[In Truncated (BodyPart)] If the value [of the Truncated element] is FALSE, or there is no Truncated element, then the body of the item has not been truncated.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S01_TC07_BodyPartPreference_NotIncludedTruncationSize
        /// <summary>
        /// This case is designed to test if the TruncationSize (BodyPartPreference) element is not included, the server will return the same response no matter whether AllOrNone is true or false.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC07_BodyPartPreference_NotIncludedTruncationSize()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email and get the non-truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPartPreference element
            Request.BodyPartPreference[] bodyPartPreferenceAllOrNoneTrue = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };

            Request.BodyPartPreference[] bodyPartPreferenceAllOrNoneFalse = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    AllOrNone = false,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            // Call Sync command with AllOrNone setting to TRUE
            DataStructures.Sync syncItemAllOrNoneTrue = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreferenceAllOrNoneTrue);

            this.VerifyBodyPartElements(syncItemAllOrNoneTrue.Email.BodyPart, true, false, false);

            // Call Sync command with AllOrNone setting to FALSE
            DataStructures.Sync syncItemAllOrNoneFalse = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreferenceAllOrNoneFalse);

            this.VerifyBodyPartElements(syncItemAllOrNoneFalse.Email.BodyPart, false, false, false);

            Site.Log.Add(
                LogEntryKind.Debug,
                "Entire content: {0}, content for AllOrNone TRUE: {1}, content for AllOrNone FALSE: {2}.",
                allContentItem.Email.BodyPart.Data,
                syncItemAllOrNoneTrue.Email.BodyPart.Data,
                syncItemAllOrNoneFalse.Email.BodyPart.Data);

            Site.Assert.IsTrue(
                allContentItem.Email.BodyPart.Data == syncItemAllOrNoneTrue.Email.BodyPart.Data && syncItemAllOrNoneTrue.Email.BodyPart.Data == syncItemAllOrNoneFalse.Email.BodyPart.Data,
                "Server should return the entire content for the request and same response no matter AllOrNone is true or false if the TruncationSize element is absent in Sync command request.");
            #endregion

            #region Verify ItemOperations command related elements
            // Call ItemOperations command with AllOrNone setting to true
            DataStructures.ItemOperations itemOperationsItemAllOrNoneTrue = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItemAllOrNoneTrue.ServerId, null, null, bodyPartPreferenceAllOrNoneTrue, null);

            this.VerifyBodyPartElements(itemOperationsItemAllOrNoneTrue.Email.BodyPart, true, false, false);

            // Call ItemOperations command with AllOrNone setting to false
            DataStructures.ItemOperations itemOperationsItemAllOrNoneFalse = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItemAllOrNoneTrue.ServerId, null, null, bodyPartPreferenceAllOrNoneFalse, null);

            this.VerifyBodyPartElements(itemOperationsItemAllOrNoneFalse.Email.BodyPart, false, false, false);

            Site.Log.Add(
                 LogEntryKind.Debug,
                 "Entire content: {0}, content for AllOrNone TRUE: {1}, content for AllOrNone FALSE: {2}.",
                 allContentItem.Email.BodyPart.Data,
                 itemOperationsItemAllOrNoneTrue.Email.BodyPart.Data,
                 itemOperationsItemAllOrNoneFalse.Email.BodyPart.Data);

            Site.Assert.IsTrue(
                allContentItem.Email.BodyPart.Data == itemOperationsItemAllOrNoneTrue.Email.BodyPart.Data && itemOperationsItemAllOrNoneTrue.Email.BodyPart.Data == itemOperationsItemAllOrNoneFalse.Email.BodyPart.Data,
                "Server should return the entire content for the request and same response no matter AllOrNone is true or false if the TruncationSize element is absent in ItemOperations command request.");
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                // Call Search command with AllOrNone setting to true
                DataStructures.Search searchItemAllNoneTrue = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItemAllOrNoneTrue.Email.ConversationId, null, bodyPartPreferenceAllOrNoneTrue);

                this.VerifyBodyPartElements(searchItemAllNoneTrue.Email.BodyPart, true, false, false);

                // Call Search command with AllOrNone setting to false
                DataStructures.Search searchItemAllNoneFalse = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItemAllOrNoneTrue.Email.ConversationId, null, bodyPartPreferenceAllOrNoneFalse);

                this.VerifyBodyPartElements(searchItemAllNoneFalse.Email.BodyPart, false, false, false);

                Site.Log.Add(
                     LogEntryKind.Debug,
                     "Entire content: {0}, content for AllOrNone TRUE: {1}, content for AllOrNone FALSE: {2}.",
                     allContentItem.Email.BodyPart.Data,
                     searchItemAllNoneTrue.Email.BodyPart.Data,
                     searchItemAllNoneFalse.Email.BodyPart.Data);

                Site.Assert.IsTrue(
                    allContentItem.Email.BodyPart.Data == searchItemAllNoneTrue.Email.BodyPart.Data && searchItemAllNoneTrue.Email.BodyPart.Data == searchItemAllNoneFalse.Email.BodyPart.Data,
                    "Server should return the entire content for the request and same response no matter AllOrNone is true or false if the TruncationSize element is absent in Search command request.");
                #endregion
            }

            #region Verify requirements
            // According to above steps, requirements MS-ASAIRS_R294 and MS-ASAIRS_R400 can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R294");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R294
            Site.CaptureRequirement(
                294,
                @"[In TruncationSize (BodyPartPreference)] If the TruncationSize element is absent, the entire content is used for the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R400");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R400
            Site.CaptureRequirement(
                400,
                @"[In AllOrNone (BodyPartPreference)]  If the TruncationSize element is not included, the server will return the same response no matter whether AllOrNone is true or false.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S01_TC08_BodyPart_Preview
        /// <summary>
        /// This case is designed to test the Preview (BodyPart) element which contains the Unicode plain text message or message part preview returned to the client.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S01_TC08_BodyPart_Preview()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Send an html email and get the none truncated data
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.HTML, subject, body);

            DataStructures.Sync allContentItem = this.GetAllContentItem(subject, this.User2Information.InboxCollectionId);
            #endregion

            #region Set BodyPartPreference element
            Request.BodyPartPreference[] bodyPartPreference = new Request.BodyPartPreference[]
            {
                new Request.BodyPartPreference()
                {
                    Type = 2,
                    Preview = 18,
                    PreviewSpecified = true
                }
            };
            #endregion

            #region Verify Sync command related elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, bodyPartPreference);

            this.VerifyBodyPartPreview(syncItem.Email, allContentItem.Email, bodyPartPreference);
            #endregion

            #region Verify ItemOperations command related elements
            DataStructures.ItemOperations itemOperationsItem = this.GetItemOperationsResult(this.User2Information.InboxCollectionId, syncItem.ServerId, null, null, bodyPartPreference, null);

            this.VerifyBodyPartPreview(itemOperationsItem.Email, allContentItem.Email, bodyPartPreference);
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Verify Search command related elements
                DataStructures.Search searchItem = this.GetSearchResult(subject, this.User2Information.InboxCollectionId, itemOperationsItem.Email.ConversationId, null, bodyPartPreference);

                this.VerifyBodyPartPreview(searchItem.Email, allContentItem.Email, bodyPartPreference);
                #endregion
            }

            #region Verify requirements
            // According to above steps, the following requirements can be captured directly
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R256");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R256
            Site.CaptureRequirement(
                256,
                @"[In Preview (BodyPart)] The Preview element MUST be present in a command response if a BodyPartPreference element (section 2.2.2.11) in the request included a Preview element and the server can honor the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R253");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R253
            Site.CaptureRequirement(
                253,
                @"[In Preview (BodyPart)] The Preview element is an optional child element of the BodyPart element (section 2.2.2.10) that contains the Unicode plain text message or message part preview returned to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R255");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R255
            Site.CaptureRequirement(
                255,
                @"[In Preview (BodyPart)] The Preview element in a response MUST contain no more than the number of characters specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R2599");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R2599
            Site.CaptureRequirement(
                2599,
                @"[In Preview (BodyPartPreference)] [The Preview element] specifies the maximum length of the Unicode plain text message or message part preview to be returned to the client.");
            #endregion

        }
        #endregion

        #region Private methods
        /// <summary>
        /// Verify elements in BodyPart element.
        /// </summary>
        /// <param name="bodyPart">The body part of item.</param>
        /// <param name="allOrNone">The value of AllOrNone element, "null" means the AllOrNone element is not present in request.</param>
        /// <param name="truncated">Whether the content is truncated.</param>
        /// <param name="includedTruncationSize">Whether includes the TruncationSize element.</param>
        private void VerifyBodyPartElements(BodyPart bodyPart, bool? allOrNone, bool truncated, bool includedTruncationSize)
        {
            Site.Assert.IsNotNull(
                bodyPart,
                "The BodyPart element should be included in response when the BodyPartPreference element is specified in request.");

            Site.Assert.AreEqual<byte>(
                1,
                bodyPart.Status,
                "The Status should be 1 to indicate the success of the response in returning Data element content given the BodyPartPreference element settings in the request.");

            // Verify elements when TruncationSize element is absent
            if (!includedTruncationSize)
            {
                Site.Assert.IsTrue(
                    !bodyPart.TruncatedSpecified || (bodyPart.TruncatedSpecified && !bodyPart.Truncated),
                    "The data should not be truncated when the TruncationSize element is absent.");

                // Since the AllOrNone will be ignored when TruncationSize is not included in request, return if includedTruncationSize is false
                return;
            }

            if (truncated)
            {
                Site.Assert.IsTrue(
                    bodyPart.Truncated,
                    "The data should be truncated when the AllOrNone element value is {0} in the request and the available data exceeds the truncation size.",
                    allOrNone);
            }
            else
            {
                Site.Assert.IsTrue(
                    !bodyPart.TruncatedSpecified || (bodyPart.TruncatedSpecified && !bodyPart.Truncated),
                    "The data should not be truncated when the AllOrNone element value is {0} in the request and the truncation size exceeds the available data.", 
                    allOrNone);
            }

            if (bodyPart.Truncated)
            {
                if (Common.IsRequirementEnabled(403, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R403");

                    // Since the EstimatedDataSize element provides an informational estimate of the size of the data associated with the parent element and the body is not null, if the Truncated value is true, the EstimatedDataSize value should not be 0, then requirement MS-ASAIRS_R403 can be captured.
                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R403
                    Site.CaptureRequirementIfAreNotEqual<uint>(
                        0,
                        bodyPart.EstimatedDataSize,
                        403,
                        @"[In Appendix B: Product Behavior] Implementation does include the EstimatedDataSize (BodyPart) element in a response message whenever the Truncated element is set to TRUE. (Exchange Server 2007 SP1 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify the Preview element in BodyPart element.
        /// </summary>
        /// <param name="email">The email item got from server.</param>
        /// <param name="allContentEmail">The email item which has full content.</param>
        /// <param name="bodyPartPreference">A BodyPartPreference object.</param>
        private void VerifyBodyPartPreview(DataStructures.Email email, DataStructures.Email allContentEmail, Request.BodyPartPreference[] bodyPartPreference)
        {
                Site.Assert.IsNotNull(
                email.BodyPart,
                "The BodyPart element should be included in command response when the BodyPartPreference element is specified in command request.");

                Site.Assert.AreEqual<byte>(
                    1,
                    email.BodyPart.Status,
                    "The Status should be 1 to indicate the success of the command response in returning Data element content given the BodyPartPreference element settings in the command request.");

                Site.Assert.IsNotNull(
                   email.BodyPart.Preview,
                   "The Preview element should be present in response if a BodyPartPreference element in the request included a Preview element and the server can honor the request.");

                Site.Assert.IsTrue(
                    email.BodyPart.Preview.Length <= bodyPartPreference[0].Preview,
                    "The Preview element in a response should contain no more than the number of characters specified in the request. The length of Preview element in response is: {0}.",
                    email.BodyPart.Preview.Length);

                Site.Assert.IsTrue(
                    allContentEmail.BodyPart.Data.Contains(email.BodyPart.Preview),
                    "The Preview element in a response should contain the message part preview returned to the client.");
 
        }
        #endregion
    }
}