namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is used to retrieve, insert, update and delete on list item.
    /// </summary>
    [TestClass]
    public class S01_ManageListItem : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion Test Suite Initialization

        #region Test Case Initialization

        /// <summary>
        /// This method will run after test case executes.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.TestCleanup();
            this.DeleteListItems(this.GeneralListName);
            this.DeleteListItems(this.DocumentLibraryName);
        }

        #endregion Test Case Initialization

        /// <summary>
        /// This test case is used to create an attachment in list item.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC01_CreateAttachmentInListItem()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert list item to list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve attachment count in list item before add attachment
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}({1})/Attachments/$count", this.GeneralListName, insertResult.Properties[Constants.IDFieldName]);
            retrieveRequest.Accept = "text/plain";
            int attachmentCountBeforeAdd = Convert.ToInt32(this.Adapter.RetrieveListItem(retrieveRequest));

            // Add attachment to list item
            Request addAttatchment = new Request();
            addAttatchment.Content = Common.GenerateResourceName(this.Site, "attachmentContent");
            addAttatchment.Slug = string.Format("{0}|{1}|{2}.txt", this.GeneralListName, insertResult.Properties[Constants.IDFieldName], Common.GenerateResourceName(this.Site, "attachmentName"));
            addAttatchment.Parameter = "Attachments";
            addAttatchment.ContentType = "text/plain";
            Entry addResult = this.Adapter.InsertListItem(addAttatchment);
            Site.Assert.IsNotNull(addResult, "Verify addResult is not null");

            // Retrieve attachment count in list item after add attachment
            int attachmentCountAfterAdd = Convert.ToInt32(this.Adapter.RetrieveListItem(retrieveRequest));

            // If create the attachment successfully, the requirements: MS-WSSREST_R105 and MS-WSSREST_R82 can be verified.
            Site.CaptureRequirementIfAreEqual<int>(
                attachmentCountAfterAdd,
                attachmentCountBeforeAdd + 1,
                105,
                @"[In Message Processing Events and Sequencing Rules] The element Attachment supports for create operation.");

            Site.CaptureRequirementIfAreEqual<int>(
                attachmentCountAfterAdd,
                attachmentCountBeforeAdd + 1,
                82,
                @"[In Attachment] Using the preceding Augmented BNF causes an attachment to be created on the protocol server.");
        }

        /// <summary>
        /// This test case is used to retrieve the value of single choice field.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC02_RetrieveSingleChoice()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert list item to list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve the single choice field
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}({1})/{2}", this.GeneralListName, insertResult.Properties[Constants.IDFieldName], Common.GetConfigurationPropertyValue("ChoiceFieldName", this.Site));
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify retrieveResult is not null");

            // Retrieve the single choice field EntitySet
            Request retrieveChoiceEntitySet = new Request();
            retrieveChoiceEntitySet.Parameter = string.Format("{0}{1}", Common.GetConfigurationPropertyValue("GeneralListName", this.Site), Common.GetConfigurationPropertyValue("ChoiceFieldName", this.Site));
            retrieveChoiceEntitySet.Accept = "application/atom+xml";
            List<Entry> retrievEntitySetResult = this.Adapter.RetrieveListItem(retrieveChoiceEntitySet) as List<Entry>;
            Site.Assert.IsNotNull(retrievEntitySetResult, "Verify retrievEntitySetResult is not null");
            string[] choiceOpertions = Common.GetConfigurationPropertyValue("ChoicefieldOptions", this.Site).Split(',');

            // If the count of the entity set returned equals to the count of the single choice options, the MS-WSSREST_R86 can be verified.
            Site.CaptureRequirementIfAreEqual<int>(
                choiceOpertions.Length,
                retrievEntitySetResult.Count,
                86,
                @"[In Choice or Multi-Choice Field] For a particular choice [or multi-choice field (2)], this EntitySet will contain Entities representing each option of the field (2) specified in the field definition of the field (2).");

            // If the retrievEntitySetResult is not null, it means the EntitySet is created for every single choice, so the MS-WSSREST_R84 is directly covered.
            Site.CaptureRequirementIfIsNotNull(
                retrievEntitySetResult,
                84,
                @"[In Choice or Multi-Choice Field] An EntitySet is created for every choice [or multi-choice] field (2) belonging to a list (1).");
        }

        /// <summary>
        /// This test case is used to retrieve the value of multiple choice field.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC03_RetrieveMultipleChoice()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert list item to list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve the multiple choice field
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}({1})/{2}", this.GeneralListName, insertResult.Properties[Constants.IDFieldName], Common.GetConfigurationPropertyValue("MultiChoiceFieldName", this.Site));
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify retrieveResult is not null");

            // Retrieve the multiple choice field EntitySet
            Request retrieveChoiceEntitySet = new Request();
            retrieveChoiceEntitySet.Parameter = string.Format("{0}{1}", Common.GetConfigurationPropertyValue("GeneralListName", this.Site), Common.GetConfigurationPropertyValue("MultiChoiceFieldName", this.Site));
            retrieveChoiceEntitySet.Accept = "application/atom+xml";
            List<Entry> retrieveEntitySetResult = this.Adapter.RetrieveListItem(retrieveChoiceEntitySet) as List<Entry>;
            Site.Assert.IsNotNull(retrieveEntitySetResult, "Verify retrieveEntitySetResult is not null");
            string[] multiChoiceOpertions = Common.GetConfigurationPropertyValue("MultiChoicefieldOptions", this.Site).Split(',');

            // If the returned entity set count equals to the multi-choice options' count, the MS-WSSREST_R87 can be verified.
            Site.CaptureRequirementIfAreEqual<int>(
                multiChoiceOpertions.Length,
                retrieveEntitySetResult.Count,
                87,
                @"[In Choice or Multi-Choice Field] For a particular [choice or] multi-choice field (2), this EntitySet will contain Entities representing each option of the field (2) specified in the field definition of the field (2).");

            // If the retrieveEntitySetResult is not null, it means the EntitySet is created for multi-choice, so the MS-WSSREST_R85 is directly covered.
            Site.CaptureRequirementIfIsNotNull(
                retrieveEntitySetResult,
                85,
                @"[In Choice or Multi-Choice Field] An EntitySet is created for [every choice or] multi-choice field (2) belonging to a list (1).");
        }

        /// <summary>
        /// This test case is used to delete an attachment from list item.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC04_DeleteAttachmentFromListItem()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert list item to list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Add attachment to list item
            Request addAttatchment = new Request();
            addAttatchment.Content = Common.GenerateResourceName(this.Site, "attachmentContent");
            addAttatchment.Slug = string.Format("{0}|{1}|{2}", this.GeneralListName, insertResult.Properties[Constants.IDFieldName], Common.GenerateResourceName(this.Site, "attachmentName"));
            addAttatchment.Parameter = "Attachments";
            addAttatchment.ContentType = "text/plain";
            Entry addResult = this.Adapter.InsertListItem(addAttatchment);
            Site.Assert.IsNotNull(addResult, "Verify addResult is not null");

            // Delete attachment from list item
            Request deleteAttachment = new Request();
            deleteAttachment.Parameter = string.Format("Attachments(EntitySet='{0}',ItemId={1},Name='{2}')", addResult.Properties["EntitySet"], addResult.Properties["ItemId"], addResult.Properties["Name"]);
            deleteAttachment.ETag = insertResult.Etag;
            bool isAttachmentDeleted = this.Adapter.DeleteListItem(deleteAttachment);

            // If the isAttachmentDeleted returns "true", it means the attachment is deleted form list item successfully, the requirements: MS-WSSREST_R68 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "If the isAttachmentDeleted returns 'true', it means the attachment is deleted form list item successfully, the requirements: MS-WSSREST_R68 can be verified.");

            Site.CaptureRequirementIfIsTrue(
                isAttachmentDeleted,
                68,
                @"[In Message Processing Events and Sequencing Rules] The element Attachment supports for delete operation.");
        }

        /// <summary>
        /// This test case is used to insert a new document into a document library.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC05_InsertDocumentToDocumentLibrary()
        {
            string documentContent = Common.GenerateResourceName(this.Site, "documentContent");
            string slug = string.Format("{0}/{1}/{2}", Common.GetConfigurationPropertyValue("DocumentLibrariesPath", this.Site), this.DocumentLibraryName, Common.GenerateResourceName(this.Site, "documentName"));

            // Retrieve document count under list before add document
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}/$count", this.DocumentLibraryName);
            retrieveRequest.Accept = "text/plain";
            int documentCountBeforeAdd = Convert.ToInt32(this.Adapter.RetrieveListItem(retrieveRequest));

            // Insert a document into document library
            Request addDocRequest = new Request();
            addDocRequest.Parameter = this.DocumentLibraryName;
            addDocRequest.Content = documentContent;
            addDocRequest.ContentType = "text/plain";
            addDocRequest.Slug = slug;
            Entry addDocResult = this.Adapter.InsertListItem(addDocRequest);
            Site.Assert.IsNotNull(addDocResult, "Verify addDocResult is not null");

            // Retrieve document count under list after add document
            int documentCountAfterAdd = Convert.ToInt32(this.Adapter.RetrieveListItem(retrieveRequest));

            // If insert the document to document library successfully, the requirements: MS-WSSREST_R70 and MS-WSSREST_R100 can be verified.
            Site.CaptureRequirementIfAreEqual<int>(
                documentCountAfterAdd,
                documentCountBeforeAdd + 1,
                70,
                @"[In Message Processing Events and Sequencing Rules] The element Document supports for inserting a new document into a document library.");

            Site.CaptureRequirementIfAreEqual<int>(
                documentCountAfterAdd,
                documentCountBeforeAdd + 1,
                100,
                @"[In Document] Using the preceding Augmented BNF causes a document to be created on the protocol server.");

            // Get document library content type id
            string documentLibraryContentTypeId = this.SutAdapter.GetDocumentLibraryContentTypeId(this.DocumentLibraryName);

            // Retrieve document content type id under document library
            retrieveRequest.Parameter = this.DocumentLibraryName;
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResults = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            string documentContentTypeId = string.Empty;
            if (retrieveResults != null && retrieveResults.Count > 0)
            {
                documentContentTypeId = retrieveResults[0].Properties["ContentTypeID"];
            }

            // If the document content type id equals to the document library content type id when the content type id is omitted, the requirement: MS-WSSREST_R98 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                documentLibraryContentTypeId.ToString(),
                documentContentTypeId.ToString(),
                98,
                @"[In Document] If the content type identifier is omitted, the document will be assigned the default content type identifier of the document library.");
        }

        /// <summary>
        /// This test case is used to retrieve a collection of list items.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC06_RetrieveACollectionOfListItems()
        {
            // Insert two list items to list
            this.InsertListItems(this.GeneralListName, 2);

            // Retrieve all the list items under list
            Request retrieveCollection = new Request();
            retrieveCollection.Parameter = this.GeneralListName;
            retrieveCollection.Accept = "application/atom+xml";
            List<Entry> retreiveCollectionResult = this.Adapter.RetrieveListItem(retrieveCollection) as List<Entry>;
            Site.Assert.IsNotNull(retreiveCollectionResult, "Verify retreiveCollectionResult is not null");
        }

        /// <summary>
        /// This test case is used to retrieve the count of a collection of list items.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC07_RetrieveTheCountOfACollectionOfListItems()
        {
            // Retrieve the count of list items under list before add list items.
            Request retrieveCount = new Request();
            retrieveCount.Parameter = string.Format("{0}/$count", this.GeneralListName);
            retrieveCount.Accept = "text/plain";
            int countOfListItemsBefore = Convert.ToInt32(this.Adapter.RetrieveListItem(retrieveCount));

            // Insert two list items to list
            this.InsertListItems(this.GeneralListName, 2);

            // Retrieve the count of list items under list after list items.
            int countOfListItemsAfter = Convert.ToInt32(this.Adapter.RetrieveListItem(retrieveCount));

            Site.Assert.AreEqual<int>(countOfListItemsBefore + 2, countOfListItemsAfter, "Verify the count of a collection of list items are retrieved");
        }

        /// <summary>
        /// This test case is used to retrieve a collection of entities with inline count. 
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC08_RetrieveACollectionOfEntitiesWithAnInlineCount()
        {
            // Insert two list items to list
            this.InsertListItems(this.GeneralListName, 2);

            // Retrieve the collection of entities with inline count. 
            Request retrieveCollection = new Request();
            retrieveCollection.Parameter = string.Format("{0}?$inlinecount=allpages", this.GeneralListName);
            retrieveCollection.Accept = "application/atom+xml";
            List<Entry> retrieveCollectionResult = this.Adapter.RetrieveListItem(retrieveCollection) as List<Entry>;
            Site.Assert.IsNotNull(retrieveCollectionResult, "Verify retrieveCollectionResult is not null");
        }

        /// <summary>
        /// This test case is used to retrieve the first two list items in a list.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC09_RetrieveTheFirstTwoListItemsInAList()
        {
            // Insert three list items to list
            this.InsertListItems(this.GeneralListName, 3);

            // Retrieve the first two list items 
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}?$top=2", this.GeneralListName);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify the retrieveResult is not null");
            Site.Assert.AreEqual<int>(2, retrieveResult.Count, "Verify retrieving the first two list items");
        }

        /// <summary>
        /// This test case is used to retrieve the last two list items in a list.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC10_RetrieveTheLastTwoListItemsInAList()
        {
            // Insert three list items to list
            this.InsertListItems(this.GeneralListName, 3);

            // Retrieve the last two list items 
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}?$skip=1", this.GeneralListName);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify the retrieveResult is not null");
            Site.Assert.AreEqual<int>(2, retrieveResult.Count, "Verify retrieving the last two list items");
        }

        /// <summary>
        /// This test case is used to retrieve the middle two list items in a list.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC11_RetrieveTheMiddleTwoListItemsInAList()
        {
            // Insert four list items to list
            this.InsertListItems(this.GeneralListName, 4);

            // Retrieve the middle two list items
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}?$skip=1&$top=2", this.GeneralListName);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify the retrieveResult is not null");
            Site.Assert.AreEqual<int>(2, retrieveResult.Count, "Verify retrieving the middle two list items");
        }

        /// <summary>
        /// This test case is used to retrieve a list item by id.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC12_RetrieveAListItemByID()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert list item to list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve list item by id
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}({1})", this.GeneralListName, insertResult.Properties[Constants.IDFieldName]);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify retrieveResult is not null");
        }

        /// <summary>
        /// This test case is used to retrieve list items that satisfy the specified criteria.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC13_RetrieveListItemsThatSatisfySpecifiedCriteria()
        {
            string listItemTitle = Common.GenerateResourceName(this.Site, "SpecifiedTitle");
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, listItemTitle);

            // Insert list items to list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResultFirst = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResultFirst, "Verify insertResultFirst is not null");

            // Retrieve the list item which contains listItemTitle
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}?$filter=Title%20eq%20'{1}'", this.GeneralListName, listItemTitle);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveCriteriaResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveCriteriaResult, "Verify the retrieveCriteriaResult is not null");
            Site.Assert.AreEqual<string>(listItemTitle, retrieveCriteriaResult[0].Properties[Constants.TitleFieldName], "Verify retrieving the satisfy specified criteria list item");
        }

        /// <summary>
        /// This test case is used to retrieve list items by id descending.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC14_RetrieveListItemsInASpecifiedOrder()
        {
            // Insert two list items to list
            this.InsertListItems(this.GeneralListName, 2);

            // Retrieve the list items by id descending
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}?$orderby=Id%20desc%20", this.GeneralListName);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify the retrieveResult is not null");

            bool isIdDescending = false;
            if (retrieveResult.Count >= 2)
            {
                isIdDescending = Convert.ToInt32(retrieveResult[0].Properties[Constants.IDFieldName]) > Convert.ToInt32(retrieveResult[1].Properties[Constants.IDFieldName]);
            }

            Site.Assert.IsTrue(isIdDescending, "Verify retrieveResult is by id descending");
        }

        /// <summary>
        /// This test case is used to insert a new list item into list.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC15_InsertANewListItem()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Retrieve the added list item before add list item.
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = this.GeneralListName;
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResultBefore = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResultBefore, "Verify the retrieveResultBefore is not null");

            // Insert a list item into list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve the added list item after add list item.
            List<Entry> retrieveResultAfter = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResultAfter, "Verify the retrieveResultAfter is not null");

            Site.Assert.AreEqual<int>(retrieveResultBefore.Count + 1, retrieveResultAfter.Count, "Verify retrieving the added list item");
        }

        /// <summary>
        /// This test case is used to update a new list item using replace-based semantics.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC16_UpdateAListItemUsingReplaceBasedSemantics()
        {
            string titleBeforeUpdate = Common.GenerateResourceName(this.Site, "title");
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, titleBeforeUpdate);

            // Insert a list item into list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve the title of list item before update
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}({1})", this.GeneralListName, insertResult.Properties[Constants.IDFieldName]);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResultBeforeUpdate = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResultBeforeUpdate, "Verify retrieveResultBeforeUpdate is not null!");
            Site.Assert.AreEqual<int>(1, retrieveResultBeforeUpdate.Count, "Verify retrieving list item before updating");

            // Update list item title
            Request updateRequest = new Request();
            updateRequest.Parameter = string.Format("{0}({1})", this.GeneralListName, retrieveResultBeforeUpdate[0].Properties[Constants.IDFieldName]);
            properties.Clear();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "titlePut"));
            updateRequest.Content = this.GenerateContent(properties);
            updateRequest.ContentType = "application/atom+xml";
            updateRequest.UpdateMethod = UpdateMethod.PUT;
            updateRequest.ETag = retrieveResultBeforeUpdate[0].Etag;
            string updateResult = this.Adapter.UpdateListItem(updateRequest);
            Site.Assert.IsNotNull(updateResult, "Verify updateResult is not null");

            // Retrieve the title of list item after update
            List<Entry> retrieveResultAfterUpdate = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResultAfterUpdate, "Verify retrieveResultAfterUpdate is not null!");
            Site.Assert.AreEqual<int>(1, retrieveResultAfterUpdate.Count, "Verify retrieving list item after updating");
            string titleAfterUpdate = retrieveResultAfterUpdate[0].Properties[Constants.TitleFieldName];
            Site.Assert.AreNotEqual<string>(titleBeforeUpdate, titleAfterUpdate, "Verify updating list item title");
        }

        /// <summary>
        /// This test case is used to update a new list item using merge-based semantics.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC17_UpdateAListItemUsingMergeBasedSemantics()
        {
            string titleBeforeUpdate = Common.GenerateResourceName(this.Site, "title");
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, titleBeforeUpdate);

            // Insert a list item into list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve the title of list item before update
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}({1})", this.GeneralListName, insertResult.Properties[Constants.IDFieldName]);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResultBeforeUpdate = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResultBeforeUpdate, "Verify retrieveResultBeforeUpdate is not null!");
            Site.Assert.AreEqual<int>(1, retrieveResultBeforeUpdate.Count, "Verify retrieving list item before updating");

            // Update list item title
            Request updateRequest = new Request();
            updateRequest.Parameter = string.Format("{0}({1})", this.GeneralListName, retrieveResultBeforeUpdate[0].Properties["Id"]);
            properties.Clear();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "titleMerge"));
            updateRequest.Content = this.GenerateContent(properties);
            updateRequest.ContentType = "application/atom+xml";
            updateRequest.UpdateMethod = UpdateMethod.MERGE;
            updateRequest.ETag = retrieveResultBeforeUpdate[0].Etag;
            string updateResult = this.Adapter.UpdateListItem(updateRequest);
            Site.Assert.IsNotNull(updateResult, "Verify updateResult is not null");

            // Retrieve the title of list item after update
            List<Entry> retrieveResultAfterUpdate = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResultAfterUpdate, "Verify retrieveResultAfterUpdate is not null!");
            Site.Assert.AreEqual<int>(1, retrieveResultAfterUpdate.Count, "Verify retrieving list item after updating");
            string titleAfterUpdate = retrieveResultAfterUpdate[0].Properties[Constants.TitleFieldName];
            Site.Assert.AreNotEqual<string>(titleBeforeUpdate, titleAfterUpdate, "Verify updating list item title");
        }

        /// <summary>
        /// This test case is used to delete a list item from list.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S01_TC18_DeleteAListItem()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert a list item into list
            Request insertRequest = new Request();
            insertRequest.Parameter = this.GeneralListName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";
            Entry insertResult = this.Adapter.InsertListItem(insertRequest);
            Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null");

            // Retrieve the added list item
            Request retrieveRequest = new Request();
            retrieveRequest.Parameter = string.Format("{0}({1})", this.GeneralListName, insertResult.Properties[Constants.IDFieldName]);
            retrieveRequest.Accept = "application/atom+xml";
            List<Entry> retrieveResult = this.Adapter.RetrieveListItem(retrieveRequest) as List<Entry>;
            Site.Assert.IsNotNull(retrieveResult, "Verify retrieveResult is not null!");
            Site.Assert.AreEqual<int>(1, retrieveResult.Count, "Verify retrieving the added list item");

            // Delete list item from list
            Request deleteRequest = new Request();
            deleteRequest.Parameter = string.Format("{0}({1})", this.GeneralListName, insertResult.Properties[Constants.IDFieldName]);
            bool deleteResult = this.Adapter.DeleteListItem(deleteRequest);
            Site.Assert.IsTrue(deleteResult, "Verify deleting the added list items");
        }
    }
}