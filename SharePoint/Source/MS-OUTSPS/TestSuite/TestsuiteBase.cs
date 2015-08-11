namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Schema;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS_OUTSPS.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// A string indicate the errorMessage Format.
        /// </summary>
        protected const string ErrorMessageTemplate = "An error occurred while {0}, the error message is {1}";

        /// <summary>
        /// Gets or sets an instance of interface IMS_OUTSPSAdapter
        /// </summary>
        protected static IMS_OUTSPSAdapter OutspsAdapter { get; set; }

        /// <summary>
        /// Gets or sets an instance of interface IMS_OUTSPSSUTControlAdapter
        /// </summary>
        protected static IMS_OUTSPSSUTControlAdapter SutControlAdapter { get; set; }

        /// <summary>
        /// Gets or sets a random generator using current time seeds.
        /// </summary>
        protected static Random RandomInstance { get; set; }

        /// <summary>
        /// Gets or sets a list type instance used to record all lists added by TestSuiteHelper
        /// </summary>
        protected static List<string> ListIdCache { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the list item number value on current test case.
        /// </summary>
        protected static uint ListitemCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the contentType number value on current test case.
        /// </summary>
        protected static uint ContentTypeCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the contentType number value on current test case.
        /// </summary>
        protected static uint ListNameCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicating the metaInfoField PropertyName number value on current test case.
        /// </summary>
        protected static uint MetaInfoFieldPropertyNameCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the folder name number value on current test case.
        /// </summary>
        protected static uint FolderNameCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the field name number value on current test case.
        /// </summary>
        protected static uint FieldNameCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the attachment name number value on current test case.
        /// </summary>
        protected static uint AttachmentNameCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the file name number value on current test case.
        /// </summary>
        protected static uint FiletNameCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the list item title's number value on current test case.
        /// </summary>
        protected static uint ListItemTitleCounterOfPerTestCases { get; set; }

        #endregion

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the variable for the test suite.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            // A method is used to initialize the variables.
            TestClassBase.Initialize(testContext);
            if (null == OutspsAdapter)
            {
                OutspsAdapter = BaseTestSite.GetAdapter<IMS_OUTSPSAdapter>();
            }

            if (null == SutControlAdapter)
            {
                SutControlAdapter = BaseTestSite.GetAdapter<IMS_OUTSPSSUTControlAdapter>();
            }

            if (null == ListIdCache)
            {
                ListIdCache = new List<string>();
            }
        }

        /// <summary>
        /// A method is used to clean up the test suite.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        /// <summary>
        /// This method will run before test case executes
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            Common.CheckCommonProperties(this.Site, true);

            #region initialization

            // If added lists are not clean up, throw an exception.
            if (ListIdCache.Count != 0)
            {
                Site.Debug.Fail("The test environment is not clean, refer the log files for details.");
            }

            // Initialize the unique resource counter
            ContentTypeCounterOfPerTestCases = 0;
            ListitemCounterOfPerTestCases = 0;
            ListNameCounterOfPerTestCases = 0;
            MetaInfoFieldPropertyNameCounterOfPerTestCases = 0;
            FolderNameCounterOfPerTestCases = 0;
            FieldNameCounterOfPerTestCases = 0;
            AttachmentNameCounterOfPerTestCases = 0;
            ListItemTitleCounterOfPerTestCases = 0;

            #endregion
        }

        /// <summary>
        /// This method will run after test case executes
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.CleanUpAddedList();
        }

        #endregion Test Suite Initialization

        #region Helper methods

        /// <summary>
        /// A method used to construct CamlViewFields instance using the specified parameters.
        /// </summary>
        /// <param name="property">A Boolean value indicate whether the properties attribute value is TRUE/FALSE</param>
        /// <param name="fieldNames">Specified the CamlViewFields instance's fields.</param>
        /// <returns>Returns the CamlViewFields instance.</returns>
        protected CamlViewFields GenerateViewFields(bool property, List<string> fieldNames)
        {
            CamlViewFields viewFields = new CamlViewFields();
            viewFields.ViewFields = new CamlViewFieldsViewFields();
            viewFields.ViewFields.Properties = property == true ? "TRUE" : "FALSE";

            int fieldCount = fieldNames.Count;
            viewFields.ViewFields.FieldRef = new CamlViewFieldsViewFieldsFieldRef[fieldCount];

            for (int i = 0; i < fieldCount; i++)
            {
                viewFields.ViewFields.FieldRef[i] = new CamlViewFieldsViewFieldsFieldRef();
                viewFields.ViewFields.FieldRef[i].Name = fieldNames[i];
            }

            return viewFields;
        }

        /// <summary>
        /// A method used to Create a list by configured list name and specified list template. 
        /// </summary>
        /// <param name="listTemplate">A parameter represents the template type of list template</param>
        /// <returns>A return value represents the added list's id</returns>
        protected string AddListToSUT(TemplateType listTemplate)
        {
            string uniqueListName = this.GetUniqueListName(listTemplate.ToString());
            string listId = this.AddListToSUT(uniqueListName, listTemplate);
            return listId;
        }

        /// <summary>
        ///  A method used to create a list by specified list name and template type.
        /// </summary>
        /// <param name="listName">A parameter represents the list name which is used in add list process.</param>
        /// <param name="listTemplate">A parameter represents the list template which the new list used.</param>
        /// <returns>A return value represents the added list's id</returns>
        protected string AddListToSUT(string listName, TemplateType listTemplate)
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw new ArgumentNullException("listName");
            }

            string listId = null;
            try
            {
                AddListResponseAddListResult addListResult = null;
                addListResult = OutspsAdapter.AddList(listName, null, (int)listTemplate);
                if (null == addListResult || string.IsNullOrEmpty(addListResult.List.ID))
                {
                    this.Site.Assert.Fail("Could not get the added list's id. List title[{0}]", listName);
                }

                // Record the added list id, so that test suite can clean up added list.
                listId = addListResult.List.ID;
                ListIdCache.Add(listId);

                // Make sure the created list contain the following field,
                // This field will be used in the following sequence.
                string expectedFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
                List<string> listNames = new List<string> { expectedFieldName };
                List<string> listTypes = new List<string> { "Text" };

                // Add field
                this.AddFieldsToList(listId, listNames, listTypes);
            }
            catch (SoapException exp)
            {
                this.Site.Assert.Fail(ErrorMessageTemplate, "creating the list " + listName, exp.Detail.InnerText);
            }

            return listId;
        }

        /// <summary>
        /// This method is used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>Returns the random generated string.</returns>
        protected string GenerateRandomString(int size)
        {
            // Sleep 10 milliseconds to prevent the random instance was created too quickly. If this method is called several times in a short timeSpan, the random instance will present same "next" value.
            System.Threading.Thread.Sleep(10);
            RandomInstance = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * RandomInstance.NextDouble()) + 65));
                ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        /// <summary>
        /// This method is used to random generate an integer in the specified range.
        /// </summary>
        /// <param name="minValue">The inclusive lower bound of the random number returned.</param>
        /// <param name="maxValue">The exclusive upper bound of the random number returned.</param>
        /// <returns>A 32-bit signed integer greater than or equal to minValue and less than maxValue</returns>
        protected string GenerateRandomNumber(int minValue, int maxValue)
        {
            // Sleep 10 milliseconds to prevent the random instance was created too quickly. If this method is called several times in a short timeSpan, the random instance will present same "next" value.
            System.Threading.Thread.Sleep(10);
            RandomInstance = new Random((int)DateTime.Now.Ticks);
            return RandomInstance.Next(minValue, maxValue).ToString();
        }

        /// <summary>
        /// A method used to Clean up added lists.
        /// Any lists added by Test Case Helper would be deleted.
        /// </summary>
        protected void CleanUpAddedList()
        {
            // Make a copy to loop in the copy array.
            List<string> listIdCopy = new List<string>(ListIdCache.ToArray());

            foreach (string listId in listIdCopy)
            {
                try
                {
                    OutspsAdapter.DeleteList(listId);
                    ListIdCache.Remove(listId);
                }
                catch (SoapException exp)
                {
                    this.Site.Log.Add(LogEntryKind.Warning, ErrorMessageTemplate, "removing the list " + listId, exp.Detail.InnerText);
                }
            }
        }

        /// <summary>
        /// A method is used to get a Unique List Name
        /// </summary>
        /// <param name="listTemplateType">A parameter represents the type of the list.</param>
        /// <returns>A return value represents the unique name that is combined with the List Object name and time stamp</returns>
        protected string GetUniqueListName(string listTemplateType)
        {
            ListNameCounterOfPerTestCases += 1;
            string listResouceName = string.Format("{0}List", listTemplateType);
            return Common.GenerateResourceName(this.Site, listResouceName, ListNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique List item Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the ListItem Object name and time stamp</returns>
        protected string GetUniqueListItemName()
        {
            ListitemCounterOfPerTestCases += 1;
            return Common.GenerateResourceName(this.Site, "ListItem", ListitemCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique ContentType Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the ContentType Object name and time stamp</returns>
        protected string GetUniqueContentTypeName()
        {
            ContentTypeCounterOfPerTestCases += 1;
            return Common.GenerateResourceName(this.Site, "ContentType", ContentTypeCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique MetaInfoField Property Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the MetaInfoField Property Object name and time stamp</returns>
        protected string GetUniqueMetaInfoFieldPropertyName()
        {
            MetaInfoFieldPropertyNameCounterOfPerTestCases += 1;
            return Common.GenerateResourceName(this.Site, "MetaInfoFieldProperty", MetaInfoFieldPropertyNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique attachment Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the attachment Object name and time stamp</returns>
        protected string GetUniqueAttachmentName()
        {
            AttachmentNameCounterOfPerTestCases += 1;
            string attachmentName = Common.GenerateResourceName(this.Site, "Attachment", AttachmentNameCounterOfPerTestCases);
            return string.Format("{0}.txt", attachmentName);
        }

        /// <summary>
        /// A method is used to get a unique name of file which is used to upload into document library.
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the file Object name and time stamp</returns>
        protected string GetUniqueUploadFileName()
        {
            FiletNameCounterOfPerTestCases += 1;
            string fileName = Common.GenerateResourceName(this.Site, "file", FiletNameCounterOfPerTestCases);
            return string.Format("{0}.txt", fileName);
        }

        /// <summary>
        /// A method is used to get a unique name of list item title.
        /// </summary>
        /// <param name="prefixOfListItem">A parameter represents the prefix of the list item. It will present as '[prefixOfListItem]Item' in the unique name.</param>
        /// <returns>A return value represents the unique name that is combined with the file Object name and time stamp</returns>
        protected string GetUniqueListItemTitle(string prefixOfListItem)
        {
            ListItemTitleCounterOfPerTestCases += 1;
            string listtitleResourceName = string.Format("{0}Item", prefixOfListItem);
            string listtitle = Common.GenerateResourceName(this.Site, listtitleResourceName, FiletNameCounterOfPerTestCases);
            return listtitle;
        }

        /// <summary>
        /// A method is used to get a valid format message data for AddDiscussionBoardItem operation. This method will generate a unique title, unique Thread-index and unique Message-ID for the message data of a DiscussionBoard item.
        /// </summary>
        /// <returns>A return value represents the valid format message data</returns>
        protected byte[] GetMessageDataForAddDiscussionBoardItem()
        {
            // read the DiscussionBoard Item message file to a string
            string messageFilePath = Common.GetConfigurationPropertyValue("MessageDataFileName", this.Site);
            string fileContentsTemp = string.Empty;
            StreamReader streasmReader = null;

            try
            {
                streasmReader = new StreamReader(File.OpenRead(messageFilePath));
                fileContentsTemp = streasmReader.ReadToEnd();

                // Set the subject value.
                string currentDiscussionTitle = this.GetUniqueListItemTitle("DiscussionItems");
                fileContentsTemp = this.SetDiscussionItemMessageProperty(fileContentsTemp, "Subject", currentDiscussionTitle);

                // Get a sub string from a combination of GUID value and random string, make it have 22 bytes.
                Guid guidTemp = Guid.NewGuid();
                string guidTempValue = guidTemp.ToString("N");
                guidTempValue = guidTempValue.Substring(0, 18);
                guidTempValue = string.Format("{0}{1}{2}", this.GenerateRandomString(2), guidTempValue, this.GenerateRandomString(2));

                // Tread-index Base64 [RFC4648] encoded string that is at least 22 bytes long when un-encoded, more detail is in [MS-LISTSWS] section 3.1.4.2.2.1.
                byte[] guidTempBytes = Encoding.ASCII.GetBytes(guidTempValue);
                string encodeTreadIndexValue = Convert.ToBase64String(guidTempBytes);

                // Set the Tread-index value.
                fileContentsTemp = this.SetDiscussionItemMessageProperty(fileContentsTemp, "Thread-Index", encodeTreadIndexValue);

                // Set the Message-ID value.
                guidTemp = Guid.NewGuid();
                fileContentsTemp = this.SetDiscussionItemMessageProperty(fileContentsTemp, "Message-ID", guidTemp.ToString("N"));
            }
            finally
            {
                if (streasmReader != null)
                {
                    streasmReader.Dispose();
                }
            }

            this.Site.Assert.IsFalse(
                            string.IsNullOrEmpty(fileContentsTemp),
                            "The file contents should have value.");

            // Convert to a byte array
            byte[] messageData = Encoding.Default.GetBytes(fileContentsTemp);
            return messageData;
        }

        /// <summary>
        /// A method used to generate unique contents for attachment.
        /// </summary>
        /// <returns>A return value represents the unique contents for attachment.</returns>
        protected byte[] GenerateUniqueAttachmentContents()
        {
            return this.GenerateUniqueAttachmentContents(5);
        }

        /// <summary>
        /// A method used to generate unique contents for attachment.
        /// </summary>
        /// <param name="attachmentStringLength">A parameter represents the length of the string content in attachment.</param>
        /// <returns>A return value represents the unique contents for attachment.</returns>
        protected byte[] GenerateUniqueAttachmentContents(int attachmentStringLength)
        {
            if (0 == attachmentStringLength)
            {
                throw new ArgumentException("The value should be large than Zero", "attachmentStringLength");
            }

            string randonString = this.GenerateRandomString(attachmentStringLength);
            string contentString = string.Format("{0}Test:{1}", this.Site.DefaultProtocolDocShortName, randonString);
            return Encoding.UTF8.GetBytes(contentString);
        }

        /// <summary>
        /// A method is used to get a Unique folder Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the folder Object name and time stamp</returns>
        protected string GetUniqueFolderName()
        {
            FolderNameCounterOfPerTestCases += 1;
            return Common.GenerateResourceName(this.Site, "Folder", FolderNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique field Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the field Object name and time stamp</returns>
        protected string GetUniqueFieldName()
        {
            FieldNameCounterOfPerTestCases += 1;
            return Common.GenerateResourceName(this.Site, "Field", FieldNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to add list items by specified number items to specified list.
        /// This method update the column value specified by ListFieldText property in the PTF configuration file, and the value of this column is specified by random generated string.
        /// </summary>
        /// <param name="list">A parameter represents list GUID or list title</param>
        /// <param name="itemCount">A parameter represents how many items would be added into the list</param>
        /// <returns>A return value represents added items id array.</returns>
        protected List<string> AddItemsToList(string list, int itemCount)
        {
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(itemCount);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(itemCount);

            for (int i = 0; i < itemCount; i++)
            {
                Dictionary<string, string> item = new Dictionary<string, string>();

                // add an item and set the field to a random string.
                item.Add(fieldName, this.GenerateRandomString(5));
                items.Add(item);
            }

            for (int i = 0; i < itemCount; i++)
            {
                cmds.Add(MethodCmdEnum.New);
            }

            UpdateListItemsUpdates updates = this.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult result = null;
            try
            {
                result = OutspsAdapter.UpdateListItems(
                                                    list,
                                                    updates);
            }
            catch (SoapException exp)
            {
                this.Site.Debug.Fail(ErrorMessageTemplate, "adding items to the list " + list + " and updating field " + fieldName, exp.Detail.InnerText);
            }

            // There MUST be one Column called ID.
            string columnNameId = string.Format("{0}{1}", "ows_", "ID");
            List<string> ids = new List<string>();
            foreach (UpdateListItemsResponseUpdateListItemsResultResult res in result.Results)
            {
                if (res.ErrorCode == "0x00000000")
                {
                    if (res.Any != null && res.Any.Count() > 0)
                    {
                        XmlNode[] lisitemRecords = this.GetZrowItems(res.Any);
                        string itemIdValue = Common.GetZrowAttributeValue(lisitemRecords, 0, columnNameId);
                        ids.Add(Convert.ToString(itemIdValue));
                    }
                    else
                    {
                        this.Site.Assert.Fail("The Result element contains more than one z:row.");
                    }
                }
                else
                {
                    this.Site.Assert.Fail(ErrorMessageTemplate, "adding items to the list " + list, this.GetElementValue(res.Any, "ErrorText"));
                }
            }

            if (ids.Count != itemCount)
            {
                this.Site.Assert.Fail("Expect Add {0} item, but actually add {1} items", itemCount, ids.Count);
            }

            return ids;
        }

        /// <summary>
        /// A method used to construct UpdateListItemsUpdates instance using the specified parameters.
        /// </summary>
        /// <param name="methodCollection">A list of MethodCmdEnum to specify the operations.</param>
        /// <param name="fieldNameValuePairs">A list of items values.</param>
        /// <param name="errorhandleType">Specify the OnError of the Batch element's value.</param>
        /// <returns>Returns the UpdateListItemsUpdates instance.</returns>
        protected UpdateListItemsUpdates CreateUpdateListItems(List<MethodCmdEnum> methodCollection, List<Dictionary<string, string>> fieldNameValuePairs, OnErrorEnum errorhandleType)
        {
            this.Site.Assert.IsNotNull(
                        methodCollection,
                        "The [methodCollection] parameter in CreateUpdateListItems cannot be null");
            this.Site.Assert.IsNotNull(
                        fieldNameValuePairs,
                        "The [methodCollection] parameter in CreateUpdateListItems cannot be null");

            this.Site.Assert.AreEqual<int>(
                            methodCollection.Count,
                            fieldNameValuePairs.Count,
                            "The element number in the methods and fieldNameValuePairs parameter List are same");

            UpdateListItemsUpdates updates = new UpdateListItemsUpdates();
            updates.Batch = new UpdateListItemsUpdatesBatch();

            updates.Batch.Method = new UpdateListItemsUpdatesBatchMethod[methodCollection.Count];
            for (int i = 0; i < methodCollection.Count; i++)
            {
                updates.Batch.OnError = errorhandleType;
                updates.Batch.OnErrorSpecified = true;
                updates.Batch.Method[i] = new UpdateListItemsUpdatesBatchMethod();
                updates.Batch.Method[i].Cmd = methodCollection[i];

                int fieldCount = fieldNameValuePairs[i].Keys.Count;
                updates.Batch.Method[i].Field = new UpdateListItemsUpdatesBatchMethodField[fieldCount];
                int j = 0;
                foreach (KeyValuePair<string, string> keyPair in fieldNameValuePairs[i])
                {
                    updates.Batch.Method[i].Field[j] = new UpdateListItemsUpdatesBatchMethodField();
                    updates.Batch.Method[i].Field[j].Name = keyPair.Key;
                    updates.Batch.Method[i].Field[j].Value = keyPair.Value;
                    j++;
                }
            }

            return updates;
        }

        /// <summary>
        /// A method used to get the value of Element from an XmlElement array by specified name.
        /// </summary>
        /// <param name="elements">A parameter represents the source elements where the method finds the value.</param>
        /// <param name="name">A parameter represents the specified element name which is used to find element's value.</param>
        /// <returns>A return value represents the value of the matched element.</returns>
        protected string GetElementValue(XmlNode[] elements, string name)
        {
            return elements.FirstOrDefault<XmlNode>(e => e.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase)).InnerText;
        }

        /// <summary>
        /// A method used to get z:row elements from a elements collection
        /// </summary>
        /// <param name="elementsCollection">A parameter represents a elements collection which contains z:row elements</param>
        /// <returns>A return value represents xml elements collection which only contains z:row elements.</returns>
        protected XmlNode[] GetZrowItems(XmlNode[] elementsCollection)
        {
            XmlNode[] elementsTemp = this.TryGetZrowItems(elementsCollection);

            if (null == elementsTemp || 0 == elementsTemp.Length)
            {
                throw new InvalidOperationException("The zrow items array should contain at least one item.");
            }

            return elementsTemp.ToArray();
        }

        /// <summary>
        /// A method used to get z:row elements from a elements collection, if there are no any zrow items return, this method will return a 0 length array.
        /// </summary>
        /// <param name="elementsCollection">A parameter represents a elements collection which contains z:row elements</param>
        /// <returns>A return value represents xml elements collection which only contains z:row elements.</returns>
        protected XmlNode[] TryGetZrowItems(XmlNode[] elementsCollection)
        {
            if (null == elementsCollection)
            {
                return new XmlNode[] { };
            }

            List<XmlNode> elementsTemp = new List<XmlNode>();
            foreach (XmlNode elementItem in elementsCollection)
            {
                if (elementItem.Name.Equals("z:row", StringComparison.OrdinalIgnoreCase))
                {
                    elementsTemp.Add(elementItem);
                }
            }

            return elementsTemp.ToArray();
        }

        /// <summary>
        /// A method used to get full URL of an attachment 
        /// </summary>
        /// <param name="listId">A parameter represents current list</param>
        /// <param name="itemId">A parameter represents the list item where the method would search attachment</param>
        /// <param name="attachmentFileName">A parameter represents the attachment fie name which is used to key work</param>
        /// <returns>A return value represents the fully URL of attachment match the condition</returns>
        protected string GetAttachmentFullUrl(string listId, string itemId, string attachmentFileName)
        {
            if (string.IsNullOrEmpty(listId) || string.IsNullOrEmpty(itemId) || string.IsNullOrEmpty(attachmentFileName))
            {
                throw new InvalidOperationException("All the input parameters should not be null or empty");
            }

            GetAttachmentCollectionResponseGetAttachmentCollectionResult attachmentResult = null;
            attachmentResult = OutspsAdapter.GetAttachmentCollection(listId, itemId);
            string attachmentUrl = this.GetAttachmentUrlFromResponse(attachmentResult, attachmentFileName);

            // If the expected attachment URL cannot be found by attachment file name, fail the case.
            if (string.IsNullOrEmpty(attachmentUrl))
            {
                this.Site.Assert.Fail("Failed to get the expected attachment URL by specified attachment file name:[{0}]", attachmentFileName);
            }

            return attachmentUrl;
        }

        /// <summary>
        /// A method used to get attachment Url from response of GetAttachmentCollection operation
        /// </summary>
        /// <param name="responseOfGetAttachmentColection">A parameter represents the response of GetAttachmentCollection operation which should contain expected attachment url.</param>
        /// <param name="attachmentName">A parameter represents the name of attachment which is used to find out the url from the response of GetAttachmentCollection operation.</param>
        /// <returns>A return value represents the url of specified attachment.</returns>
        protected string GetAttachmentUrlFromResponse(GetAttachmentCollectionResponseGetAttachmentCollectionResult responseOfGetAttachmentColection, string attachmentName)
        {
            if (null == responseOfGetAttachmentColection || null == responseOfGetAttachmentColection.Attachments || responseOfGetAttachmentColection.Attachments.Length == 0)
            {
                throw new ArgumentException("The response of GetAttachmentColection operation should contain valid attachment data.", "responseOfGetAttachmentColection");
            }

            // Get the attachment URL from GetAttachmentCollection response.
            // The URL value is contained in string array, which is defined in [MS-LISTSWS].
            // There is at least one text item which contain the "attachmentFileName" in the Text array.
            string attachmentUrl = string.Empty;
            var urlItems = from urlItem in responseOfGetAttachmentColection.Attachments
                           where urlItem.IndexOf(attachmentName, StringComparison.OrdinalIgnoreCase) >= 0
                           select urlItem;

            if (urlItems.Count() > 0)
            {
                attachmentUrl = urlItems.ElementAt(0);
            }
            else
            {
                this.Site.Assert.Fail(
                        "The response of GetAttachmentColection operation does not contain the url information for the attachment[{0}]",
                        attachmentName);
            }

            return attachmentUrl;
        }

        /// <summary>
        /// A method used to add a folder into specified document library.
        /// </summary>
        /// <param name="listId">A parameter represents the id of a list where the folder is added.</param>
        /// <param name="folderItemName">A parameter represents the name of folder item which is added to specified list.</param>
        /// <returns>A return value represents the folder name which is added into the specified document library.</returns>
        protected string AddFolderIntoList(string listId, string folderItemName)
        {
            if (string.IsNullOrEmpty(listId))
            {
                throw new ArgumentException("The value should not be null or empty.", "listId");
            }

            if (string.IsNullOrEmpty(folderItemName))
            {
                throw new ArgumentException("The value should not be null or empty.", "folderItemName");
            }

            UpdateListItemsUpdates listItemUpdates = new UpdateListItemsUpdates();
            listItemUpdates.Batch = new UpdateListItemsUpdatesBatch();
            listItemUpdates.Batch.Method = new UpdateListItemsUpdatesBatchMethod[1];
            listItemUpdates.Batch.Method[0] = new UpdateListItemsUpdatesBatchMethod();
            listItemUpdates.Batch.Method[0].ID = (uint)0;
            listItemUpdates.Batch.Method[0].Cmd = MethodCmdEnum.New;
            listItemUpdates.Batch.Method[0].Field = new UpdateListItemsUpdatesBatchMethodField[2];
            listItemUpdates.Batch.Method[0].Field[0] = new UpdateListItemsUpdatesBatchMethodField();
            listItemUpdates.Batch.Method[0].Field[0].Name = "FSObjType";
            listItemUpdates.Batch.Method[0].Field[0].Value = "1";
            listItemUpdates.Batch.Method[0].Field[1] = new UpdateListItemsUpdatesBatchMethodField();
            listItemUpdates.Batch.Method[0].Field[1].Name = "BaseName";
            listItemUpdates.Batch.Method[0].Field[1].Value = folderItemName;
            UpdateListItemsResponseUpdateListItemsResult updateListResult = null;

            try
            {
                updateListResult = OutspsAdapter.UpdateListItems(listId, listItemUpdates);
            }
            catch (SoapException)
            {
                this.Site.Log.Add(LogEntryKind.Debug, "Add folder[{0}] into List[{0}] fail.", folderItemName, listId);
                throw;
            }

            if (null == updateListResult || null == updateListResult.Results || null == updateListResult.Results
                || 1 != updateListResult.Results.Length)
            {
                this.Site.Assert.Fail("The response of UpdateListItems operation for adding a folder[{0}] should include update result.", folderItemName);
            }

            XmlNode[] zrowsdata = this.GetZrowItems(updateListResult.Results[0].Any);

            // In response of UpdateListItems, only contain one folder added result.
            string addedFolderItemIdValue = Common.GetZrowAttributeValue(zrowsdata, 0, "ows_Id");
            int folderItemId;
            if (!int.TryParse(addedFolderItemIdValue, out folderItemId))
            {
                this.Site.Assert.Fail("The list item id should be integer format value, current value is [{0}]", addedFolderItemIdValue);
            }

            return addedFolderItemIdValue;
        }

        /// <summary>
        /// Delete the specified items in the specified list.
        /// </summary>
        /// <param name="list">A parameter represents the list title or list id.</param>
        /// <param name="deleteIDList">Specify the delete items id.</param>
        /// <param name="errorEnum">Specify how to affect the following sequence when one delete item method fails.</param>
        protected void DeleteListItems(string list, List<string> deleteIDList, OnErrorEnum errorEnum)
        {
            List<Dictionary<string, string>> deletedItems = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            foreach (string id in deleteIDList)
            {
                Dictionary<string, string> itemId = new Dictionary<string, string>();
                itemId.Add("ID", id);
                deletedItems.Add(itemId);
                cmds.Add(MethodCmdEnum.Delete);
            }

            UpdateListItemsUpdates updates = this.CreateUpdateListItems(
                                                                cmds,
                                                                deletedItems,
                                                                errorEnum);

            UpdateListItemsResponseUpdateListItemsResult result = OutspsAdapter.UpdateListItems(
                                                                                    list,
                                                                                    updates);

            if (deleteIDList.Count != result.Results.Length)
            {
                this.Site.Assert.Fail("Expect delete {0} item, but actually add {1} items", deleteIDList.Count, result.Results.Length);
            }

            foreach (UpdateListItemsResponseUpdateListItemsResultResult res in result.Results)
            {
                if (res.ErrorCode != "0x00000000")
                {
                    this.Site.Debug.Fail(ErrorMessageTemplate, "deleting item from the list " + list, this.GetElementValue(res.Any, "ErrorText"));
                }
            }
        }

        /// <summary>
        /// A method used to get a recurrence XML string, which is used to setting the recurrence appointment item on event list. 
        /// </summary>
        /// <param name="recurrenceXMLData">A parameter represents the RecurrenceXML type instance which include the setting the settings of recurrence appointment.</param>
        /// <returns>A return value represents the serialized xml string from specified RecurrenceXML type instance.</returns>
        protected string GetRecurrenceXMLString(RecurrenceXML recurrenceXMLData)
        {
            string serializedXMLstring = string.Empty;
            StringBuilder strBuilder = new StringBuilder();
            using (StringWriter stringWriter = new StringWriter(strBuilder))
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(RecurrenceXML));
                xmlSerializer.Serialize(stringWriter, recurrenceXMLData);
                serializedXMLstring = strBuilder.ToString();
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(serializedXMLstring);
            var recurrenceXMLElementitems = from XmlNode childNodeItem in xmlDoc.ChildNodes
                                            where childNodeItem.LocalName.Equals("RecurrenceXML", StringComparison.OrdinalIgnoreCase)
                                            select childNodeItem;

            if (0 == recurrenceXMLElementitems.Count())
            {
                this.Site.Assert.Fail("The de-serialize XML string should contain expected RecurrenceXML element. Current XML string:\r\n[{0}]", serializedXMLstring);
            }

            XmlElement recurrenceXMLElement = (XmlElement)recurrenceXMLElementitems.ElementAt<XmlNode>(0);
            serializedXMLstring = recurrenceXMLElement.InnerXml;
            serializedXMLstring = serializedXMLstring.Replace(@" xmlns=""http://schemas.microsoft.com/sharepoint/soap/""", string.Empty);
            return serializedXMLstring;
        }

        /// <summary>
        /// A method used to get a TimeZoneXML XML string, which is used to setting the recurrence appointment item on event list. 
        /// </summary>
        /// <param name="timeZoneXMLData">A parameter represents the TimeZoneXML type instance which include the setting the settings of recurrence appointment.</param>
        /// <returns>A return value represents the serialized xml string from specified RecurrenceXML type instance.</returns>
        protected string GetTimeZoneXMLString(TimeZoneXML timeZoneXMLData)
        {
            string serializedXMLstring = string.Empty;
            StringBuilder strBuilder = new StringBuilder();
            using (StringWriter stringWriter = new StringWriter(strBuilder))
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(TimeZoneXML));
                xmlSerializer.Serialize(stringWriter, timeZoneXMLData);
                serializedXMLstring = strBuilder.ToString();
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(serializedXMLstring);
            var recurrenceXMLElementitems = from XmlNode childNodeItem in xmlDoc.ChildNodes
                                            where childNodeItem.LocalName.Equals("TimeZoneXML", StringComparison.OrdinalIgnoreCase)
                                            select childNodeItem;

            if (0 == recurrenceXMLElementitems.Count())
            {
                this.Site.Assert.Fail("The de-serialize XML string should contain expected TimeZoneXML element. Current XML string:\r\n[{0}]", serializedXMLstring);
            }

            XmlElement recurrenceXMLElement = (XmlElement)recurrenceXMLElementitems.ElementAt<XmlNode>(0);
            serializedXMLstring = recurrenceXMLElement.InnerXml;
            serializedXMLstring = serializedXMLstring.Replace(@" xmlns=""http://schemas.microsoft.com/sharepoint/soap/""", string.Empty);
            return serializedXMLstring;
        }

        /// <summary>
        /// A method used to get fields setting of recurrence event.
        /// </summary>
        /// <param name="eventTitle">A parameter represents the title of event.</param>
        /// <param name="eventDate">A parameter represents the event date of event.</param>
        /// <param name="endDate">A parameter represents the end date of event.</param>
        /// <param name="recurrenceXmlString">A parameter represents the xml string of recurrence setting.</param>
        /// <returns>A return value represents the fields setting of recurrence event.</returns>
        protected Dictionary<string, string> GetGeneralRecurrenceEventFieldsSetting(string eventTitle, DateTime eventDate, DateTime endDate, string recurrenceXmlString)
        {
            if (string.IsNullOrEmpty(eventTitle))
            {
                throw new ArgumentException("Value should not be null or empty", "eventTitle");
            }

            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If EventType is "1", this property MUST contain a valid RecurrenceXML.
            recurEventFieldsSetting.Add("RecurrenceData", recurrenceXmlString);

            // If EventType is "1", then this property MUST contain a valid TimeZoneXML.
            TimeZoneXML timeZoneXml = this.GetCustomPacificTimeZoneXmlSetting();
            string timeZoneXmlString = this.GetTimeZoneXMLString(timeZoneXml);
            recurEventFieldsSetting.Add("XMLTZone", timeZoneXmlString);

            // Setting title field's value
            recurEventFieldsSetting.Add("Title", eventTitle);

            // If fRecurrence is TRUE, this property MUST contain a valid stringGUID.
            string uidValue = Guid.NewGuid().ToString();
            recurEventFieldsSetting.Add("UID", uidValue);

            return recurEventFieldsSetting;
        }

        /// <summary>
        /// Generate a daily recurrence setting with specified dayFrequency, specified title and duration.
        /// </summary>
        /// <param name="eventTitle">A parameter represents the title of event.</param>
        /// <param name="eventDate">A parameter represents the event date of event. The end date of the event is computed from this parameter by this formula: eventDate + 1 hour.</param>
        /// <param name="dayFrequencyValue">A parameter represents the value of daily recurrent frequency.</param>
        /// <param name="durationOfRecurrence">A parameter represents the duration of the recurrence event </param>
        /// <returns>A return value represents the daily recurrence setting with computed windowEnd value. windowEnd value is computed from durationOfRecurrence and eventDate parameters.</returns>
        protected Dictionary<string, string> GetDailyRecurrenceSettingWithwindowEnd(string eventTitle, DateTime eventDate, string dayFrequencyValue, double durationOfRecurrence)
        {
            if (string.IsNullOrEmpty(eventTitle))
            {
                throw new ArgumentException("Value should not be null or empty", "eventTitle");
            }

            // Recurring data
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Set the windowEnd value, the recurrence will be end on durationOfRecurrence + 1 days from current date, and the recurrence duration should be equal to durationOfRecurrence.
            DateTime endDate = eventDate.AddHours(1);
            DateTime winEndDate = eventDate.Date.AddDays(durationOfRecurrence);
            recurrenceXMLData.recurrence.rule.Item = winEndDate;

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();
            RepeatPatternDaily repeatPatternDailyData = new RepeatPatternDaily();

            // Setting the dayFrequencyValue
            repeatPatternDailyData.dayFrequency = dayFrequencyValue;
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDailyData;

            string recurrenceXMLString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Generate a daily recurrence setting with specified dayFrequency, specified title and durationOfRecurrence.
            Dictionary<string, string> dailyRecurrenceSetting = this.GetGeneralRecurrenceEventFieldsSetting(eventTitle, eventDate, endDate, recurrenceXMLString);
            return dailyRecurrenceSetting;
        }

        /// <summary>
        /// A method used to get fields' setting of exception item of an existing recurrence item.
        /// </summary>
        /// <param name="exceptionItemEventDate">A parameter represents the event date of the exception item.</param>
        /// <param name="recurrenceID">A parameter represents the RecurrenceID field value of this exception item. The RecurrenceID must be equal to the starting date and time of one instance of a recurrence.</param>
        /// <param name="exceptionTitle">A parameter represents the title of the exception item.</param>
        /// <param name="masterSeriesItemID">A parameter represents the MasterSeriesItemID field value of this exception item. The RecurrenceID must be equal to the list item id of the recurring item that the exception belongs to.</param>
        /// <param name="recurrenceEventSetting">A parameter represents the fields' setting of an existing recurrence item which the exception item belong to.</param>
        /// <returns>A return value represents the fields' setting of exception item.</returns>
        protected Dictionary<string, string> GetExceptionsItemSettingForRecurrenceEvent(DateTime exceptionItemEventDate, DateTime recurrenceID, string exceptionTitle, string masterSeriesItemID, Dictionary<string, string> recurrenceEventSetting)
        {
            if (string.IsNullOrEmpty(masterSeriesItemID))
            {
                throw new ArgumentException("Must specified the 'MasterSeriesItemID' field for an exception appointment item.", "masterSeriesItemID");
            }

            if (string.IsNullOrEmpty(exceptionTitle))
            {
                throw new ArgumentException("Must specified the 'Title' field for an exception appointment item.", "exceptionTitle");
            }

            Dictionary<string, string> exceptionEventFieldsSetting = new Dictionary<string, string>();

            string exceptionEndDateValue = this.GetGeneralFormatTimeString(exceptionItemEventDate.AddHours(1));
            string exceptionEventDateValue = this.GetGeneralFormatTimeString(exceptionItemEventDate);
            string recurrenceIDValue = this.GetGeneralFormatTimeString(recurrenceID);

            // The ending date and time of the appointment.
            exceptionEventFieldsSetting.Add("EndDate", exceptionEndDateValue);

            // The starting date and time of the appointment.
            exceptionEventFieldsSetting.Add("EventDate", exceptionEventDateValue);

            // "4" means this item is an exception to a recurrence item.
            exceptionEventFieldsSetting.Add("EventType", "4");

            // "1" means this is a recurrence item, and exception item is belong to recurrence item instances.
            exceptionEventFieldsSetting.Add("fRecurrence", "1");

            // RecurrenceID MUST be equal to the starting date and time of one instance of a recurrence when the EventType indicates an exception or deleted instance.
            exceptionEventFieldsSetting.Add("RecurrenceID", recurrenceIDValue);

            // "0" means this is not an all-day event.
            exceptionEventFieldsSetting.Add("Title", exceptionTitle);

            // Setting MasterSeriesItemID field's value
            exceptionEventFieldsSetting.Add("MasterSeriesItemID", masterSeriesItemID);

            // The UID should equal to the recurrence item's UID whose instance will be overwrote by exception item.
            if (recurrenceEventSetting.ContainsKey("UID"))
            {
                exceptionEventFieldsSetting.Add("UID", recurrenceEventSetting["UID"]);
            }
            else
            {
                this.Site.Assert.Fail("The recurrence event Setting should contain [UID] value.");
            }

            // The RecurrenceData should equal to the recurrence item's RecurrenceData whose instance will be overwrote by exception item.
            if (recurrenceEventSetting.ContainsKey("RecurrenceData"))
            {
                exceptionEventFieldsSetting.Add("RecurrenceData", recurrenceEventSetting["RecurrenceData"]);
            }
            else
            {
                this.Site.Assert.Fail("The recurrence event Setting should contain [RecurrenceData] value.");
            }

            return exceptionEventFieldsSetting;
        }

        /// <summary>
        /// A method used to get list item id collection from the operation of UpdateListItems operation. If the [expectedUpdatedListItem] parameter specified value and the response does not contain expected list items' id record, method will throw Assert exception.
        /// </summary>
        /// <param name="updateListItemResponse">A parameter represents the response of UpdateListItems operation.</param>
        /// <param name="expectedUpdatedListItem">A parameter represents the expected number of updated list items. The method will perform a check: if the actual number of updated list items does not equal to this value, this method will throw a Assert Exception. Specified this value to null means the method does not perform the check process.</param>
        /// <returns>A return value presents all updated list items' id</returns>
        protected List<string> GetListItemIdsFromUpdateListItemsResponse(UpdateListItemsResponseUpdateListItemsResult updateListItemResponse, int? expectedUpdatedListItem)
        {
            List<string> listItemIdsTemp = new List<string>();
            if (null == updateListItemResponse)
            {
                throw new ArgumentNullException("updateListItemResponse");
            }

            if (null == updateListItemResponse.Results || 0 == updateListItemResponse.Results.Length)
            {
                this.Site.Assert.Fail("The response of UpdateListItem operation should contain at least one record of updated list item.");
            }

            foreach (UpdateListItemsResponseUpdateListItemsResultResult updateResultItem in updateListItemResponse.Results)
            {
                XmlNode[] zrowsItem = this.GetZrowItems(updateResultItem.Any);

                for (int updatedItemIndex = 0; updatedItemIndex < zrowsItem.Length; updatedItemIndex++)
                {
                    string idvalue = Common.GetZrowAttributeValue(zrowsItem, updatedItemIndex, "ows_ID");
                    listItemIdsTemp.Add(idvalue);
                }
            }

            if (expectedUpdatedListItem.HasValue)
            {
                // Each result item means a update process for a list item.
                this.Site.Assert.AreEqual(
                    expectedUpdatedListItem.Value,
                    listItemIdsTemp.Count,
                    "The actual updated list items' number should match the expected number:[{0}].",
                    expectedUpdatedListItem.Value);
            }

            return listItemIdsTemp;
        }

        /// <summary>
        /// A method used to update the fields' setting by specified value.
        /// </summary>
        /// <param name="originalRecurrenceSetting">A parameter represents the fields' setting expected to update.</param>
        /// <param name="fieldName">A parameter represents the target field's name which the method will update.</param>
        /// <param name="fieldValue">A parameter represents the updated value.</param>
        /// <returns>A return value represents the updated fields' setting.</returns>
        protected Dictionary<string, string> GetUpdatedRecurrenceSetting(Dictionary<string, string> originalRecurrenceSetting, string fieldName, string fieldValue)
        {
            if (null == originalRecurrenceSetting || 0 == originalRecurrenceSetting.Count)
            {
                throw new ArgumentException("The original recurrence setting should contain valid setting records.", "originalRecurrenceSetting");
            }

            if (string.IsNullOrEmpty(fieldName))
            {
                throw new ArgumentNullException("fieldName");
            }

            if (string.IsNullOrEmpty(fieldValue))
            {
                throw new ArgumentNullException("fieldValue");
            }

            if (originalRecurrenceSetting.ContainsKey(fieldName))
            {
                originalRecurrenceSetting[fieldName] = fieldValue;
            }
            else
            {
                this.Site.Assert.Fail("The recurrence event Setting should contain [{0}] setting.", fieldName);
            }

            return originalRecurrenceSetting;
        }

        /// <summary>
        /// A method used to get fields' setting which trigger the exception items deletion.
        /// </summary>
        /// <param name="originalRecurrenceSetting">A parameter represents the fields' setting which is existing recurrence appointment item on SUT.</param>
        /// <param name="updatedEventDate">A parameter represents the updated value of EventDate field.</param>
        /// <returns>A return value represents the fields' setting which is used to update an existing recurrence appointment item, and then trigger the exception items deletion.</returns>
        protected Dictionary<string, string> GetExceptionDeletionSettingWithEventUpdate(Dictionary<string, string> originalRecurrenceSetting, DateTime updatedEventDate)
        {
            string updatedEventDateValue = this.GetUTCFormatTimeString(updatedEventDate);

            // Update the event field value.
            Dictionary<string, string> updatedSettings = this.GetUpdatedRecurrenceSetting(originalRecurrenceSetting, "EventDate", updatedEventDateValue);

            // Append the "deleteExceptions" Element.
            updatedSettings = this.AppenddeleteExceptionsElement(updatedSettings);
            return updatedSettings;
        }

        /// <summary>
        /// A method used to append "deleteExceptions" element on a value of "RecurrenceData" field. This element must be present if and only if protocol client update "EndDate", "EventDate", "RecurrenceData", "UID", "XMLTZone" fields for an recurrence appointment item and expect protocol SUT trigger the exception deletion.
        /// </summary>
        /// <param name="originalRecurrenceSetting">A parameter represents the fields' setting which expected to append the "deleteExceptions" element on "RecurrenceData" field.</param>
        /// <returns>A return value represents the fields' setting which append the "deleteExceptions" element on "RecurrenceData" field.</returns>
        protected Dictionary<string, string> AppenddeleteExceptionsElement(Dictionary<string, string> originalRecurrenceSetting)
        {
            if (null == originalRecurrenceSetting)
            {
                throw new ArgumentNullException("originalRecurrenceSetting");
            }

            string recurrenceDataFieldName = "RecurrenceData";
            if (originalRecurrenceSetting.ContainsKey(recurrenceDataFieldName))
            {
                string recurrenceDataValue = originalRecurrenceSetting[recurrenceDataFieldName];

                // If the string have not contained <deleteExceptions>..</deleteExceptions> element, append the <deleteExceptions> element xml string.
                if (recurrenceDataValue.IndexOf(@"<deleteExceptions>", StringComparison.OrdinalIgnoreCase) < 0
                    || !recurrenceDataValue.EndsWith(@"</deleteExceptions>", StringComparison.OrdinalIgnoreCase))
                {
                    recurrenceDataValue = string.Format("{0}{1}", recurrenceDataValue, @"<deleteExceptions>true</deleteExceptions>");
                }

                // update the RecurrenceData field with updated value.
                originalRecurrenceSetting = this.GetUpdatedRecurrenceSetting(originalRecurrenceSetting, recurrenceDataFieldName, recurrenceDataValue);
            }
            else
            {
                this.Site.Assert.Fail("The recurrence event Setting should contain [{0}] setting.", "RecurrenceData");
            }

            return originalRecurrenceSetting;
        }

        /// <summary>
        /// A method used to get a UTC format time string from specified DateTime type.
        /// </summary>
        /// <param name="sourceDateTime">A parameter represents the DateTime type instance which will be converted to a UTC format string.</param>
        /// <returns>A return value represents the time string.</returns>
        protected string GetUTCFormatTimeString(DateTime sourceDateTime)
        {
            // ISO8601 UTC time format
            string timeFormatPattern = @"yyyy-MM-ddTHH:mm:ssZ";
            string timeString = sourceDateTime.ToUniversalTime().ToString(timeFormatPattern);
            return timeString;
        }

        /// <summary>
        /// A method used to get a general format time string with zero seconds from specified DateTime type. The format is look like this: "2012-12-21 23:59:00"
        /// </summary>
        /// <param name="sourceDateTime">A parameter represents the DateTime type instance which will be converted to a "2012-12-21 HH:mm:00" format string.</param>
        /// <returns>A return value represents the time string.</returns>
        protected string GetGeneralFormatTimeString(DateTime sourceDateTime)
        {
            // General time format
            string timeFormatPattern = @"yyyy-MM-dd HH:mm:00";
            string timeString = sourceDateTime.ToString(timeFormatPattern);
            return timeString;
        }

        /// <summary>
        /// A method used to verify whether the zrow collection contain an expected list item whose id equal to specified value.
        /// </summary>
        /// <param name="listItemIdValue">A parameter represents the expected list item id value.</param>
        /// <param name="zrowCollection">A parameter represents the zrow collection which contain list items.</param>
        /// <returns>Return 'true' indicating the zrow collection contain the expected list item.</returns>
        protected bool VerifyContainExpectedListItemById(string listItemIdValue, XmlNode[] zrowCollection)
        {
            if (null == zrowCollection || 0 == zrowCollection.Length)
            {
                throw new ArgumentException("The zrow collection should contain at least one updated list item record.", "zrowCollection");
            }

            if (string.IsNullOrEmpty(listItemIdValue))
            {
                throw new ArgumentNullException("listItemIdValue");
            }

            int listItemId;
            if (!int.TryParse(listItemIdValue, out listItemId))
            {
                throw new ArgumentException("The list item id should be string value which can convert to int. actual:[{0}]", "listItemIdValue");
            }

            bool isContainExceptedListItem = false;
            for (int zrowIndex = 0; zrowIndex < zrowCollection.Length; zrowIndex++)
            {
                string actualIdValue = Common.GetZrowAttributeValue(zrowCollection, zrowIndex, "ows_ID");
                if (listItemIdValue.Equals(actualIdValue, StringComparison.OrdinalIgnoreCase))
                {
                    isContainExceptedListItem = true;
                    break;
                }
            }

            return isContainExceptedListItem;
        }

        /// <summary>
        /// A method used to verify simple types defined in MS-OUTSPS for each zrow items. The "Priority" field's schema definition is only valid for zrow item which indicate the Appointment list item's data. If there are any schema validation errors, method will throw a XmlSchemaValidationException exception.
        /// </summary>
        /// <param name="zrowItems">A parameter represents the zrow items' collection which are used to perform schema check.</param>
        /// <returns>Return true indicating the validation is succeed.</returns>
        protected bool VerifySimpleTypeSchema(XmlNode[] zrowItems)
        {
            if (null == zrowItems || 0 == zrowItems.Count())
            {
                throw new ArgumentException("The zrow items should contain at least one zrow record of a list item.", "zrowItems");
            }

            foreach (XmlNode zrowItem in zrowItems)
            {
                string zrowXmlString = zrowItem.OuterXml;

                // Verify simple types' schema validation.
                SchemaValidation.ValidateXml(this.Site, zrowXmlString);
            }

            if (SchemaValidation.ValidationResult != ValidationResult.Success)
            {
                string validationErrorMessage = SchemaValidation.GenerateValidationResult();
                throw new XmlSchemaValidationException(validationErrorMessage);
            }

            return true;
        }

        /// <summary>
        /// A method used to verify complex types in MS-OUTSPS, only support RecurrenceXML, TimeZoneXML, AttachProps 3 types.If there are any validation errors, method will throw method will throw a XmlSchemaValidationException exception.
        /// </summary>
        /// <param name="complexTypeFieldValue">A parameter represents the value of complex type.</param>
        /// <param name="typeOfComplexType">A parameter represents the type of the value specified by complexTypeFieldValue parameter.</param>
        /// <returns>Return true indicating the validation is succeed.</returns>
        protected bool VerifyComplexTypesSchema(string complexTypeFieldValue, Type typeOfComplexType)
        {
            if (string.IsNullOrEmpty(complexTypeFieldValue))
            {
                throw new ArgumentNullException("complexTypeFieldValue");
            }

            if (null == typeOfComplexType)
            {
                throw new ArgumentNullException("typeOfComplexType");
            }

            // Construct a xml string of MSOUTSPSComplexTypes elements, and perform the schema validation by using MS-OUTSPS.wsdl 
            string subElementString = string.Empty;
            string rootElementString = @"<MSOUTSPSComplexTypes xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">{0}</MSOUTSPSComplexTypes>";

            if (typeOfComplexType.Equals(typeof(RecurrenceXML)))
            {
                subElementString = string.Format(@"<RecurrenceDataTypeField>{0}</RecurrenceDataTypeField>", complexTypeFieldValue);
            }
            else if (typeOfComplexType.Equals(typeof(TimeZoneXML)))
            {
                subElementString = string.Format(@"<TimeZoneXMLTypeField>{0}</TimeZoneXMLTypeField>", complexTypeFieldValue);
            }
            else if (typeOfComplexType.Equals(typeof(AttachProps)))
            {
                subElementString = string.Format(@"<AttachPropsTypeField>{0}</AttachPropsTypeField>", complexTypeFieldValue);
            }
            else
            {
                this.Site.Assert.Fail("The complex type should be one of types: [RecurrenceXML], [TimeZoneXML], [AttachProps]");
            }

            // All complex types are under "http://schemas.microsoft.com/sharepoint/soap/" namespace.
            string fullyXmlString = string.Format(rootElementString, subElementString);

            SchemaValidation.ValidateXml(this.Site, fullyXmlString);
            if (SchemaValidation.ValidationResult != ValidationResult.Success)
            {
                string validationErrorMessage = SchemaValidation.GenerateValidationResult();
                throw new XmlSchemaValidationException(validationErrorMessage);
            }

            return true;
        }

        /// <summary>
        /// A method used to get the index of zrow items array by specified list item id. If there are no any match zrow items found in zrow items array, this method will throw Assert exception.
        /// </summary>
        /// <param name="zrowItems">A parameter represents the zrow items array.</param>
        /// <param name="listItemId">A parameter represents the list item id, which indicates a list item. each zrow item means a list item in protocol SUT.</param>
        /// <returns>A return value represents the index of the specified zrow items array.</returns>
        protected int GetZrowItemIndexByListItemId(XmlNode[] zrowItems, string listItemId)
        {
            int matchedIndex = this.TryGetZrowItemIndexByListItemId(zrowItems, listItemId);
            if (-1 == matchedIndex)
            {
                this.Site.Assert.Fail(
                                "The zrow item array does not contain a zrow item which match specified list item id. Expected list item id[{0}]",
                                listItemId);
            }

            return matchedIndex;
        }

        /// <summary>
        /// A method used to get the index of zrow items array by specified list item id. If there are no any match zrow items found in zrow items array, this method will return -1.
        /// </summary>
        /// <param name="zrowItems">A parameter represents the zrow items array.</param>
        /// <param name="listItemId">A parameter represents the list item id, which indicates a list item. each zrow item means a list item in protocol SUT.</param>
        /// <returns>A return value represents the index of the specified zrow items array.</returns>
        protected int TryGetZrowItemIndexByListItemId(XmlNode[] zrowItems, string listItemId)
        {
            #region Verification of parameter
            if (null == zrowItems)
            {
                throw new ArgumentNullException("zrowItems");
            }

            if (string.IsNullOrEmpty(listItemId))
            {
                throw new ArgumentNullException("listItemId");
            }

            int listItemIdValue;
            if (!int.TryParse(listItemId, out listItemIdValue))
            {
                throw new ArgumentException("The value should be a integer format string.", "listItemId");
            }

            if (listItemIdValue <= 0)
            {
                throw new ArgumentException("The value should be larger than zero.", "listItemId");
            }

            if (0 == zrowItems.Length)
            {
                throw new ArgumentException("Should contain at least one zrow item record.", "zrowItems");
            }
            #endregion Verification of parameter

            int indexOfZrowItem = 0;
            int matchedIndex = -1;
            for (; indexOfZrowItem < zrowItems.Length; indexOfZrowItem++)
            {
                string listItemIdTemp = Common.GetZrowAttributeValue(zrowItems, indexOfZrowItem, "ows_ID");
                if (listItemId.Equals(listItemIdTemp, StringComparison.OrdinalIgnoreCase))
                {
                    matchedIndex = indexOfZrowItem;
                    break;
                }
            }

            return matchedIndex;
        }

        /// <summary>
        /// A method used to try to get the index of zrow items array by specified end with value for fileRef field. If there are no any match zrow items found in zrow items array, this method will throw Assert exception.
        /// </summary>
        /// <param name="zrowItems">A parameter represents the zrow items array.</param>
        /// <param name="fileRefEndWithValue">A parameter represent the value which the fileRef field should end with.For a document library item, the fileRef field value should end with the list item's actual name.</param>
        /// <returns>A return value represents the index of the specified zrow items array.</returns>
        protected int GetZrowItemIndexByFileRef(XmlNode[] zrowItems, string fileRefEndWithValue)
        {
            #region Verification of parameter
            if (null == zrowItems)
            {
                throw new ArgumentNullException("zrowItems");
            }

            if (string.IsNullOrEmpty(fileRefEndWithValue))
            {
                throw new ArgumentNullException("fileRefEndWithValue");
            }

            if (0 == zrowItems.Length)
            {
                throw new ArgumentException("Should contain at least one zrow item record.", "zrowItems");
            }
            #endregion Verification of parameter

            int indexOfZrowItem = 0;
            int matchedIndex = -1;
            for (; indexOfZrowItem < zrowItems.Length; indexOfZrowItem++)
            {
                // The fileRef field value means the fielServer-relative URL of the full path of the item's related file. For a document library item, the fileRef field value should end with this list item's actual name. 
                string fileDirRefTemp = Common.GetZrowAttributeValue(zrowItems, indexOfZrowItem, "ows_fileRef");
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(fileDirRefTemp),
                                    @"The fileDirRef field should have value.");

                string expectedEndwithValue = string.Format(@"/{0}", fileRefEndWithValue);
                if (fileDirRefTemp.EndsWith(expectedEndwithValue, StringComparison.OrdinalIgnoreCase))
                {
                    matchedIndex = indexOfZrowItem;
                    break;
                }
            }

            if (-1 == matchedIndex)
            {
                this.Site.Assert.Fail(
                                @"The zrow item array does not contain a zrow item whose ""fileDirRef"" field should end with specified value[{0}].",
                                fileRefEndWithValue);
            }

            return matchedIndex;
        }

        /// <summary>
        /// A method used to get the status code from WebException instance.
        /// </summary>
        /// <param name="webException">A parameter represents the WebException instance which contains the status code.</param>
        /// <returns>A return value represents the HttpStatusCode enum type value which is contained in WebException instance.</returns>
        protected HttpStatusCode GetStatusCodeFromWebException(WebException webException)
        {
            HttpWebResponse lowLevelResponse = webException.Response as HttpWebResponse;
            if (null == lowLevelResponse)
            {
                this.Site.Assert.Fail("The current webException instance does not contain expected HttpWebResponse data.");
            }

            return lowLevelResponse.StatusCode;
        }

        /// <summary>
        /// A method used to verify the response of GetListItemChangesSinceToken operation whether contain zrow data structure under "listitems" element. If the response does not contain the structure, this method will throw Assert exception.
        /// </summary>
        /// <param name="listitemChangesRes">A parameter represents the response of GetListItemChangesSinceToken operation which expected to contain zrow data structure under "listitems" element.</param>
        protected void VerifyContainZrowDataStructure(GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listitemChangesRes)
        {
            if (null == listitemChangesRes)
            {
                throw new ArgumentNullException("listitemChangesRes");
            }

            this.Site.Assert.IsNotNull(
                         listitemChangesRes.listitems.data,
                         "The response of GetListItemChangesSinceToken operation should contain [zrow] data structure under [listitems] element.");
        }

        /// <summary>
        /// A method used to verify the response of GetListItemChanges operation whether contain at least one rs:data structure item under "listitems" element. If the response does not contain any rs:data item under "listitems" element, this method will throw Assert exception.
        /// </summary>
        /// <param name="listitemChangesRes">A parameter represents the response of GetListItemChangesSinceToken operation which expected to contain zrow data structure under "listitems" element.</param>
        protected void VerifyContainZrowDataStructure(GetListItemChangesResponseGetListItemChangesResult listitemChangesRes)
        {
            if (null == listitemChangesRes)
            {
                throw new ArgumentNullException("listitemChangesRes");
            }

            this.Site.Assert.IsNotNull(
                           listitemChangesRes.listitems.data,
                           "The response of GetListItemChanges operation should contain [zrow] data structure under [listitems] element.");

            this.Site.Assert.IsTrue(
                            listitemChangesRes.listitems.data.Length > 0,
                            "The response of GetListItemChanges operation should contain at least one rs:data element.");

            // There can be a maximum of two rs:data elements. The first rs:data element contains all the inserted and updated list items that have occurred subsequent to the specified since parameter. The second rs:data element contains all of the list items currently in the list. 
            this.Site.Assert.IsTrue(
                            listitemChangesRes.listitems.data.Length <= 2,
                            "The response of GetListItemChanges operation should contain maximum of two rs:data elements.");
        }

        /// <summary>
        /// A method used to verify the response of UpdateListItems operation whether contain "Result" item under "Results" element. If the response does not contain any  "Result" items, this method will throw Assert exception.
        /// </summary>
        /// <param name="responseOfUpdateListItems">A parameter represents the response of UpdateListItems operation which expected to contain "Result" item under "Results" element.</param>
        /// <param name="expectedItemNumber">A parameter represents the how many items the response of UpdateListItems operation expected to contain. If does not specified the expectedItemNumber value, method will not perform the expected items' number verification.</param>
        protected void VerifyContainResultItem(UpdateListItemsResponseUpdateListItemsResult responseOfUpdateListItems, int? expectedItemNumber)
        {
            if (null == responseOfUpdateListItems)
            {
                throw new ArgumentNullException("responseOfUpdateListItems");
            }

            int expectedItemNumberValue = -1;
            if (expectedItemNumber.HasValue)
            {
                expectedItemNumberValue = expectedItemNumber.Value;
                if (expectedItemNumberValue < 0)
                {
                    string errorMessage = string.Format(@"The value should be larger than or equal to zero. actual value:[{0}]", expectedItemNumberValue);
                    throw new ArgumentException(errorMessage, "expectedItemNumber");
                }
            }

            this.Site.Assert.IsNotNull(
                    responseOfUpdateListItems.Results,
                    "The response of UpdateListItems operation should contain [Results] element.");

            this.Site.Assert.AreNotEqual<int>(
                        0,
                        responseOfUpdateListItems.Results.Length,
                        "The response of UpdateListItems operation should contain at least one [Result] item.");

            // If does not specified the expectedItemNumber value, method will not perform the verification.
            if (!expectedItemNumber.HasValue)
            {
                return;
            }

            this.Site.Assert.AreEqual<int>(
                                    expectedItemNumberValue,
                                    responseOfUpdateListItems.Results.Length,
                                    "The response of UpdateListItems operation should contain expected number of [Result] items.");
        }

        /// <summary>
        /// A method is used to verify the result for UpdateListItem operation. If not all the update operation succeed, this method will throw Assert exception.
        /// </summary>
        /// <param name="updateResult">The response of the UpdateListItem operation.</param>
        protected void VerifyResponseOfUpdateListItem(UpdateListItemsResponseUpdateListItemsResult updateResult)
        {
            if (null == updateResult)
            {
                throw new ArgumentNullException("updateResult");
            }

            this.Site.Assert.IsNotNull(updateResult.Results, "The response of UpdateListItems operation should contain the Results element.");

            foreach (UpdateListItemsResponseUpdateListItemsResultResult res in updateResult.Results)
            {
                // If the UpdateListItem success, the ErrorCode is "0x00000000".
                if ("0x00000000".Equals(res.ErrorCode, StringComparison.OrdinalIgnoreCase))
                {
                    if (res.Any != null && res.Any.Count() > 0)
                    {
                        this.Site.Log.Add(LogEntryKind.Debug, "UpdateListItem operation successfully.");
                    }
                    else
                    {
                        this.Site.Assert.Fail("The Result element should contains at least one zrow item.");
                    }
                }
                else
                {
                    if (res.Any != null && res.Any.Count() > 0)
                    {
                        foreach (XmlElement elementItem in res.Any)
                        {
                            if (elementItem.Name.Equals("ErrorText", StringComparison.OrdinalIgnoreCase))
                            {
                                // Export the error message for the result of UpdateListItem.
                                this.Site.Log.Add(LogEntryKind.Warning, ErrorMessageTemplate, "UpdateListItem operation fail:" + elementItem.InnerText);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// A method used to get TimeZoneXML setting instance of Pacific time(UTC-8:00). The daylight saving setting is that: UTC-7:00 start at "3th month/second week/Sunday/2:00 AM" and end at "11th month/first week/Sunday/2:00 AM"
        /// </summary>
        /// <returns>A return value represents the TimeZoneXML instance of Pacific time</returns>
        protected TimeZoneXML GetCustomPacificTimeZoneXmlSetting()
        {
            TimeZoneXML timeZoneXml = new TimeZoneXML();
            timeZoneXml.timeZoneRule = new TimeZoneRule();

            // standardDate element
            timeZoneXml.timeZoneRule.standardDate = new TransitionDate();
            timeZoneXml.timeZoneRule.standardDate.transitionRule = new TransitionDateTransitionRule();
            timeZoneXml.timeZoneRule.standardDate.transitionRule.month = "11";
            timeZoneXml.timeZoneRule.standardDate.transitionRule.day = DayOfWeek.su;
            timeZoneXml.timeZoneRule.standardDate.transitionRule.daySpecified = true;
            timeZoneXml.timeZoneRule.standardDate.transitionRule.weekdayOfMonth = WeekdayOfMonth.first;
            timeZoneXml.timeZoneRule.standardDate.transitionRule.weekdayOfMonthSpecified = true;
            timeZoneXml.timeZoneRule.standardDate.transitionTime = "2:0:0";

            // daylightDate element
            timeZoneXml.timeZoneRule.daylightDate = new TransitionDate();
            timeZoneXml.timeZoneRule.daylightDate.transitionRule = new TransitionDateTransitionRule();
            timeZoneXml.timeZoneRule.daylightDate.transitionRule.month = "3";
            timeZoneXml.timeZoneRule.daylightDate.transitionRule.day = DayOfWeek.su;
            timeZoneXml.timeZoneRule.daylightDate.transitionRule.daySpecified = true;
            timeZoneXml.timeZoneRule.daylightDate.transitionRule.weekdayOfMonth = WeekdayOfMonth.second;
            timeZoneXml.timeZoneRule.daylightDate.transitionRule.weekdayOfMonthSpecified = true;
            timeZoneXml.timeZoneRule.daylightDate.transitionTime = "2:0:0";

            timeZoneXml.timeZoneRule.standardBias = "480";
            timeZoneXml.timeZoneRule.additionalDaylightBias = "-60";

            return timeZoneXml;
        }

        /// <summary>
        /// A method used to verify whether a version string is equal or larger than specified version string. Only support the version string whose sub versions values all are integer format.
        /// </summary>
        /// <param name="currentVersion">A parameter represents the version value which is compare to specified version value.</param>
        /// <param name="specifiedVersion">A parameter represents the version value which is used as comparing standard.</param>
        /// <returns>Return 'true' indicating the version value is equal or larger than the specified version value which is specified by specifiedVersion parameter.</returns>
        protected bool VerifyEqualOrLargerThanSpecifiedVersionString(string currentVersion, string specifiedVersion)
        {
            if (string.IsNullOrEmpty(currentVersion))
            {
                throw new ArgumentNullException("currentVersion");
            }

            if (string.IsNullOrEmpty(currentVersion))
            {
                throw new ArgumentNullException("specifiedVersion");
            }

            // Verify whether equal
            if (currentVersion.Equals(specifiedVersion, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            else
            {
                string[] versionValuesOfCurrenVersionTemp = currentVersion.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                List<string> versionValuesOfCurrenVersion = new List<string>(versionValuesOfCurrenVersionTemp);

                string[] versionValusOfSpecifiedTemp = specifiedVersion.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                List<string> versionValusOfSpecified = new List<string>(versionValusOfSpecifiedTemp);

                bool isLargerThanSpecifiedVersion = this.CompareTwoVersionString(versionValuesOfCurrenVersion, versionValusOfSpecified);
                return isLargerThanSpecifiedVersion;
            }
        }

        /// <summary>
        /// A method used to get list items changes from protocol SUT, it will use "GetListItemChangesSinceToken" operation as the first choice. If the protocol SUT does not support the "GetListItemChangesSinceToken" operation, this method will use "GetListItemChanges" operation. Support GetListItemChangesSinceToken" operation is determined by "R106802Enabled" property on product option behaviors configuration file. The method will set an empty "ViewFields" element in request of SOAP operation, in order to show all fields' value of a list. This method will set "EnumRecurrencePatternXMLVersion.v3" in request, so that it could receive recurrence XML for certain types of recurrences appointment items.
        /// </summary>
        /// <param name="listId">A parameter represents the id of the list, which contain list items' changes.</param>
        /// <returns>A return value presents the list items' changes data. each item mapping a list item's data on protocol SUT.</returns>
        protected XmlNode[] GetListItemsChangesFromSUT(string listId)
        {
            XmlNode[] listItemsChanges = this.GetListItemsChangesFromSUT(listId, null);

            return listItemsChanges;
        }

        /// <summary>
        /// A method used to get list items changes from protocol SUT, it will use "GetListItemChangesSinceToken" operation as the first choice. If the protocol SUT does not support the "GetListItemChangesSinceToken" operation, this method will use "GetListItemChanges" operation. Support GetListItemChangesSinceToken" operation is determined by "R106802Enabled" property on product option behaviors configuration file. This method will set "EnumRecurrencePatternXMLVersion.v3" in request, so that it could receive recurrence XML for certain types of recurrences appointment items. If there are no any list items get by this method, it will throw Assert exception.
        /// </summary>
        /// <param name="listId">A parameter represents the id of the list, which contain list items' changes.</param>
        /// <param name="viewfieds">A parameter represents the CamlViewFields instance which will be set into the request of SOAP operation, if this value is null, this method will set a "ViewFields" element in request of SOAP operation.</param>
        /// <returns>A return value presents the list items' changes data. each item mapping a list item's data on protocol SUT.</returns>
        protected XmlNode[] GetListItemsChangesFromSUT(string listId, CamlViewFields viewfieds)
        {
            if (string.IsNullOrEmpty(listId))
            {
                throw new ArgumentNullException("listId");
            }

            XmlNode[] listItemsChanges = this.TryGetListItemsChangesFromSUT(listId, viewfieds);

            // If there are no any list items return, throw assert exception.
            this.Site.Assert.AreNotEqual<int>(
                                   0,
                       listItemsChanges.Length,
                       "The list identified by list id[{0}] should contain at least one list items.",
                       listId);

            return listItemsChanges;
        }

        /// <summary>
        /// A method used to get list items changes from protocol SUT, it will use "GetListItemChangesSinceToken" operation as the first choice. If the protocol SUT does not support the "GetListItemChangesSinceToken" operation, this method will use "GetListItemChanges" operation. Support GetListItemChangesSinceToken" operation is determined by "R106802Enabled" property on product option behaviors configuration file. This method will set "EnumRecurrencePatternXMLVersion.v3" in request, so that it could receive recurrence XML for certain types of recurrences appointment items.
        /// </summary>
        /// <param name="listId">A parameter represents the id of the list, which contain list items' changes.</param>
        /// <param name="viewfieds">A parameter represents the CamlViewFields instance which will be set into the request of SOAP operation, if this value is null, this method will set a "ViewFields" element in request of SOAP operation.</param>
        /// <returns>A return value presents the list items' changes data. each item mapping a list item's data on protocol SUT. If there are no any list items in the list, method will return a zero length zrow item array.</returns>
        protected XmlNode[] TryGetListItemsChangesFromSUT(string listId, CamlViewFields viewfieds)
        {
            if (string.IsNullOrEmpty(listId))
            {
                throw new ArgumentNullException("listId");
            }

            XmlNode[] listItemsChanges = null;

            // Set "<ViewFields />" in order to show all fields' value of a list, if the viewfieds parameter does not have value.
            if (null == viewfieds)
            {
                viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();
            }

            // If the protocol SUT support GetListItemChangesSinceToken changes, this method will use this GetListItemChangesSinceToken operation to get the list items' changes.
            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesResOfSinceToken = OutspsAdapter.GetListItemChangesSinceToken(
                                                        listId,
                                                        null,
                                                        null,
                                                        viewfieds,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

                this.VerifyContainZrowDataStructure(listItemChangesResOfSinceToken);
                listItemsChanges = this.TryGetZrowItems(listItemChangesResOfSinceToken.listitems.data.Any);
                this.Site.Log.Add(LogEntryKind.Debug, "Get the list items' changes by using GetListItemChangesSinceToken operation.");
            }
            else
            {
                // Call GetListItemChanges operation to get list items change.
                GetListItemChangesResponseGetListItemChangesResult listitemChangesRes = null;
                listitemChangesRes = OutspsAdapter.GetListItemChanges(
                                         listId,
                                         viewfieds,
                                         null,
                                         null);

                // Get the list items change data.
                this.VerifyContainZrowDataStructure(listitemChangesRes);
                listItemsChanges = this.TryGetZrowItems(listitemChangesRes.listitems.data[0].Any);
                this.Site.Log.Add(LogEntryKind.Debug, "Get the list items' changes by using GetListItemChanges operation.");
            }

            return listItemsChanges;
        }

        #endregion Helper methods

        #region Private Methods

        /// <summary>
        ///  A method used to verify whether a version string is larger than specified version string. Only support the version string whose sub versions values all are integer format.
        /// </summary>
        /// <param name="versionValuesOfCurrentVersion"> parameter represents the version value which is compare to specified version value.</param>
        /// <param name="versionValusOfSpecified">A parameter represents the version value which is used as comparing standard.</param>
        /// <returns>Return 'true' indicating the version value is larger than the specified version value which is specified by specifiedVersion parameter.</returns>
        private bool CompareTwoVersionString(List<string> versionValuesOfCurrentVersion, List<string> versionValusOfSpecified)
        {
            if (null == versionValuesOfCurrentVersion)
            {
                throw new ArgumentNullException("versionValuesOfCurrentVersion");
            }

            if (null == versionValusOfSpecified)
            {
                throw new ArgumentNullException("versionValusOfSpecified");
            }

            int indexOfCurrentItem = 0;
            string subVersionOfCurrentVersion = versionValuesOfCurrentVersion[indexOfCurrentItem];
            string subVersionOfSpecifiedVersion = versionValusOfSpecified[indexOfCurrentItem];
            if (string.IsNullOrEmpty(subVersionOfCurrentVersion) || string.IsNullOrEmpty(subVersionOfSpecifiedVersion))
            {
                this.Site.Assert.Fail("The sub version should have value.");
            }

            int subVersionOfCurrenVersionValue = -1;
            int subVersionOfSpecifiedVersionValue = -1;
            if (!int.TryParse(subVersionOfCurrentVersion, out subVersionOfCurrenVersionValue) || !int.TryParse(subVersionOfSpecifiedVersion, out subVersionOfSpecifiedVersionValue))
            {
                this.Site.Assert.Fail("The sub version should be valid integer value.");
            }

            if (subVersionOfCurrenVersionValue > subVersionOfSpecifiedVersionValue)
            {
                return true;
            }
            else if (subVersionOfCurrenVersionValue.Equals(subVersionOfSpecifiedVersionValue))
            {
                versionValuesOfCurrentVersion.RemoveAt(indexOfCurrentItem);
                versionValusOfSpecified.RemoveAt(indexOfCurrentItem);

                int currentLeftCounterOfCurrenVersion = versionValuesOfCurrentVersion.Count;
                int currentLeftCounterOfversionValusOfSpecified = versionValusOfSpecified.Count;

                if (currentLeftCounterOfCurrenVersion > 0 && currentLeftCounterOfversionValusOfSpecified > 0)
                {
                    return this.CompareTwoVersionString(versionValuesOfCurrentVersion, versionValusOfSpecified);
                }

                // If one of them does not have more sub version, and the current subversion length larger than 0,  that means specified sub version have no anymore sub versions.
                if (currentLeftCounterOfCurrenVersion > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// This method is used to construct UpdateListFieldsRequest instance for UpdateList operation's add fields parameter.
        /// </summary>
        /// <param name="fieldNames">A parameter represents the fields to be added.</param>
        /// <param name="fieldTypes">A parameter represents the field types to be added.</param>
        /// <param name="isRequired">A parameter represents the whether added field required or not.</param>
        /// <param name="viewName">A parameter represents the view to which the field should be added. If it is null, the field will not be added to any view. If it is empty string, the field will be added to default view</param>
        /// <returns>This method will return the UpdateListFieldsRequest instance.</returns>
        private UpdateListFieldsRequest CreateAddListFieldsRequest(List<string> fieldNames, List<string> fieldTypes, bool isRequired, string viewName)
        {
            UpdateListFieldsRequest newFields = new UpdateListFieldsRequest();
            newFields.Fields = new UpdateListFieldsRequestFields();
            newFields.Fields.Method = new UpdateListFieldsRequestFieldsMethod[fieldNames.Count];

            for (int i = 0; i < fieldNames.Count; i++)
            {
                newFields.Fields.Method[i] = new UpdateListFieldsRequestFieldsMethod();
                newFields.Fields.Method[i].ID = Guid.NewGuid().ToString();
                newFields.Fields.Method[i].Field = new FieldDefinition();
                newFields.Fields.Method[i].Field.Name = fieldNames[i];
                newFields.Fields.Method[i].Field.DisplayName = fieldNames[i];
                newFields.Fields.Method[i].Field.Type = fieldTypes[i];
                newFields.Fields.Method[i].Field.Required = isRequired == true ? "TRUE" : "FALSE";
                newFields.Fields.Method[i].AddToView = viewName;
            }

            return newFields;
        }

        /// <summary>
        /// A method used to add specified fields to the specified list. If any error occurs, this will fail the test case.
        /// </summary>
        /// <param name="listId">A parameter represents the id or title of a list where the fields expected to add.</param>
        /// <param name="fieldNames">A parameter represents the added fields names.</param>
        /// <param name="fieldTypes">A parameter represents the added fields types.</param>
        private void AddFieldsToList(string listId, List<string> fieldNames, List<string> fieldTypes)
        {
            this.Site.Assert.AreEqual<int>(
                                fieldNames.Count,
                                fieldTypes.Count,
                                "The element number in the field name array MUST be equal the element number in the field value array");

            // Construct UpdateListFieldsRequest instance.
            UpdateListFieldsRequest newFields = this.CreateAddListFieldsRequest(fieldNames, fieldTypes, false, null);

            try
            {
                UpdateListResponseUpdateListResult result = null;
                result = OutspsAdapter.UpdateList(
                                                listId,
                                                null,
                                                newFields,
                                                null,
                                                null,
                                                null);

                foreach (UpdateListFieldResultsMethod method in result.Results.NewFields)
                {
                    if (method.ErrorCode != "0x00000000")
                    {
                        this.Site.Assert.Fail(ErrorMessageTemplate, "adding field to the list " + listId, method.ErrorText);
                    }
                }
            }
            catch (SoapException exp)
            {
                this.Site.Assert.Fail(ErrorMessageTemplate, "adding field to the list " + listId, exp.Detail.InnerText);
            }
        }

        /// <summary>
        /// A method used to set the property value of a message data which is used for Discussion Board item. and the message data must follow this rule:The message MUST be in MIME [RFC2045] format and then Base64 [RFC4648] encoded. Message headers [RFC2822] MUST contain the following text properties. More detail are described in [MS-LISTSWS] section 3.1.4.2.2.1 
        /// </summary>
        /// <param name="discussionMessageString">A parameter represents the raw string of a message data.</param>
        /// <param name="propertyName">A parameter represents the name of property in message data.</param>
        /// <param name="expectedValue">A parameter represents the value which is set to a property which is indicated by propertyName parameter.</param>
        /// <returns>A return value represents the string of a message data which have been set the value into the specified propertyName.</returns>
        private string SetDiscussionItemMessageProperty(string discussionMessageString, string propertyName, string expectedValue)
        {
            propertyName = string.Format("{0}:", propertyName);
            int keyWordPositionOfProperty = discussionMessageString.IndexOf(propertyName, StringComparison.OrdinalIgnoreCase);
            if (-1 == keyWordPositionOfProperty)
            {
                this.Site.Assert.Fail("The [{0}] message data file should contain the [{0}] property.", propertyName);
            }

            int rowBreakerPositionOfProperty = discussionMessageString.IndexOf("\r\n", keyWordPositionOfProperty, StringComparison.OrdinalIgnoreCase);
            if (-1 == rowBreakerPositionOfProperty)
            {
                this.Site.Assert.Fail(
                                "The [{0}] message data file should contain the [{0}] property value like this format [Subject: *****\r\n].",
                                 propertyName);
            }

            int subStringStartPosition = keyWordPositionOfProperty + propertyName.Length;
            int subStringLength = rowBreakerPositionOfProperty - subStringStartPosition;
            discussionMessageString = discussionMessageString.Remove(subStringStartPosition, subStringLength);
            string currentPropertyValue = string.Format(" {0}", expectedValue);
            discussionMessageString = discussionMessageString.Insert(subStringStartPosition, currentPropertyValue);

            return discussionMessageString;
        }

        #endregion Private Methods
    }
}