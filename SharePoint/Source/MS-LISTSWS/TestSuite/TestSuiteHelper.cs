//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that provides helper methods used in test case level.
    /// </summary>
    internal static class TestSuiteHelper
    {
        /// <summary>
        /// Error message template
        /// </summary>
        private const string ErrorMessageTemplate = "An error occurred while {0}, the error message is {1}";

        /// <summary>
        /// A IMS_LISTSWSAdapter instance
        /// </summary>
        private static IMS_LISTSWSAdapter listswsAdapter;

        /// <summary>
        /// A random generator using current time seeds.
        /// </summary>
        private static Random random;

        /// <summary>
        /// A ITestSite instance
        /// </summary>
        private static ITestSite testSite;

        /// <summary>
        /// A list used to record all lists added by TestSuiteHelper
        /// </summary>
        private static List<string> listIdCache = new List<string>();

        /// <summary>
        /// An uint indicate the list item number value on current test case.
        /// </summary>
        private static uint listitemCounterOfPerTestCases;

        /// <summary>
        /// An uint indicate the contentType number value on current test case.
        /// </summary>
        private static uint contentTypeCounterOfPerTestCases;

        /// <summary>
        ///  An uint indicate the contentType number value on current test case.
        /// </summary>
        private static uint listNameCounterOfPerTestCases;

        /// <summary>
        ///  An uint indicate the metaInfoField PropertyName number value on current test case.
        /// </summary>
        private static uint metaInfoFieldPropertyNameCounterOfPerTestCases;

        /// <summary>
        /// An uint indicate the folder name number value on current test case.
        /// </summary>
        private static uint folderNameCounterOfPerTestCases;

        /// <summary>
        /// An uint indicate the field name number value on current test case.
        /// </summary>
        private static uint fieldNameCounterOfPerTestCases;

        /// <summary>
        /// A method used to Initialize the TestSuiteHelper with specified ITestSite instance and IMS_LISTSWSAdapter instance
        /// </summary>
        /// <param name="site">A parameter represents ITestSite instance.</param>
        /// <param name="adapter">A parameter represents IMS_LISTSWSAdapter instance.</param>
        public static void Initialize(ITestSite site, IMS_LISTSWSAdapter adapter)
        {
            // Reset the counter
            contentTypeCounterOfPerTestCases = 0;
            listitemCounterOfPerTestCases = 0;
            listNameCounterOfPerTestCases = 0;
            metaInfoFieldPropertyNameCounterOfPerTestCases = 0;
            folderNameCounterOfPerTestCases = 0;
            fieldNameCounterOfPerTestCases = 0;
            listswsAdapter = adapter;
            testSite = site;
        }

        /// <summary>
        /// This method is used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>Returns the random generated string.</returns>
        public static string GenerateRandomString(int size)
        {
            random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {   
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
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
        public static string GenerateRandomNumber(int minValue, int maxValue)
        {
            random = new Random((int)DateTime.Now.Ticks);
            return random.Next(minValue, maxValue).ToString();
        }

        /// <summary>
        /// A method used to get the  valid OwsHiddenVersion for specified list item
        /// </summary>
        /// <param name="list">A parameter represents the list where the method will target</param>
        /// <returns>A return value represents the validOwsHiddenVersion for specified list item</returns>
        public static string GetOwsHiddenVersion(string list)
        {
            // Subtract 1 day for the time, so as to return all the changes of list items.
            string since = System.DateTime.Now.AddDays(-1).ToString("s");

            // Get the ViewFields whose Properties is true and reference field is MetaInfo.
            GetListItemChangesResponseGetListItemChangesResult addItems = null;
            addItems = listswsAdapter.GetListItemChanges(
                                            list,
                                            null,
                                            since,
                                            null);

            string owsVersion = null;

            if (addItems.listitems.data != null)
            {
                // If data is not null, the array data exist, select the first raw 0 to test.
                DataTable data = AdapterHelper.ExtractData(addItems.listitems.data[0].Any);
                if (data != null && data.Rows.Count == 1)
                {
                    string columnName = AdapterHelper.PrefixOws + AdapterHelper.FieldOwshiddenversionName;
                    owsVersion = Convert.ToString(data.Rows[0][columnName]);
                }
                else
                {
                    testSite.Debug.Fail("The GetListItemChanges MUST return one row");
                }
            }
            else
            {
                testSite.Debug.Fail("The GetListItemChanges MUST return one data element");
            }

            return owsVersion;
        }

        /// <summary>
        /// A method is used to generate a random version number in the format X.0.0.XXX.
        /// </summary>
        /// <returns>The value represents a random version number.</returns>
        public static string GenerateVersionNumber()
        {
            string firstNum = GenerateRandomNumber(0, 9);
            string middleNum = ".0.0.";
            string lastNum = Convert.ToInt32(GenerateRandomNumber(0, 9999)).ToString("0000");
            string combineString = string.Concat(firstNum, middleNum, lastNum);
            return combineString;
        }

        /// <summary>
        /// A method used to construct CamlViewFields instance using the specified parameters.
        /// </summary>
        /// <param name="property">A Boolean value indicate whether the Prosperities attribute value is TRUE/FALSE</param>
        /// <param name="fieldNames">Specified the CamlViewFields instance's fields.</param>
        /// <returns>Returns the CamlViewFields instance.</returns>
        public static CamlViewFields CreateViewFields(bool property, List<string> fieldNames)
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
        /// A method used to construct CamlQueryRoot instance using specified field name and value 
        /// with equal condition.
        /// </summary>
        /// <param name="fieldName">A parameter represents the field name.</param>
        /// <param name="fieldValue">A parameter represents the field value.</param>
        /// <returns>Returns the CamlQueryRoot instance.</returns>
        public static CamlQueryRoot CreateQueryRoot(string fieldName, string fieldValue)
        {
            CamlQueryRoot root = new CamlQueryRoot();
            root.Where = new LogicalJoinDefinition();

            LogicalTestDefinition equal = new LogicalTestDefinition();
            equal.FieldRef = new FieldRefDefinitionQueryTest();
            equal.FieldRef.Name = fieldName;

            ValueDefinition valueDef = new ValueDefinition();
            valueDef.Type = "Text";
            valueDef.Text = new string[] { fieldValue };
            equal.Value = valueDef;

            root.Where.Items = new object[1];
            root.Where.Items[0] = equal;

            root.Where.ItemsElementName = new ItemsChoiceType1[1];
            root.Where.ItemsElementName[0] = ItemsChoiceType1.Eq;

            return root;
        }

        /// <summary>
        /// Get the content in the file of specified file name.
        /// </summary>
        /// <param name="fileName">The specified file name</param>
        /// <returns>Return the file content if there is no error, otherwise returns null.</returns>
        public static byte[] GetAttachmentContent(string fileName)
        {
            byte[] data = null;
            FileStream fs = null;
            try
            {
                fs = File.Open(fileName, FileMode.Open);
                int len = (int)fs.Length;
                data = new byte[len];
                fs.Read(data, 0, len);
            }
            catch (IOException exp)
            {
                testSite.Debug.Fail(ErrorMessageTemplate, "reading content in the file " + fileName, exp.InnerException);
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }

            return data;
        }

        /// <summary>
        /// Extract the error code from the SoapException response.
        /// </summary>
        /// <param name="exp">The specified SoapException.</param>
        /// <returns>
        /// If there is an ErrorCode in the specified SoapException then return the error code, 
        /// otherwise return null.
        /// </returns>
        public static string GetErrorCode(SoapException exp)
        {
            return AdapterHelper.GetErrorCodeFromSoapException(exp);
        }

        /// <summary>
        /// Extract the error string from the SoapException response.
        /// </summary>
        /// <param name="exp">The specified SoapException.</param>
        /// <returns>
        /// If there is an ErrorString in the specified SoapException then return the error string, otherwise return null.
        /// </returns>
        public static string GetErrorString(SoapException exp)
        {
            // Get the SOAP error code
            if (exp.Detail.FirstChild == null)
            {
                return null;
            }
            else
            {
                if ((string.Compare(exp.Detail.FirstChild.Name, "errorstring", StringComparison.OrdinalIgnoreCase) == 0)
                        && (exp.Detail.FirstChild.InnerText != null))
                {
                    return exp.Detail.FirstChild.InnerText;
                }
            }

            return null;
        }

        /// <summary>
        /// A method used to construct UpdateListItemsUpdates instance using the specified parameters.
        /// </summary>
        /// <param name="methods">A list of MethodCmdEnum to specify the operations.</param>
        /// <param name="fieldNameValuePairs">A list of items values.</param>
        /// <returns>Returns the UpdateListItemsUpdates instance.</returns>
        public static UpdateListItemsUpdates CreateUpdateListItems(
                        List<MethodCmdEnum> methods,
                        List<Dictionary<string, string>> fieldNameValuePairs)
        {
            return CreateUpdateListItems(methods, fieldNameValuePairs, OnErrorEnum.Continue);
        }

        /// <summary>
        /// A method used to construct UpdateListItemsUpdates instance using the specified parameters.
        /// </summary>
        /// <param name="methodCollection">A list of MethodCmdEnum to specify the operations.</param>
        /// <param name="fieldNameValuePairs">A list of items values.</param>
        /// <param name="errorhandleType">Specify the OnError of the Batch element's value.</param>
        /// <returns>Returns the UpdateListItemsUpdates instance.</returns>
        public static UpdateListItemsUpdates CreateUpdateListItems(
                        List<MethodCmdEnum> methodCollection,
                        List<Dictionary<string, string>> fieldNameValuePairs,
                        OnErrorEnum errorhandleType)
        {
            testSite.Debug.IsNotNull(
                        methodCollection,
                        "The [methodCollection] parameter in CreateUpdateListItems can not be null");
            testSite.Debug.IsNotNull(
                        fieldNameValuePairs,
                        "The [methodCollection] parameter in CreateUpdateListItems can not be null");

            testSite.Debug.AreEqual<int>(
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
        /// A method used to construct UpdateListItemsWithKnowledgeUpdates instance using the specified parameters.
        /// </summary>
        /// <param name="methods">A list of MethodCmdEnum to specify the operations.</param>
        /// <param name="fieldNameValuePairs">A list of items values.</param>
        /// <returns>A return value represents the UpdateListItemsUpdates instance.</returns>
        public static UpdateListItemsWithKnowledgeUpdates CreateUpdateListWithKnowledgeItems(
            List<MethodCmdEnum> methods, List<Dictionary<string, string>> fieldNameValuePairs)
        {
            return CreateUpdateListWithKnowledgeItems(methods, fieldNameValuePairs, OnErrorEnum.Continue);
        }

        /// <summary>
        /// A method used to construct UpdateListItemsWithKnowledgeUpdates instance using the specified parameters.
        /// </summary>
        /// <param name="methods">A list of MethodCmdEnum to specify the operations.</param>
        /// <param name="fieldNameValuePairs">A list of items values.</param>
        /// <param name="errorhandleType">Specify the OnError of the Batch element's value.</param>
        /// <returns>Returns the UpdateListItemsUpdates instance.</returns>
        public static UpdateListItemsWithKnowledgeUpdates CreateUpdateListWithKnowledgeItems(
                        List<MethodCmdEnum> methods,
                        List<Dictionary<string, string>> fieldNameValuePairs,
                        OnErrorEnum errorhandleType)
        {
            testSite.Assert.IsNotNull(
                        methods,
                        "The parameter methods in CreateUpdateListWithKnowledgeItems cannot be null");
            testSite.Assert.IsNotNull(
                        fieldNameValuePairs,
                        "The parameter methods in CreateUpdateListWithKnowledgeItems cannot be null");

            testSite.Assert.AreEqual<int>(
                            methods.Count,
                            fieldNameValuePairs.Count,
                            "The element number in the methods and fieldNameValuePairs parameter List are same");

            UpdateListItemsWithKnowledgeUpdates updates = new UpdateListItemsWithKnowledgeUpdates();
            updates.Batch = new UpdateListItemsWithKnowledgeUpdatesBatch();

            updates.Batch.Method = new UpdateListItemsWithKnowledgeUpdatesBatchMethod[methods.Count];
            for (int i = 0; i < methods.Count; i++)
            {
                updates.Batch.OnError = errorhandleType;
                updates.Batch.OnErrorSpecified = true;
                updates.Batch.Method[i] = new UpdateListItemsWithKnowledgeUpdatesBatchMethod();
                updates.Batch.Method[i].Cmd = methods[i];

                int fieldCount = fieldNameValuePairs[i].Keys.Count;
                updates.Batch.Method[i].Field = new UpdateListItemsWithKnowledgeUpdatesBatchMethodField[fieldCount];
                int j = 0;
                foreach (KeyValuePair<string, string> keyPair in fieldNameValuePairs[i])
                {
                    updates.Batch.Method[i].Field[j] = new UpdateListItemsWithKnowledgeUpdatesBatchMethodField();
                    updates.Batch.Method[i].Field[j].Name = keyPair.Key;
                    updates.Batch.Method[i].Field[j].Value = keyPair.Value;
                    j++;
                }
            }

            return updates;
        }

        /// <summary>
        /// A method is used to add list items by specified number items to specified list.
        /// This method update the column value specified by ListFieldText property in the PTF configuration file, and the value of this column is specified by random generated string.
        /// </summary>
        /// <param name="list">A parameter represents list GUID or list title</param>
        /// <param name="itemCount">A parameter represents how many items would be added into the list</param>
        /// <returns>A return value represents added items id array.</returns>
        public static List<string> AddListItems(string list, int itemCount)
        {
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", testSite);

            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(itemCount);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(itemCount);

            for (int i = 0; i < itemCount; i++)
            {
                Dictionary<string, string> item = new Dictionary<string, string>();

                // We make the Text field value always "TextField" append a counter number.
                item.Add(fieldName, TestSuiteHelper.GenerateRandomString(5));
                items.Add(item);
            }

            for (int i = 0; i < itemCount; i++)
            {
                cmds.Add(MethodCmdEnum.New);
            }

            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items);
            UpdateListItemsResponseUpdateListItemsResult result = null;
            try
            {
                result = listswsAdapter.UpdateListItems(
                                                    list,
                                                    updates);
            }
            catch (SoapException exp)
            {
                testSite.Debug.Fail(ErrorMessageTemplate, "adding items to the list " + list + " and updating field " + fieldName, exp.Detail.InnerText);
            }

            // There MUST be one Column called ID.
            string columnNameId = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            List<string> ids = new List<string>();
            foreach (UpdateListItemsResponseUpdateListItemsResultResult res in result.Results)
            {
                if (res.ErrorCode == "0x00000000")
                {
                    DataTable data = AdapterHelper.ExtractData(res.Any);
                    if (data != null && data.Rows.Count == 1)
                    {
                        ids.Add(Convert.ToString(data.Rows[0][columnNameId]));
                    }
                    else
                    {
                        testSite.Assert.Fail("The Result element contains more than one z:row.");
                    }
                }
                else
                {              
                    testSite.Assert.Fail(ErrorMessageTemplate, "adding items to the list " + list, AdapterHelper.GetElementValue(res.Any, "ErrorText"));
                }
            }

            if (ids.Count != itemCount)
            {
                testSite.Assert.Fail("Expect Add {0} item, but actually add {1} items", itemCount, ids.Count);
            }

            return ids;
        }
 
        /// <summary>
        /// A method is used to insert one item which only contains "MetaInfo" field value.
        /// This method is constraint to only add one property to the lookup type "MetaInfo" field.
        /// </summary>
        /// <param name="list">A parameter represents list GUID or list title.</param>
        /// <param name="propertyName">A parameter represents one property name for the lookup type "MetaInfo" field.</param>
        /// <param name="propertyValue">A parameter represents the property value for the corresponding property name.</param>
        /// <returns>A return value represents the item's id.</returns>
        public static string AddListItemWithMetaInfoProperty(string list, string propertyName, string propertyValue)
        {
            UpdateListItemsUpdates updates = new UpdateListItemsUpdates();
            updates.Batch = new UpdateListItemsUpdatesBatch();
            updates.Batch.Method = new UpdateListItemsUpdatesBatchMethod[1];
            updates.Batch.Method[0] = new UpdateListItemsUpdatesBatchMethod();
            updates.Batch.Method[0].ID = 0;
            updates.Batch.Method[0].Cmd = MethodCmdEnum.New;
            updates.Batch.Method[0].Field = new UpdateListItemsUpdatesBatchMethodField[1];
            updates.Batch.Method[0].Field[0] = new UpdateListItemsUpdatesBatchMethodField();
            updates.Batch.Method[0].Field[0].Name = "MetaInfo";
            updates.Batch.Method[0].Field[0].Property = propertyName;
            updates.Batch.Method[0].Field[0].Value = propertyValue;

            return AddListItem(list, updates);
        }

        /// <summary>
        /// Get the content type id in the specified list name and use the specified content type name.
        /// </summary>
        /// <param name="listName">The specified list name</param>
        /// <param name="name">The specified content type name</param>
        /// <returns>Return the content type id if the content type name is in the specified list, otherwise returns null.</returns>
        public static string GetContentTypeId(string listName, string name)
        {
            // Get the parent content type
            GetListContentTypesResponseGetListContentTypesResult allTypes = null;
            allTypes = listswsAdapter.GetListContentTypes(
                                listName,
                                null);

            foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType ct in allTypes.ContentTypes.ContentType)
            {
                if (ct.Name == name)
                {
                    return ct.ID;
                }
            }

            return null;
        }

        /// <summary>
        /// Update the specified items' List.FieldText configured column in the specified list to the "NewTextField" value.
        /// </summary>
        /// <param name="list">A parameter represents the list title or list id.</param>
        /// <param name="updateIds">A parameter represents the update items ids.</param>
        /// <param name="errorEnum">A parameter represents when error occurs, the following sequence will continue or return.</param>
        public static void UpdateListItems(string list, List<string> updateIds, OnErrorEnum errorEnum)
        {
            List<Dictionary<string, string>> updatedItems = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            foreach (string id in updateIds)
            {
                Dictionary<string, string> itemId = new Dictionary<string, string>();
                itemId.Add("ID", id);

                // We make the Text field value always update to "NewTextField" append a counter number.
                itemId.Add(
                    Common.GetConfigurationPropertyValue("ListFieldText", testSite),
                    string.Format("{0}{1}", "NewTextField", id));
                updatedItems.Add(itemId);
                cmds.Add(MethodCmdEnum.Update);
            }

            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(
                                                                cmds,
                                                                updatedItems,
                                                                errorEnum);

            UpdateListItemsResponseUpdateListItemsResult result = null;

            try
            {
                result = listswsAdapter.UpdateListItems(
                                            list,
                                            updates);
            }
            catch (SoapException exp)
            {
                testSite.Debug.Fail(ErrorMessageTemplate, "updating list items in the list " + list, exp.Detail.InnerText);
            }

            if (updateIds.Count != result.Results.Length)
            {
                testSite.Assert.Fail("Expect update {0} item, but actually add {1} items", updateIds.Count, result.Results.Length);
            }

            foreach (UpdateListItemsResponseUpdateListItemsResultResult res in result.Results)
            {
                if (res.ErrorCode != "0x00000000")
                {               
                    testSite.Debug.Fail(ErrorMessageTemplate, "updating list items in the list " + list, res.Any);
                }
            }
        }

        /// <summary>
        /// A method used to Create a list by configured list name and configured list template 
        /// </summary>
        /// <returns>A return value represents the added list's id</returns>
        public static string CreateList()
        {
            string listName = GetUniqueListName();
            int templateID = (int)TemplateType.Generic_List;
            return CreateList(listName, templateID);
        }

        /// <summary>
        /// A method used to Create a list by configured list name and specified list template. 
        /// </summary>
        /// <param name="templateId">A parameter represents the templateId of list template</param>
        /// <returns>A return value represents the added list's id</returns>
        public static string CreateList(int templateId)
        {
            string listName = GetUniqueListName();
            return CreateList(listName, templateId);
        }

        /// <summary>
        /// A method used to Create a list by specified list name and configured list template.
        /// </summary>
        /// <param name="listName">A parameter represents the list name.</param>
        /// <returns>A return value represents the added list's id</returns>
        public static string CreateList(string listName)
        {
            int templateID = (int)TemplateType.Generic_List;
            return CreateList(listName, templateID);
        }

        /// <summary>
        /// A method used to Create a list by specified list name and template Id
        /// </summary>
        /// <param name="listName">A parameter represents the list name</param>
        /// <param name="templateId">A parameter represents the templateId of list template</param>
        /// <returns>A return value represents the added list's id</returns>
        public static string CreateList(string listName, int templateId)
        {
            string listId = null;
            try
            {
                AddListResponseAddListResult addListResult = null;
                addListResult = listswsAdapter.AddList(listName, null, templateId);

                // Keep the create list id
                listId = addListResult.List.ID;
                listIdCache.Add(listId);

                // Make sure the created list contain the following fields,
                // This two fields will be used in the following sequence.
                List<string> listNames = new List<string> { Common.GetConfigurationPropertyValue("ListFieldText", testSite), Common.GetConfigurationPropertyValue("ListFieldCounter", testSite) };
                List<string> listTypes = new List<string> { "Text", "Counter" };

                // Add this two fields.
                AddFieldsToList(
                            listId,
                            listNames,
                            listTypes,
                            new List<string> { null, null });
            }
            catch (SoapException exp)
            {               
                testSite.Assert.Fail(ErrorMessageTemplate, "creating the list " + listName, exp.Detail.InnerText);
            }

            return listId;
        }

        /// <summary>
        /// A method used to Clean up lists added by Test Case Helper. 
        /// Any lists added by Test Case Helper would be deleted.
        /// </summary>
        public static void CleanUp()
        {
            // Make a copy to loop in the copy array.
            List<string> listIdCopy = new List<string>(listIdCache.ToArray());

            foreach (string id in listIdCopy)
            {
                RemoveList(id);
            }
        }

        /// <summary>
        /// A method used to verify whether current environment is clean status
        /// </summary>
        /// <returns>A return value represents whether the current environment is clean status, 
        /// “true” means current environment is clean, “false” means not. </returns>
        public static bool GuardEnviromentClean()
        {
            return listIdCache.Count == 0;
        }

        /// <summary>
        /// A method is used to deep compare two object.
        /// </summary>
        /// <param name="objectLeft">Specify the left object instance.</param>
        /// <param name="objectRight">Specify the right object instance.</param>
        /// <returns>Return true if the left and right object are same, otherwise false.</returns>
        public static bool DeepCompare(object objectLeft, object objectRight)
        {
            if (objectLeft != null && objectRight != null)
            {
                if (objectLeft.GetType() == objectRight.GetType())
                {
                    string left = SerializerHelp(objectLeft, objectLeft.GetType());
                    string right = SerializerHelp(objectRight, objectRight.GetType());

                    return string.Compare(left, right, StringComparison.OrdinalIgnoreCase) == 0;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                if (objectLeft == null && objectRight == null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Delete the specified items in the specified list.
        /// </summary>
        /// <param name="list">A parameter represents the list title or list id.</param>
        /// <param name="deleteIDList">Specify the delete items id.</param>
        /// <param name="errorEnum">Specify how to affect the following sequence when one delete item method fails.</param>
        public static void RemoveListItems(string list, List<string> deleteIDList, OnErrorEnum errorEnum)
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

            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(
                                                                cmds,
                                                                deletedItems,
                                                                errorEnum);

            UpdateListItemsResponseUpdateListItemsResult result = listswsAdapter.UpdateListItems(
                                                                                    list,
                                                                                    updates);

            if (deleteIDList.Count != result.Results.Length)
            {
                testSite.Assert.Fail("Expect delete {0} item, but actually add {1} items", deleteIDList.Count, result.Results.Length);
            }

            foreach (UpdateListItemsResponseUpdateListItemsResultResult res in result.Results)
            {
                if (res.ErrorCode != "0x00000000")
                {                   
                    testSite.Debug.Fail(ErrorMessageTemplate, "deleting item from the list " + list, AdapterHelper.GetElementValue(res.Any, "ErrorText"));
                }
            }
        }

        /// <summary>
        /// A method used to remove list by specified list id
        /// </summary>
        /// <param name="listId">A parameter represents the list id</param>
        /// <returns>A return value represents whether the list is removed successfully, “true” means the list is removed successfully, “false” means not. 
        /// </returns>
        public static bool RemoveList(string listId)
        {
            try
            {
                listswsAdapter.DeleteList(listId);
                listIdCache.Remove(listId);
            }
            catch (SoapException exp)
            {              
                testSite.Log.Add(LogEntryKind.Warning, ErrorMessageTemplate, "removing the list " + listId, exp.Detail.InnerText);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Create a content type on the list.
        /// </summary>
        /// <param name="list">The list title or list id which the content type is created on.</param>
        /// <param name="displayName">The contented display name.</param>
        /// <param name="fieldNames">The fields of the created content type.</param>
        /// <returns>Return the content type id if successful, otherwise will fail the test case and log the reason</returns>
        public static string CreateContentType(string list, string displayName, List<string> fieldNames)
        {
            // Try to Get the first exist basic content type
            string contentTypeName = GetFirstExistContentTypeName(list);

            // Try to Get the parent id by using the content name.
            string parentTypeId = GetContentTypeId(list, contentTypeName);
            if (parentTypeId == null)
            {
                testSite.Debug.Fail("Failed to find the content type {0} in the list {1}", contentTypeName, list);
            }

            AddOrUpdateFieldsDefinition addField = CreateAddContentTypeFields(fieldNames.ToArray());
            return CreateContentType(list, displayName, displayName, displayName, parentTypeId, addField);
        }

        /// <summary>
        /// This method is used to add specified fields to the specified list.
        /// All the added field will be not required fields.
        /// If any error occurs, this will fail the test case.
        /// </summary>
        /// <param name="listId">Specify the list which will be added fields.</param>
        /// <param name="fieldNames">Specify the added fields names.</param>
        /// <param name="fieldTypes">Specify the added fields types.</param>
        /// <param name="viewNames">A parameter represents the whether added field to a view. If it is null, the field will not be added to any view. If it is empty string, the field will not be added to default view</param>
        public static void AddFieldsToList(string listId, List<string> fieldNames, List<string> fieldTypes, List<string> viewNames)
        {
            AddFieldsToList(listId, fieldNames, fieldTypes, false, viewNames);
        }

        /// <summary>
        /// This method is used to add specified fields to the specified list.
        /// If any error occurs, this will fail the test case.
        /// </summary>
        /// <param name="listId">Specify the list which will be added fields.</param>
        /// <param name="fieldNames">Specify the added fields names.</param>
        /// <param name="fieldTypes">Specify the added fields types.</param>
        /// <param name="isRequired">Specify all the added fields whether required or not.</param>
        /// <param name="viewNames">A parameter represents the view to which the field should be added. If it is null, the field will not be added to any view. If it is empty string, the field will be added to default view</param>
        public static void AddFieldsToList(string listId, List<string> fieldNames, List<string> fieldTypes, bool isRequired, List<string> viewNames)
        {
            testSite.Assert.AreEqual<int>(
                                fieldNames.Count,
                                fieldTypes.Count,
                                "The element number in the field name array MUST be equal the element number in the field value array");

            // Construct UpdateListFieldsRequest instance.
            UpdateListFieldsRequest newFields = CreateAddListFieldsRequest(fieldNames, fieldTypes, isRequired, viewNames);

            try
            {
                UpdateListResponseUpdateListResult result = null;
                result = listswsAdapter.UpdateList(
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
                        testSite.Assert.Fail(ErrorMessageTemplate, "adding field to the list " + listId, method.ErrorText);
                    }
                }
            }
            catch (SoapException exp)
            {                
                testSite.Assert.Fail(ErrorMessageTemplate, "adding field to the list " + listId, exp.Detail.InnerText);
            }
        }

        /// <summary>
        /// This method is used to construct UpdateListFieldsRequest instance for UpdateList operation's add fields parameter.
        /// All the fields will be constructed as not required fields.
        /// </summary>
        /// <param name="fieldNames">A parameter represents the fields to be added.</param>
        /// <param name="fieldTypes">A parameter represents the field types to be added.</param>
        /// <param name="viewNames">A parameter represents the view to which the field should be added. If it is null, the field will not be added to any view. If it is empty string, the field will be added to default view</param>
        /// <returns>This method will return the UpdateListFieldsRequest instance.</returns>
        public static UpdateListFieldsRequest CreateAddListFieldsRequest(List<string> fieldNames, List<string> fieldTypes, List<string> viewNames)
        {
            return CreateAddListFieldsRequest(fieldNames, fieldTypes, false, viewNames);
        }

        /// <summary>
        /// This method is used to construct UpdateListFieldsRequest instance for UpdateList operation's delete fields parameter.
        /// </summary>
        /// <param name="fieldNames">A parameter represents the fields to be deleted.</param>
        /// <returns>The UpdateListFieldsRequest instance.</returns>
        public static UpdateListFieldsRequest CreateDeleteListFieldsRequest(List<string> fieldNames)
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
            }

            return newFields;
        }

        /// <summary>
        /// This method is used to get the specified list definition.
        /// </summary>
        /// <param name="listId">The specified list id which needs to be retrieved list definition.</param>
        /// <returns>This method returns the list id specified list definition if success, otherwise will fail the test case.</returns>
        public static ListDefinitionSchema GetListDefinition(string listId)
        {
            ListDefinitionSchema listDef = null;

            try
            {
                GetListResponseGetListResult result = listswsAdapter.GetList(listId);

                testSite.Assert.IsTrue(
                            result != null && result.List != null,
                            "The GetList Result in the method GetListDefinition MUST contain list definition");

                listDef = result.List;
            }
            catch (SoapException exp)
            {               
                testSite.Assert.Fail(ErrorMessageTemplate, "getting the list definition of " + listId, exp.Detail.InnerText);
            }

            return listDef;
        }

        /// <summary>
        /// Construct updated fields for the UpdateContentType operation using the specified field names.
        /// This will update the field display name add a word "New" in the prefix before old name.
        /// </summary>
        /// <param name="fieldNames">The updated field names.</param>
        /// <returns>The AddOrUpdateFieldsDefinition type instance used in the UpdateContentType operation.</returns>
        public static AddOrUpdateFieldsDefinition CreateUpdateContentTypeFields(params string[] fieldNames)
        {
            AddOrUpdateFieldsDefinition updateField = new AddOrUpdateFieldsDefinition();
            updateField.Fields = new AddOrUpdateFieldsDefinitionMethod[fieldNames.Length];
            for (int i = 0; i < fieldNames.Length; i++)
            {
                updateField.Fields[i] = new AddOrUpdateFieldsDefinitionMethod();
                updateField.Fields[i].ID = Guid.NewGuid().ToString();
                updateField.Fields[i].Field = new AddOrUpdateFieldDefinition();
                updateField.Fields[i].Field.Name = fieldNames[i];
                updateField.Fields[i].Field.DisplayName = "New" + fieldNames[i];
            }

            return updateField;
        }

        /// <summary>
        /// Construct deleted fields for the UpdateContentType operation using the specified field names.
        /// </summary>
        /// <param name="fieldNames">The deleted field names.</param>
        /// <returns>The DeleteFieldsDefinition type instance used in the UpdateContentType operation.</returns>
        public static DeleteFieldsDefinition CreateDeleteContentTypeFields(params string[] fieldNames)
        {
            DeleteFieldsDefinition deleteFields = new DeleteFieldsDefinition();
            deleteFields.Fields = new DeleteFieldsDefinitionMethod[fieldNames.Length];
            for (int i = 0; i < fieldNames.Length; i++)
            {
                deleteFields.Fields[i] = new DeleteFieldsDefinitionMethod();
                deleteFields.Fields[i].ID = Guid.NewGuid().ToString();
                deleteFields.Fields[i].Field = new DeleteFieldDefinition();
                deleteFields.Fields[i].Field.Name = fieldNames[i];
            }

            return deleteFields;
        }

        /// <summary>
        /// Construct added fields for the UpdateContentType operation using the specified field names.
        /// </summary>
        /// <param name="fieldNames">The added field names.</param>
        /// <returns>The AddOrUpdateFieldsDefinition type instance used in the UpdateContentType operation.</returns>
        public static AddOrUpdateFieldsDefinition CreateAddContentTypeFields(params string[] fieldNames)
        {
            AddOrUpdateFieldsDefinition addField = new AddOrUpdateFieldsDefinition();
            addField.Fields = new AddOrUpdateFieldsDefinitionMethod[fieldNames.Length];
            for (int i = 0; i < fieldNames.Length; i++)
            {
                addField.Fields[i] = new AddOrUpdateFieldsDefinitionMethod();
                addField.Fields[i].ID = Guid.NewGuid().ToString();
                addField.Fields[i].Field = new AddOrUpdateFieldDefinition();
                addField.Fields[i].Field.Name = fieldNames[i];
            }

            return addField;
        }

        /// <summary>
        /// This method is used to construct a new document XmlNode.
        /// </summary>
        /// <param name="elementName">A parameter represents he Xml node's element name.</param>
        /// <param name="uri">A parameter represents he Xml node's element namespace URI.</param>
        /// <param name="innerXml">A parameter represents he Xml node's inner xml content.</param>
        /// <returns>This method returns the new document XmlNode instance.</returns>
        public static XmlNode CreateNewDocument(string elementName, string uri, string innerXml)
        {
            XmlDocument document = new XmlDocument();
            XmlNode newDocument = document.CreateElement(elementName, uri);
            newDocument.InnerXml = innerXml;
            return newDocument;
        }

        /// <summary>
        /// A method used to Get an invalid format GUID string, it is not corresponds to any list's title
        /// </summary>
        /// <returns>The method returns an invalid format GUID string.</returns>
        public static string GetInvalidGuidAndNocorrespondString()
        {
            Guid guidTemp = Guid.NewGuid();
            string invalidStringTemp = guidTemp.ToString("D");

            // remove the first symbol
            invalidStringTemp = invalidStringTemp.Remove(0, 1);

            // remove the last symbol
            invalidStringTemp = invalidStringTemp.Remove(invalidStringTemp.Length - 1, 1);
            return invalidStringTemp;
        }

        /// <summary>
        /// A method used to get full URL of an attachment 
        /// </summary>
        /// <param name="listId">A parameter represents current list</param>
        /// <param name="itemId">A parameter represents the list item where the method would search attachment</param>
        /// <param name="attachmentFileName">A parameter represents the attachment fie name which is used to key work</param>
        /// <returns>A return value represents the fully URL of attachment match the condition</returns>
        public static string GetAttachmentFullUrl(string listId, string itemId, string attachmentFileName)
        {
            if (string.IsNullOrEmpty(listId) || string.IsNullOrEmpty(itemId) || string.IsNullOrEmpty(attachmentFileName))
            {
                testSite.Assert.Fail("All the input parameters should not be null or empty");
            }

            GetAttachmentCollectionResponseGetAttachmentCollectionResult attachmentResult = null;
            attachmentResult = listswsAdapter.GetAttachmentCollection(listId, itemId);
            if (null == attachmentResult || null == attachmentResult.Attachments || attachmentResult.Attachments.Length != 1)
            {
                testSite.Assert.Fail("Failed to get the attachment of List {0}", listId);
            }

            // Get the attachment URL from GetAttachmentCollection response.
            // The URL value is contained in string array,  which is defined in [MS-LISTSWS].
            // There is at least one text item which contain the "attachmentFileName" in the Text array.
            string attachmentUrl = string.Empty;
            if (attachmentResult.Attachments != null && attachmentResult.Attachments.Length > 0)
            {
                var urlItems = from urlItem in attachmentResult.Attachments
                               where urlItem.IndexOf(attachmentFileName, StringComparison.OrdinalIgnoreCase) >= 0
                               select urlItem;
                if (urlItems.Count() > 0)
                {
                    attachmentUrl = urlItems.ElementAt(0);
                }
            }

            // If the expected attachment URL cannot be found by attachment file name, fail the case.
            if (string.IsNullOrEmpty(attachmentUrl))
            {
                testSite.Assert.Fail("Failed to get the expected attachment URL by specified attachment file name:[{0}]", attachmentFileName);
            }

            return attachmentUrl;
        }

        /// <summary>
        /// A method is used to get a Unique List Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the List Object name and time stamp</returns>
        public static string GetUniqueListName()
        {
            listNameCounterOfPerTestCases += 1;
            return GetUniqueNameByObjectName("List", listNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique List item Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the ListItem Object name and time stamp</returns>
        public static string GetUniqueListItemName()
        {
            listitemCounterOfPerTestCases += 1;
            return GetUniqueNameByObjectName("ListItem", listitemCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique ContentType Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the ContentType Object name and time stamp</returns>
        public static string GetUniqueContentTypeName()
        {
            contentTypeCounterOfPerTestCases += 1;
            return GetUniqueNameByObjectName("ContentType", contentTypeCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique MetaInfoField Property Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the MetaInfoField Property Object name and time stamp</returns>
        public static string GetUniqueMetaInfoFieldPropertyName()
        {
            metaInfoFieldPropertyNameCounterOfPerTestCases += 1;
            return GetUniqueNameByObjectName("MetaInfoFieldProperty", metaInfoFieldPropertyNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to Get a valid format message data for AddDiscussionBoardItem operation
        /// </summary>
        /// <returns>A return value represents the valid format message data</returns>
        public static byte[] GetMessageDataForAddDiscussionBoardItem()
        {
            // read the DiscussionBoard Item message file to a string
            string messageFilePath = Common.GetConfigurationPropertyValue("MessageDataFileName", testSite);
            string messageTemp = File.ReadAllText(messageFilePath);

            // Convert to a byte array
            byte[] messageData = Encoding.Default.GetBytes(messageTemp);
            return messageData;
        }

        /// <summary>
        /// A method is used to get a Unique folder Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the folder Object name and time stamp</returns>
        public static string GetUniqueFolderName()
        {
            folderNameCounterOfPerTestCases += 1;
            return GetUniqueNameByObjectName("Folder", folderNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method is used to get a Unique field Name
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the field Object name and time stamp</returns>
        public static string GetUniqueFieldName()
        {
            fieldNameCounterOfPerTestCases += 1;
            return GetUniqueNameByObjectName("Field", fieldNameCounterOfPerTestCases);
        }

        /// <summary>
        /// A method used to get a existed Content Type, it is used in set a parent type value in CreateContentType operation.
        /// </summary>
        /// <param name="listId">A parameter represents the identification of a list</param>
        /// <returns>A return value represents the name of a existed Content Type</returns>
        public static string GetFirstExistContentTypeName(string listId)
        {
            if (string.IsNullOrEmpty(listId))
            {
                testSite.Assert.Fail("The listId parameter should not be null or empty");
            }

            GetListContentTypesResponseGetListContentTypesResult result = null;
            try
            {
                result = listswsAdapter.GetListContentTypes(listId, null);
            }
            catch (SoapException soapEx)
            {                
                testSite.Debug.Fail(ErrorMessageTemplate, "getting list with list id " + listId, soapEx.Detail.InnerText);
            }

            string existedParentTypeName = string.Empty;
            if (result == null || result.ContentTypes == null || result.ContentTypes.ContentType == null
                || result.ContentTypes.ContentType.Length < 1)
            {
                string errorMsg = string.Format("The content type does not exist in current list [{0}]", listId);
                testSite.Assert.Fail(errorMsg);
            }

            existedParentTypeName = result.ContentTypes.ContentType[0].Name;
            return existedParentTypeName;
        }

        #region private methods

        /// <summary>
        /// This method is used to serialize an object to string.
        /// </summary>
        /// <param name="targetObject">The serialized target object.</param>
        /// <param name="type">The serialized type.</param>
        /// <returns>The serialized result string.</returns>
        private static string SerializerHelp(object targetObject, Type type)
        {
            XmlSerializer serializer = new XmlSerializer(type);
            using (StringWriter sw = new StringWriter())
            {
                XmlTextWriter writer = new XmlTextWriter(sw);
                serializer.Serialize(writer, targetObject);

                return sw.ToString();
            }
        }

        /// <summary>
        /// This method is used to Create a content type using the specified information.
        /// </summary>
        /// <param name="list">The list title or list id which the content type is created on.</param>
        /// <param name="displayName">The content type display name.</param>
        /// <param name="title">The content type title.</param>
        /// <param name="description">The content type description</param>
        /// <param name="parentTypeId">Specify the content type's parent content type ID.</param>
        /// <param name="addField">Specify the exist fields that adds into the list content type.</param>
        /// <returns>Return the content type id if successful, otherwise will fail the test case and log the reason</returns>
        private static string CreateContentType(string list, string displayName, string title, string description, string parentTypeId, AddOrUpdateFieldsDefinition addField)
        {
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Description = description;
            addProperties.ContentType.Title = title;

            string contentTypeId = null;

            try
            {
                contentTypeId = listswsAdapter.CreateContentType(
                                                    list,
                                                    displayName,
                                                    parentTypeId,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            }
            catch (SoapException exp)
            {              
                testSite.Assert.Fail(ErrorMessageTemplate, "creating content type in the list " + list, exp.Detail.InnerText);
                throw;
            }

            return contentTypeId;
        }

        /// <summary>
        /// A method is used to get a Unique Name by specified object name and counter.
        /// </summary>
        /// <param name="objectName">A parameter represents object name which is used to combine the unique name</param>
        /// <param name="counter">>A parameter represents the counter which is used to combine the unique name </param>
        /// <returns>>A return value represents the unique name combined with object name, counter and time stamp </returns>
        private static string GetUniqueNameByObjectName(string objectName, uint counter)
        {
            string protocolPrefix = "MSLISTSWS";
            return string.Format(
                              @"{0}_{1}{2}_{3}",
                              protocolPrefix,
                              objectName,
                              counter,
                              Common.FormatCurrentDateTime());
        }

        /// <summary>
        /// A method is used to insert one item to the specified list.
        /// </summary>
        /// <param name="list">A parameter represents list GUID or list title.</param>
        /// <param name="updates">A parameter represents UpdateListItemsUpdates instances.</param>
        /// <returns>This method will return the item's ID.</returns>
        private static string AddListItem(string list, UpdateListItemsUpdates updates)
        {
            UpdateListItemsResponseUpdateListItemsResult result = null;
            try
            {
                result = listswsAdapter.UpdateListItems(
                                                    list,
                                                    updates);
            }
            catch (SoapException exp)
            {              
                testSite.Debug.Fail(ErrorMessageTemplate, "adding items in the list " + list, exp.Detail.InnerText);
            }

            testSite.Assert.AreEqual<int>(
                                1,
                                result.Results.Count(),
                                "The result must contain one result");

            DataTable data = AdapterHelper.ExtractData(result.Results[0].Any);
            string columnNameId = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");

            testSite.Assert.AreEqual<int>(
                                1,
                                data.Rows.Count,
                                "There MUST be one z:row");

            // There MUST be one column called ID
            return Convert.ToString(data.Rows[0][columnNameId]);
        }

        /// <summary>
        /// This method is used to construct UpdateListFieldsRequest instance for UpdateList operation's add fields parameter.
        /// </summary>
        /// <param name="fieldNames">A parameter represents the fields to be added.</param>
        /// <param name="fieldTypes">A parameter represents the field types to be added.</param>
        /// <param name="isRequired">A parameter represents the whether added field required or not.</param>
        /// <param name="viewNames">A parameter represents the view to which the field should be added. If it is null, the field will not be added to any view. If it is empty string, the field will be added to default view</param>
        /// <returns>This method will return the UpdateListFieldsRequest instance.</returns>
        private static UpdateListFieldsRequest CreateAddListFieldsRequest(List<string> fieldNames, List<string> fieldTypes, bool isRequired, List<string> viewNames)
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
                newFields.Fields.Method[i].AddToView = viewNames[i];
            }

            return newFields;
        }
        #endregion
    }
}