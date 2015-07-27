//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Contain test cases designed to test [MS-WSSREST] protocol.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// Adapter instance.
        /// </summary>
        private IMS_WSSRESTAdapter adapter;

        /// <summary>
        /// MS_WSSRESTSUTAdapter instance.
        /// </summary>
        private IMS_WSSRESTSUTControlAdapter sutAdapter;

        /// <summary>
        /// The GeneralList name.
        /// </summary>
        private string generalListName;

        /// <summary>
        /// The DocumentLibrary name.
        /// </summary>
        private string documentLibraryName;

        /// <summary>
        /// Gets the adapter instance.
        /// </summary>
        protected IMS_WSSRESTAdapter Adapter
        {
            get
            {
                return this.adapter;
            }
        }

        /// <summary>
        /// Gets the MS_WSSRESTSUTAdapter instance.
        /// </summary>
        protected IMS_WSSRESTSUTControlAdapter SutAdapter
        {
            get
            {
                return this.sutAdapter;
            }
        }

        /// <summary>
        /// Gets the GeneralList name.
        /// </summary>
        protected string GeneralListName
        {
            get
            {
                return this.generalListName;
            }
        }

        /// <summary>
        /// Gets the DocumentLibrary name.
        /// </summary>
        protected string DocumentLibraryName
        {
            get
            {
                return this.documentLibraryName;
            }
        }

        #endregion Variables

        #region Test Case Initialization

        /// <summary>
        /// A test case's level initialization method for TestSuiteBase class. It will perform before each test case.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.TestInitialize();

            this.adapter = BaseTestSite.GetAdapter<IMS_WSSRESTAdapter>();
            this.sutAdapter = BaseTestSite.GetAdapter<IMS_WSSRESTSUTControlAdapter>();
            this.documentLibraryName = Common.GetConfigurationPropertyValue("DoucmentLibraryListName", this.Site);
            this.generalListName = Common.GetConfigurationPropertyValue("GeneralListName", this.Site);
          
            // Check if MS-WSSREST service is supported in current SUT.
            if (!Common.GetConfigurationPropertyValue<bool>("MS-WSSREST_Supported", this.Site))
            {
                SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", this.Site);
                this.Site.Assert.Inconclusive("This test suite does not supported under current SUT, because MS-WSSREST_Supported value set to false in MS-WSSREST_{0}_SHOULDMAY.deployment.ptfconfig file.", currentSutVersion);
            }
        }

        #endregion Test Case Initialization

        #region Methods

        /// <summary>
        /// Delete all list items in the special list.
        /// </summary>
        /// <param name="listName">The list name.</param>
        protected void DeleteListItems(string listName)
        {
            // Retrieve List
            Request retrieve = new Request();
            retrieve.Parameter = listName;
            List<Entry> retrieveResults = this.adapter.RetrieveListItem(retrieve) as List<Entry>;

            if (retrieveResults != null && retrieveResults.Count > 0)
            {
                foreach (Entry item in retrieveResults)
                {
                    Request deleteRequest = new Request();
                    deleteRequest.Parameter = string.Format("{0}({1})", listName, item.Properties["Id"]);
                    this.adapter.DeleteListItem(deleteRequest);
                }
            }
        }

        /// <summary>
        /// Generate the http request body.
        /// </summary>
        /// <param name="properties">The dictionary that contains field name and field value.</param>
        /// <returns>The content of http request.</returns>
        protected string GenerateContent(Dictionary<string, string> properties)
        {
            if (properties != null && properties.Count > 0)
            {
                XmlDocument doc = new XmlDocument();
                XmlElement rootElement = doc.CreateElement("entry");
                rootElement.SetAttribute("xmlns:d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
                rootElement.SetAttribute("xmlns:m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");
                rootElement.SetAttribute("xmlns", "http://www.w3.org/2005/Atom");

                XmlElement content = doc.CreateElement("content");
                content.SetAttribute("type", "application/xml");

                XmlElement property = doc.CreateElement("m", "properties", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");
                foreach (KeyValuePair<string, string> item in properties)
                {
                    XmlElement propertyItem = doc.CreateElement("d", item.Key, "http://schemas.microsoft.com/ado/2007/08/dataservices");
                    propertyItem.InnerText = item.Value;
                    property.AppendChild(propertyItem);
                }

                doc.AppendChild(rootElement);
                rootElement.AppendChild(content);
                content.AppendChild(property);

                return doc.OuterXml;
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Insert multiple list items to list.
        /// </summary>
        /// <param name="listName">The list name.</param>
        /// <param name="insertNumber">The number of inserted list items.</param>
        protected void InsertListItems(string listName, int insertNumber)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert list items to list
            Request insertRequest = new Request();
            insertRequest.Parameter = listName;
            insertRequest.Content = this.GenerateContent(properties);
            insertRequest.ContentType = "application/atom+xml";

            for (int i = 1; i <= insertNumber; i++)
            {
                Entry insertResult = this.adapter.InsertListItem(insertRequest);
                Site.Assert.IsNotNull(insertResult, "Verify insertResult is not null!");
            }
        }

        /// <summary>
        /// Get properties from entity type.
        /// </summary>
        /// <param name="xnl">The metadata from server.</param>
        /// <param name="entityTypeName">The entity type name.</param>
        /// <returns>The properties list.</returns>
        protected List<string> GetPropertiesOfEntityType(XmlNodeList xnl, string entityTypeName)
        {
            List<string> properites = new List<string>();

            foreach (XmlNode node in xnl)
            {
                string listName = node.Attributes["Name"].Value;
                if (listName.Equals(entityTypeName, StringComparison.OrdinalIgnoreCase) && null != node.ChildNodes)
                {
                    foreach (XmlNode itemNode in node.ChildNodes)
                    {
                        if (itemNode.Name.Equals("Property", StringComparison.OrdinalIgnoreCase))
                        {
                            properites.Add(itemNode.Attributes["Name"].Value);
                        }
                    }
                }
            }

            return properites;
        }

        /// <summary>
        /// Check whether the special entity type exist in the metadata.
        /// </summary>
        /// <param name="xnl">The metadata from server.</param>
        /// <param name="entityTypeName">The entity type name.</param>
        /// <returns>True if the specified entity type is contained in the metadata, otherwise false.</returns>
        protected bool IsExistEntityType(XmlNodeList xnl, string entityTypeName)
        {
            bool result = false;

            foreach (XmlNode node in xnl)
            {
                string listName = node.Attributes["Name"].Value;
                if (listName.Equals(entityTypeName, StringComparison.OrdinalIgnoreCase))
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        #endregion
    }
}
