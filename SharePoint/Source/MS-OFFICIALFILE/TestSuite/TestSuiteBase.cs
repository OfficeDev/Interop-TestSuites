namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is the base class for all the test classes in the MS-OFFICIALFILE test suite.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Gets or sets the instance of IProtocolAdapter.
        /// </summary>
        protected IMS_OFFICIALFILEAdapter Adapter { get; set; }

        /// <summary>
        /// Gets or sets the instance of ISUTControlAdapter.
        /// </summary>
        protected IMS_OFFICIALFILESUTControlAdapter SutControlAdapter { get; set; }
        #endregion

        #region variables storing ptfconfig values
        /// <summary>
        /// Gets or sets the webService address of a repository which is not configured for routing content.
        /// </summary>
        protected string DisableRoutingFeatureRecordsCenterServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets the webService address of a repository which is configured for routing content and user have permissions to store content.
        /// </summary>
        protected string EnableRoutingFeatureRecordsCenterServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets the webService address of a repository configured for configured for paring enable.
        /// </summary>
        protected string EnableRoutingFeatureDocumentsCenterServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets the URL of file submitted to list not configured for routing content
        /// </summary>
        protected string DefaultLibraryUrl { get; set; }

        /// <summary>
        /// Gets or sets the URL of file submitted to a list configured for document set file type only.
        /// </summary>
        protected string DocumentSetUrl { get; set; }

        /// <summary>
        /// Gets or sets the URL of file submitted to a list where routing is not enforced.
        /// </summary>
        protected string NoEnforceLibraryUrl { get; set; }

        /// <summary>
        /// Gets or sets the user have no summit permission.
        /// </summary>
        protected string LimitedUserName { get; set; }

        /// <summary>
        /// Gets or sets the password of User have no summit permission.
        /// </summary>
        protected string LimitedUserPassword { get; set; }

        /// <summary>
        /// Gets or sets the user have submit permission
        /// </summary>
        protected string SubmitUserName { get; set; }

        /// <summary>
        /// Gets or sets the name of OfficialFile server domain
        /// </summary>
        protected string DomainName { get; set; }

        /// <summary>
        /// Gets or sets the password of user
        /// </summary>
        protected string Password { get; set; }

        /// <summary>
        /// Gets or sets the document content Type
        /// </summary>
        protected string DocumentContentTypeName { get; set; }

        /// <summary>
        /// Gets or sets the picture content Type
        /// </summary>
        protected string NotSupportedContentTypeName { get; set; }

        /// <summary>
        /// Gets or sets the URL of file submitted to list configured to overwrite existing files. 
        /// </summary>
        protected string DocumentLibraryUrlOfSharePointVersion { get; set; }

        /// <summary>
        /// Gets or sets the URL of file submitted to list not configured to overwrite existing files. 
        /// </summary>
        protected string DocumentLibraryUrlOfAppendUniqueSuffix { get; set; }

        #endregion variables storing ptfconfig values

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear up the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Case Initialization
        /// <summary>
        /// Initialize the test.
        /// </summary>
        [TestInitialize]
        public void TestSuiteBaseInitialize()
        {
            Site.DefaultProtocolDocShortName = "MS-OFFICIALFILE";
            this.Adapter = Site.GetAdapter<IMS_OFFICIALFILEAdapter>();
            this.SutControlAdapter = Site.GetAdapter<IMS_OFFICIALFILESUTControlAdapter>();

            // Initial WebService address
            this.DisableRoutingFeatureRecordsCenterServiceUrl = Common.Common.GetConfigurationPropertyValue("DisableRoutingFeatureRecordsCenterServiceUrl", this.Site);
            this.EnableRoutingFeatureRecordsCenterServiceUrl = Common.Common.GetConfigurationPropertyValue("EnableRoutingFeatureRecordsCenterServiceUrl", this.Site);
            this.EnableRoutingFeatureDocumentsCenterServiceUrl = Common.Common.GetConfigurationPropertyValue("EnableRoutingFeatureDocumentsCenterServiceUrl", this.Site);

            // Initial URL of file submitted
            this.DocumentLibraryUrlOfAppendUniqueSuffix = Common.Common.GetConfigurationPropertyValue("DocumentLibraryUrlOfAppendUniqueSuffix", this.Site);
            this.DocumentLibraryUrlOfSharePointVersion = Common.Common.GetConfigurationPropertyValue("DocumentLibraryUrlOfSharePointVersion", this.Site);
            this.DefaultLibraryUrl = Common.Common.GetConfigurationPropertyValue("DefaultLibraryUrl", this.Site);
            this.DocumentSetUrl = Common.Common.GetConfigurationPropertyValue("DocumentSetUrl", this.Site);
            this.NoEnforceLibraryUrl = Common.Common.GetConfigurationPropertyValue("NoEnforceLibraryUrl", this.Site);

            // Initial the content type.
            this.DocumentContentTypeName = Common.Common.GetConfigurationPropertyValue("SupportedContentTypeName", this.Site);
            this.NotSupportedContentTypeName = Common.Common.GetConfigurationPropertyValue("NotSupportedContentTypeName", this.Site);

            // Get user info
            this.LimitedUserName = Common.Common.GetConfigurationPropertyValue("NoRecordsCenterSubmittersPermissionUserName", this.Site);
            this.LimitedUserPassword = Common.Common.GetConfigurationPropertyValue("NoRecordsCenterSubmittersPermissionPassword", this.Site);
            this.SubmitUserName = Common.Common.GetConfigurationPropertyValue("UserName", this.Site);
            this.DomainName = Common.Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.Password = Common.Common.GetConfigurationPropertyValue("Password", this.Site);

            // Check if MS-OFFICIALFILE service is supported in current SUT.
            if (!Common.Common.GetConfigurationPropertyValue<bool>("MS-OFFICIALFILE_Supported", this.Site))
            {
                Common.SutVersion currentSutVersion = Common.Common.GetConfigurationPropertyValue<Common.SutVersion>("SutVersion", this.Site);
                this.Site.Assert.Inconclusive("This test suite does not supported under current SUT, because MS-OFFICIALFILE_Supported value set to false in MS-OFFICIALFILE_{0}_SHOULDMAY.deployment.ptfconfig file.", currentSutVersion);
            }
        }

        /// <summary>
        /// Clean up the test.
        /// </summary>
        protected override void TestCleanup()
        {
            this.Adapter.Reset();
        }
        #endregion

        /// <summary>
        /// This method is used to construct the required record repository properties that is configured in the PropertyConfig.xml file.
        /// </summary>
        /// <returns>Return the required record repository properties.</returns>
        protected RecordsRepositoryProperty[] ConstructAllRequiredProperties()
        {
            var configures = this.DeserializePropertyConfig();
            return configures.RecordsRepositoryProperties.ToArray();
        }

        /// <summary>
        /// This method is used to construct the required record repository properties that is configured in the PropertyConfig.xml file.
        /// But not all the record repository properties will be returned, the first property will be skipped to return partial required properties.
        /// </summary>
        /// <returns>Return the partial required record repository properties.</returns>
        protected RecordsRepositoryProperty[] ConstructPartialRequiredProperties()
        {
            var configures = this.DeserializePropertyConfig();

            if (configures.RecordsRepositoryProperties.Count <= 1)
            {
                Site.Assert.Fail("In the PropertyConfig.xml file, at least contains two sub RecordsRepositoryProperty elements.");
            }

            // Get rid of the first element and returns the rest.
            return configures.RecordsRepositoryProperties.Skip(1).ToArray();
        }

        /// <summary>
        /// This method is used to construct the required record repository properties that is configured in the PropertyConfig.xml file and
        /// all the common properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery and _dlc_hold_searchcontexturl.
        /// </summary>
        /// <returns>Return the record repository properties that contains required and common ones.</returns>
        protected RecordsRepositoryProperty[] ConstructAllProperties()
        {
            List<RecordsRepositoryProperty> allProperties = new List<RecordsRepositoryProperty>();
            allProperties.AddRange(this.ConstructAllCommonProperties());
            var configures = this.DeserializePropertyConfig();
            allProperties.AddRange(configures.RecordsRepositoryProperties);

            return allProperties.ToArray();
        }

        /// <summary>
        /// This method is used to construct  all the common properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery and _dlc_hold_searchcontexturl.
        /// </summary>
        /// <returns>Return the record repository properties that contains common ones</returns>
        protected RecordsRepositoryProperty[] ConstructAllCommonProperties()
        {
            List<RecordsRepositoryProperty> allProperties = new List<RecordsRepositoryProperty>();

            var property = new RecordsRepositoryProperty();
            property.Name = "_dlc_hold_url";
            property.DisplayName = "url";
            property.Other = string.Empty;
            property.Type = "OfficialFileCustomType";
            property.Value = Common.Common.GetConfigurationPropertyValue("HoldUrl", this.Site);
            allProperties.Add(property);

            property = new RecordsRepositoryProperty();
            property.Name = "_dlc_hold_id";
            property.DisplayName = "id";
            property.Other = string.Empty;
            property.Type = "OfficialFileCustomType";
            property.Value = Common.Common.GetConfigurationPropertyValue("HoldId", this.Site);
            allProperties.Add(property);

            // The comments does not affect the testing result, so just hard code.
            property = new RecordsRepositoryProperty();
            property.Name = "_dlc_hold_comments";
            property.DisplayName = "comments";
            property.Other = string.Empty;
            property.Type = "OfficialFileCustomType";
            property.Value = "Hold testing";
            allProperties.Add(property);

            property = new RecordsRepositoryProperty();
            property.Name = "_dlc_hold_searchqquery";
            property.DisplayName = "searchqquery";
            property.Other = string.Empty;
            property.Type = "OfficialFileCustomType";
            property.Value = Common.Common.GetConfigurationPropertyValue("HoldSearchQuery", this.Site);
            allProperties.Add(property);

            property = new RecordsRepositoryProperty();
            property.Name = "_dlc_hold_searchcontexturl";
            property.DisplayName = "searchcontexturl";
            property.Other = string.Empty;
            property.Type = "OfficialFileCustomType";
            property.Value = Common.Common.GetConfigurationPropertyValue("HoldSearchContextUrl", this.Site);
            allProperties.Add(property);

            return allProperties.ToArray();
        }

        /// <summary>
        /// This method is used to construct the full file URL in the give library with a random file name. 
        /// </summary>
        /// <param name="libraryUrl">Specify the library URL in which the file exists.</param>
        /// <returns>Return the full file URL.</returns>
        protected string GetOriginalSaveLocation(string libraryUrl)
        {
            return this.GetOriginalSaveLocation(libraryUrl, null);
        }

        /// <summary>
        /// This method is used to construct the full file URL in the give library with the suggested file name. 
        /// </summary>
        /// <param name="libraryUrl">Specify the library URL in which the file exists.</param>
        /// <param name="suggestedFileName">Specify the suggested file name.</param>
        /// <returns>Return the full file URL.</returns>
        protected string GetOriginalSaveLocation(string libraryUrl, string suggestedFileName)
        {
            string fileName = suggestedFileName ?? this.GenerateRandomTextFileName();
            string slashUrl = libraryUrl.EndsWith(@"/") ? libraryUrl : libraryUrl + @"/";
            return new Uri(new Uri(slashUrl), fileName).AbsoluteUri;
        }

        /// <summary>
        /// This method is used to generate random file name with "TXT" suffix.
        /// </summary>
        /// <returns>Return the random file name.</returns>
        protected string GenerateRandomTextFileName()
        {
            return Common.Common.GenerateResourceName(this.Site, "TextFile") + ".txt";
        }

        /// <summary>
        /// This method is used de-serialize the properties from the configurable file PropertyConfig.xml.
        /// </summary>
        /// <returns>Return the configurable records repository properties.</returns>
        private PropertyConfig DeserializePropertyConfig()
        {
            XmlSerializer serializer = new XmlSerializer(typeof(PropertyConfig));

            using (FileStream fs = File.Open("PropertyConfig.xml", FileMode.Open))
            {
                return serializer.Deserialize(fs) as PropertyConfig;
            }
        }
    }
}