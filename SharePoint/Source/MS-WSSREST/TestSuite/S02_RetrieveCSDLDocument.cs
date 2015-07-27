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
    /// This scenario is used to retrieve the conceptual schema definition language (CSDL) document.
    /// </summary>
    [TestClass]
    public class S02_RetrieveCSDLDocument : TestSuiteBase
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

        /// <summary>
        /// This test case is used to retrieve a CSDL document.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S02_TC01_RetrieveACSDLDocument()
        {
            string choiceFieldName = Common.GetConfigurationPropertyValue("ChoiceFieldName", this.Site);
            string mutilChoiceFieldName = Common.GetConfigurationPropertyValue("MultiChoiceFieldName", this.Site);

            // Retrieve the metadata
            Request retrieveMetadata = new Request();
            retrieveMetadata.Parameter = "$metadata";
            retrieveMetadata.Accept = "application/xml";
            XmlDocument retrieveMetadataResult = this.Adapter.RetrieveListItem(retrieveMetadata) as XmlDocument;
            Site.Assert.IsNotNull(retrieveMetadataResult, "Verify retrieveMetadataResult is not null");

            XmlNodeList entityTypes = retrieveMetadataResult.GetElementsByTagName("EntityType");

            // Get entity type properties of single choice field
            string choicefieldEntityType = string.Format("{0}{1}Value", this.GeneralListName, choiceFieldName);
            List<string> choiceProperites = this.GetPropertiesOfEntityType(entityTypes, choicefieldEntityType);

            // If the entity type of single choice field contains a single property "Value", the MS-WSSREST_R89 can be verified
            Site.Log.Add(LogEntryKind.Debug, "If the entity type of single choice field contains a single property 'Value', the MS-WSSREST_R89 can be verified.");
            bool isVerifyR89 = choiceProperites.Count == 1 && choiceProperites[0].Equals("Value", StringComparison.OrdinalIgnoreCase);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR89,
                89,
                @"[In Choice or Multi-Choice Field] The EntityType for this EntitySet [of choice] as specified in [MC-CSDL] section 2.1.17 contains a single property ""Value"", which also serves as its [EntitySet's] EntityKey as specified in [MC-CSDL] section 2.1.5.");

            // Get entity type properties of multi-choice field
            string mutilChoicefieldEntityType = string.Format("{0}{1}Value", this.GeneralListName, mutilChoiceFieldName);
            List<string> mutilChoiceProperites = this.GetPropertiesOfEntityType(entityTypes, mutilChoicefieldEntityType);
            Site.Assert.IsNotNull(mutilChoiceProperites, "Verify mutilChoiceProperites is not null");

            // If the entity type of multi-choice field contains a single property "Value", the MS-WSSREST_R90 can be verified
            Site.Log.Add(LogEntryKind.Debug, "If the entity type of multi-choice field contains a single property 'Value', the MS-WSSREST_R90 can be verified.");
            bool isVerifyR90 = mutilChoiceProperites.Count == 1 && mutilChoiceProperites[0].Equals("Value", StringComparison.OrdinalIgnoreCase);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR90,
                90,
                @"[In Choice or Multi-Choice Field] The EntityType for this EntitySet [of multi-choice] as specified in [MC-CSDL] section 2.1.17 contains a single property ""Value"", which also serves as its [EntitySet's] EntityKey as specified in [MC-CSDL] section 2.1.5.");

            // If the entity type exist for document library, the MS-WSSREST_R94 can be verified 
            Site.Log.Add(LogEntryKind.Debug, "If the entity type exist for document library, the MS-WSSREST_R94 can be verified ");
            bool isVerifyR94 = this.IsExistEntityType(entityTypes, string.Format("{0}Item", Common.GetConfigurationPropertyValue("DoucmentLibraryListName", this.Site)));
            Site.CaptureRequirementIfIsTrue(
                isVerifyR94,
                94,
                @"[In Document] Document libraries are represented as EntityTypes that have an associated media resource, as specified in [RFC5023].");
        }
    }
}