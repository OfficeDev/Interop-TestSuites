namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test SyncFolderItems operation with/without optional elements in request on multiple items.
    /// </summary>
    [TestClass]
    public class S04_OperateSyncFolderItemsOptionalElements : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="testContext">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// Client calls SyncFolderItems operation without all optional elements to get the synchronization information of multiple items.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S04_TC01_SyncFolderItems_WithoutOptionalElements()
        {
            // Add drafts folder into list
            this.FolderIdNameType.Add(DistinguishedFolderIdNameType.drafts);
            SyncFolderItemsType[] request = new SyncFolderItemsType[this.FolderIdNameType.Count];

            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                request[i] = this.CreateSyncFolderItemsRequestWithoutOptionalElements(this.FolderIdNameType[i], DefaultShapeNamesType.AllProperties);
            }

            this.VerifySyncFolderItemsOperation(request, false);
        }

        /// <summary>
        /// Client calls SyncFolderItems operation with all optional elements to get the synchronization information of multiple items.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S04_TC02_SyncFolderItems_WithAllOptionalElements()
        {
            // Add drafts folder into list
            this.FolderIdNameType.Add(DistinguishedFolderIdNameType.drafts);
            SyncFolderItemsType[] request = new SyncFolderItemsType[this.FolderIdNameType.Count];

            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                request[i] = this.CreateSyncFolderItemsRequestWithoutOptionalElements(this.FolderIdNameType[i], DefaultShapeNamesType.IdOnly);

                if (Common.IsRequirementEnabled(347, this.Site))
                {
                    request[i].SyncScopeSpecified = true;
                    request[i].SyncScope = SyncFolderItemsScopeType.NormalItems;
                }

                // Set the value of SyncState element and Ignore element
                request[i].SyncState = string.Empty;

                // It will be rewrite to a new value in test case
                request[i].Ignore = null;

                // Set the value of AdditionalProperties element
                request[i].ItemShape.AdditionalProperties = new BasePathToElementType[] 
                { 
                    new PathToUnindexedFieldType
                    {
                        FieldURI = UnindexedFieldURIType.itemSubject
                    } 
                };

                if (Common.IsRequirementEnabled(37809, this.Site))
                {
                    request[i].ItemShape.InlineImageUrlTemplate = "Test Template";
                }

                // Configure the SOAP header to cover the case that the header contains all optional parts before calling operations.
                this.ConfigureSOAPHeader();
            }

            this.VerifySyncFolderItemsOperation(request, true);
        }
        #endregion
    }
}