namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System.Web.Services.Protocols;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the GetViewCollection operation.
    /// </summary>
    [TestClass]
    public class S02_GetAllViews : TestSuiteBase
    {
        #region Initialization and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }
        #endregion

        #region Test Cases
        /// <summary>
        /// A test case used to test GetViewCollection method successfully with valid parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S02_TC01_GetViewCollection_Success()
        {            
            // Call AddView method twice to add two views for the specified list on the server.           
            this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);
            this.AddView(false, Query.AvailableQueryInfo, ViewType.Calendar);

            // Call GetViewCollection to retrieve the collection of list views of a specified list in the server.
            string listName = TestSuiteBase.ListGUID;

            GetViewCollectionResponseGetViewCollectionResult getViewCollection = Adapter.GetViewCollection(listName);
            this.Site.Assert.IsNotNull(getViewCollection, "There SHOULD be a view collection returned from a successful GetViewCollection operation.");

            foreach (GetViewCollectionResponseGetViewCollectionResultView view in getViewCollection.Views)
            {
                // The existence of the attribute Name of the view indicates the attribute group of type ViewAttributeGroup exists, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    view.Name,
                    113,
                    @"[In GetViewCollectionResponse] It [the GetViewCollectionResult element] MUST include the collection of View elements of the specified list, which includes an attribute group of type ViewAttributeGroup.");

                // If the value of FrameState is set to "Normal" or "Minimized", then the following requirement can be captured.
                bool isValidated = view.FrameState == "Normal" || view.FrameState == "Minimized";
                Site.CaptureRequirementIfIsTrue(
                    isValidated,
                    "MS-WSSCAML",
                    810,
                    @"[In Attributes] [The attribute of ViewDefinition type]FrameState: MUST be set to ""Normal"" or ""Minimized"". ");
            }
        }

        /// <summary>
        /// A test case used to test GetViewCollection method with invalid listName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S02_TC02_GetViewCollection_InvalidListName()
        {
            bool caughtSoapException = false; 

            // Call GetViewCollection with invalid listName.
            try
            {
                string invalidListName = this.GenerateRandomString(10);
                Adapter.GetViewCollection(invalidListName);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If the server returns SOAP fault, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    13,
                    "[In listName] If the value of listName element is not the name or GUID of a list, the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }
        #endregion
    }
}