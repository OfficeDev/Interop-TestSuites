namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the following operations with valid or invalid parameters.
    /// <list type="bullet">
    ///     <item>AddList</item>
    ///     <item>DeleteList</item>
    ///     <item>GetList</item>
    ///     <item>UpdateList</item>
    ///     <item>GetListAndView</item>
    ///     <item>AddListFromFeature</item>
    /// </list>
    /// </summary>
    [TestClass]
    public class S01_OperationOnList : TestClassBase
    {
        #region Private member variables

        /// <summary>
        /// Protocol adapter
        /// </summary>
        private IMS_LISTSWSAdapter listswsAdapter;

        #endregion

        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        #region AddListFromFeature

        /// <summary>
        /// This test case is used to verify the successful status of AddListFromFeature operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC01_AddListFromFeature_Succeed()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            bool addListSucceeded = false;
            try
            {
                // Call method AddListFromFeature with valid parameters.
                int templateId = (int)TemplateType.Generic_List;
                AddListFromFeatureResponseAddListFromFeatureResult result = this.listswsAdapter.AddListFromFeature(
                                        listName,
                                        string.Empty,
                                        Common.GetConfigurationPropertyValue("ListFeatureId", this.Site),
                                        templateId);
                addListSucceeded = result != null && result.List != null && !string.IsNullOrEmpty(result.List.ID);

                Site.Log.Add(
                  LogEntryKind.Debug,
                  "The actual value: List.ID[{0}] for requirement #R366",
                  string.IsNullOrEmpty(result.List.ID) ? "NullOrEmpty" : result.List.ID);

                Site.CaptureRequirementIfIsTrue(
                    addListSucceeded,
                    366,
                    @"[In AddListFromFeature] If there are no other errors, a new list MUST be created on the site by using the listName, description, featureID, and templateID specified in the AddListSoapIn request message.");

                // Verify requirement R3521.
                // If there are no other errors, it means implementation does support this AddListFromFeature method. R3521 can be captured.
                if (Common.IsRequirementEnabled(3521, this.Site))
                {
                    Site.CaptureRequirementIfIsTrue(
                        addListSucceeded,
                        3521,
                        @"Implementation does support this method[AddListFromFeature]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
            finally
            {
                if (addListSucceeded)
                {
                    this.listswsAdapter.DeleteList(listName);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify the negative status of AddListFromFeature operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC02_AddListFromFeature_EmptyFeatureID()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            string newListName = TestSuiteHelper.GetUniqueListName();
            string errorCode = string.Empty;
            bool addListSucceeded = false;
            bool addUnmatchTemplateIDSucceeded = false;
            bool addFeatureSucceeded = false;
            bool addListWithExistentNameSucceeded = false;
            bool isExistentFaultFirst = false;
            bool isExistentFaultSecond = false;
            int templateId = (int)TemplateType.Generic_List;

            try
            {
                // Add a list with empty featureID.
                try
                {
                    string emptyFeatureID = string.Empty;
                    AddListFromFeatureResponseAddListFromFeatureResult result = this.listswsAdapter.AddListFromFeature(listName, string.Empty, emptyFeatureID, templateId);
                    addListSucceeded = result != null && result.List != null && !string.IsNullOrEmpty(result.List.ID);
                }
                catch (SoapException exp)
                {
                    errorCode = TestSuiteHelper.GetErrorCode(exp);
                    isExistentFaultFirst = true;
                }

                this.Site.Assert.IsTrue(isExistentFaultFirst, "The server response should contain the SOAP fault during the AddList operation when the feature id is empty.");

                Site.CaptureRequirementIfIsFalse(
                    addListSucceeded,
                    360,
                    @"[In AddListFromFeature operation] If the featureID tag is specified and the value is not a GUID or empty, the protocol server MUST return a SOAP fault.");

                Site.CaptureRequirementIfIsNull(
                    errorCode,
                    1604,
                    @"[In AddListFromFeature operation] [If the featureID tag is specified and the value is not a GUID or empty, the protocol server MUST return a SOAP fault.]There is no error code returned for this fault.");
                try
                {
                    // Add a list with unmatched featureID and templateID
                    int unmatchTemplateId = (int)TemplateType.Grid;
                    AddListFromFeatureResponseAddListFromFeatureResult addResult = this.listswsAdapter.AddListFromFeature(
                                        newListName,
                                        string.Empty,
                                        Common.GetConfigurationPropertyValue("ListFeatureId", this.Site),
                                        unmatchTemplateId);
                    addUnmatchTemplateIDSucceeded = addResult != null && addResult.List != null && !string.IsNullOrEmpty(addResult.List.ID);
                }
                catch (SoapException exp)
                {
                    errorCode = TestSuiteHelper.GetErrorCode(exp);
                    isExistentFaultSecond = true;
                }

                this.Site.Assert.IsTrue(isExistentFaultSecond, "The server response does not contain the SOAP fault during the AddList operation when the template id is not valid.");

                // If the error code is 0x81072101, then capture R1605 and R1606.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x81072101",
                    errorCode,
                    1605,
                    @"[AddListFromFeature]If the provided templateID cannot be used with the provided featureID, the protocol server MUST return a SOAP fault with error code 0x81072101");
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x81072101",
                    errorCode,
                    1606,
                    @"[AddListFromFeature] [If the provided templateID cannot be used with the provided featureID, the protocol server MUST return a SOAP fault with error code 0x81072101] This indicates that the SOAP protocol failed to add a list.");

                // Add a list with correct value.
                try
                {
                    AddListFromFeatureResponseAddListFromFeatureResult addFeatureListResult = this.listswsAdapter.AddListFromFeature(
                                        listName,
                                        null,
                                        Common.GetConfigurationPropertyValue("ListFeatureId", this.Site),
                                        templateId);
                    addFeatureSucceeded = addFeatureListResult != null && addFeatureListResult.List != null && !string.IsNullOrEmpty(addFeatureListResult.List.ID);
                }
                catch (SoapException)
                {
                    this.Site.Assert.Fail("Test suite should add the list successfully.");
                }

                bool isExistentFaultThird = false;

                try
                {
                    // Add a list with the listName that already exist.
                    AddListFromFeatureResponseAddListFromFeatureResult addFeatureResult = this.listswsAdapter.AddListFromFeature(listName, null, null, templateId);
                    addListWithExistentNameSucceeded = addFeatureResult != null && addFeatureResult.List != null && !string.IsNullOrEmpty(addFeatureResult.List.ID);
                }
                catch (SoapException exp)
                {
                    errorCode = TestSuiteHelper.GetErrorCode(exp);
                    isExistentFaultThird = true;
                }

                this.Site.Assert.IsTrue(isExistentFaultThird, "The server response should contain the SOAP fault when the listName is used by another list.");

                // If the error code is 0x81020012, then capture R365 and R1608.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x81020012",
                    errorCode,
                    365,
                    @"[In AddListFromFeature] If the listName is already used by another list then the protocol server MUST return a SOAP fault with error code 0x81020012.");
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x81020012",
                    errorCode,
                    1608,
                    @"[In AddListFromFeature] [If the listName is already used by another list then the protocol server MUST return a SOAP fault with error code 0x81020012.] This indicates that another list has the specified listName.");
            }
            finally
            {
                if (addListSucceeded || addFeatureSucceeded || addListWithExistentNameSucceeded)
                {
                    this.listswsAdapter.DeleteList(listName);
                }

                if (addUnmatchTemplateIDSucceeded)
                {
                    this.listswsAdapter.DeleteList(newListName);
                }
            }
        }

        /// <summary>
        /// This test case is used to validate the AddListFromFeature operation with invalid templateID whose value is less than 0.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC03_AddListFromFeature_InvalidTemplateId()
        {
            #region Add list with less than 0 Template ID, try to capture R357

            string listNameCustom = TestSuiteHelper.GetUniqueListName();
            bool isAddListFailed = false;
            string featureId = Common.GetConfigurationPropertyValue("ListFeatureId", this.Site);
            try
            {
                // Add list from feature using a less than 0 template id.
                this.listswsAdapter.AddListFromFeature(
                                            listNameCustom,
                                            null,
                                            featureId,
                                            (int)TemplateType.Invalid);
            }
            catch (SoapException exp)
            {
                string errorCode = TestSuiteHelper.GetErrorCode(exp);
                isAddListFailed = true;

                bool isR357Captured = string.IsNullOrEmpty(errorCode);

                Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorCode[{0}] for requirement #R357",
                        string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                // If the server return a SOAP fault without error code, capture R357.
                Site.CaptureRequirementIfIsTrue(
                    isR357Captured,
                    357,
                    @"[In AddListFromFeature operation] If the templateID provided is less than 0, the protocol server MUST return a SOAP fault. There is no error code for this fault.");
            }
            finally
            {
                if (!isAddListFailed)
                {
                    this.listswsAdapter.DeleteList(listNameCustom);
                }

                Site.Assert.IsTrue(
                    isAddListFailed,
                    "MSLISTSWS_S01_TC03_AddListFromFeature_InvalidTemplateId, AddListFromFeature operation should fail with TemplateId[{0}]!",
                    (int)TemplateType.Invalid);
            }
            #endregion
        }

        /// <summary>
        ///  This test case is used to validate the AddListFromFeature operation with the templateID which is not one of the known list template identifiers.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC04_AddListFromFeature_UnknownTemplateId()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(3581, this.Site) || Common.IsRequirementEnabled(2771, this.Site),
                @"Test is executed only when R3581Enabled is set to true or R2771Enabled is set to true.");

            string listNameCustom = TestSuiteHelper.GetUniqueListName();
            bool isAddListFailed = false;
            bool caughtSoapException = false;
            string featureId = Common.GetConfigurationPropertyValue("ListFeatureId", this.Site);
            string expectederrorString = "Parameter {0} is missing or invalid.";
            try
            {
                // Add list from feature using an unknown template id.
                this.listswsAdapter.AddListFromFeature(
                                            listNameCustom,
                                            null,
                                            featureId,
                                            (int)TemplateType.Unkown);
            }
            catch (SoapException ex)
            {
                caughtSoapException = true;
                string errorCode = TestSuiteHelper.GetErrorCode(ex);
                string errorString = TestSuiteHelper.GetErrorString(ex);
                isAddListFailed = true;

                if (Common.IsRequirementEnabled(3581, this.Site))
                {
                    // If the server return a SOAP fault with error string "Parameter {0} is missing or invalid.", capture R3581.
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorString[{0}] for requirement #R3581",
                        string.IsNullOrEmpty(errorString) ? "NullOrEmpty" : errorString);

                    Site.CaptureRequirementIfIsTrue(
                    expectederrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase),
                    3581,
                    @"[In AddListFromFeature operation] Implementation does return a SOAP fault with the error string ""Parameter {0} is missing or invalid"", if the templateID provided is not one of the known list template identifiers.(SharePoint Foundation 2010 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled(2771, this.Site))
                {
                    // If the server return a SOAP fault with error code 0x81072101, capture R2771.
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorCode[{0}] for requirement #R2771",
                        string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                    Site.CaptureRequirementIfAreEqual<string>(
                    "0x81072101",
                    errorCode,
                    2771,
                    @"[In Appendix B: Product Behavior][In AddListFromFeature operation] Implementation does returns a SOAP fault with error code 0x81072101, if the templateID provided is not one of the known list template identifiers . (<28> Section 3.1.4.4:  Windows SharePoint Services 3.0 returns a SOAP fault with error code 0x81072101.)");
                }
            }
            finally
            {
                Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'AddListFromFeature' with an unknown template ID.");

                if (!isAddListFailed)
                {
                    this.listswsAdapter.DeleteList(listNameCustom);
                }

                Site.Assert.IsTrue(
                    isAddListFailed,
                    "MSLISTSWS_S01_TC03_AddListFromFeature_InvalidTemplateId, AddListFromFeature operation should fail with TemplateId[{0}]!",
                    (int)TemplateType.Unkown);
            }
        }

        #endregion

        #region AddList

        /// <summary>
        /// This test case is used to verify the successful status of AddList operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC05_AddList_Succeed()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            bool addListSucceeded = false;
            try
            {
                // Add a list with valid parameters.
                int templateId = (int)TemplateType.Generic_List;
                AddListResponseAddListResult result = this.listswsAdapter.AddList(listName, string.Empty, templateId);
                addListSucceeded = result != null && result.List != null && !string.IsNullOrEmpty(result.List.ID);
                this.Site.Assert.IsTrue(addListSucceeded, "Test suite should add the list successfully.");

                Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: List.ID[{0}] for requirement #R347",
                    string.IsNullOrEmpty(result.List.ID) ? "NullOrEmpty" : result.List.ID);

                Site.CaptureRequirementIfIsTrue(
                    addListSucceeded,
                    347,
                    @"[In AddList operation] Otherwise [if the templateID provided matches a known template which can create a new list and the listName is not used by another list], a new list named listName MUST be created on the site, using the template with an identification matching the provided templateID and the list description MUST be the description passed in.");
            }
            finally
            {
                if (addListSucceeded)
                {
                    this.listswsAdapter.DeleteList(listName);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify the negative status of AddList operation with existing list name and unknown templateID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC06_AddList_AlreadyUsedListName_UnknownTemplateID()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            string errorCode = string.Empty;
            string newListName = string.Empty;
            bool addListSucceeded = false;
            bool addInvalidListSucceeded = false;
            bool isExistentFaultfirst = false;
            bool isExistentFaultForUnknownTemplateID = false;
            try
            {
                // Add a list with correct value.
                int templateId = (int)TemplateType.Generic_List;
                AddListResponseAddListResult result = this.listswsAdapter.AddList(listName, null, templateId);
                addListSucceeded = result != null && result.List != null && !string.IsNullOrEmpty(result.List.ID);
                this.Site.Assert.IsTrue(addListSucceeded, "Test suite should add the list successfully");
                try
                {
                    int newTemplateId = (int)TemplateType.Grid;

                    // Add a list with above listName
                    this.listswsAdapter.AddList(listName, null, newTemplateId);
                }
                catch (SoapException exp)
                {
                    errorCode = TestSuiteHelper.GetErrorCode(exp);
                    isExistentFaultfirst = true;
                }

                this.Site.Assert.IsTrue(isExistentFaultfirst, "The server response does not contain the SOAP fault during the AddList operation when the list name is used by another list.");

                // If the error code is 0x81020012, then capture R346 and R1589.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x81020012",
                    errorCode,
                    346,
                    @"[In AddList operation] If the listName is already used by another list then the protocol server MUST return a SOAP fault with error code 0x81020012.");

                Site.CaptureRequirementIfAreEqual<string>(
                    "0x81020012",
                    errorCode,
                    1589,
                    @"[In AddList operation] [If the listName is already used by another list then the protocol server MUST return a SOAP fault with error code 0x81020012.] This indicates that another list has the specified listName.");
                try
                {
                    // Add a list with an invalid template id.
                    int invalidTemplateId = (int)TemplateType.Unkown;
                    newListName = TestSuiteHelper.GetUniqueListName();
                    AddListResponseAddListResult resultInvalid = this.listswsAdapter.AddList(newListName, null, invalidTemplateId);
                    addInvalidListSucceeded = resultInvalid != null && resultInvalid.List != null && !string.IsNullOrEmpty(result.List.ID);
                }
                catch (SoapException exp)
                {
                    errorCode = TestSuiteHelper.GetErrorCode(exp);
                    isExistentFaultForUnknownTemplateID = true;
                }

                this.Site.Assert.IsTrue(isExistentFaultForUnknownTemplateID, "The server response should contain the SOAP fault during the AddList operation when template id is unknown.");

                // If the error code is 0x8102007b, then capture R342.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x8102007b",
                    errorCode,
                    342,
                    @"[In AddList operation] If the templateID provided is not one of the known list template identifiers, the protocol server MUST return a SOAP fault with error code 0x8102007b.This indicates that the list template is invalid.");
            }
            finally
            {
                if (addListSucceeded)
                {
                    this.listswsAdapter.DeleteList(listName);
                }

                if (addInvalidListSucceeded)
                {
                    this.listswsAdapter.DeleteList(newListName);
                }
            }
        }

        /// <summary>
        /// The test case is used to verify the AddList operation when the templateID is less than 0. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC07_AddList_Negative()
        {
            #region Add list with less than 0 Template ID, try to capture R341

            string listNameCustom = TestSuiteHelper.GetUniqueListName();
            bool isAddListFailed = false;
            string errorCode = null;
            try
            {
                // Add a list using less than 0 template id.
                this.listswsAdapter.AddList(listNameCustom, "Add list with less than 0 Template ID.", (int)TemplateType.Invalid);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isAddListFailed = true;

                // If the server return a SOAP fault without error code, capture R341.
                Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorCode[{0}] for requirement #R341 while the templateID is less than 0",
                        string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.CaptureRequirementIfIsTrue(
                    string.IsNullOrEmpty(errorCode),
                    341,
                    @"[In AddList operation] If the templateID provided is less than 0, the protocol server MUST return a SOAP fault. There is no error code for this fault.");
            }
            finally
            {
                if (!isAddListFailed)
                {
                    this.listswsAdapter.DeleteList(listNameCustom);
                }

                Site.Assert.IsTrue(
                            isAddListFailed,
                            "MSLISTSWS_S01_TC06_AddList_Negative, AddList operation should be fail with templateID[{0}]!",
                            (int)TemplateType.Invalid);
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to verify adding a list on the protocol server with a list based on the template specified by templateID that already exists, and the template is marked as unique in AddList operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC08_AddList_UniqueTemplate()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            string errorCode = string.Empty;
            bool addListSucceeded = false;
            bool isExistentFault = false;
            int templateId = (int)TemplateType.UserInfo;

            // Verify whether the "User_Information_List" template's list is existing in current protocol SUT.
            GetListCollectionResponseGetListCollectionResult getlistCollectionResult = this.listswsAdapter.GetListCollection();
            if (null == getlistCollectionResult || null == getlistCollectionResult.Lists
                || 0 == getlistCollectionResult.Lists.Length)
            {
                this.Site.Assert.Fail("Could not get the valid response of GetListCollection operation.");
            }

            // The "User_Information_List" template's value is 112, as described in [MS-WSSFO2] section 2.2.3.12.
            var userInfomatTemplateList = from listItem in getlistCollectionResult.Lists
                                          where listItem.ServerTemplate.Equals("112")
                                          select listItem;

            // If "User_Information_List" template's list is not existing in current protocol SUT, create one
            if (0 == userInfomatTemplateList.Count())
            {
                TestSuiteHelper.CreateList(templateId);
            }

            try
            {
                try
                {
                    // Call method AddList operation to add a list with unique template "User_Information_List" on the server.
                    AddListResponseAddListResult result = this.listswsAdapter.AddList(listName, string.Empty, templateId);
                    addListSucceeded = result != null && result.List != null && !string.IsNullOrEmpty(result.List.ID);
                }
                catch (SoapException exp)
                {
                    errorCode = TestSuiteHelper.GetErrorCode(exp);
                    isExistentFault = true;
                }

                this.Site.Assert.IsTrue(isExistentFault, "The server response should contain the SOAP fault during the AddList request operation when the based template id which is marked as unique has been already used by another list.");

                // There should be only one User Information List in the site, if an error code
                // 0x8102003c is returned, then the following requirements can be captured.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x8102003c",
                    errorCode,
                    344,
                    @"[In AddList operation] If the templateID provided matches a known template, but a list based on that template already exists and the template is marked as unique, the protocol server MUST return a SOAP fault with error code 0x8102003c.");
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x8102003c",
                    errorCode,
                    1587,
                    @"[In AddList operation] [ If the templateID provided matches a known template, but a list based on that template already exists and the template is marked as unique, the protocol server MUST return a SOAP fault with error code 0x8102003c.] This indicates that templateID provided is marked unique and that a list with the specified templateID already exists.");
            }
            finally
            {
                if (addListSucceeded)
                {
                    this.listswsAdapter.DeleteList(listName);
                }
            }
        }

        /// <summary>
        /// This test case is used to add a list on the protocol server with invalid password. The server will notify the client authorization faults using HTTP status codes.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC09_AuthorizationFaultsTest()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            string description = TestSuiteHelper.GenerateRandomString(10);
            bool issoapFaultGenerated = false;

            // Create list with valid credential first.
            string listGuid = TestSuiteHelper.CreateList(listName);
            Assert.IsNotNull(listGuid, "Create list with correct credential should be successfully.");

            // Set invalid credential.
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GenerateInvalidPassword(Common.GetConfigurationPropertyValue("Password", this.Site));
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);

            this.listswsAdapter.Credentials = new NetworkCredential(userName, password, domain);

            // Create list with invalid credential.
            try
            {
                this.listswsAdapter.AddList(listName, description, (int)TemplateType.Generic_List);
            }
            catch (SoapException)
            {
                this.Site.Log.Add(LogEntryKind.Debug, "Exception happened but was ignored. The exception will not impact the test case");
                issoapFaultGenerated = true;
            }
            catch (WebException)
            {
                this.Site.Log.Add(LogEntryKind.Debug, "Exception happened but was ignored. The exception will not impact the test case");
                issoapFaultGenerated = true;
            }
            finally
            {
                // Restore default credential.
                userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
                password = Common.GetConfigurationPropertyValue("Password", this.Site);
                domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
                this.listswsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            }

            if (!issoapFaultGenerated)
            {
                // when no exception threw,  this AuthorizationFaultsTest case fails.
                Site.Assert.Fail("Authorization Failed when using the user {0} to call operation AddList", userName);
            }
            else
            {
                // If threw exception, then capture R283.
                Site.CaptureRequirement(
                    283,
                    @"[This protocol allows protocol servers to perform implementation-specific authorization checks] and notify protocol clients of authorization faults either by using HTTP Status Codes or by using SOAP faults as specified previously in this section.");
            }
        }

        #endregion

        #region DeleteList

        /// <summary>
        /// This test case is used to verify the successful status of DeleteList operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC10_DeleteList_Succeed()
        {
            // Add a generic list.
            string listName = TestSuiteHelper.GetUniqueListName();
            int templateId = (int)TemplateType.Generic_List;
            AddListResponseAddListResult addResult = this.listswsAdapter.AddList(listName, string.Empty, templateId);
            bool addListSucceeded = addResult != null && addResult.List != null && !string.IsNullOrEmpty(addResult.List.ID);
            this.Site.Assert.IsTrue(addListSucceeded, "Test suite should add the list successfully.");

            // Delete the list.
            this.listswsAdapter.DeleteList(addResult.List.ID);
            XmlNode responseXmlNode = (XmlNode)SchemaValidation.LastRawResponseXml;
            string reponse = responseXmlNode.FirstChild.FirstChild.LocalName;

            // Call GetListCollection when there is not any list in the server. The output lists should not contain the added list information.
            GetListCollectionResponseGetListCollectionResult result = this.listswsAdapter.GetListCollection();
            bool isExist = result.Lists.Any(founder => founder.Title.Equals(listName, StringComparison.OrdinalIgnoreCase));

            // If the DeleteListResponse is returned ,then capture R520 and R521.
            Site.CaptureRequirementIfAreEqual<string>(
                "DeleteListResponse",
                reponse,
                520,
                @"[In DeleteList operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, that list MUST be deleted and the protocol server MUST return a DeleteListResponse element.");

            Site.CaptureRequirementIfAreEqual<string>(
                "DeleteListResponse",
                reponse,
                521,
                @"[In DeleteList operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, that list MUST be deleted and the protocol server MUST return a DeleteListResponse element.");

            // If the new added list exists in the server, then capture R524.
            Site.CaptureRequirementIfIsFalse(
                isExist,
                524,
                @"[In DeleteList operation] If there are no other errors, the list MUST be deleted.");
        }

        /// <summary>
        /// This test case is used to verify the negative status of DeleteList operation with a listName which does not correspond to any lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC11_DeleteList_NonExistentListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2757, this.Site), @"Test is executed only when R2757Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string errorCode = string.Empty;
            bool isExistentFault = false;

            // Delete a non-existent list
            try
            {
                this.listswsAdapter.DeleteList(listName);
            }
            catch (SoapException exp)
            {
                isExistentFault = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain the SOAP fault during the DeleteList operation when the specified list name is not exist.");

            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000006",
                errorCode,
                2757,
                @"[In DeleteList operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.) ");
        }

        /// <summary>
        /// This test case is used to verify the DeleteList operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC12_DeleteNonExistentList_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2448, this.Site), @"Test is executed only when R2448Enabled is set to true.");

            // Delete the list with a list name that is nonexistent.
            string invalidGuid = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            string soapFault = "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            bool issoapFaultGenerated = false;
            string errorString = string.Empty;
            string errorCode = string.Empty;
            try
            {
                // Delete the list
                this.listswsAdapter.DeleteList(invalidGuid);
            }
            catch (SoapException exp)
            {
                issoapFaultGenerated = true;
                errorString = TestSuiteHelper.GetErrorString(exp);
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }

            this.Site.Assert.IsTrue(issoapFaultGenerated, "There should be a Soap fault generated when call DeleteList operation with invalid Guid");
            this.Site.Assert.IsTrue(errorString.Equals(soapFault, StringComparison.OrdinalIgnoreCase), "The returned SoapFault should match with TD.");
            this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorCode), "The Soap exception contains error Code.");
            
            // If the protocol server returns the SOAP fault with no error code: 
            // "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)", then capture R2448.
            Site.CaptureRequirement(
                            2448,
                            @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<43> Section 3.1.4.13: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
        }

        #endregion

        #region GetListAndView

        /// <summary>
        /// This test case is used to verify the successful status of GetListAndView operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC13_GetListAndView_Succeed()
        {
            // Add a generic list with correct value.
            string listName = TestSuiteHelper.GetUniqueListName();
            bool addListSucceeded = false;

            try
            {
                // Add a generic list.
                int templateId = (int)TemplateType.Generic_List;
                AddListResponseAddListResult addResult = this.listswsAdapter.AddList(listName, string.Empty, templateId);
                addListSucceeded = addResult != null && addResult.List != null && !string.IsNullOrEmpty(addResult.List.ID);
                this.Site.Assert.IsTrue(addListSucceeded, "Test suite should add the list successfully.");

                // Set list validation enabled.
                UpdateListListProperties properties = new UpdateListListProperties();
                properties.List = new UpdateListListPropertiesList();
                properties.List.Validation = new UpdateListListPropertiesListValidation();
                properties.List.Validation.Message = TestSuiteHelper.GenerateRandomString(1024);
                this.listswsAdapter.UpdateList(addResult.List.Name, properties, null, null, null, null);

                // Call method GetListAndView with correct list name and set the viewName parameter to null. 
                GetListAndViewResponseGetListAndViewResult getResultValid = this.listswsAdapter.GetListAndView(addResult.List.Name, null);

                // If the GetListAndViewResponse element is returned, then capture R568.
                Site.CaptureRequirementIfIsNotNull(
                    getResultValid,
                    568,
                    @"[In GetListAndView operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, the protocol server MUST return a GetListAndViewResponse element.");

                // If the default view properties is returned, then capture R1777.
                Site.CaptureRequirementIfIsTrue(
                    bool.Parse(getResultValid.ListAndView.View.DefaultView),
                    1777,
                    @"[GetListAndView]If the specified viewName is not specified[ or is an empty string], the default view properties MUST be returned.");

                // Call method GetListAndView with correct list title which is an invalid GUID.
                GetListAndViewResponseGetListAndViewResult getResultInvalid = this.listswsAdapter.GetListAndView(addResult.List.Title, null);

                // If the GetListAndViewReponse element is returned, then capture R570.
                Site.CaptureRequirementIfIsNotNull(
                    getResultInvalid,
                    570,
                    @"[In GetListAndView operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, the protocol server MUST return a GetListAndViewReponse element.");

                // If the GetListAndViewReponse element is returned, then capture R569.
                Site.CaptureRequirementIfIsNotNull(
                    getResultInvalid,
                    569,
                    @"[In GetListAndView operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, the protocol server MUST return a GetListAndViewReponse element.");

                // Call method GetListAndView by specifying a viewName with an empty string.
                GetListAndViewResponseGetListAndViewResult getResultEmptyViewName = this.listswsAdapter.GetListAndView(addResult.List.Name, string.Empty);
                this.Site.Assert.IsNotNull(getResultEmptyViewName, "GetListAndView operation failed.");

                // If the default view properties is returned, then capture R2257.
                Site.CaptureRequirementIfIsTrue(
                    bool.Parse(getResultEmptyViewName.ListAndView.View.DefaultView),
                    2257,
                    @"[GetListAndView]If the specified viewName [is not specified or] is an empty string, the default view properties MUST be returned.");
            }
            finally
            {
                if (addListSucceeded)
                {
                    this.listswsAdapter.DeleteList(listName);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify the negative status of GetListAndView operation with invalid viewName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC14_GetListAndView_InvalidGUID()
        {
            // Add a generic list.
            string listGuid = TestSuiteHelper.CreateList();
            bool createListSucceeded = !string.IsNullOrEmpty(listGuid);
            this.Site.Assert.IsTrue(createListSucceeded, "Test suite should add the list successfully.");
            string errorCode = string.Empty;
            string invalidViewName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            bool isExistentFault = false;

            // Call method GetListAndView with an invalid viewName
            try
            {
                this.listswsAdapter.GetListAndView(listGuid, invalidViewName);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault while requesting GetListAndView with invalid view name.");

            // If error code error code 0x82000001 is returned, then capture R3014 and R3015.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000001",
                errorCode,
                3014,
                @"[GetListAndView] If the specified viewName is not a valid GUID and is not an empty string, the protocol server MUST return a SOAP fault with error code 0x82000001.");

            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000001",
                errorCode,
                3015,
                @"[GetListAndView] [If the specified viewName is not a valid GUID and is not an empty string, the protocol server MUST return a SOAP fault with error code 0x82000001.] This indicates that a required parameter is missing or invalid.");
        }

        /// <summary>
        /// This test case is used to validate the GetListAndView operations with invalid listName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC15_GetListAndView_InvalidListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2760, this.Site), @"Test is executed only when R2760Enabled is set to true.");

            string invalidListName = TestSuiteHelper.GetUniqueListName();
            string viewName = Guid.NewGuid().ToString();
            bool caughtSoapException = false;
            string errorCode = null;
            try
            {
                // Get list and view by using invalid list name.
                this.listswsAdapter.GetListAndView(invalidListName, viewName);
            }
            catch (SoapException exp)
            {
                caughtSoapException = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }

            Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'GetListAndView' with invalid list name.");

            Site.CaptureRequirementIfAreEqual(
                    "0x82000006",
                    errorCode,
                    2760,
                    @"[In GetListAndView operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.SharePoint Foundation 2010 and above follow this behavior.) ");
        }

        /// <summary>
        /// This test case is used to validate the GetListAndView operations with invalid viewName but valid list name.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC16_GetListAndView_InvalidViewName()
        {
            #region Get a list and view using a valid list, but nonexistent view name in the corresponding list, try capture R1778

            string listGuid = TestSuiteHelper.CreateList();
            bool isGetListAndViewFail = false;
            string errorCode = null;
            string invaildViewName = Guid.NewGuid().ToString();
            try
            {
                // Get list and view using a nonexistent view name.
                this.listswsAdapter.GetListAndView(listGuid, invaildViewName);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isGetListAndViewFail = true;
            }

            Site.Assert.IsTrue(
                isGetListAndViewFail,
                "MSLISTSWS_S01_TC15_GetListAndView_InvalidViewName, GetListAndView operation is failed!");

            // If  the protocol server returns a SOAP fault with error code 0x82000005, capture R1778.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000005",
                errorCode,
                1778,
                @"[GetListAndView]If the specified viewName does not correspond to an existing viewName for the given list and it is not an empty string, the protocol server MUST return a SOAP fault with error code 0x82000005.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify the GetListAndView operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC17_GetListAndView_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2454, this.Site), @"Test is executed only when R2454Enabled is set to true.");

            // Add a list
            string soapFault = "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            bool isFault = false;
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            try
            {
                // Delete the list which is nonexistent.
                this.listswsAdapter.GetListAndView(invalidListName, null);
            }
            catch (SoapException exp)
            {
                isFault = true;
                string errorString = TestSuiteHelper.GetErrorString(exp);
                string errorCode = TestSuiteHelper.GetErrorCode(exp);
                this.Site.Assert.IsTrue(errorString.Equals(soapFault, StringComparison.OrdinalIgnoreCase), "The returned SoapFault does not match with TD.");
                this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorCode), "The Soap exception contains error Code.");

                // If the protocol server returns the SOAP fault with no error code: 
                // "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)", then capture R2454.
                Site.CaptureRequirement(
                                    2454,
                                    @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<47> Section 3.1.4.16: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            }

            this.Site.Assert.IsTrue(isFault, "The operation doesn't catch any Soap Fault messages.");
        }

        #endregion

        #region GetListCollection

        /// <summary>
        /// This test case is used to verify the successful status of GetListCollection operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC18_GetListCollection_Succeed()
        {
            string listGuid = TestSuiteHelper.CreateList();
            bool createListSucceeded = !string.IsNullOrEmpty(listGuid);
            this.Site.Assert.IsTrue(createListSucceeded, "Test suite should add the list successfully.");

            // If the operation is successful, the lists will contain valid list information.
            GetListCollectionResponseGetListCollectionResult result = this.listswsAdapter.GetListCollection();
            bool iscontianAddList = result.Lists.Any(founder => founder.ID.Equals(listGuid, StringComparison.OrdinalIgnoreCase));
            this.Site.Assert.IsTrue(iscontianAddList, " List set must contain a valid list when create list succeeded.");
        }

        #region GetList

        /// <summary>
        /// This test case is used to verify the successful status of GetList operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC19_GetList_Succeed()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            bool addListSucceeded = false;
            IMS_LISTSWSSUTControlAdapter sutControlAdapter = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();

            // Call GetListItems to retrieve details about "User Information List" items
            string userInfoListName = Common.GetConfigurationPropertyValue("UserInfoListName", this.Site);
            GetListItemsQuery query = new GetListItemsQuery();
            query.Query = new CamlQueryRoot();
            query.Query.OrderBy = new OrderByDefinition();

            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            getListItemsResult = this.listswsAdapter.GetListItems(
                userInfoListName,
                null,
                query,
                null,
                null,
                null,
                null);

            // extract a row element to a DataTable, it is used for "z:row" data
            System.Data.DataTable data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

            int listItemsCount = data.Rows.Count;
            Dictionary<string, string> userList = new Dictionary<string, string>();
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "Name");

            for (int i = 0; i < listItemsCount; i++)
            {
                userList.Add(Convert.ToString(data.Rows[i][AdapterHelper.PrefixOws + AdapterHelper.FieldIDName]), Convert.ToString(data.Rows[i][columnName]));
            }

            try
            {
                // Add a Generic list.
                int templateId = (int)TemplateType.Generic_List;
                AddListResponseAddListResult addResult = this.listswsAdapter.AddList(listName, string.Empty, templateId);
                addListSucceeded = addResult != null && addResult.List != null && !string.IsNullOrEmpty(addResult.List.ID);
                this.Site.Assert.IsTrue(addListSucceeded, "Test suite should add the list successfully.");

                // Get Presence and RecycleBinEnable value of WebApp.
                bool presenceEnabled = sutControlAdapter.GetWebAppPresence();
                bool recycleBinEnable = sutControlAdapter.GetWebAppRecycleBin();

                // Set Presence and RecycleBinEnable value to True.
                sutControlAdapter.SetWebAppPresence(true);
                System.Threading.Thread.Sleep(10000);
                sutControlAdapter.SetWebAppRecycleBin(true);
                System.Threading.Thread.Sleep(10000);

                // Call method GetList to get the list from server and check Presence and RecycleBinEnable value.
                GetListResponseGetListResult getResult = this.listswsAdapter.GetList(addResult.List.ID);

                // Verify R1381
                bool isContainInUserList = userList.ContainsKey(getResult.List.Author);

                Site.Assert.IsTrue(isContainInUserList, "The user identifier doesn't contain in the user information list.");

                string currentUser = string.Empty;
                string currentUserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site);

                if (currentUserDomain.Contains("."))
                {
                    currentUser = currentUserDomain.Substring(0, currentUserDomain.IndexOf(".", StringComparison.OrdinalIgnoreCase)) + "\\" + Common.GetConfigurationPropertyValue("UserName", this.Site);
                }
                else
                {
                    currentUser = currentUserDomain + "\\" + Common.GetConfigurationPropertyValue("UserName", this.Site);
                }

                // On O15 Farm, if the identity of the application pool is that of a particular administrator or if the administrator logs in as the farm account
                // or the application pool account, then SharePoint will show that person as System Account(SHAREPOINT\system).
                // So if the list author is System Account, it is the create user and the isCreateUser variable should be "True".
                bool isCreateUser = currentUser.Equals(userList[getResult.List.Author], StringComparison.CurrentCultureIgnoreCase)
                    || currentUser.Equals(userList[getResult.List.Author].Substring(userList[getResult.List.Author].IndexOf('|') + 1), StringComparison.CurrentCultureIgnoreCase)
                    || userList[getResult.List.Author].Equals("SHAREPOINT\\system", StringComparison.CurrentCultureIgnoreCase);

                Site.CaptureRequirementIfIsTrue(
                    isContainInUserList && isCreateUser,
                    1381,
                    @"[ListDefinitionCT.Author:] The user identifier of the user who created the list, which is contained in the user information list.");

                // If the value of Presence equals to the value set by sutControlAdapter(true), capture R1427.
                Site.CaptureRequirementIfIsTrue(
                    bool.Parse(getResult.List.RegionalSettings.Presence),
                    1427,
                    @"[ListDefinitionSchema.Presence:] Specifies that presence is enabled if set to True; [otherwise Presence is not enabled.]");

                // If the value of RecycleBinEnabled equals to the value set by sutControlAdapter(true), capture R1429.
                Site.CaptureRequirementIfIsTrue(
                    bool.Parse(getResult.List.ServerSettings.RecycleBinEnabled),
                    1429,
                    @"[ListDefinitionSchema.RecycleBinEnabled: ]Specifies that the Recycle Bin is enabled if set to True;[ otherwise, the Recycle Bin is not enabled.]");

                // Set Presence and RecycleBinEnable value to True.
                sutControlAdapter.SetWebAppPresence(false);
                System.Threading.Thread.Sleep(10000);
                sutControlAdapter.SetWebAppRecycleBin(false);
                System.Threading.Thread.Sleep(10000);

                // Call method GetList to get the list from server and check Presence and RecycleBinEnable value.
                getResult = this.listswsAdapter.GetList(addResult.List.ID);

                // If the value of Presence equals to the value set by sutControlAdapter(false), capture R2249.
                Site.CaptureRequirementIfIsFalse(
                    bool.Parse(getResult.List.RegionalSettings.Presence),
                    2249,
                    @"ListDefinitionSchema.Presence: [Specifies that presence is enabled if set to True; ]otherwise Presence is not enabled.");

                // If the value of RecycleBinEnabled equals to the value set by sutControlAdapter(false), capture R2248.
                Site.CaptureRequirementIfIsFalse(
                    bool.Parse(getResult.List.ServerSettings.RecycleBinEnabled),
                    2248,
                    @"ListDefinitionSchema.RecycleBinEnabled: [Specifies that the Recycle Bin is enabled if set to True;] otherwise, the Recycle Bin is not enabled.");

                // Restore Presence and RecycleBinEnable settings for WebApp.
                sutControlAdapter.SetWebAppPresence(presenceEnabled);
                System.Threading.Thread.Sleep(10000);
                sutControlAdapter.SetWebAppRecycleBin(recycleBinEnable);
                System.Threading.Thread.Sleep(10000);

                // Call method GetList to get the list from server.
                // If the GetListResponse is returned, then capture R555 and R557.
                getResult = this.listswsAdapter.GetList(addResult.List.ID);
                Site.CaptureRequirementIfIsNotNull(
                    getResult,
                    555,
                    @"[In GetList operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, the protocol server MUST return a GetListResponse element.");

                GetListResponseGetListResult getTitleResult = this.listswsAdapter.GetList(addResult.List.Title);
                Site.CaptureRequirementIfIsNotNull(
                    getTitleResult,
                    557,
                    @"[In GetList operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site, and if so the protocol server MUST return a GetListResponse element.");

                // Call method GetList with ListName does not have a valid GUID.
                // If the GetListResponse is returned, then capture R556.
                GetListResponseGetListResult getResultInvalidName = this.listswsAdapter.GetList(listName);
                Site.CaptureRequirementIfIsNotNull(
                    getResultInvalidName,
                    556,
                    @"[In GetList operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site, and if so the protocol server MUST return a GetListResponse element.");
            }
            finally
            {
                if (addListSucceeded)
                {
                    this.listswsAdapter.DeleteList(listName);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify the negative status of GetList operation with a listName which does not correspond to any lists. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC20_GetList_NonExistentListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2759, this.Site), @"Test is executed only when R2759Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string errorCode = string.Empty;
            bool isExistentFault = false;

            // Call method GetList with a non-exist listName
            try
            {
                this.listswsAdapter.GetList(listName);
            }
            catch (SoapException exp)
            {
                isExistentFault = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain the SOAP fault during the GetList operation when the list name is not exist.");

            // If the error code is 0x82000006, then capture R2759.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000006",
                errorCode,
                2759,
                @"[In GetList operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.) ");
        }

        /// <summary>
        /// This test case is used to verify the GetList operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC21_GetList_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2452, this.Site), @"Test is executed only when R2452Enabled is set to true.");

            // Add a list
            string soapFault = "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            bool caughtSoapException = false;
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            try
            {
                // Get the list with a list name that is nonexistent.
                this.listswsAdapter.GetList(invalidListName);
            }
            catch (SoapException exp)
            {
                caughtSoapException = true;
                string errorString = TestSuiteHelper.GetErrorString(exp);
                string errorCode = TestSuiteHelper.GetErrorCode(exp);
                this.Site.Assert.IsTrue(errorString.Equals(soapFault, StringComparison.OrdinalIgnoreCase), "The returned SoapFault does not match with TD.");
                this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorCode), "The Soap exception contains error Code.");

                // If the protocol server returns the SOAP fault with no error code: 
                // "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)", then capture R2452.
                Site.CaptureRequirement(
                                    2452,
                                    @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<46> Section 3.1.4.15: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'GetList' with a list name that is nonexistent.");
        }

        #endregion

        #endregion

        #region UpdateList

        /// <summary>
        /// This test case is used to verify the ErrorCode element in the response of UpdateList operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC22_UpdateList_FailureErrorCodeInUpdateListFieldResults()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            // This field does not exist in current list, and construct a method in request to delete this field.
            string nonExistentField = Guid.NewGuid().ToString("N");
            UpdateListFieldsRequest deleteFields = TestSuiteHelper.CreateDeleteListFieldsRequest(new List<string> { nonExistentField });

            this.Site.Assert.IsNotNull(deleteFields, "CreateDeleteListFieldsRequest operation should succeed.");
            UpdateListResponseUpdateListResult updateListResult = null;
            updateListResult = this.listswsAdapter.UpdateList(listName, null, null, null, deleteFields, null);

            if (null == updateListResult || null == updateListResult.Results || null == updateListResult.Results.DeleteFields
                || updateListResult.Results.DeleteFields.Length != 1)
            {
                this.Site.Assert.Fail("Could not get the proper response from UpdateList operation when performing field deletion.");
            }

            // If the ErrorCode element in the server response is not equal to 0x00000000, then the following requirement can be captured.
            Site.CaptureRequirementIfAreNotEqual<string>(
                "0x00000000",
                updateListResult.Results.DeleteFields[0].ErrorCode,
                2250,
                @"[Success is indicated by 0x00000000 and] failure is indicated by any other value.");
        }

        /// <summary>
        /// This test case is used to test operation UpdateList 
        /// when listName element is invalid value (Not valid List GUID, Not valid List Title).
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC23_UpdateList_InvalidListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(8861, this.Site), @"Test is executed only when R8861Enabled is set to true.");

            #region  Invoke AddList operation to create a new generic list.
            string listId = TestSuiteHelper.CreateList();
            #endregion

            #region Invoke GetList operation to get the current list version of the generic list.
            string listVersion = string.Empty;
            GetListResponseGetListResult getListResult = null;
            getListResult = this.listswsAdapter.GetList(listId);
            Site.Assert.IsNotNull(getListResult, "The object \"getListResult\" is null!");
            Site.Assert.IsNotNull(getListResult.List, "The object \"getListResult.List\" is null!");
            listVersion = getListResult.List.Version.ToString();
            #endregion

            #region Invoke UpdateList with correct value of list version, and set the value of listName element to invalid value (Not valid List GUID, Not valid List Title)
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Capture requirements #886 and #2039 if we get a SOAP fault with error code 0x82000006 
            // in response of UpdateList operation.
            bool isSoapFaultGenerated = false;
            string errorcode = string.Empty;
            try
            {
                this.listswsAdapter.UpdateList(invalidListName, null, null, null, null, listVersion);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultGenerated = true;
                errorcode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            #region Capture requirements #8861 when we get a SOAP fault with error code 0x82000006 in response of UpdateList operation.

            this.Site.Log.Add(
                            LogEntryKind.Debug,
                            "The actual value: isSoapFaultGenerated[{0}], error code:[{0}] for requirement #R8861",
                            isSoapFaultGenerated,
                            string.IsNullOrEmpty(errorcode) ? "NullOrEmpty" : errorcode);

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated && "0x82000006".Equals(errorcode, StringComparison.OrdinalIgnoreCase),
                8861,
                @"[In UpdateList operation] Implementation does return  a SOAP fault with error code 0x82000006 if listName does not correspond to a list from either of these checks.(In Microsoft® Office 2010 suites/Microsoft® SharePoint® Foundation 2010 and above follow this behavior.)");

            #endregion
        }

        /// <summary>
        /// This test case is used to test operation UpdateList 
        /// when the value of element listVersion cannot be converted to an integer 
        /// by product server.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC24_UpdateList_InvalidListVersion_CanNotBeConvertedToInteger()
        {
            #region  Invoke AddList operation to create a new generic list.
            string listId = TestSuiteHelper.CreateList();
            #endregion

            #region Invoke UpdateList with the value of element listVersion that cannot be converted to an integer.
            string invalidIntegerValue = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            bool isSoapFaultGenerated = false;
            string errorCode = string.Empty;

            // Capture requirements #889 and #2040 if we get a SOAP fault in the response of UpdateList.
            try
            {
                this.listswsAdapter.UpdateList(listId, null, null, null, null, invalidIntegerValue);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultGenerated = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: isSoapFaultGenerated[{0}], errorCode[{1}] for requirement #R889 #R2040",
                        isSoapFaultGenerated,
                        string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

            #region Capture requirements #889 when we get a SOAP fault in the response of UpdateList.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated && string.IsNullOrEmpty(errorCode),
                889,
                @"[In UpdateList operation] If the listVersion string cannot be converted to an integer, the protocol server "
                + "MUST return a SOAP fault.");
            #endregion

            #region Capture requirements #2040 when we get a SOAP fault without error code.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated && string.IsNullOrEmpty(errorCode),
                2040,
                @"[In UpdateList operation] [If the listVersion string cannot be converted to an integer, the protocol server "
                + "MUST return a SOAP fault.]There is no error code for this fault.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test operation UpdateList 
        /// when the value of element listVersion string is numeric 
        /// but does not match the version of the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC25_UpdateList_InvalidListVersion_MismatchedNumeric()
        {
            #region  Invoke AddList operation to create a new generic list.
            string listId = TestSuiteHelper.CreateList();
            #endregion

            #region Invoke GetList operation to get the current list version of the generic list.
            GetListResponseGetListResult getListResult = null;
            getListResult = this.listswsAdapter.GetList(listId);
            Site.Assert.IsNotNull(getListResult, "The object \"getListResult\" is null!");
            Site.Assert.IsNotNull(getListResult.List, "The object \"getListResult.List\" is null!");
            int listVersion = getListResult.List.Version;
            #endregion

            #region Invoke UpdateList with the value of element listVersion that is numeric, but does not match the version of the list.
            string mismatchedListVersion = (listVersion + 1).ToString();
            bool isSoapFaultGenerated = false;
            string errorCode = string.Empty;

            // Capture requirements #890 if we get a SOAP fault with error code 0x81020015 in the response of UpdateList.
            try
            {
                this.listswsAdapter.UpdateList(listId, null, null, null, null, mismatchedListVersion);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultGenerated = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            this.Site.Log.Add(
                            LogEntryKind.Debug,
                            "The actual value: isSoapFaultGenerated[{0}], errorCode[{1}] for requirement #R891 #R2042 #899 #2060",
                            isSoapFaultGenerated,
                            string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

            #region Capture requirements #890 when we get a SOAP fault with error code 0x81020015 in the response of UpdateList.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated && "0x81020015".Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                891,
                @"[In UpdateList operation] If the listVersion does not match the version of the list, "
                + "the protocol server MUST return a SOAP fault with error code 0x81020015.");
            Site.CaptureRequirementIfIsTrue(
                 isSoapFaultGenerated && "0x81020015".Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                2042,
                @"[In UpdateList operation] [If the listVersion does not match the version of the list, "
                + "the protocol server MUST return a SOAP fault with error code 0x81020015.] This indicates "
                + "that the list changes are in conflict with those made by another user.");
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated && "0x81020015".Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                899,
                @"[In UpdateList operation] [In UpdateList element] [In listVersion element] If the listVersion value "
                + "sent up in the UpdateList request does not match the current version then the protocol server "
                + "MUST respond with a SOAP fault with error code 0x81020015.");
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated && "0x81020015".Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                2060,
                @"[In UpdateList operation] [In UpdateList element] [In listVersion element] [If the listVersion value "
                + "sent up in the UpdateList request does not match the current version then the protocol server "
                + "MUST respond with a SOAP fault with error code 0x81020015.] This indicates that the changes being "
                + "requested via UpdateList will conflict with whatever changes have already been made to this list that "
                + "resulted in the protocol server listVersion being higher than the one sent by the protocol client.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test operation UpdateList 
        /// when the value of element listVersion string is numeric 
        /// but not within the range of an unsigned 32-bit integer.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC26_UpdateList_InvalidListVersion_OutOfUInt32()
        {
            #region  Invoke AddList operation to create a new generic list.
            string listId = TestSuiteHelper.CreateList();
            #endregion

            #region Invoke UpdateList with the value of element listVersion that is numeric, but not within the range of an unsigned 32-bit integer.
            string outOfUInt32RangValue = (ulong.MaxValue - 1).ToString();
            bool isSoapFaultGenerated = false;
            string errorCode = string.Empty;

            // Capture requirements #890 if we get a SOAP fault in the response of UpdateList.
            // Capture requirements #2041 if we get a SOAP fault without error code in the response of UpdateList.
            try
            {
                this.listswsAdapter.UpdateList(listId, null, null, null, null, outOfUInt32RangValue);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultGenerated = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            this.Site.Log.Add(
                       LogEntryKind.Debug,
                       "The actual value: isSoapFaultGenerated[{0}], errorCode[{1}] for requirement #R890 #R2041",
                       isSoapFaultGenerated,
                       string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

            #region Capture requirements #890 when we get a SOAP fault in the response of UpdateList.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated && string.IsNullOrEmpty(errorCode),
                890,
                @"[In UpdateList operation] If the listVersion string is numeric but not within the range "
                + "of an unsigned 32-bit integer, UpdateList MUST return a SOAP fault.");
            #endregion

            #region Capture requirements #2041 when we get a SOAP fault without error code in the response of UpdateList.
            Site.CaptureRequirementIfIsTrue(
                 isSoapFaultGenerated && string.IsNullOrEmpty(errorCode),
                2041,
                @"[In UpdateList operation] [If the listVersion string is numeric but not within the range of "
                + "an unsigned 32-bit integer, UpdateList MUST return a SOAP fault.]There is no error code for this fault.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the UpdateList operation with the Method.AddToView element in the UpdateListFieldsRequest.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC27_UpdateList_InvalidViewNameInUpdateListFieldResults()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            bool isNotGUIDFieldInView = false;
            bool isEmptyStringFieldInView = false;
            bool isNullFieldInView = false;
            bool isDefaultView = false;
            
            List<string> fieldNames = new List<string> { };
            List<string> fieldTypes = new List<string> { };
            List<string> viewNames = new List<string> { };

            string fieldNameOfNotGUIDAddToView = TestSuiteHelper.GenerateRandomString(4);
            string fieldNameOfEmptyAddToView = TestSuiteHelper.GenerateRandomString(6);
            string fieldNameOfNullAddToView = TestSuiteHelper.GenerateRandomString(8);
 
            // New field with Method.AddToView element is not a GUID.
            // The new fields will be added to default view, the UpdateList operation with attribute AddToView is not a GUID
            // does not support Windows SharePoint Services 3.0 and SharePointServer2007.
            if (Common.IsRequirementEnabled(267001, this.Site))
            {
                fieldNames.Add(fieldNameOfNotGUIDAddToView);
                fieldTypes.Add("Text");
                viewNames.Add(TestSuiteHelper.GetInvalidGuidAndNocorrespondString());
            }
            else
            {
                this.Site.Log.Add(LogEntryKind.Debug, "Current protocol SUT does not support below behaviors: [Support UpdateList operation and assume the default view if the Method.AddToView attribute present and its value is not a GUID.]");
            }

            // New field with Method.AddToView element is an empty string.
            // The new fields will be added to default view.
            fieldNames.Add(fieldNameOfEmptyAddToView);
            fieldTypes.Add("Text");
            viewNames.Add(string.Empty);

            // New field with Method.AddToView element is not presented.
            // The new fields does not be added to any view.
            // Due to there is only default view in the list, the field does not be added into default view.
            fieldNames.Add(fieldNameOfNullAddToView);
            fieldTypes.Add("Text");
            viewNames.Add(null);

            UpdateListFieldsRequest newFields = TestSuiteHelper.CreateAddListFieldsRequest(fieldNames, fieldTypes, viewNames);
            this.listswsAdapter.UpdateList(listName, null, newFields, null, null, null);

            GetListAndViewResponseGetListAndViewResult getResultValid = this.listswsAdapter.GetListAndView(listName, string.Empty);
            isDefaultView = bool.Parse(getResultValid.ListAndView.View.DefaultView);
            this.Site.Assert.IsTrue(isDefaultView, "The response of GetListAndView operation should contain the default view when the viewName parameter is not set or empty value.");
            FieldRefDefinitionView[] viewFields = getResultValid.ListAndView.View.ViewFields;

            if (Common.IsRequirementEnabled(267001, this.Site))
            {
                // Verify whether a field with AddToView attribute is not a GUID in default view.
                isNotGUIDFieldInView = viewFields.Any(fieldFounder => fieldNameOfNotGUIDAddToView.Equals(fieldFounder.Name, StringComparison.OrdinalIgnoreCase));
                this.Site.CaptureRequirementIfIsTrue(
                                                    isNotGUIDFieldInView,
                                                    267001,
                                                    @"[In Appendix B: Product Behavior] Implementation does support this method[UpdateList] and assume the default view will be used, if the Method.AddToView attribute present and its value is not a GUID.(SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Capture R265.
            isEmptyStringFieldInView = viewFields.Any(fieldFounder => fieldNameOfEmptyAddToView.Equals(fieldFounder.Name, StringComparison.OrdinalIgnoreCase));
            Site.CaptureRequirementIfIsTrue(
                isEmptyStringFieldInView & isDefaultView,
                265,
                @"[In UpdateListFieldsRequest] [Method.AddToView] This is an optional parameter, and the protocol server MUST assume the default view if the value of the parameter is an empty string.");

            // Due to there is only default view in the list. If field does not be added into default view, then capture R266.
            isNullFieldInView = viewFields.Any(fieldFounder => fieldNameOfNullAddToView.Equals(fieldFounder.Name, StringComparison.OrdinalIgnoreCase));
            Site.CaptureRequirementIfIsFalse(
                isNullFieldInView,
                266,
                @"[In UpdateListFieldsRequest] [Method.AddToView] If the parameter is not presented, the protocol server MUST NOT add the fields to any view.");
        }

        /// <summary>
        /// This test case is used to verify the ListDefinitionCT complex type in UpdateList operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC28_UpdateList_ListDefinitionCT()
        {
            IMS_LISTSWSSUTControlAdapter sutControlAdapter = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();

            // Create a list.
            string listName = TestSuiteHelper.GetUniqueListItemName();
            string listId = TestSuiteHelper.CreateList(listName);

            #region Verify RootFolder in the list

            GetListResponseGetListResult getResultForNewList = this.listswsAdapter.GetList(listName);

            string listRootFolder = sutControlAdapter.GetListRootFolder(listName);

            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"^\d+$");

            int listFirstVersion = getResultForNewList.List.Version;
            bool isMajorVersion = regex.IsMatch(listFirstVersion.ToString());

            Site.Assert.IsTrue(
                isMajorVersion,
                "The numeric major revision of the list is {0}.",
                listFirstVersion);

            Site.Assert.AreEqual<string>(
                listRootFolder,
                getResultForNewList.List.RootFolder,
                "The RootFolder of the list is {0}.",
                getResultForNewList.List.RootFolder);

            // Verify R1375
            Site.CaptureRequirement(
                1375,
                @"[ListDefinitionCT.RootFolder:] The root folder of the list.");

            #endregion

            // Construct a UpdateListListProperties instance.
            UpdateListListProperties properties = new UpdateListListProperties();
            properties.List = new UpdateListListPropertiesList();

            // Generate a 10 length random string 
            properties.List.Description = TestSuiteHelper.GenerateRandomString(10);
            properties.List.Hidden = "True";
            properties.List.EnableVersioning = "True";
            properties.List.Ordered = "True";

            // Call UpdateListResponseUpdateListResult operation to update list properties.
            UpdateListResponseUpdateListResult updateListResult = null;
            updateListResult = this.listswsAdapter.UpdateList(
                                                              listId,
                                                              properties,
                                                              null,
                                                              null,
                                                              null,
                                                              null);

            Site.Assert.IsNotNull(updateListResult, "UpdateList operation succeeded.");

            // Verify R1364
            isMajorVersion = regex.IsMatch(updateListResult.Results.ListProperties.Version.ToString());
            bool isNumeric = updateListResult.Results.ListProperties.Version.Equals(listFirstVersion + 1);
            Site.CaptureRequirementIfIsTrue(
                    isMajorVersion && isNumeric,
                    1364,
                    @"[ListDefinitionCT.Version: ]The numeric major revision of the list.");

            // Verify R1353
            Site.CaptureRequirementIfAreEqual<string>(
                updateListResult.Results.ListProperties.Title,
                listName,
                1353,
                @"[ListDefinitionCT.Title:] The display name of the list.");

            // Verify R1354
            Site.CaptureRequirementIfAreEqual<string>(
                updateListResult.Results.ListProperties.Description,
                properties.List.Description,
                1354,
                @"[ListDefinitionCT.Description:] The description of the list.");

            #region Capture Requirements R159

            Site.CaptureRequirementIfIsTrue(
                 Convert.ToBoolean(updateListResult.Results.ListProperties.ShowUser),
                 159,
                 "[ListDefinitionCT.ShowUser] True if this list is a survey list and user names are included in responses.");

            #endregion

            #region Capture requirement R156, R157, R158

            Site.CaptureRequirementIfIsTrue(
                Convert.ToBoolean(updateListResult.Results.ListProperties.EnableVersioning),
                156,
                "[ListDefinitionCT.EnableVersioning] True if this list is a document library and version control is enabled.");

            Site.CaptureRequirementIfIsTrue(
                Convert.ToBoolean(updateListResult.Results.ListProperties.Hidden),
                157,
                "[ListDefinitionCT.Hidden] True if this list is hidden.");

            Site.CaptureRequirementIfIsTrue(
            Convert.ToBoolean(updateListResult.Results.ListProperties.Hidden),
            158,
            "[ListDefinitionCT.Ordered] True if list items can be explicitly re-ordered.");

            #endregion

            // Create a document library.
            listId = TestSuiteHelper.CreateList(Convert.ToInt32(TemplateType.Document_Library));

            // Set Custom Send To Destination Name and Url for document library.
            sutControlAdapter.SetSendToNameAndUrl(listId, AdapterHelper.SendToDestinationName, AdapterHelper.SendToDestinationUrl);

            // Enable the versioning of the list.
            bool isSetVersionLimitSuccess = sutControlAdapter.SetVersionLimit(listId, AdapterHelper.MajorVersionLimitValue, AdapterHelper.MajorWithMinorVersionsLimitValue);
            Site.Assert.IsTrue(
                isSetVersionLimitSuccess,
                "SetVersioning operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isSetVersionLimitSuccess);

            // Call UpdateListResponseUpdateListResult operation to update list properties.
            updateListResult = this.listswsAdapter.UpdateList(
                                                            listId,
                                                              null,
                                                              null,
                                                              null,
                                                              null,
                                                              null);
            Site.Assert.IsNotNull(updateListResult, "UpdateList operation succeeded.");
            Site.Assert.IsNotNull(updateListResult.Results.ListProperties.SendToLocation, "The SendToLocation property should not be null when SendToDestinationName and SendToDestinationUrl is set.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1390
            // If the SendToLocation can be split to two string, then the following 
            // requirement can be captured.
            string[] strSendTo = updateListResult.Results.ListProperties.SendToLocation.Split('|');

            Site.CaptureRequirementIfAreEqual<int>(
                2,
                strSendTo.Length,
                1390,
                @"[ListDefinitionCT.SendToLocation:]These two items are returned as a string with "
                + @"a '|' character in between them.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1393
            // If the MajorVersionLimit can be parsed to a int value, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                AdapterHelper.MajorVersionLimitValue,
                updateListResult.Results.ListProperties.MajorVersionLimit,
                1393,
                @"[ListDefinitionCT.MajorVersionLimit: ]The maximum number of major versions allowed for a document in a document library that uses version control with major versions only.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1394
            // If the MajorWithMinorVersionsLimit can be parsed to a int value, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                AdapterHelper.MajorWithMinorVersionsLimitValue,
                updateListResult.Results.ListProperties.MajorWithMinorVersionsLimit,
                1394,
                @"[ListDefinitionCT.MajorWithMinorVersionsLimit:] The maximum number of major versions that are allowed for a document in a document library that uses version control with both major versions and minor versions.");

            #region Capture Requirements R1357

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)BaseType.Document_Library).ToString(),
                 updateListResult.Results.ListProperties.BaseType,
                 1357,
                 "[ListDefinitionCT.BaseType:] The base type of the list");

            #endregion

            #region Capture Requirements R1360

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)TemplateType.Document_Library).ToString(),
                 updateListResult.Results.ListProperties.ServerTemplate,
                 1360,
                 "[ListDefinitionCT.ServerTemplate:] The value[of ListDefinitionCT.ServerTemplate] corresponding to the template that the list is based on.");

            #endregion

            // Create an issues list.   
            listId = TestSuiteHelper.CreateList(Convert.ToInt32(TemplateType.Issues));

            // Call UpdateListResponseUpdateListResult operation to update list properties.
            updateListResult = this.listswsAdapter.UpdateList(
                                                            listId,
                                                              null,
                                                              null,
                                                              null,
                                                              null,
                                                              null);
            Site.Assert.IsNotNull(updateListResult, "UpdateList operation succeeded.");

            #region Capture Requirements R1357

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)BaseType.Issues).ToString(),
                 updateListResult.Results.ListProperties.BaseType,
                 1357,
                 "[ListDefinitionCT.BaseType:] The base type of the list");

            #endregion

            #region Capture Requirements R1360

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)TemplateType.Issues).ToString(),
                 updateListResult.Results.ListProperties.ServerTemplate,
                 1360,
                 "[ListDefinitionCT.ServerTemplate:] The value[of ListDefinitionCT.ServerTemplate] corresponding to the template that the list is based on.");

            #endregion

            #region Capture Requirement R5641

            Site.CaptureRequirementIfIsNotNull(
                updateListResult.Results.ListProperties.EnableAssignedToEmail,
                5641,
                "[ListDefinitionCT.EnableAssignedToEmail] This attribute is present if the current list is an issues list.");

            #endregion

            listId = TestSuiteHelper.CreateList(Convert.ToInt32(TemplateType.Generic_List));

            // Call UpdateListResponseUpdateListResult operation to update list properties.
            updateListResult = this.listswsAdapter.UpdateList(
                                                            listId,
                                                              null,
                                                              null,
                                                              null,
                                                              null,
                                                              null);
            Site.Assert.IsNotNull(updateListResult, "UpdateList operation succeeded.");

            #region Capture Requirements R1357

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)BaseType.Generic_List).ToString(),
                 updateListResult.Results.ListProperties.BaseType,
                 1357,
                 "[ListDefinitionCT.BaseType:] The base type of the list");

            #endregion

            #region Capture Requirements R1360

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)TemplateType.Generic_List).ToString(),
                 updateListResult.Results.ListProperties.ServerTemplate,
                 1360,
                 "[ListDefinitionCT.ServerTemplate:] The value[of ListDefinitionCT.ServerTemplate] corresponding to the template that the list is based on.");

            #endregion

            #region Capture Requirement R5642

            Site.CaptureRequirementIfIsNull(
                updateListResult.Results.ListProperties.EnableAssignedToEmail,
                5642,
                "[ListDefinitionCT.EnableAssignedToEmail] This attribute is not present if the current list is not an issues list.");

            #endregion

            // Create a Survey list
            listId = TestSuiteHelper.CreateList(Convert.ToInt32(TemplateType.Survey));

            // Call UpdateListResponseUpdateListResult operation to update list properties.
            updateListResult = this.listswsAdapter.UpdateList(
                                                            listId,
                                                              null,
                                                              null,
                                                              null,
                                                              null,
                                                              null);
            Site.Assert.IsNotNull(updateListResult, "UpdateList operation succeeded.");

            #region Capture Requirements R1357

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)BaseType.Survey).ToString(),
                 updateListResult.Results.ListProperties.BaseType,
                 1357,
                 "[ListDefinitionCT.BaseType:] The base type of the list");

            #endregion

            #region Capture Requirements R1360

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)TemplateType.Survey).ToString(),
                 updateListResult.Results.ListProperties.ServerTemplate,
                 1360,
                 "[ListDefinitionCT.ServerTemplate:] The value[of ListDefinitionCT.ServerTemplate] corresponding to the template that the list is based on.");

            #endregion

            // Create a Discussion board list
            listId = TestSuiteHelper.CreateList(Convert.ToInt32(TemplateType.Discussion_Board));

            // Call UpdateListResponseUpdateListResult operation to update list properties.
            updateListResult = this.listswsAdapter.UpdateList(
                                                            listId,
                                                              null,
                                                              null,
                                                              null,
                                                              null,
                                                              null);
            Site.Assert.IsNotNull(updateListResult, "UpdateList operation succeeded.");

            #region Capture Requirements R1357

            Site.CaptureRequirementIfAreEqual<string>(
                ((int)BaseType.Generic_List).ToString(),
                updateListResult.Results.ListProperties.BaseType,
                1357,
                "[ListDefinitionCT.BaseType:] The base type of the list");

            #endregion

            #region Capture Requirements R1360

            Site.CaptureRequirementIfAreEqual<string>(
                 ((int)TemplateType.Discussion_Board).ToString(),
                 updateListResult.Results.ListProperties.ServerTemplate,
                 1360,
                 "[ListDefinitionCT.ServerTemplate:] The value[of ListDefinitionCT.ServerTemplate] corresponding to the template that the list is based on.");

            #endregion
        }

        /// <summary>
        /// This test case is used to validate the successful status of UpdateList operation when the parameters are separately set to different values.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC29_UpdateList_Succeed()
        {
            #region Add a generic list

            string listName = TestSuiteHelper.GetUniqueListName();
            int templateListID = (int)TemplateType.Generic_List;
            string listGuid = TestSuiteHelper.CreateList(listName, templateListID);

            #endregion

            #region Call method UpdateList with correct listname. If succeeds, then capture R884

            UpdateListResponseUpdateListResult updateListResult = this.listswsAdapter.UpdateList(listGuid, null, null, null, null, null);
            bool isListUpdateSuccessfully = updateListResult != null && updateListResult.Results.ListProperties != null && !string.IsNullOrEmpty(updateListResult.Results.ListProperties.ID);

            // If UpdateList operation succeeds, then capture R884.
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: updateListResult[{0}],ListProperties[{1}], ListProperties.ID[{2}] for Requirement #R884",
                    updateListResult != null ? "NotNull" : "Null",
                    updateListResult.Results.ListProperties != null ? "NotNull" : "Null",
                    string.IsNullOrEmpty(updateListResult.Results.ListProperties.ID) ? "NullOrEmpty" : updateListResult.Results.ListProperties.ID);

            Site.CaptureRequirementIfIsTrue(
                isListUpdateSuccessfully,
                884,
                "[In UpdateList operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            // Add a field to the list.
            List<string> fieldNames = new List<string> { TestSuiteHelper.GenerateRandomString(5) };
            List<string> fieldTypes = new List<string> { "Text" };
            UpdateListFieldsRequest newFields = TestSuiteHelper.CreateAddListFieldsRequest(fieldNames, fieldTypes, new List<string> { null });
            UpdateListResponseUpdateListResult result = this.listswsAdapter.UpdateList(listGuid, null, newFields, null, null, null);

            // If the ErrorCode equals "0x00000000", then capture R1449.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000000",
                result.Results.NewFields[0].ErrorCode,
                1449,
                @"[ErrorCode]Success is indicated by 0x00000000 [and failure is indicated by any other value.]");

            #endregion

            #region Call method UpdateList with non-GUID listname if success capture R885

            // Update list with non-GUID list name
            updateListResult = this.listswsAdapter.UpdateList(listName, null, null, null, null, null);
            isListUpdateSuccessfully = updateListResult != null && updateListResult.Results.ListProperties != null && !string.IsNullOrEmpty(updateListResult.Results.ListProperties.ID);

            // If UpdateList operation success, capture R885
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value:updateListResult.Results.ListProperties.ID[{0}],expectedValue[{1}] for requirement #R885",
                    string.IsNullOrEmpty(updateListResult.Results.ListProperties.ID) ? "NullOrEmpty" : updateListResult.Results.ListProperties.ID,
                    listGuid);

            Site.CaptureRequirementIfIsTrue(
            listGuid.Equals(updateListResult.Results.ListProperties.ID, StringComparison.OrdinalIgnoreCase),
                    885,
                    "[In UpdateList operation] If the specified listName is not a valid GUID or does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            #endregion

            #region Call method UpdateList with listVersion which is set to 'null'. Requirements covered in this step:R887, R888

            // Update list with listVersion set to null.
            updateListResult = this.listswsAdapter.UpdateList(listName, null, null, null, null, null);
            isListUpdateSuccessfully = updateListResult != null && updateListResult.Results.ListProperties != null && !string.IsNullOrEmpty(updateListResult.Results.ListProperties.ID);

            // If UpdateList operation succeeds, then capture R887 and R888.
            Site.Log.Add(
                   LogEntryKind.Debug,
                   "The actual value: updateListResult[{0}],ListProperties[{1}], ListProperties.ID[{2}] for Requirement #R887 and #R888",
                   updateListResult != null ? "NotNull" : "Null",
                   updateListResult.Results.ListProperties != null ? "NotNull" : "Null",
                   string.IsNullOrEmpty(updateListResult.Results.ListProperties.ID) ? "NullOrEmpty" : updateListResult.Results.ListProperties.ID);

            Site.CaptureRequirementIfIsTrue(
                 isListUpdateSuccessfully,
                 887,
                 @"[In UpdateList operation] If the listVersion is null, the list MUST be updated.");

            Site.CaptureRequirementIfIsNotNull(
                updateListResult,
                888,
                @"[In UpdateList operation] If the listVersion is null, the protocol server MUST return an UpdateListReponse element.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify the UpdateList operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S01_TC30_UpdateList_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2504, this.Site), @"Test is executed only when R2504Enabled is set to true.");

            // Add a list
            string soapFault = "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            bool isFault = false;
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            try
            {
                // Update the list which is nonexistent
                this.listswsAdapter.UpdateList(invalidListName, null, null, null, null, null);
            }
            catch (SoapException exp)
            {
                isFault = true;

                string errorString = TestSuiteHelper.GetErrorString(exp);
                string errorCode = TestSuiteHelper.GetErrorCode(exp);
                this.Site.Assert.IsTrue(errorString.Equals(soapFault, StringComparison.OrdinalIgnoreCase), "The returned SoapFault does not match with TD.");
                this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorCode), "The Soap exception contains error Code.");

                // If the protocol server returns the SOAP fault with no error code: 
                // "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)", then capture R2504
                Site.CaptureRequirement(
                                2504,
                                @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<78> Section 3.1.4.30: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            }

            this.Site.Assert.IsTrue(isFault, "The operation doesn't catch any Soap Fault messages.");
        }

        #endregion

        #endregion

        #region protected override method
        /// <summary>
        /// This method will run before test case executes
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.listswsAdapter = this.Site.GetAdapter<IMS_LISTSWSAdapter>();
            Common.CheckCommonProperties(this.Site, true);

            #region new initialization
            if (!TestSuiteHelper.GuardEnviromentClean())
            {
                Site.Debug.Fail("The test environment is not clean, refer the log files for details.");
            }

            // Initialize the TestSuiteHelper
            TestSuiteHelper.Initialize(this.Site, this.listswsAdapter);
            #endregion
        }

        /// <summary>
        /// This method will run after test case executes
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            #region new clean up
            TestSuiteHelper.CleanUp();
            #endregion
        }

        #endregion
    }
}