namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.Collections;
    using System.Net;
    using System.Text;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    
    /// <summary>
    /// This test class is used to verify all test cases of scenario 01 about CopyIntoItems operation and GetItem operation.
    /// </summary>
    [TestClass]
    public class S01_CopyIntoItems : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void TestClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static void TestClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }

        /// <summary>
        /// This method will run before test case executes.
        /// </summary>
        [TestInitialize]
        public void CopyIntoItemsTestCaseInitialize()
        {
            // Verify the source computer is configurated correctly, it is determined by property "SourceSutComputerName". If the property value is null or empty, taht means the computer is not configurated correctly.
            string sourceComputerName = Common.GetConfigurationPropertyValue("SourceSutComputerName", this.Site);
            if (string.IsNullOrEmpty(sourceComputerName))
            {
                this.Site.Assume.Inconclusive(@"This test case runs only when the property ""SourceSutComputerName"" property in ""MS-COPYS_TestSuite.deployment.ptfconfig"" file is not set to empty.");
            }
 
            // Set the target service location to the source SUT for CopyIntoItems test cases.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.SourceSUT);
        }

        #endregion

        /// <summary>
        /// This test case is used to verify the GetItem operation sequence.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC01_GetItem_Success()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);
            
            string fileContents = Encoding.UTF8.GetString(getitemsResponse.StreamRawValues);

            // Verify MS-COPYS requirement: MS-COPYS_R165
            this.Site.CaptureRequirementIfAreEqual(
                Common.GetConfigurationPropertyValue("SourceFileContents", this.Site),
                fileContents,
                165,
                @"[In Message Processing Events and Sequencing Rules] GetItem: Retrieves a file and metadata for that file from the source location.");

            // Verify MS-COPYS requirement: MS-COPYS_R167
            this.Site.CaptureRequirementIfAreEqual(
                Common.GetConfigurationPropertyValue("SourceFileContents", this.Site),
                fileContents,
                167,
                @"[In GetItem] The GetItem operation retrieves content and metadata for a file that is stored in a source location.");

            // Verify MS-COPYS requirement: MS-COPYS_R178
            this.Site.CaptureRequirementIfAreEqual(
               Common.GetConfigurationPropertyValue("SourceFileContents", this.Site),
               fileContents,
                178,
                @"[In GetItem] [The protocol server returns results based on the following conditions:] If the source location 
                points to an existing file in the source location and the file can be read based on the permission settings for
                the file, the source location MUST return the content and metadata of the file.");

            if (getitemsResponse.Fields.Length > 0)
            {
                // Verify MS-COPYS requirement: MS-COPYS_R197
                // If the Fields element is present, the Stream element is present as well;
                this.Site.CaptureRequirementIfIsNotNull(
                    getitemsResponse.Stream,
                    197,
                    @"[In GetItemResponse] [Fields] If the Fields element is present, the Stream element MUST be present as well.");
            }
        }

        /// <summary>
        /// This test case is used to verify the GetItems operation when the source location is an invalid URL.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC02_GetItem_Fail()
        {
            if (!Common.IsRequirementEnabled(1048, this.Site))
            {
                this.Site.Assume.Inconclusive(@"The test case is only executed when the SHOULDMAY switch ""R1048Enabled"" is set to true.");
            }

            // Generate invalid file URL by construct a unique and random string.
            string invalidSourceFileUrl = Common.GenerateResourceName(this.Site, "URL");
            bool isCoughtException = false;

            try
            {
                // Retrieve content and metadata for a file that is stored in a source location with an invalid URL format string.
                GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(invalidSourceFileUrl);
            }
            catch (SoapException soapExp)
            {
                this.Site.Log.Add(LogEntryKind.Debug, "When the source location is not a valid URL format string, the server returns SoapException:{0}", soapExp.Message);
                isCoughtException = true;
            }

            // Verify MS-COPYS requirement: MS-COPYS_R1048
            this.Site.CaptureRequirementIfIsTrue(
                isCoughtException,
                1048,
                @"[In Appendix B: Product Behavior] Implementation does return a SOAP fault if the URL parameter is an invalid URI format string. (Microsoft SharePoint Foundation 2010 and Microsoft SharePoint Foundation 2013 follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to verify if there have some permission settings for the server or the file at the source
        /// location, the protocol server MUST report a failure by using HTTP Status-Code 401 Unauthorized.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC03_GetItem_NoPermission()
        {
            // Change the user which is no permission.
            MSCopysAdapter.SwitchUser(
                                      Common.GetConfigurationPropertyValue("MSCOPYSNoPermissionUser", this.Site),
                                      Common.GetConfigurationPropertyValue("PasswordOfNoPermissionUser", this.Site),
                                      Common.GetConfigurationPropertyValue("Domain", this.Site));

            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            int statusCode = 0;

            bool isCoughtException = false;
            try
            {
                // Retrieve content and metadata for a file that is stored in a source location.
                GetItemResponse getitemsResponseChangeUser = MSCopysAdapter.GetItem(sourceFileUrl);
            }
            catch (WebException webExp)
            {
                HttpWebResponse errorResponse = webExp.Response as HttpWebResponse;
                statusCode = (int)errorResponse.StatusCode;
                isCoughtException = true;
            }

            this.Site.Assert.IsTrue(isCoughtException, "Server should return Exception when the user has no permission to read the file at the source location.");

            // Verify MS-COPYS requirement: MS-COPYS_R177
            this.Site.CaptureRequirementIfAreEqual(
                401,
                statusCode,
                177,
                @"[In GetItem] [The protocol server returns results based on the following conditions:] If the protocol server
                implements permission settings for files and the file at the source location cannot be read based on the permission
                settings for the file, the protocol server MUST report a failure by using HTTP Status-Code 401 Unauthorized, as defined in [RFC2616].");
       
            // Verify MS-COPYS requirement: MS-COPYS_R8
            this.Site.CaptureRequirementIfAreEqual(
                401,
                statusCode,
                8,
                @"[In Transport] Protocol server faults MUST be returned by using either HTTP Status-Codes, as specified in 
                [RFC2616] section 10"); 
        }

        /// <summary>
        /// This test case is used to verify the InternalName attribute values are unique.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC04_GetItem_InternalNameUnique()
        {
             // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Verify MS-COPYS requirement: MS-COPYS_R73
            bool isVerifiedR73 = false;
            
            ArrayList internalNameList = new ArrayList();

            for (int i = 0; i < getitemsResponse.Fields.Length; i++)
            {
                if (internalNameList.Contains(getitemsResponse.Fields[i].InternalName))
                {
                    isVerifiedR73 = false;
                    break;
                }
                else
                {
                    internalNameList.Add(getitemsResponse.Fields[i].InternalName);
                }
            }

            if (getitemsResponse.Fields.Length == internalNameList.Count)
            {
                isVerifiedR73 = true;
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR73,
                73,
                @"[In FieldInformationCollection] The InternalName attribute values MUST be unique  for a sample of N 
                (default N=10) across all FieldInformation elements in the collection.");
        }

        /// <summary>
        /// This test case is used to verify the Id attribute values are unique.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC05_GetItem_IdUnique()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Verify MS-COPYS requirement: MS-COPYS_R75
            bool isVerifiedR75 = false;

            ArrayList idlist = new ArrayList();

            for (int i = 0; i < getitemsResponse.Fields.Length; i++)
            {
                if (idlist.Contains(getitemsResponse.Fields[i].Id))
                {
                    isVerifiedR75 = false;
                    break;
                }
                else
                {
                    idlist.Add(getitemsResponse.Fields[i].Id);
                }
            }

            if (getitemsResponse.Fields.Length == idlist.Count)
            {
                isVerifiedR75 = true;
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR75,
                75,
                @"[In FieldInformationCollection] The Id attribute values MUST be unique for a sample of N (default N=10) 
                across FieldInformation elements in the collection.");
        }

        /// <summary>
        /// This test case is used to verify if the CopyIntoItems operation executes successfully, error code "Success" should be returned.  
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC06_CopyIntoItems_ErrorCodeForSuccess()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results, "The element Results should be return if CopyIntoItems operation executes successfully");
            this.Site.Assert.IsTrue(2 == copyIntoItemsResponse.Results.Length, "The Results element should contain two CopyResult elements.");

            // If the error code is equal to Success when call CopyIntoItems operation with correct configuration, then R207 and R87 should be covered.
            this.Site.Assert.AreEqual(CopyErrorCode.Success, copyIntoItemsResponse.Results[0].ErrorCode, "CopyIntoItems operation should succeed.");
            this.Site.Assert.AreEqual(CopyErrorCode.Success, copyIntoItemsResponse.Results[1].ErrorCode, "CopyIntoItems operation should succeed.");
            
            // Verify MS-COPYS requirement: MS-COPYS_R207
            this.Site.CaptureRequirement(
                207,
                @"[In CopyIntoItems] The protocol server MUST report the status of the operation inside the Results collection 
                (see section 3.1.4.2.2.2) for each destination location that is passed.");

            // Verify MS-COPYS requirement: MS-COPYS_R87
            this.Site.CaptureRequirement(
                87,
                @"[In CopyErrorCode] Success: This value is used when the CopyIntoItems operation succeeds for the specified 
                destination location.");

            // Verify MS-COPYS requirement: MS-COPYS_R50
            bool isVerifiedR50 = string.IsNullOrEmpty(copyIntoItemsResponse.Results[0].ErrorMessage);
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR50,
                50,
                @"[In CopyResult] [ErrorMessage] [For CopyIntoItems operation] If the value of ErrorCode is ""Success,"" 
                the attribute MUST NOT be present.");
        }

        /// <summary>
        /// This test case is used to verify "Unknown" error code for CopyIntoItems operation when one of fields is not invalid value.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC07_CopyIntoItems_UnknownForInvalidField()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1 };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // WorkflowVersion field is an integer field, it must be returned for any document type, and any document library. Refer [MS-WSSTS]
            FieldInformation workflowVersionField = this.SelectFieldBySpecifiedAtrribute(getitemsResponse.Fields, "WorkflowVersion", FieldAttributeType.DisplayName);

            // Change WorkflowVersion field to an invalid value.
            workflowVersionField.Value = "invalidFieldValue";

            // Copy a file to the destination server with invalid field value.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);
            
            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results, "The element Results should be return if CopyIntoItems operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsResponse.Results.Length, "The Results element should contain one CopyResult element.");

            if (CopyErrorCode.Success == copyIntoItemsResponse.Results[0].ErrorCode)
            {
                // Collect files from specified file URLs.
                this.CollectFileByUrl(desUrls);
            }

            // The error code is equal to Unknown and the ErrorMessage is not null, then R210, R53 and R110 should be covered.
            this.Site.Assert.AreEqual(CopyErrorCode.Unknown, copyIntoItemsResponse.Results[0].ErrorCode, "Should return Unknown error code when the field value is invalid for CopyIntoItems operation.");
            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(copyIntoItemsResponse.Results[0].ErrorMessage), "ErrorMessage should not be empty when the error code is Unknown.");

            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results, "The element Results should be return if CopyIntoItems operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsResponse.Results.Length, "The Results element should contain one CopyResult element.");

            // Verify MS-COPYS requirement: MS-COPYS_R210          
            this.Site.CaptureRequirement(
                210,
                @"[In CopyIntoItems] If the protocol server detects a Value attribute that is not one of the valid values of 
                the FieldInformation elements of the Fields collection (see section 3.1.4.2.2.1), the protocol server MUST 
                report a failure of the operation by setting the ErrorCode attribute to ""Unknown"" for each CopyResult element,
                and provide a string value that denotes the error in the ErrorMessage attribute.");
        
            // Verify MS-COPYS requirement: MS-COPYS_R53
            this.Site.CaptureRequirement(
                53,
                @"[In CopyResult] [ErrorMessage] [For CopyIntoItems operation] Otherwise[If the value of ErrorCode is not ""Success"" ], the ErrorMessage attribute MUST be present and the value MUST be a non-empty Unicode string.");
        
            // Verify MS-COPYS requirement: MS-COPYS_R110
            this.Site.CaptureRequirement(
                110,
                @"[In CopyErrorCode] [For CopyIntoItems operation] Unknown This value is used to indicate an error for all other error conditions for a given destination location.");
        }

        /// <summary>
        /// This test case is used to verify "DestinationCheckedOut" error code for CopyIntoItems operation.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC08_CopyIntoItems_DestinationCheckedOut()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Upload a txt file to the destination SUT.
            this.UploadTxtFileByFileUrl(desFileUrl1);

            string[] desUrls = new string[] { desFileUrl1 };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Check out the file by the specified user.
            string checkOutUser = Common.GetConfigurationPropertyValue("MSCOPYSCheckOutUserName", this.Site);
            string passwordOfCheckoutUser = Common.GetConfigurationPropertyValue("PasswordOfCheckOutUser", this.Site);
            string domainValue = Common.GetConfigurationPropertyValue("Domain", this.Site);
            bool checkOutResult = MSCOPYSSutControlAdapter.CheckOutFileByUser(
                                                        desFileUrl1,
                                                        checkOutUser,
                                                        passwordOfCheckoutUser,
                                                        domainValue);

            this.Site.Assert.IsTrue(
                                     checkOutResult,
                                    "The check out action for the file[{0}] should be successful by using credentials: User[{1}] password[{2}] domain[{3}]",
                                    desFileUrl1,
                                    checkOutUser,
                                    passwordOfCheckoutUser,
                                    domainValue);

            CopyIntoItemsResponse copyIntoItemsResponse;
            try
            {
                // Copy a file to the destination server.
                copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                sourceFileUrl,
                                                                desUrls,
                                                                getitemsResponse.Fields,
                                                                getitemsResponse.StreamRawValues);
            }
            finally
            {
                // Undo checkout for a file by specified user credential.
                MSCOPYSSutControlAdapter.UndoCheckOutFileByUser(
                                                              desFileUrl1,
                                                              checkOutUser,
                                                              passwordOfCheckoutUser,
                                                              domainValue);

                // Collect a file by specified file URL. 
                this.CollectFileByUrl(desFileUrl1);
            }
 
            // Verify MS-COPYS requirement: MS-COPYS_R219
            this.Site.Assert.AreEqual(
                                     CopyErrorCode.DestinationCheckedOut, 
                                     copyIntoItemsResponse.Results[0].ErrorCode,
                                     "The CopyIntoItems operation should fail and return DestinationCheckedOut error when the file on destination is checked out.");
            
            // Verify MS-COPYS requirement: MS-COPYS_R102           
            this.Site.CaptureRequirement(
                102,
                @"[In CopyErrorCode] [For CopyIntoItems operation] DestinationCheckedOut This value is used to indicate an 
                error when the file on the destination location is checked out and cannot be overridden.");

            bool isContainErrorString = !string.IsNullOrEmpty(copyIntoItemsResponse.Results[0].ErrorMessage);  
            this.Site.CaptureRequirementIfIsTrue(
                isContainErrorString,
                219,
                @"[In CopyIntoItems] If the file on the protocol server is checked out and cannot be updated, the protocol 
                server MUST report a failure of the copy operation by setting the value of the ErrorCode attribute of the 
                corresponding CopyResult element to ""DestinationCheckedOut"", and provide a string value that specifies 
                the error in the ErrorMessage attribute.");

            // Verify MS-COPYS requirement: MS-COPYS_R154            
            this.Site.CaptureRequirementIfIsTrue(
                isContainErrorString,
                154,
                @"[In Abstract Data Model] In this case[files as checked out], the CopyIntoItems operations take into account 
                the checked-out status when accessing files at the destination locations.");
        }

        /// <summary>
        /// This test case is used to verify "DestinationMWS" error code for CopyIntoItems operation.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC09_CopyIntoItems_DestinationMWS()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrlMWS = this.GetDestinationFileUrl(DestinationFileUrlType.MWSLibraryOnDestinationSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrlMWS };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to Des SUT
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);

            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results, "The element Results should be return if CopyIntoItems operation executes successfully");
            this.Site.Assert.IsTrue(2 == copyIntoItemsResponse.Results.Length, "The Results element should contain two CopyResult elements.");
            
            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // If one of the destination locations is a Meeting Workspace site, the server return the error code is DestinationMWS for this site,
            // and other is Success, then R162 ,R158 and R96 should be covered.
            this.Site.Assert.AreEqual(
                                    CopyErrorCode.Success,
                                    copyIntoItemsResponse.Results[0].ErrorCode,
                                    "The CopyIntoItems operation should success.");

            this.Site.Assert.AreEqual(
                                    CopyErrorCode.DestinationMWS, 
                                    copyIntoItemsResponse.Results[1].ErrorCode,
                                    "The CopyIntoItems operation should return DestinationMWS error when the destination location is a Meeting Workspace site");

            // Verify MS-COPYS requirement: MS-COPYS_R162
            this.Site.CaptureRequirement(
                162,
                @"[In Abstract Data Model]The protocol server can proceed with CopyIntoItems operation which attempting to use locations that are part of a Meeting Workspace site as a destination.");
            
            // Verify MS-COPYS requirement: MS-COPYS_R158
            this.Site.CaptureRequirement(
                158,
                @"[In Abstract Data Model] Although such locations[Some locations on a protocol server can be part of a Meeting 
                Workspace site] are valid file locations, attempts to use such a location as a destination for a CopyIntoItems 
                operation will fail.");

            // Verify MS-COPYS requirement: MS-COPYS_R96
            this.Site.CaptureRequirement(
                96,
                @"[In CopyErrorCode] [For CopyIntoItems operation] DestinationMWS: This value is used to indicate a failure to
                copy the file because the destination location is inside a Meeting Workspace site.");
            
            // Verify MS-COPYS requirement: MS-COPYS_R220
            bool isVerifyR220 = string.IsNullOrEmpty(copyIntoItemsResponse.Results[1].ErrorMessage);

            this.Site.CaptureRequirementIfIsFalse(
                isVerifyR220,
                220,
                @"[In CopyIntoItems] If the destination location is part of a Meeting Workspace site, the protocol server MUST 
                report a failure of the copy operation by setting the value of the ErrorCode attribute of the corresponding 
                CopyResult element to ""DestinationMWS"", and provide a string value that specifies the error in the ErrorMessage attribute.");
        }

        /// <summary>
        /// This test case is used to verify if the destination location is a malformed IRI, server should have different responses.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC10_CopyIntoItems_malformedIRI()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // An invalid destination location.
            string invalidDesFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);
            invalidDesFileUrl = this.GenerateInvalidFileUrl(invalidDesFileUrl);

            string[] desUrls = new string[] { invalidDesFileUrl };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                        sourceFileUrl,
                                                                                        desUrls,
                                                                                        getitemsResponse.Fields,
                                                                                        getitemsResponse.StreamRawValues);

            // Verify MS-COPYS requirement: MS-COPYS_R104
            this.Site.CaptureRequirementIfAreEqual(
                CopyErrorCode.InvalidUrl,
                copyIntoItemsResponse.Results[0].ErrorCode,
                104,
                @"[In CopyErrorCode] InvalidUrl:  This value is used to indicate an error when the IRI of a destination location is malformed.");
        }

        /// <summary>
       /// This test case is used to verify the ContentTypeId map to field type on different products.
       /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC11_GetItem_ContentTypeId()
       {
           // Get the value of the source file URL.
           string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

           // Retrieve content and metadata for a file that is stored in a source location.
           GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

           // Select a field by specified field attribute value.
           FieldInformation fieldInfoDes = this.SelectFieldBySpecifiedAtrribute(
                                                                               getitemsResponse.Fields,
                                                                               "ContentTypeId",
                                                                               FieldAttributeType.InternalName);

           if (Convert.ToBoolean(Common.IsRequirementEnabled(124, this.Site)))
           {
               // Verify MS-COPYS requirement: MS-COPYS_R124
               this.Site.CaptureRequirementIfAreEqual(
                   FieldType.Text,
                   fieldInfoDes.Type,
                   124,
                   @"[In Appendix B: Product Behavior] Implementation does ContentTypeId field type map to the Text field type.
                   ( Microsoft SharePoint Foundation 2013 follow this behavior)");
           }
           else if (Convert.ToBoolean(Common.IsRequirementEnabled(125, this.Site)))
           {
               // Verify MS-COPYS requirement: MS-COPYS_R125
               this.Site.CaptureRequirementIfAreEqual(
                   FieldType.Error,
                   fieldInfoDes.Type,
                   125,
                   @"[In Appendix B: Product Behavior] Implementation does ContentTypeId field type map to error field type.
                   (Windows SharePoint Services 3.0 follow this behavior.)");
           }
           else if (Convert.ToBoolean(Common.IsRequirementEnabled(126, this.Site)))
           {
               // Verify MS-COPYS requirement: MS-COPYS_R126
               this.Site.CaptureRequirementIfAreEqual(
                   FieldType.Text,
                   fieldInfoDes.Type,
                   126,
                   @"[In Appendix B: Product Behavior] Implementation does ContentTypeId field type map to Text field type.
                   (SharePoint Foundation 2010 SP2 follow this behavior.)");
           }
           else
           {
               this.Site.Assume.Inconclusive("This test case is used to verify on SharePoint Foundation 2013, Windows SharePoint Services 3.0 and SharePoint Foundation 2010");
           }
       }

        /// <summary>
        /// This test case is used to verify if the Fields collection is empty, the CopyIntoItems operation can succeed.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC12_CopyIntoItems_EmptyFieldsCollection()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1 };

            string fieldName = Common.GetConfigurationPropertyValue("FieldNameOfTestReadOnly", this.Site);

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfo = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponse.Fields,
                                                                                fieldName,
                                                                                FieldAttributeType.InternalName);
            string fieldValueBeforeCopy = fieldInfo.Value;

            this.Site.Assert.AreEqual(
                                    Common.GetConfigurationPropertyValue("FieldDefaultValueOfTestReadOnlyOnSourceDocLibrary", this.Site),
                                    fieldValueBeforeCopy,
                                    "The value of the TestReadOnlyField field type should equal to the property of FieldDefaultValueOfTestReadOnlyOnSourceDocLibrary");

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server with an empty fields collection.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(sourceFileUrl, desUrls, new FieldInformation[] { }, getitemsResponse.StreamRawValues);

            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results, "The element Results should be return if CopyIntoItems operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsResponse.Results.Length, "The Results element should contain one CopyResult element.");

            // Verify whether all copy results are successful.
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItems operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // Retrieve content and metadata for a file that is stored in a source location.
            getitemsResponse = MSCopysAdapter.GetItem(copyIntoItemsResponse.Results[0].DestinationUrl);

            // Select a field by specified field attribute value.
            fieldInfo = this.SelectFieldBySpecifiedAtrribute(
                                                             getitemsResponse.Fields,
                                                             fieldName,
                                                             FieldAttributeType.InternalName);
            string fieldValueAfterCopy = fieldInfo.Value;

            this.Site.Assert.AreEqual(
                                    Common.GetConfigurationPropertyValue("FieldDefaultValueOfTestReadOnlyOnDesDocLibrary", this.Site),
                                    fieldValueAfterCopy,
                                    "The value of the TestReadOnlyField field type should equal to the property of FieldDefaultValueOfTestReadOnlyOnDesDocLibrary");

            // Verify MS-COPYS requirement: MS-COPYS_R250
            this.Site.CaptureRequirementIfAreNotEqual(
                fieldValueBeforeCopy,
                fieldValueAfterCopy,
                250,
                @"[In CopyIntoItems] [Fields] The protocol server MUST support an empty Fields collection [by copying the 
                destination stream, and] by using implementation-specific default values for the metadata.");

            // Verify MS-COPYS requirement: MS-COPYS_R249
            this.Site.CaptureRequirementIfAreEqual(
                copyIntoItemsResponse.Results[0].ErrorCode,
                CopyErrorCode.Success,
                249,
                @"[In CopyIntoItems] [Fields] The protocol server MUST support an empty Fields collection by copying the 
                destination stream[, and by using implementation-specific default values for the metadata].");
        }

        /// <summary>
        /// This test case is used to verify the field EncodedAbsUrl does not be copied when call CopyIntoItems operation on Windows SharePoint Services 3.0.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC13_CopyIntoItems_EncodedAbsUrlField()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(230, this.Site), @"This is executed only when R230Enable is set to true.");

            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            this.Site.Assert.IsNotNull(getitemsResponse, "GetItem operation should succeed");

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoScource = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponse.Fields, 
                                                                                "EncodedAbsUrl",
                                                                                FieldAttributeType.InternalName);

            // Switch to Des SUT
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItems operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseDes = MSCopysAdapter.GetItem(copyIntoItemsResponse.Results[0].DestinationUrl);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoDes = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponseDes.Fields,
                                                                                "EncodedAbsUrl",
                                                                                FieldAttributeType.InternalName);

            // Verify MS-COPYS requirement: MS-COPYS_R230           
            this.Site.CaptureRequirementIfAreNotEqual(
                fieldInfoScource.Value.ToLower(),
                fieldInfoDes.Value.ToLower(),
                230,
                @"[In Appendix B: Product Behavior] CopyIntoItems operation does not copy the EncodedAbsUrl field.(Windows SharePoint Services 3.0, SharePoint Foundation 2010 and SharePoint Foundation 2013 follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to verify CopyIntoItems operation the value of field with internal name _CopySource 
        /// does equal to the value of source location on SharePoint Foundation 2010.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC14_CopyIntoItems_CopySourceField()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(148, this.Site), @"This is executed only when R148Enable is set to true.");

            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            this.Site.Assert.IsNotNull(getitemsResponse, "GetItem operation should succeed");

            // Switch to Des SUT
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItems operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseDes = MSCopysAdapter.GetItem(copyIntoItemsResponse.Results[0].DestinationUrl);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoDes = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponseDes.Fields,
                                                                                "_CopySource",
                                                                                FieldAttributeType.InternalName);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-COPYS_R148");
        
            // Verify MS-COPYS requirement: MS-COPYS_R148
            this.Site.CaptureRequirementIfAreEqual(
                sourceFileUrl.ToLower(),
                fieldInfoDes.Value.ToLower(),
                148,
                @"[In Appendix B: Product Behavior] [For CopyIntoItems operation] Implementation [the value of field with internal name _CopySource ] does equal to the value of source location.(Windows SharePoint Services 3.0, SharePoint Foundation 2010 and SharePoint Foundation 2013 follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to verify the WorkflowEventType map to Error field type.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC15_GetItem_WorkflowEventType()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            string fieldName = Common.GetConfigurationPropertyValue("FieldNameOfTestWorkflowEventType", this.Site);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoDes = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponse.Fields,
                                                                                fieldName,
                                                                                FieldAttributeType.InternalName);
            
            // Verify MS-COPYS requirement: MS-COPYS_R119
            this.Site.CaptureRequirementIfAreEqual(
                FieldType.Error,
                fieldInfoDes.Type,
                119,
                @"[In FieldType] Field Type WorkflowEventType MUST map to error field type in this protocol.");
        }

        /// <summary>
        /// This test case is used to verify the protocol server should accept the Value attribute either be present and have 
        /// an empty value, or not be present.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC16_CopyIntoItems_ValueAttribute()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1 };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to Des SUT
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoSource = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponse.Fields,
                                                                                "_ModerationComments",
                                                                                FieldAttributeType.InternalName);

            #region Value is null

            // Set the value of the Value attribute is null.
            fieldInfoSource.Value = null;

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponseValueIsnull = MSCopysAdapter.CopyIntoItems(
                                                                                                sourceFileUrl,
                                                                                                desUrls,
                                                                                                getitemsResponse.Fields,
                                                                                                getitemsResponse.StreamRawValues);

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.AreEqual(CopyErrorCode.Success, copyIntoItemsResponseValueIsnull.Results[0].ErrorCode, "CopyIntoItems operation should succeed.");

            this.Site.Assert.IsNotNull(copyIntoItemsResponseValueIsnull.Results, "The element Results should be return if CopyIntoItems operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsResponseValueIsnull.Results.Length, "The Results element should contain one CopyResult element.");

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseValueIsNull = MSCopysAdapter.GetItem(copyIntoItemsResponseValueIsnull.Results[0].DestinationUrl);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoDesValueIsNull = this.SelectFieldBySpecifiedAtrribute(
                                                                                            getitemsResponseValueIsNull.Fields,
                                                                                            "_ModerationComments",
                                                                                            FieldAttributeType.InternalName);

            #endregion Value is null

            #region Value is empty

            // Set the value of the Value attribute is empty.
            fieldInfoSource.Value = string.Empty;

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponseValueIsEmpty = MSCopysAdapter.CopyIntoItems(
                                                                                            sourceFileUrl,
                                                                                            desUrls,
                                                                                            getitemsResponse.Fields,
                                                                                            getitemsResponse.StreamRawValues);

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.AreEqual(CopyErrorCode.Success, copyIntoItemsResponseValueIsEmpty.Results[0].ErrorCode, "CopyIntoItems operation should succeed.");
            this.Site.Assert.IsNotNull(copyIntoItemsResponseValueIsEmpty.Results, "The element Results should be return if CopyIntoItems operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsResponseValueIsEmpty.Results.Length, "The Results element should contain one CopyResult element.");

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseValueIsEmpty = MSCopysAdapter.GetItem(copyIntoItemsResponseValueIsEmpty.Results[0].DestinationUrl);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoDesValueIsEmpty = this.SelectFieldBySpecifiedAtrribute(
                                                                                            getitemsResponseValueIsEmpty.Fields,
                                                                                            "_ModerationComments",
                                                                                            FieldAttributeType.InternalName);

            #endregion Value is emopty

            this.Site.Assert.IsTrue(string.IsNullOrEmpty(fieldInfoDesValueIsNull.Value), "The value of the fieldInfoDesValueIsEmpty attribute should be empty.");

            this.Site.Assert.IsTrue(string.IsNullOrEmpty(fieldInfoDesValueIsEmpty.Value), "The value of the fieldInfoDesValueIsEmpty attribute should be empty.");

            // Verify MS-COPYS requirement: MS-COPYS_R215
            this.Site.CaptureRequirement(
                215,
                @"[In CopyIntoItems] [If the value of a field is empty and the base field type is something other than 
                ""Integer"", ""Number"", ""Boolean"", or ""DateTime"", ] A protocol server MUST accept both choices[the 
                Value attribute MUST either be present and have an empty value, or not be present] as an empty value.");
        }

        /// <summary>
        /// This test case is used to verify when the source location does not point to an existing file on the protocol server,
        /// the Fields and Stream elements in the GetItemResponse element will be removed.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC17_GetItems_FileNotExist()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Generate invalid file URL by construct a not-existing file name. 
            string invalidSourceFileUrl = this.GenerateInvalidFileUrl(sourceFileUrl);

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(invalidSourceFileUrl);

            this.Site.Assert.IsTrue(string.IsNullOrEmpty(getitemsResponse.Stream), "The element Stream must be null when the source location does not point to an existing file for GetItems operation.");
            this.Site.Assert.IsNull(getitemsResponse.Fields, "The element Fields must be null when the source location does not point to an existing file for GetItems operation.");

            // Verify MS-COPYS requirement: MS-COPYS_R176
            this.Site.CaptureRequirement(
                176,
                @"[In GetItem] [The protocol server returns results based on the following conditions:] If the source location does not point to an existing file on the protocol server, the protocol server MUST omit the Fields and Stream elements in the GetItemResponse element (section 3.1.4.1.2.2).");
        }

        /// <summary>
        /// This test case is used to verify if CopyIntoItems operation executes successfully, the file should be copied from 
        /// source location to all destination locations.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC18_CopyIntoItems_CheckFileContent()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItems operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.AreEqual(2, copyIntoItemsResponse.Results.Length, "CopyIntoItems operation return a collection must contain two items if the DestinationUrls have two items in the request.");

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseDes1 = MSCopysAdapter.GetItem(copyIntoItemsResponse.Results[0].DestinationUrl);

            // Retrieve content and metadata for a file that is stored in a source location with second destination URL.
            GetItemResponse getitemsResponseDes2 = MSCopysAdapter.GetItem(copyIntoItemsResponse.Results[1].DestinationUrl);

            this.Site.Assert.IsNotNull(getitemsResponseDes1.StreamRawValues, "The element StreamRawValues should be returned.");
            this.Site.Assert.IsNotNull(getitemsResponseDes2.StreamRawValues, "The element StreamRawValues should be returned.");

            // If the contents of the file are equal to the property which is configured, then R200, R166 and R218 are captured.
            this.Site.Assert.AreEqual(Common.GetConfigurationPropertyValue("SourceFileContents", this.Site), Encoding.UTF8.GetString(getitemsResponseDes1.StreamRawValues), "The file content should equal to the source which is copied to the destination location.");
            this.Site.Assert.AreEqual(Common.GetConfigurationPropertyValue("SourceFileContents", this.Site), Encoding.UTF8.GetString(getitemsResponseDes2.StreamRawValues), "The file content should equal to the source which is copied to the destination location.");

            // Verify MS-COPYS requirement: MS-COPYS_R200
            this.Site.CaptureRequirement(
                200,
                @"[In CopyIntoItems] The CopyIntoItems operation copies a file to the destination server. ");

            // Verify MS-COPYS requirement: MS-COPYS_R166
            this.Site.CaptureRequirement(
                166,
                @"[In Message Processing Events and Sequencing Rules] CopyIntoItems:  Copies a file to a destination server that is different from the source location.");

            // Verify MS-COPYS requirement: MS-COPYS_R218
            this.Site.CaptureRequirement(
                218,
                @"[In CopyIntoItems] The protocol server MUST attempt to copy the file to all destination locations that are 
                specified in the request.");

            // Must call SwitchTargetServiceLocation method which is used to change the target service location
            // before call the CopyIntoItems operation.
            this.Site.CaptureRequirement(
                201,
                @"[In CopyIntoItems] This operation can be used when the destination server is different from the source location.");
        }

        /// <summary>
        /// This test case is used to verify the CopyIntoItems operation executes successfully, the collection MUST have 
        /// exactly one record for each destination location.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC19_CopyIntoItems_CheckResultNumber()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItems operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results, "The Result of CopyIntoItems should be returned.");

            // If server should return results for each destinations location, then R254 and R260 should be covered.
            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results[0], "The Result should be return if CopyIntoItems operation succeed.");
            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results[1], "The Result should be return if CopyIntoItems operation succeed.");

            // Verify MS-COPYS requirement: MS-COPYS_R254
            this.Site.CaptureRequirement(
                254,
                @"[In CopyIntoItemsResponse] It contains a collection of results for each destination location that was passed
                to the protocol server in the CopyIntoItems request.");

            // Verify MS-COPYS requirement: MS-COPYS_R260
            this.Site.CaptureRequirement(
                260,
                @"[In CopyIntoItemsResponse] [Results] The collection MUST have exactly one record for each destination location
                that is passed into the request, as specified in section 3.1.4.2.");
        }

        /// <summary>
        /// This test case is used to verify if CopyIntoItems operation executes successfully, the return value in the Results collection MUST be in the same order as the items 
        /// in the destination locations collection.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC20_CopyIntoItems_CheckResultOrder()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };

            // Retrieve content and metadata for a file that is stored in a source location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            // Switch to destination SUT.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                                                                                    sourceFileUrl,
                                                                                    desUrls,
                                                                                    getitemsResponse.Fields,
                                                                                    getitemsResponse.StreamRawValues);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItems operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsResponse.Results, "The Result element should be returned.");
            this.Site.Assert.AreEqual(2, copyIntoItemsResponse.Results.Length, "CopyIntoItems operation return a collection must contain two items if the DestinationUrls have two items in the request.");

            // If the value of DestinationUrl in result is equal to the item in the collection of locations, then R208, R209 should be captured.
            this.Site.Assert.AreEqual(desUrls[0], copyIntoItemsResponse.Results[0].DestinationUrl, "The value of DestinalUrl in first item should equal to the first index of DestinationUrlCollection");
            this.Site.Assert.AreEqual(desUrls[1], copyIntoItemsResponse.Results[1].DestinationUrl, "The value of DestinalUrl in second item should equal to the second index of DestinationUrlCollection");

            // Verify MS-COPYS requirement: MS-COPYS_R208
            this.Site.CaptureRequirement(
                208,
                @"[In CopyIntoItems] The CopyResult element in the Results collection MUST be in the same order as the items 
                in the destination locations collection.");

            // Verify MS-COPYS requirement: MS-COPYS_R209
            this.Site.CaptureRequirement(
                209,
                @"[In CopyIntoItems] The DestinationUrl attribute of the CopyResult element (section 2.2.4.2) that corresponds to the destination location MUST be set to the value of the destination location.");
        }

        /// <summary>
        /// This test case is used to verify if the file cannot be created at the given destination location,
        /// protocol server should return "Unknown" error code and provide a string value that describes the error in the
        /// ErrorMessage attribute.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S01_TC21_CopyIntoItems_UnknowForCannotCreateFileAtDestination()
        {
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);
            string invalieFileUrl = this.GenerateInvalidFolderPathForFileUrl(desFileUrl1);

            string[] desUrls = new string[] { invalieFileUrl };

            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);

            // Copy a file to the destination server when one of the destination URL is invalid.
            CopyIntoItemsResponse copyIntoItemsResponse = MSCopysAdapter.CopyIntoItems(
                sourceFileUrl,
                desUrls,
                getitemsResponse.Fields,
                getitemsResponse.StreamRawValues);

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-COPYS_R224, returned error code is '{0}'", copyIntoItemsResponse.Results[0].ErrorCode);

            bool isR224Verified = copyIntoItemsResponse.Results[0].ErrorCode == CopyErrorCode.Unknown
                && !string.IsNullOrEmpty(copyIntoItemsResponse.Results[0].ErrorMessage);

            // Verify MS-COPYS requirement: MS-COPYS_R224  
            this.Site.CaptureRequirementIfIsTrue(
                isR224Verified,
                224,
                @"[In CopyIntoItems] If the file cannot be created at the given destination location, the protocol server MUST report a failure for this destination location by setting the ErrorCode attribute of the corresponding CopyResult element to ""Unknown"" and provide a string value that describes the error in the ErrorMessage attribute.");
        }
    }
}