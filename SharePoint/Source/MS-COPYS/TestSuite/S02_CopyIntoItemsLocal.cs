namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test whether the CopyIntoItemsLocal behaviors follow the Open Spec definitions.
    /// </summary>
    [TestClass]
    public class S02_CopyIntoItemsLocal : TestSuiteBase
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
        public void CopyIntoItemsLocalTestCaseInitialize()
        {
            // Set the target service location to the destination SUT for CopyIntoItemsLocal test cases.
            MSCopysAdapter.SwitchTargetServiceLocation(ServiceLocation.DestinationSUT);
        }
        #endregion

        #region Test cases

        /// <summary>
        /// This test case is used to verify if CopyIntoItemsLocal operation executes successfully, error code "Success" should be returned.  
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC01_CopyIntoItemsLocal_ErrorCodeForSucess()
        {
            // Get resource from properties,normal source file and normal destination library.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };

            // Copy files in same SUT.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrlOnDesSut,
                                                                                    desUrls);

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results, "The element Results should be return if CopyIntoItemsLocal operation executes successfully");
            this.Site.Assert.IsTrue(2 == copyIntoItemsLocalResponse.Results.Length, "The Results element should contain two CopyResult elements.");

            // If the error code is equal to Success when call CopyIntoItemsLocal operation with correct configuration, then R281, R277 and R88 should be covered.
            this.Site.Assert.AreEqual(
                                      CopyErrorCode.Success,
                                      copyIntoItemsLocalResponse.Results[0].ErrorCode,
                                      "The CopyIntoItemsLocal operation should succeed.");

            this.Site.Assert.AreEqual(
                                      CopyErrorCode.Success,
                                      copyIntoItemsLocalResponse.Results[1].ErrorCode,
                                      "The CopyIntoItemsLocal operation should succeed.");

            // Verify MS-COPYS requirement: MS-COPYS_R281
            this.Site.CaptureRequirement(
                281,
                @"[In CopyIntoItemsLocal] The protocol server MUST report the status of the operation inside the Results 
                collection for each destination location that is passed");

            // Verify MS-COPYS requirement: MS-COPYS_R88
            this.Site.CaptureRequirement(
                88,
                @"[In CopyErrorCode] Success: This value is used when the CopyIntoItemsLocal operation succeeds for the specified destination location.");

            // Verify MS-COPYS requirement: MS-COPYS_R277
            this.Site.CaptureRequirement(
                277,
                @"[In CopyIntoItemsLocal] The protocol server MUST perform the copy operation for the file and construct a response.");

            this.Site.Assert.IsTrue(2 == copyIntoItemsLocalResponse.Results.Length, "The Results element should contain two CopyResult elements.");

            // Verify MS-COPYS requirement: MS-COPYS_R51
            bool isVerifiedR51 = string.IsNullOrEmpty(copyIntoItemsLocalResponse.Results[0].ErrorMessage);
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR51,
                51,
                @"[In CopyResult] [ErrorMessage] [For CopyIntoItemsLocal operation] If the value of ErrorCode is ""Success,"" the attribute MUST NOT be present.");
        }

        /// <summary>
        /// This test case is used to verify CopyIntoItemsLocal operation when a folder location that is not valid on the destination server.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC02_CopyIntoItemsLocal_DestinationInvalid()
        {
            // Get resource from properties,normal source file and normal destination library.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get a folder location that is not valid on the destination server. 
            string desFileUrl = this.GenerateInvalidFolderPathForFileUrl(desFileUrl1);      
            string[] desUrls = new string[] { desFileUrl };

            // Copy files in same SUT with the invalid destination folder.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrlOnDesSut,
                                                                                    desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results, "The element Results should be return if CopyIntoItemsLocal operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsLocalResponse.Results.Length, "The Results element should contain one CopyResult element.");

            // Verify MS-COPYS requirement: MS-COPYS_R92
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.DestinationInvalid,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                92,
                @"[In CopyErrorCode] DestinationInvalid: This value is used to indicate  the destination location points to 
                a folder location that is not valid on the destination server.");

            // Verify MS-COPYS requirement: MS-COPYS_R279
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.DestinationInvalid,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                279,
                @"[In CopyIntoItemsLocal] [If the source location and the destination location refer to different protocol servers, or ]if the destination location points to a non-existing folder, the protocol server MUST report a failure by returning the CopyResult element with the ErrorCode attribute set to ""DestinationInvalid"".");
        }

        /// <summary>
        /// This test case is used to verify CopyIntoItemsLocal operation when destination location is inside a Meeting Workspace site.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC03_CopyIntoItemsLocal_DestinationMWS()
        {
            // Get resource from properties,normal source file and a Meeting Workspace site destination library.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrlMWS = this.GetDestinationFileUrl(DestinationFileUrlType.MWSLibraryOnDestinationSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrlMWS };

            // Copy files in same SUT with  Meeting Workspace library.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrlOnDesSut,
                                                                                    desUrls);

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // If one of the destination locations is a Meeting Workspace site, the server return the error code is DestinationMWS for this site,
            // and other is Success, then R163, R159, R97 should be covered.
            this.Site.Assert.AreEqual(
                                     CopyErrorCode.Success, 
                                     copyIntoItemsLocalResponse.Results[0].ErrorCode,
                                     "The CopyIntoItemsLocal operation should succeed.");

            this.Site.Assert.AreEqual(
                                     CopyErrorCode.DestinationMWS, 
                                     copyIntoItemsLocalResponse.Results[1].ErrorCode,
                                     "The CopyIntoItemsLocal operation should fail when the destination Location is a Meeting Workspace site.");
            
            // Verify MS-COPYS requirement: MS-COPYS_R163
            this.Site.CaptureRequirement(
                163,
                @"[In Abstract Data Model]The protocol server can proceed with CopyIntoItemsLocal operation which attempting to use locations that are part of a Meeting Workspace site as a destination.");

            // Verify MS-COPYS requirement: MS-COPYS_R159
            this.Site.CaptureRequirement(
                159,
                @"[In Abstract Data Model] Although such locations[Some locations on a protocol server can be part of a Meeting Workspace site] are valid file locations, attempts to use such a location as a destination for a CopyIntoItemsLocal operation will fail.");

            // Verify MS-COPYS requirement: MS-COPYS_R97
            this.Site.CaptureRequirement(
                97,
                @"[In CopyErrorCode] [For CopyIntoItemsLocal operation] DestinationMWS: This value is used to indicate a failure to copy the file because the destination location is inside a Meeting Workspace site.");

             // Verify MS-COPYS requirement: MS-COPYS_R54
            bool isVerifyR54 = string.IsNullOrEmpty(copyIntoItemsLocalResponse.Results[1].ErrorMessage);

            this.Site.CaptureRequirementIfIsFalse(
                isVerifyR54,
                54,
                @"[In CopyResult] [ErrorMessage] [For CopyIntoItemsLocal operation] Otherwise[If the value of ErrorCode is not ""Success"" ], the ErrorMessage attribute MUST be present and the value MUST be a non-empty Unicode string.");

            // Verify MS-COPYS requirement: MS-COPYS_R287
            bool isVerifyR287 = isVerifyR54;
            this.Site.CaptureRequirementIfIsFalse(
                isVerifyR287,
                287,
                @"[In CopyIntoItemsLocal] If the destination location is part of a Meeting Workspace site, the protocol 
                server MUST report a failure of the copy operation by setting the value of the ErrorCode attribute of 
                the corresponding CopyResult element to ""DestinationMWS"", and provide a string value that specifies
                the error in the ErrorMessage attribute.");
        }

        /// <summary>
        /// This test case is used to verify CopyIntoItemsLocal operation when destination file is checked out.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC04_CopyIntoItemsLocal_DestinationCheckedOut()
        {
            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Upload a txt file to the destination SUT.
            this.UploadTxtFileByFileUrl(desFileUrl1);

            string[] desUrls = new string[] { desFileUrl1 };

            // Check out the file by the specified user.
            MSCOPYSSutControlAdapter.CheckOutFileByUser(
                                                        desFileUrl1,
                                                        Common.GetConfigurationPropertyValue("MSCOPYSCheckOutUserName", this.Site),
                                                        Common.GetConfigurationPropertyValue("PasswordOfCheckOutUser", this.Site),
                                                        Common.GetConfigurationPropertyValue("Domain", this.Site));

            // Copy files in same SUT with check out source file.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrl,
                                                                                    desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results, "The element Results should be return if CopyIntoItemsLocal operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsLocalResponse.Results.Length, "The Results element should contain one CopyResult element.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // Verify MS-COPYS requirement: MS-COPYS_R155
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.DestinationCheckedOut,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                155,
                @"[In Abstract Data Model] In this case[files as checked out], the CopyIntoItemsLocal operations take into account the checked-out status when accessing files at the destination locations.");

            // Verify MS-COPYS requirement: MS-COPYS_R286
            bool isVerifyR286 = !string.IsNullOrEmpty(copyIntoItemsLocalResponse.Results[0].ErrorMessage);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifyR286,
                286,
                @"[In CopyIntoItemsLocal] If the file on the protocol server is checked out and cannot be updated, the protocol 
                server MUST report a failure of the copy operation by setting the value of the ErrorCode attribute of the 
                corresponding CopyResult element to ""DestinationCheckedOut"", and provide a string value that specifies the error
                in the ErrorMessage attribute.");

            // Verify MS-COPYS requirement: MS-COPYS_R103
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
               CopyErrorCode.DestinationCheckedOut,
               copyIntoItemsLocalResponse.Results[0].ErrorCode,
                103,
                @"[In CopyErrorCode] [For CopyIntoItemsLocal operation] DestinationCheckedOut: This value is used to indicate an error when the file on the destination location is checked out and cannot be overridden.");

            // Undo checkout for a file by specified user credential.
            MSCOPYSSutControlAdapter.UndoCheckOutFileByUser(
                                                            desFileUrl1,
                                                            Common.GetConfigurationPropertyValue("MSCOPYSCheckOutUserName", this.Site),
                                                            Common.GetConfigurationPropertyValue("PasswordOfCheckOutUser", this.Site),
                                                            Common.GetConfigurationPropertyValue("Domain", this.Site));
        }

        /// <summary>
        /// This test case is used to verify CopyIntoItemsLocal operation when source folder is not existent.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC05_CopyIntoItemsLocal_SourceNotPointExistFolder()
        {
            // Get resource from properties,a non-existent source folder and normal destination library.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Generate invalid file URL by confusing the folder path.
            string sourceFolder = this.GenerateInvalidFolderPathForFileUrl(sourceFileUrlOnDesSut);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1 };

            // Copy files in same SUT with non-existent source folder.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFolder,
                                                                                    desUrls);

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results, "The element Results should be return if CopyIntoItemsLocal operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsLocalResponse.Results.Length, "The Results element should contain one CopyResult element.");

            // Verify MS-COPYS requirement: MS-COPYS_R111
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.Unknown,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                111,
                @"[In CopyErrorCode] [For CopyIntoItemsLocal operation] Unknown: This value is used to indicate an error for all other error conditions for a given destination location.");
            
            // Verify MS-COPYS requirement: MS-COPYS_R271
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.Unknown,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                271,
                @"[In CopyIntoItemsLocal] If the source location does not point to an existing file, then if the destination location does not point to a existing folder or file, the protocol server MUST report a failure by returning the CopyResult element (section 2.2.4.2) with the ErrorCode attribute set to ""Unknown"" for this destination location.");

            // Verify MS-COPYS requirement: MS-COPYS_R289
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.Unknown,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                289,
                @"[In CopyIntoItemsLocal] If the file cannot be created at the given destination location,  the protocol server 
                MUST report a failure for this destination location by setting the ErrorCode attribute of the corresponding 
                CopyResult element to ""Unknown"" and provide a string value that describes the error in the ErrorMessage attribute.");
        }

        /// <summary>
        /// This test case is used to verify CopyIntoItemsLocal operation when a destination location is a malformed IRI.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC06_CopyIntoItemsLocal_DesMalformedIRI()
        {
            // Get resource from properties,normal source and malformed IRI.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);
            
            // Invalid destination URL.
            string invalidDesFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);
            invalidDesFileUrl = this.GenerateInvalidFileUrl(invalidDesFileUrl);

            string[] desUrls = new string[] { invalidDesFileUrl };
 
            // Copy files in same SUT with malformed IRI.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrlOnDesSut,
                                                                                    desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results, "The element Results should be return if CopyIntoItemsLocal operation executes successfully");
            this.Site.Assert.IsTrue(1 == copyIntoItemsLocalResponse.Results.Length, "The Results element should contain one CopyResult element.");

            if (Convert.ToBoolean(Common.IsRequirementEnabled(1043, this.Site)))
            {
                // Verify MS-COPYS requirement: MS-COPYS_R1043
                this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                    CopyErrorCode.InvalidUrl,
                    copyIntoItemsLocalResponse.Results[0].ErrorCode,
                    1043,
                    @"[In CopyErrorCode] [For CopyIntoItemsLocal operation] Implementation does return an ErrorCode of ""InvalidUrl"" when a destination location is a malformed IRI.(SharePoint Foundation 2013 follow this behavior.)");
            }
            else if (Convert.ToBoolean(Common.IsRequirementEnabled(1044, this.Site)))
            {
                // Verify MS-COPYS requirement: MS-COPYS_R1044
                this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                    CopyErrorCode.InvalidUrl,
                    copyIntoItemsLocalResponse.Results[0].ErrorCode,
                    1044,
                    @"[In CopyErrorCode] [For CopyIntoItemsLocal operation] Implementation does return an ErrorCode of ""Unknown"" when a destination location  is a malformed IRI on Windows SharePoint Services 3.0 and SharePoint Foundation 2010.");
            }
            else 
            {
                this.Site.Assume.Inconclusive("This test case is used to verify on SharePoint Foundation 2013, Windows SharePoint Services 3.0 and SharePoint Foundation 2010");
            }
        }

        /// <summary>
        /// This test case is used to verify the field EncodedAbsUrl does not be copy when call CopyIntoItemsLocal operation on Windows SharePoint Services 3.0.
        /// EncodedAbsUrl: Absolute server-relative URL of the related file of an item. This value is computed by destination location.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC07_CopyIntoItemsLocal_EncodedAbsUrlField()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(232, this.Site), @"This is executed only when R232Enable is set to true.");

            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl };

            // Retrieve the contents and metadata for a file from the specified location.
            GetItemResponse getitemsResponse = MSCopysAdapter.GetItem(sourceFileUrl);

            this.Site.Assert.IsNotNull(getitemsResponse, "GetItem operation should succeed");

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoScource = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponse.Fields,
                                                                                "EncodedAbsUrl",
                                                                                FieldAttributeType.InternalName);

            // Copy a file to the destination server.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrl,
                                                                                    desUrls);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsLocalResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItemsLocal operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseDes = MSCopysAdapter.GetItem(copyIntoItemsLocalResponse.Results[0].DestinationUrl);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoDes = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponseDes.Fields,
                                                                                "EncodedAbsUrl",
                                                                                FieldAttributeType.InternalName);

            if (Common.IsRequirementEnabled(232, this.Site))
            {
                // Verify MS-COPYS requirement: MS-COPYS_R232           
                this.Site.CaptureRequirementIfAreNotEqual(
                fieldInfoScource.Value.ToLower(),
                fieldInfoDes.Value.ToLower(),
                232,
                @"[In Appendix B: Product Behavior] CopyIntoItemsLocal operation does not copy the EncodedAbsUrl field.(Windows SharePoint Services 3.0, SharePoint Foundation 2010 and SharePoint Foundation 2013 follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify CopyIntoItemsLocal operation the value of field with internal name _CopySource 
        /// does equal to the value of source location on SharePoint Foundation 2010.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC08_CopyIntoItemsLocal_CopySourceField()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(149, this.Site), @"This is executed only when R149Enable is set to true.");

            // Get the value of the source file URL.
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl };

            // Copy a file to the destination server.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrl,
                                                                                    desUrls);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsLocalResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItemsLocal operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseDes = MSCopysAdapter.GetItem(copyIntoItemsLocalResponse.Results[0].DestinationUrl);

            // Select a field by specified field attribute value.
            FieldInformation fieldInfoDes = this.SelectFieldBySpecifiedAtrribute(
                                                                                getitemsResponseDes.Fields,
                                                                                "_CopySource",
                                                                                FieldAttributeType.InternalName);

            if (Common.IsRequirementEnabled(149, this.Site))
            {
                // Verify MS-COPYS requirement: MS-COPYS_R149
                this.Site.CaptureRequirementIfAreEqual(
                sourceFileUrl.ToLower(),
                fieldInfoDes.Value.ToLower(),
                149,
                @"[In Appendix B: Product Behavior] [For CopyIntoItemsLocation operation] Implementation [the value of field with internal name _CopySource ] does equal to the value of source location.(Windows SharePoint Services 3.0, SharePoint Foundation 2010 and SharePoint Foundation 2013 follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify if CopyIntoItemsLocal operation executes successfully, the file should be copied from 
        /// source location to all destination locations.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC09_CopyIntoItemsLocal_CheckFileContent()
        {
            // Get resource from properties,normal source file and normal destination library.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };
 
            // Copy files in same SUT.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrlOnDesSut,
                                                                                    desUrls);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsLocalResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItemsLocal operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.AreEqual(2, copyIntoItemsLocalResponse.Results.Length, "CopyIntoItemsLocal operation return a collection must contain two items if the DestinationUrls have two items in the request.");

            // Retrieve content and metadata for a file that is stored in a source location with first destination URL.
            GetItemResponse getitemsResponseDes1 = MSCopysAdapter.GetItem(copyIntoItemsLocalResponse.Results[0].DestinationUrl);

            // Retrieve content and metadata for a file that is stored in a source location with second destination URL.
            GetItemResponse getitemsResponseDes2 = MSCopysAdapter.GetItem(copyIntoItemsLocalResponse.Results[1].DestinationUrl);

            this.Site.Assert.IsNotNull(getitemsResponseDes1.StreamRawValues, "The element StreamRawValues should be returned.");
            this.Site.Assert.IsNotNull(getitemsResponseDes2.StreamRawValues, "The element StreamRawValues should be returned.");

            // If the contents of the file are equal to the property which is configured, then R168, R285, R265 and R261 should be covered.
            this.Site.Assert.AreEqual(
                                      Common.GetConfigurationPropertyValue("SourceFileContents", this.Site),
                                      Encoding.UTF8.GetString(getitemsResponseDes1.StreamRawValues),
                                      "The file content should equal to the source which is copied to the destination location.");

            this.Site.Assert.AreEqual(
                                      Common.GetConfigurationPropertyValue("SourceFileContents", this.Site),
                                      Encoding.UTF8.GetString(getitemsResponseDes2.StreamRawValues),
                                      "The file content should equal to the source which is copied to the destination location.");

            // Verify MS-COPYS requirement: MS-COPYS_R168
            this.Site.CaptureRequirement(
                168,
                @"[In Message Processing Events and Sequencing Rules] CopyIntoItemsLocal: Copies a file when the destination 
                of the operation is on the same protocol server as the source location.");

            // Verify MS-COPYS requirement: MS-COPYS_R285
            this.Site.CaptureRequirement(
                285,
                @"[In CopyIntoItemsLocal] The protocol server MUST attempt to copy the file to all destination locations that 
                are specified in the request.");

            // Because the server is never be changed, the source location and destination server is the same server all the time.
            // Verify MS-COPYS requirement: MS-COPYS_R265
            this.Site.CaptureRequirement(
                265,
                @"[In CopyIntoItemsLocal] The source location and the destination server refer to the same protocol server for 
                this operation.");

            // Verify MS-COPYS requirement: MS-COPYS_R261
            this.Site.CaptureRequirement(
                261,
                @"[In CopyIntoItemsLocal] The CopyIntoItemsLocal operation copies a file, and the associated metadata, from one 
                location to one or more locations on the same protocol server.");
        }

        /// <summary>
        /// This test case is used to verify the CopyIntoItemsLocal operation executes successfully, the collection MUST have 
        /// exactly one record for each destination location.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC10_CopyIntoItemsLocal_CheckResultNumber()
        {
            // Get resource from properties,normal source file and normal destination library.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };
 
            // Copy files in same SUT.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrlOnDesSut,
                                                                                    desUrls);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsLocalResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItemsLocal operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results, "The Result of CopyIntoItemsLocal should be returned.");
            this.Site.Assert.IsTrue(2 == copyIntoItemsLocalResponse.Results.Length, "The Results element should contain two CopyResult elements.");

            // If server should return results for each destinations location, then R311 should be covered.
            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results[0], "The Result should be return if CopyIntoItemsLocal operation succeed.");
            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results[1], "The Result should be return if CopyIntoItemsLocal operation succeed.");

            // Verify MS-COPYS requirement: MS-COPYS_R311
            this.Site.CaptureRequirement(
                311,
                @"[In CopyIntoItemsLocalResponse] [Results] The collection MUST have exactly one entry for each destination 
                location that is passed in the request, as specified in section 3.1.4.2.");
        }

        /// <summary>
        /// This test case is used to verify if CopyIntoItemsLocal operation executes successfully, the return value in the Results collection MUST be in the same order as the items 
        /// in the destination locations collection.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC11_CopyIntoItemsLocal_CheckResultOrder()
        {
            // Get resource from properties,normal source file and normal destination library.
            string sourceFileUrlOnDesSut = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);

            // Get the first destination location.
            string desFileUrl1 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            // Get the section destination location.
            string desFileUrl2 = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl1, desFileUrl2 };
 
            // Copy files in same SUT.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                                                                                    sourceFileUrlOnDesSut,
                                                                                    desUrls);

            // Verify whether all copy results are successful. 
            bool isVerifyCopySuccess = VerifyAllCopyResultsSuccess(copyIntoItemsLocalResponse.Results, true);

            this.Site.Assert.IsTrue(isVerifyCopySuccess, "CopyIntoItemsLocal operation should succeed.");

            // Collect files from specified file URLs.
            this.CollectFileByUrl(desUrls);

            this.Site.Assert.IsNotNull(copyIntoItemsLocalResponse.Results, "The Result element should be returned.");
            this.Site.Assert.AreEqual(2, copyIntoItemsLocalResponse.Results.Length, "CopyIntoItemsLocal operation return a collection must contain two items if the DestinationUrls have two items in the request.");

            // If the value of DestinationUrl is equal to the item in the collection of locations, then R282 should be covered.
            this.Site.Assert.AreEqual(
                                      desFileUrl1,
                                      copyIntoItemsLocalResponse.Results[0].DestinationUrl,
                                      "The CopyResult element in the Results collection should be the same order as the items in the destination collection.");

            this.Site.Assert.AreEqual(
                                      desFileUrl2,
                                      copyIntoItemsLocalResponse.Results[1].DestinationUrl,
                                      "The CopyResult element in the Results collection should be the same order as the items in the destination collection.");

            // Verify MS-COPYS requirement: MS-COPYS_R282
            this.Site.CaptureRequirement(
                282,
                @"[In CopyIntoItemsLocal] The CopyResult element in the Results collection MUST be in the same order as the 
                items in the destination locations collection.");
        }

        /// <summary>
        /// This test case is used to verify empty result is returned if the source location and the destination locaiotn refer 
        /// to different server.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC12_CopyIntoItemsLocal_SourceDestinationOnDifferentServer()
        {
            if (Common.GetConfigurationPropertyValue("SourceSutComputerName", this.Site) == string.Empty)
            {
                Site.Assert.Inconclusive("This case runs only when the Source system under test exists.");
            }
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnSourceSUT);
            string desFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);

            string[] desUrls = new string[] { desFileUrl };

            // Copy files in same SUT.
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                sourceFileUrl,
                desUrls);

            if (Common.IsRequirementEnabled(2781, this.Site))
            {
                // Verify MS-COPYS requirement: MS-COPYS_R2781
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0,
                    copyIntoItemsLocalResponse.Results.Length,
                    2781,
                    @"[In Appendix B: Product Behavior] Implementation does return empty results if the source location and the destination location refer to different protocol servers. (<5> Section 3.1.4.3:  The server returns empty results when the source location and the destination location refer to different protocol servers.)");
            }
        }

        /// <summary>
        /// This test case is used to verify if the protocol client does not have permission to access the source file and the destionation
        /// location does not point to a existing folder or file, "Unknown" error code should be returned.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC13_CopyIntoItemsLocal_NoPermisonAndDestinationLocationNotExist()
        {
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);
            string desFileUrl = this.GetDestinationFileUrl(DestinationFileUrlType.NormalDesLibraryOnDesSUT);
            string[] desUrls = new string[] { desFileUrl };

            MSCopysAdapter.SwitchUser(
                Common.GetConfigurationPropertyValue("MSCOPYSNoPermissionUser", this.Site),
                Common.GetConfigurationPropertyValue("PasswordOfNoPermissionUser", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site));
            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                sourceFileUrl,
                desUrls);

            // Verify MS-COPYS requirement: MS-COPYS_R270
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.Unknown,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                270,
                @"[In CopyIntoItemsLocal] If the source location points to a file whose permission setting does not allow access by the protocol client, then if the destination location does not point to a existing folder or file, the protocol server MUST report a failure by returning the CopyResult element (section 2.2.4.2) with the ErrorCode attribute set to ""Unknown"" for this destination location.");
        }

        /// <summary>
        /// This test case is used to verify if the source location does not point to an existing file and destination location points to 
        /// an existing file, "SourceInvalid" error code should be returned.
        /// </summary>
        [TestCategory("MSCOPYS"), TestMethod()]
        public void MSCOPYS_S02_TC14_CopyIntoItemsLocal_SourceLoationNotExistAndDestinationLocationExist()
        {
            string sourceFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);
            string invalidSourceFileUrl = sourceFileUrl.Insert(sourceFileUrl.LastIndexOf("."), DateTime.Now.Ticks.ToString());

            string desFileUrl = this.GetSourceFileUrl(SourceFileUrlType.SourceFileOnDesSUT);
            string[] desUrls = new string[] { desFileUrl };

            CopyIntoItemsLocalResponse copyIntoItemsLocalResponse = MSCopysAdapter.CopyIntoItemsLocal(
                invalidSourceFileUrl,
                desUrls);

            // Verify MS-COPYS requirement: MS-COPYS_R273
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.SourceInvalid,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                273,
                @"[In CopyIntoItemsLocal] If the source location does not point to an existing file, then if the destination location points to an existing file, the protocol server MUST report a failure by returning the CopyResult element with the ErrorCode attribute set to ""SourceInvalid"" for this destination location. ");

            // Verify MS-COPYS requirement: MS-COPYS_R98
            this.Site.CaptureRequirementIfAreEqual<CopyErrorCode>(
                CopyErrorCode.SourceInvalid,
                copyIntoItemsLocalResponse.Results[0].ErrorCode,
                98,
                @"[In CopyErrorCode] SourceInvalid: This value is used to indicate an error when the source location for the copy operation does not reference an existing file in the source location.");
        }
        #endregion Test cases
    }
}