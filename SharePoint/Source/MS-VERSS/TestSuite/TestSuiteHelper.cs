//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// A class that contains the helper methods used by MS-VERSS test cases.
    /// </summary>
    public class TestSuiteHelper 
    {
        #region Variables
        /// <summary>
        /// A string indicates the file name which will be uploaded to site.
        /// </summary>
        public const string UploadFileName = "MS-VERSS_Test.txt";

        /// <summary>
        /// The comments for checking in file.
        /// </summary>
        public const string FileComments = "File comments.";

        /// <summary>
        /// The instance of the SUT control adapter.
        /// </summary>
        private IMS_VERSSSUTControlAdapter sutControlAdapterInstance;

        /// <summary>
        /// The instance of the protocol adapter.
        /// </summary>
        private IMS_VERSSAdapter protocolAdapterInstance;

        /// <summary>
        /// The instance of ILISTSWSSUTControlAdapter interface.
        /// </summary>
        private IMS_LISTSWSSUTControlAdapter listsSutControlAdaterInstance;

         /// <summary>
        /// The name of list in the site.
        /// </summary>
        private string documentLibrary;

        /// <summary>
        /// The name of file in the list.
        /// </summary>
        private string fileName;

        /// <summary>
        /// The absolute URL for the site collection.
        /// </summary>
        private string requestUrl;

        /// <summary>
        /// Transfer ITestSite into adapter, make adapter can use ITestSite's function.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// The id of list in the site.
        /// </summary>
        private string listId = string.Empty;

        /// <summary>
        /// A Dictionary object indicates the attribute of specified file and specified version.
        /// </summary>
        private Dictionary<string, string> fileVersionAttributes = new Dictionary<string, string>();
        #endregion

        /// <summary>
        /// Initializes a new instance of the TestSuiteHelper class.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite.</param>
        /// <param name="listName">The name of list in the site.</param>
        /// <param name="fileName">The name of file in the list.</param>
        /// <param name="listsSutControlAdapter">The instance of ILISTSWSSUTControlAdapter interface.</param>
        /// <param name="protocolAdapter">The instance of the protocol adapter.</param>
        /// <param name="sutControlAdapter">The instance of the SUT control adapter.</param>
        public TestSuiteHelper(
            ITestSite testSite,
            string listName,
            string fileName,
            IMS_LISTSWSSUTControlAdapter listsSutControlAdapter,
            IMS_VERSSAdapter protocolAdapter,
            IMS_VERSSSUTControlAdapter sutControlAdapter)
        {
            this.site = testSite;
            this.listsSutControlAdaterInstance = listsSutControlAdapter;
            this.sutControlAdapterInstance = sutControlAdapter;
            this.protocolAdapterInstance = protocolAdapter;

            this.requestUrl = Common.GetConfigurationPropertyValue("RequestUrl", testSite);
            this.documentLibrary = listName;

            this.fileName = fileName;
            this.listId = listsSutControlAdapter.GetListID(listName);
            this.fileVersionAttributes.Clear();
        }

        /// <summary>
        /// Add versions to the file. 
        /// </summary>
        public void AddFileVersions()
        {
            // Call AddOneFileVersion three times to add three new versions to the file.
            this.AddOneFileVersion(this.fileName);

            this.AddOneFileVersion(this.fileName);

            this.AddOneFileVersion(this.fileName);
        }

        /// <summary>
        /// Add one version to the file. 
        /// </summary>
        /// <param name="specifiedFileName">A string indicates file name.</param>
        public void AddOneFileVersion(string specifiedFileName)
        {
            Uri requestUri = new Uri(this.requestUrl);
            Uri fileUrl = AdapterHelper.ConstructDocFileFullUrl(requestUri, this.documentLibrary, specifiedFileName);
            bool isCheckOutFile = this.listsSutControlAdaterInstance.CheckoutFile(fileUrl);
            this.site.Assert.IsTrue(
                isCheckOutFile,
                "CheckOutFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isCheckOutFile);

            bool isCheckInFile = this.listsSutControlAdaterInstance.CheckInFile(
                fileUrl,
                TestSuiteHelper.FileComments, 
                ((int)VersionType.MinorCheckIn).ToString());
            this.site.Assert.IsTrue(
                isCheckInFile, 
                "CheckInFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isCheckInFile);
        }

        /// <summary>
        /// Verify the response results.
        /// </summary>
        /// <param name="results">The results information from server.</param>
        /// <param name="operationName">The operation name for the server response results.</param>
        /// <param name="versioningEnabled">A Boolean indicates whether the versioning is enabled.</param>
        public void VerifyResultsInformation(Results results, OperationName operationName, bool versioningEnabled)
        {
            this.site.Assert.AreNotEqual<string>(
                string.Empty,
                this.listId,
                "GetListID operation returns {0}, the non-empty value means the operation was executed successfully" +
                " and the empty value means the operation failed",
                this.listId);

            // Verify MS-VERSS requirement: MS-VERSS_R39
            this.site.CaptureRequirementIfAreEqual<string>(
                this.listId.ToUpper(System.Globalization.CultureInfo.CurrentCulture),
                results.list.id.ToUpper(System.Globalization.CultureInfo.CurrentCulture),
                39,
                @"[In Results] list.id: Specifies the GUID of the document library in which the file resides.");

            if (versioningEnabled == true)
            {
                // Verify MS-VERSS requirement: MS-VERSS_R44
                this.site.CaptureRequirementIfAreEqual<byte>(
                    1,
                    results.versioning.enabled,
                    44,
                    @"[In Results] versioning.enabled: A value of ""1"" indicates that versioning is enabled.");
            }
            else
            {
                // Verify MS-VERSS requirement: MS-VERSS_R43
                this.site.CaptureRequirementIfAreEqual<byte>(
                    0,
                    results.versioning.enabled,
                    43,
                    @"[In Results] versioning.enabled: A value of ""0"" indicates that versioning is disabled.");
            }

            // Verify version data.
            this.VerifyVersionData(results.result, operationName);
        }

        /// <summary>
        /// Clean up test environment.
        /// </summary>
        public void CleanupTestEnvironment()
        {
            #region Clean the server
            bool isDeleteList = this.listsSutControlAdaterInstance.DeleteList(this.documentLibrary);
            this.site.Assert.IsTrue(
                isDeleteList,
                "DeleteList operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isDeleteList);

            bool isDeleteFromRecycleBin = this.sutControlAdapterInstance.DeleteItemsInListFromRecycleBin(this.documentLibrary);
            this.site.Assert.IsTrue(
                isDeleteFromRecycleBin,
                "DeleteItemsInListFromRecycleBin operation returns {0}," +
                " TRUE means the operation was executed successfully, FALSE means the operation failed",
                isDeleteFromRecycleBin);
            #endregion
        }

        /// <summary>
        /// Get the version information from server response results.
        /// </summary>
        /// <param name="versionData">The version data from server response results.</param>
        /// <returns>The version information in VersionData of response result.</returns>
        private static string TransformVersionDataToString(VersionData[] versionData)
        {
            string responseVersionData = string.Empty;
            bool firstCount = true;

            foreach (VersionData data in versionData)
            {
                if (firstCount)
                {
                    responseVersionData = data.version;
                    firstCount = false;
                }
                else
                {
                    responseVersionData += "^" + data.version;
                }
            }

            return responseVersionData;
        }

        /// <summary>
        /// Verify the version data information from server response results.
        /// </summary>
        /// <param name="versionData">The version data from server response results.</param>
        /// <param name="operationName">The operation name for the server response results.</param>
        private void VerifyVersionData(VersionData[] versionData, OperationName operationName)
        {
            bool areVersionsResultEqual = false;

            string fileVersionsFromSUT = string.Empty;

            // The file versions get from GetVersions response
            VersionData[] fileVersionsFromGetVersions = null;

            // If the VersionData are got from GetVersions operation, use SUT method GetFileVersions to
            // verify the VersionData. Else use protocol method GetVersions to verify the VersionData.
            if (operationName == OperationName.GetVersions)
            {
                fileVersionsFromSUT = this.sutControlAdapterInstance.GetFileVersions(this.documentLibrary, this.fileName);
                this.site.Assert.AreNotEqual<string>(
                    string.Empty,
                    fileVersionsFromSUT,
                    "GetFileVersions operation returns {0}, the not empty value means the operation executed" +
                    " successfully and the empty value means the operation failed",
                    fileVersionsFromSUT);

                // Verify that the result element returned from server using
                // Protocol Adapter equals the result element returned from server using SUT Control Adapter.
                areVersionsResultEqual = AdapterHelper.AreVersionsResultEqual(fileVersionsFromSUT, versionData);
            }
            else
            {
                // Get the relative filename of the file.
                string docRelativeUrl = this.documentLibrary + "/" + this.fileName;

                // Get the versions information from server.
                GetVersionsResponseGetVersionsResult getVersionsResponse = 
                    this.protocolAdapterInstance.GetVersions(docRelativeUrl);
                fileVersionsFromGetVersions = getVersionsResponse.results.result;

                // Verify that the result element returned from server equals 
                // the result element returned from GetVersions response.
                areVersionsResultEqual = AdapterHelper.AreVersionsResultEqual(fileVersionsFromGetVersions, versionData);
            }

            // Verify MS-VERSS requirement: MS-VERSS_R48 
            this.site.CaptureRequirementIfIsTrue(
                areVersionsResultEqual,
                48,
                @"[In Results] result: A separate result element MUST exist for each version of the file that the user can access.");

            string responseVersionData = TransformVersionDataToString(versionData);

            switch (operationName)
            {
                case OperationName.GetVersions:

                    // Add the debug information
                    this.site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual response of GetVersions is {0} and the expected response of GetVersions is {1}",
                        responseVersionData,
                        fileVersionsFromSUT);

                    // Verify MS-VERSS requirement: MS-VERSS_R131
                    this.site.CaptureRequirementIfIsTrue(
                        areVersionsResultEqual,
                        131,
                        @"[In GetVersionsResponse] GetVersionsResult: An XML node that conforms to the structure specified in section 2.2.4.1 and that contains the details about all the versions of the specified file that the user can access.");
                    break;

                case OperationName.DeleteVersion:
                    // Add the debug information 
                    this.site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual response of DeleteVersion is {0} and the expected response of DeleteVersion is {1}",
                        responseVersionData,
                        TransformVersionDataToString(fileVersionsFromGetVersions));

                    // Verify MS-VERSS requirement: MS-VERSS_R112
                    this.site.CaptureRequirementIfIsTrue(
                        areVersionsResultEqual,
                        112,
                        @"[In DeleteVersionResponse] DeleteVersionResult: An XML node that conforms to the structure specified in section 2.2.4.1 and that contains the details about all the versions of the specified file that the user can access.");
                    break;

                case OperationName.RestoreVersion:
                    // Add the debug information
                    this.site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual response of RestoreVersion is {0} and the expected response of RestoreVersion is {1}",
                        responseVersionData,
                        TransformVersionDataToString(fileVersionsFromGetVersions));

                    // Verify MS-VERSS requirement: MS-VERSS_R152
                    this.site.CaptureRequirementIfIsTrue(
                        areVersionsResultEqual,
                        152,
                        @"[In RestoreVersionResponse] RestoreVersionResult: MUST return an XML node that conforms to the structure specified in section 2.2.4.1, which contains the details about all the versions of the specified file that the user can access.");
                    break;

                case OperationName.DeleteAllVersions:
                    // Add the debug information
                    this.site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual response of DeleteAllVersions is {0} and the expected response of DeleteAllVersions is {1}",
                        responseVersionData,
                        TransformVersionDataToString(fileVersionsFromGetVersions));

                    // Verify MS-VERSS requirement: MS-VERSS_R92
                    this.site.CaptureRequirementIfIsTrue(
                        areVersionsResultEqual,
                        92,
                        @"[In DeleteAllVersionsResponse] DeleteAllVersionsResult: An XML node that conforms to the structure specified in section 2.2.4.1 and that contains the  details about all the versions of the specified file that the user can access.");
                    break;
            }

            foreach (VersionData data in versionData)
            {
                string specifiedVersionFileInformation = string.Empty;
                if (this.fileVersionAttributes.ContainsKey(data.version.Replace("@", string.Empty)) == false)
                {
                    // Get data version information from server.
                    specifiedVersionFileInformation = this.sutControlAdapterInstance.GetFileVersionAttributes(
                        this.documentLibrary,
                        this.fileName,
                        data.version.Replace("@", string.Empty));

                    this.fileVersionAttributes.Add(data.version.Replace("@", string.Empty), specifiedVersionFileInformation);
                }
                else
                {
                    specifiedVersionFileInformation = this.fileVersionAttributes[data.version.Replace("@", string.Empty)];
                }

                this.site.Assert.AreNotEqual<string>(
                    string.Empty,
                    specifiedVersionFileInformation,
                    "GetSpecifiedVersionFileInformation operation returns {0}, the non-empty value means the" +
                    " operation was executed successfully and the empty value means the operation failed", 
                    specifiedVersionFileInformation);

                string[] attributes = specifiedVersionFileInformation.Split(new string[] { "^" }, StringSplitOptions.None);

                string expectedCreatedByName = attributes[0];
                expectedCreatedByName = expectedCreatedByName.Substring(expectedCreatedByName.IndexOf("\\", System.StringComparison.CurrentCulture) + 1);
                ulong expectedSize = ulong.Parse(attributes[1]);

                // According to Open Specification section 2.2.4.3, the createdByName attribute is an optional attribute.
                if (!string.IsNullOrEmpty(data.createdByName))
                {
                    // Verify MS-VERSS requirement: MS-VERSS_R165
                    this.site.CaptureRequirementIfAreEqual<string>(
                        expectedCreatedByName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                        data.createdByName.Substring(data.createdByName.IndexOf("\\", System.StringComparison.CurrentCulture) + 1).ToLower(System.Globalization.CultureInfo.CurrentCulture),
                        165,
                        @"[In VersionData] createdByName: The display name of the creator of the version of the file.");
                }

                // Verify MS-VERSS requirement: MS-VERSS_R64
                this.site.CaptureRequirementIfAreEqual<ulong>(
                    expectedSize,
                    data.size,
                    64,
                    @"[In VersionData] size: The size, in bytes, of the version of the file.");
            }
        }
    }
}