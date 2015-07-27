//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System;
    using System.Collections;
    using System.Collections.Specialized;
    using System.Net;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Traditional tests for PropFindExtension scenario.
    /// </summary>
    [TestClass]
    public class S03_PropFindExtension : TestSuiteBase
    {
        #region Constant members

        /// <summary>
        /// The default time stamp that is defined in the open specification
        /// </summary>
        private const string DefaultTimeStamp = "1969-01-01T12:00:00Z";

        #endregion Constant members

        #region ClassInitialize method

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        #endregion

        #region ClassCleanup method

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion

        #region PropFindExtension test cases

        /// <summary>
        /// When the server receives a PROPFIND request with the "Repl:collblob" element set to a timestamp, 
        /// it includes a "response" element for each resource in the "multistatus" element. 
        /// This test case is used to verify that each "response" element is a descendant of the "Request-URI" in the PROPFIND request. 
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S03_TC01_PropFindExtension_Resource()
        {
            // Get the request URI from the property "Server_DefaultDocLibUri", the URI should be collection resource end with "/";
            // such as: http://SUT01/sites/WDVMODUU/shared20%documents/
            string requestUri = Common.GetConfigurationPropertyValue("Server_DefaultDocLibUri", this.Site);

            // Construct HTTP headers, set the value of "Depth" is "infinity".
            NameValueCollection headersCollection = S03_PropFindExtension.ConstructHttpHeaders("infinity");

            // Construct HTTP body with default time stamp "1969-01-01T12:00:00Z".
            string body = S03_PropFindExtension.ConstructHttpBody(DefaultTimeStamp);

            // Call HTTP PROPFIND method with above settings
            WDVMODUUResponse response = this.Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            XmlDocument xmlDoc = response.BodyXmlData;

            // Get valid resource from the response.
            ArrayList resourceList = this.GetValidResourceUnderMultistatusElement(xmlDoc);

            // Capture MS-WDVMODUU_R101, if in the HTTP response, each "response" element is a descendant of the request URI in the PROPFIND request.
            this.Site.Assert.IsTrue((resourceList != null) && (resourceList.Count > 0), "There should be some valid resource in the response!");
            bool doesCaptureRequirement_101 = true;
            IEnumerator enumeratorResInfo = resourceList.GetEnumerator();
            while (enumeratorResInfo.MoveNext())
            {
                ResourceInfo resInfo = (ResourceInfo)enumeratorResInfo.Current;
                string href = resInfo.Href;

                // In some protocol server, the return URI for collection resource end with "/"; 
                // In some protocol server, the return URI for collection resource end without "/";
                //  The example for such URI can be: 
                //              http://SUT01/sites/WDVMODUU/shared20%documents/
                //  or          http://SUT01/sites/WDVMODUU/shared20%documents
                //  The two URI identify the same collection resource.
                // So following codes use two check results to make sure if the return resource is under the request URI in the PROPFIND request.
                int index1 = href.IndexOf(requestUri, StringComparison.OrdinalIgnoreCase);
                int index2 = (href + "/").IndexOf(requestUri, StringComparison.OrdinalIgnoreCase);
                if ((index1 < 0) && (index2 < 0))
                {
                    // If the return resource is NOT under the request URI in the PROPFIND request, report the error.
                    doesCaptureRequirement_101 = false;
                    this.Site.Log.Add(
                        LogEntryKind.TestError,
                        "The return resource '{0}' is not under the request URI '{1}' in the PROPFIND request",
                        href,
                        requestUri);
                }
            }

            this.Site.Assert.IsTrue(doesCaptureRequirement_101, "The return resource should under the request URI in the PROPFIND request.");

            // If the return resource under the request URI in the PROPFIND request, the MS-WDVMODUU_R101 should be directly covered.
            this.Site.CaptureRequirement(
                101,
                @"[In Repl:collblob and Repl:repl] When the server receives a PROPFIND request with the Repl:collblob element set to a timestamp, [it includes a response element for each resource in the multistatus element] that is a descendant of the Request-URI");
        }

        /// <summary>
        /// When the server receives a PROPFIND request with the "Repl:collblob" element set to a timestamp, 
        /// it includes a "response" element for each resource in the "multistatus" element. 
        /// And these return resources should conform to following rules:
        ///    Rule 1: The resource was last modified later than 5 minutes before the timestamp. 
        /// OR Rule 2: The resource is a descendant of a resource that was last modified later than 5 minutes before the timestamp.
        /// This test case is used to verify the Rule 1.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S03_TC02_PropFindExtension_Resource_5Minutes()
        {
            // Step 1. 
            // Call HTTP PROPFIND method with above settings
            //    Set "Depth" header to "0";
            //    Set the time stamp to the default time stamp "1969-01-01T12:00:00Z";
            //    Set the "Request-URI" from the property "Server_NewFile001Uri", it is a non-collection resource;
            string requestURI = Common.GetConfigurationPropertyValue("Server_NewFile001Uri", this.Site);
            NameValueCollection headersCollection = S03_PropFindExtension.ConstructHttpHeaders("0");
            string body = S03_PropFindExtension.ConstructHttpBody(DefaultTimeStamp);
            WDVMODUUResponse response = this.Adapter.PropFind(requestURI, body, headersCollection);

            // Get XML data from the response of Step 1.
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            XmlDocument xmlDoc = response.BodyXmlData;

            // Get valid resource from the response of Step 1.
            ArrayList resourceList = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList != null) && (resourceList.Count > 0), "There should be some valid resource in the response!");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList.Count = {0}", resourceList.Count);
            this.Site.Assert.IsTrue(resourceList.Count == 1, "There should be only one valid resource in the response!");

            // Step2. Record the last modified date time for the non-collection resource.
            IEnumerator enumratorResInfo = resourceList.GetEnumerator();
            enumratorResInfo.MoveNext();
            ResourceInfo resInfo = (ResourceInfo)enumratorResInfo.Current;
            DateTime timeStamp_LastModified = resInfo.LastModifiedDateTime;

            // Step 3.
            // Call HTTP PROPFIND request again with following settings:
            //    Set "Depth" header to "0";
            //    Set the time stamp to the last modified date time of the non-collection resource plus 5 minutes;
            //    Set the "Request-URI" from the property "Server_NewFile001Uri", it is a non-collection resource;
            DateTime timeStamp_LastModified_Plus5Mins = timeStamp_LastModified.AddMinutes(5);
            requestURI = Common.GetConfigurationPropertyValue("Server_NewFile001Uri", this.Site);
            body = S03_PropFindExtension.ConstructHttpBody(S03_PropFindExtension.GetUtcFormatString(timeStamp_LastModified_Plus5Mins));
            response = this.Adapter.PropFind(requestURI, body, headersCollection);

            // Get XML data from the response of Step 3.
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            xmlDoc = response.BodyXmlData;

            // Get valid resource from the response of Step 3.
            resourceList = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList != null) && (resourceList.Count > 0), "There should be some valid resource in the response!");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList.Count = {0}", resourceList.Count);
            this.Site.Assert.IsTrue(resourceList.Count == 1, "There should be only one valid resource in the response!");

            // Step 4.
            // In the response of step 3, make sure the non-collection resource is returned.
            enumratorResInfo = resourceList.GetEnumerator();
            enumratorResInfo.MoveNext();
            ResourceInfo resInfo_5Min = (ResourceInfo)enumratorResInfo.Current;
            bool findTheResource_5Min = false;
            if (string.Compare(requestURI, resInfo_5Min.Href, true) == 0)
            {
                findTheResource_5Min = true;
            }

            // Step 5.
            // Call HTTP PROPFIND request in the third time with following settings:
            //    Set "Depth" header to "0";
            //    Set the time stamp to the last modified date time of the non-collection resource plus 5 minutes and 1 second;
            //    Set the "Request-URI" from the property "Server_NewFile001Uri", it is a non-collection resource;
            DateTime timeStamp_LastModified_Plus5MinsPlus1Second = timeStamp_LastModified.AddSeconds((5 * 60) + 1);
            requestURI = Common.GetConfigurationPropertyValue("Server_NewFile001Uri", this.Site);
            body = S03_PropFindExtension.ConstructHttpBody(S03_PropFindExtension.GetUtcFormatString(timeStamp_LastModified_Plus5MinsPlus1Second));
            response = this.Adapter.PropFind(requestURI, body, headersCollection);

            // Get XML data from the response of Step 5.
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            xmlDoc = response.BodyXmlData;

            // Get valid resource from the response of Step 5.
            resourceList = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList != null) && (resourceList.Count >= 0), "The valid resource list should be return.");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList.Count = {0}", resourceList.Count);
            this.Site.Assert.IsTrue(resourceList.Count == 0, "There should be no valid resource in the response!");

            // Step 6.
            // In the response of step 5, make sure the non-collection resource is NOT returned.
            bool doesNotFindTheResource_5MinPlus1Sec = false;
            if (resourceList.Count == 0)
            {
                doesNotFindTheResource_5MinPlus1Sec = true;
            }

            // Step 7.
            // If the non-collection resource is returned in step 3 and NOT returned in step 5, then capture MS-WDVMODUU_R99, and MS-WDVMODUU_R104.
            this.Site.CaptureRequirementIfIsTrue(
                findTheResource_5Min && doesNotFindTheResource_5MinPlus1Sec,
                99,
                @"[In Repl:collblob and Repl:repl] The existence of a Repl:collblob element in a PROPFIND request restricts the set of results returned by the server.");
            this.Site.CaptureRequirementIfIsTrue(
                findTheResource_5Min && doesNotFindTheResource_5MinPlus1Sec,
                104,
                @"[In Repl:collblob and Repl:repl] [When the server receives a PROPFIND request with the Repl:collblob element set to a timestamp, it includes a response element for each resource in the multistatus element that is a descendant of the Request-URI (limited by the Depth header specified in [RFC2518]) and that has changed according to the rule:] The resource was last modified later than 5 minutes before the timestamp.");
        }

        /// <summary>
        /// This test case is used to verify that when the server receives a PROPFIND request with the "Repl:collblob" element set to a timestamp, 
        /// it includes a "response" element for each resource in the "multistatus" element, 
        /// and the resource is limited by the Depth header specified in [RFC2518].
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S03_TC03_PropFindExtension_Resource_Depth()
        {
            // In the server, the following resource structure is existed before running this test case.
            // [Root-Folder]
            // --- New_File001.txt
            // --- New_File002.txt
            // --- [Sub-Folder]
            //     --- New_File003.txt

            // Get all expected URI from the properties.
            string expectedRootFolderUri = Common.GetConfigurationPropertyValue("Server_DefaultDocLibUri", this.Site);
            string expectedSubFolderUri = Common.GetConfigurationPropertyValue("Server_SubFolderUri", this.Site);
            string expectedNewFile1Uri = Common.GetConfigurationPropertyValue("Server_NewFile001Uri", this.Site);
            string expectedNewFile2Uri = Common.GetConfigurationPropertyValue("Server_NewFile002Uri", this.Site);
            string expectedNewFile3Uri = Common.GetConfigurationPropertyValue("Server_NewFile003Uri", this.Site);

            ArrayList expectedSourceList = new ArrayList();
            bool doesCaptureRequirement_100 = true;
            bool doesCaptureRequirement_102 = true;

            // Step1.
            // Call HTTP PROPFIND request with following settings:
            //  Set "Depth" header to "Infinity";
            //  Set the time stamp to the default time stamp "1969-01-01T12:00:00Z";
            //  Set the "Request-URI" to "[Root-Folder]".
            string requestUri = Common.GetConfigurationPropertyValue("Server_DefaultDocLibUri", this.Site);
            NameValueCollection headersCollection = S03_PropFindExtension.ConstructHttpHeaders("infinity");
            string body = S03_PropFindExtension.ConstructHttpBody(DefaultTimeStamp);
            WDVMODUUResponse response = this.Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response of Step 1.
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            XmlDocument xmlDoc = response.BodyXmlData;

            // Get valid resource list from the response of Step 1.
            ArrayList resourceList = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList != null) && (resourceList.Count > 0), "There should be some valid resource in the response!");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList.Count = {0}", resourceList.Count);
            if ((resourceList == null) || (resourceList.Count == 0))
            {
                // If there is no valid resource in the response, then we can't capture MS-WDVMODUU_R100.
                doesCaptureRequirement_100 = false;
            }

            // Step2.
            // In the response of step 1, make sure following "resource" are returned:
            //   [Root-Folder]
            //   New_File001.txt
            //   New_File002.txt
            //   [Sub-Folder]
            //   New_File003.txt
            bool isSuccessful = false;
            expectedSourceList.Clear();
            expectedSourceList.Add(expectedRootFolderUri);
            expectedSourceList.Add(expectedSubFolderUri);
            expectedSourceList.Add(expectedNewFile1Uri);
            expectedSourceList.Add(expectedNewFile2Uri);
            expectedSourceList.Add(expectedNewFile3Uri);
            foreach (object obj in expectedSourceList)
            {
                string expectedSource = (string)obj;

                // Make sure the expected source is existed in the valid resource list.
                isSuccessful = this.FindSpecialResource(resourceList, expectedSource);
                if (!isSuccessful)
                {
                    doesCaptureRequirement_102 = false;
                    break;
                }
            }

            // Assert the expected source is existed in the valid resource list.
            this.Site.Assert.IsTrue(isSuccessful, "All expected resource should be existed in the return valid resource list.");

            // Step 3.
            // Call HTTP PROPFIND request with following settings:
            //   Set "Depth" header to "1";
            //   Set the time stamp to the default time stamp "1969-01-01T12:00:00Z";
            //   Set the "Request-URI" to "[Root-Folder]".
            requestUri = Common.GetConfigurationPropertyValue("Server_DefaultDocLibUri", this.Site);
            headersCollection = S03_PropFindExtension.ConstructHttpHeaders("1");
            body = S03_PropFindExtension.ConstructHttpBody(DefaultTimeStamp);
            response = this.Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response of Step 3.
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            xmlDoc = response.BodyXmlData;

            // Get valid resource list from the response of Step 3.
            resourceList = null;
            resourceList = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList != null) && (resourceList.Count > 0), "There should be some valid resource in the response!");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList.Count = {0}", resourceList.Count);
            if ((resourceList == null) || (resourceList.Count == 0))
            {
                // If there is no valid resource in the response, then we can't capture MS-WDVMODUU_R100.
                doesCaptureRequirement_100 = false;
            }

            // Step 4.
            // In the response of step 3, make sure following "resource" are returned:
            //   [Root-Folder]
            //   New_File001.txt
            //   New_File002.txt
            //   [Sub-Folder]
            expectedSourceList.Clear();
            expectedSourceList.Add(expectedRootFolderUri);
            expectedSourceList.Add(expectedNewFile1Uri);
            expectedSourceList.Add(expectedNewFile2Uri);
            expectedSourceList.Add(expectedSubFolderUri);
            foreach (object obj in expectedSourceList)
            {
                string expectedSource = (string)obj;

                // Make sure the expected source is existed in the valid resource list.
                isSuccessful = this.FindSpecialResource(resourceList, expectedSource);
                if (!isSuccessful)
                {
                    doesCaptureRequirement_102 = false;
                    break;
                }
            }

            // Assert the expected source is existed in the valid resource list.
            this.Site.Assert.IsTrue(isSuccessful, "All expected resource should be existed in the return valid resource list.");

            // Step 5.
            // Call HTTP PROPFIND request with following settings:
            //  Set "Depth" header to "0";
            //  Set the time stamp to the default time stamp "1969-01-01T12:00:00Z";
            //  Set the "Request-URI" to "New_File001.txt".
            requestUri = Common.GetConfigurationPropertyValue("Server_NewFile001Uri", this.Site);
            body = S03_PropFindExtension.ConstructHttpBody(DefaultTimeStamp);
            headersCollection = S03_PropFindExtension.ConstructHttpHeaders("0");
            response = this.Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response of Step 5.
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            xmlDoc = response.BodyXmlData;

            // Get valid resource list from the response of Step 5.
            resourceList = null;
            resourceList = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList != null) && (resourceList.Count > 0), "There should be some valid resource in the response!");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList.Count = {0}", resourceList.Count);
            if ((resourceList == null) || (resourceList.Count == 0))
            {
                // If there is no valid resource in the response, then we can't capture MS-WDVMODUU_R100.
                doesCaptureRequirement_100 = false;
            }

            // Assert there is only one valid resource in the response.
            this.Site.Assert.IsTrue(resourceList.Count == 1, "There should be only one valid resource in the response.");

            // Step 5.
            // In the response of step 5, make sure only "resource" "New_File001.txt" is returned.
            if (resourceList.Count != 1)
            {
                doesCaptureRequirement_102 = false;
            }
            else
            {
                IEnumerator enumratorResInfo = resourceList.GetEnumerator();
                enumratorResInfo.MoveNext();
                ResourceInfo resInfo = (ResourceInfo)enumratorResInfo.Current;
                if (string.Compare(resInfo.Href, requestUri, true) != 0)
                {
                    doesCaptureRequirement_102 = false;
                }
            }

            // Step 7.
            // If all responses each resource is under one "response" element , then capture MS-WDVMODUU_R100.
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureRequirement_100,
                100,
                @"[In Repl:collblob and Repl:repl] When the server receives a PROPFIND request with the Repl:collblob element set to a timestamp, it includes a response element for each resource in the multistatus element");

            // If all responses in step1, step3 and step5 are same as expected, then capture MS-WDVMODUU_R102.
            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureRequirement_102,
                102,
                @"[In Repl:collblob and Repl:repl] When the server receives a PROPFIND request with the Repl:collblob element set to a timestamp, [it includes a response element for each resource in the multistatus element that is a descendant of the Request-URI] (limited by the Depth header specified in [RFC2518])");
        }

        /// <summary>
        /// When the server receives a PROPFIND request with the "Repl:collblob" element set to a timestamp, 
        /// it includes a response element for each resource in the "multistatus" element.
        /// And these return resources should conform to following rules:
        ///    Rule 1: The resource was last modified later than 5 minutes before the timestamp.
        /// OR Rule 2: The resource is a descendant of a resource that was last modified later than 5 minutes before the timestamp.
        /// This test case is used to verify the Rule 2.
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S03_TC04_PropFindExtension_Resource_Descendant()
        {
            // In the server, the following resource structure is existed before running this test case.
            //   [Root-Folder]
            //    --- [Sub-Folder]
            //          --- New_File003.txt
            bool isSuccessful = false;
            bool doesCaptureRequirement_105 = false;
            string destinationUri_File004 = string.Empty;
            string newFilePath_File004 = string.Empty;

            // Step1. 
            // Call HTTP PUT request with following settings:
            //  - Set the "Request-URI" to "[Sub-Folder]\New_File004.txt".
            destinationUri_File004 = Common.GetConfigurationPropertyValue("Server_NewFile004Uri", this.Site);
            newFilePath_File004 = Common.GetConfigurationPropertyValue("Client_NewFile004Name", this.Site);
            isSuccessful = this.PutNewFileIntoServer(destinationUri_File004, newFilePath_File004);
            Site.Assert.IsTrue(isSuccessful, "Upload file New_File004 should be successful.");
            this.ArrayListForDeleteFile.Add((object)destinationUri_File004);

            // Step2.
            // Call HTTP PROPFIND request with following settings:
            //  - Set "Depth" header to "1";
            //  - Set the time stamp to the default time stamp "1969-01-01T12:00:00Z";
            //  - Set the "Request-URI" to "[Sub-Folder]".
            string requestUri = Common.GetConfigurationPropertyValue("Server_SubFolderUri", this.Site);
            NameValueCollection headersCollection = S03_PropFindExtension.ConstructHttpHeaders("1");
            string body = S03_PropFindExtension.ConstructHttpBody(DefaultTimeStamp);
            WDVMODUUResponse response = Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            XmlDocument xmlDoc = response.BodyXmlData;

            // Get valid resource list from the response of Step 2.
            ArrayList resourceList_1 = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList_1 != null) && (resourceList_1.Count > 0), "There should be some valid resource in the response!");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList_1.Count = {0}", resourceList_1.Count);

            // Step3.
            // Based on the response of step 2, calculate the critical time stamp for resource "[Sub-Folder]" and resource "New_File003.txt". 
            // The critical time stamp is the last modified date time of the resource plus 5 minutes and 1 second.
            DateTime criticalTimeStamp_SubFolder = DateTime.MinValue;
            DateTime criticalTimeStamp_New_File003 = DateTime.MinValue;
            string expectedHref_SubFolder = requestUri;
            string expectedHref_New_File003 = Common.GetConfigurationPropertyValue("Server_NewFile003Uri", this.Site);
            IEnumerator enumratorResInfo = resourceList_1.GetEnumerator();
            while (enumratorResInfo.MoveNext())
            {
                ResourceInfo resInfo = (ResourceInfo)enumratorResInfo.Current;

                // Try to find the resource "New_File003.txt",and calculate the critical time stamp.
                if (string.Compare(resInfo.Href, expectedHref_New_File003, true) == 0)
                {
                    this.Site.Log.Add(LogEntryKind.Comment, "The last modified data time for 'NewFile003' is {0}.", resInfo.LastModifiedDateTime);
                    criticalTimeStamp_New_File003 = resInfo.LastModifiedDateTime.AddSeconds((5 * 60) + 1);
                }

                // Try to find the resource "[Sub-Folder]", and calculate the critical time stamp.
                // In some protocol server, the return URI for collection resource end with "/"; 
                // In some protocol server, the return URI for collection resource end without "/";
                //  The example for such URI can be: 
                //              http://SUT01/sites/WDVMODUU/shared20%documents/SubFolder/
                //  or          http://SUT01/sites/WDVMODUU/shared20%documents/SubFolder
                //  The two URI identify the same collection resource.
                // So following codes use two compare results to check if the resource "[Sub-Folder]" is existed in the resource list.
                if ((string.Compare(resInfo.Href, expectedHref_SubFolder, true) == 0)
                    || (string.Compare(resInfo.Href + "/", expectedHref_SubFolder, true) == 0))
                {
                    this.Site.Log.Add(LogEntryKind.Comment, "The last modified data time for '[Sub-Folder]' is {0}.", resInfo.LastModifiedDateTime);
                    criticalTimeStamp_SubFolder = resInfo.LastModifiedDateTime.AddSeconds((5 * 60) + 1);
                }
            }

            this.Site.Assert.IsTrue(criticalTimeStamp_SubFolder != DateTime.MinValue, "The resource \"[Sub-Folder]\" should be  existed in the resource list.");
            this.Site.Assert.IsTrue(criticalTimeStamp_New_File003 != DateTime.MinValue, "The resource \"New_File003.txt\" should be  existed in the resource list");

            // Step4.
            //  Call HTTP PROPFIND request with following settings:
            //    - Set "Depth" header to "1";
            //    - Set the time stamp to the critical time stamp for resource "New_File003.txt";
            //    - Set the "Request-URI" to "[Sub-Folder]".
            requestUri = Common.GetConfigurationPropertyValue("Server_SubFolderUri", this.Site);
            headersCollection = S03_PropFindExtension.ConstructHttpHeaders("1");
            body = S03_PropFindExtension.ConstructHttpBody(S03_PropFindExtension.GetUtcFormatString(criticalTimeStamp_New_File003));
            response = this.Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            xmlDoc = response.BodyXmlData;

            // Get valid resource list from the response of Step 4.
            ArrayList resourceList_2 = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Assert.IsTrue((resourceList_2 != null) && (resourceList_2.Count > 0), "There should be some valid resource in the response!");
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList_2.Count = {0}", resourceList_2.Count);

            // Step5.
            // In the response of step 4, make sure the resource "New_File003.txt" is still returned based on Rule 2. 
            // Because its parent resource "[Sub-Folder]" is returned based on Rule 1.
            bool findNewFile003DueToRule2 = false;
            enumratorResInfo = resourceList_2.GetEnumerator();
            while (enumratorResInfo.MoveNext())
            {
                ResourceInfo resInfo = (ResourceInfo)enumratorResInfo.Current;
                if (string.Compare(resInfo.Href, expectedHref_New_File003, true) == 0)
                {
                    findNewFile003DueToRule2 = true;
                    break;
                }
            }

            this.Site.Assert.IsTrue(findNewFile003DueToRule2, "The resource \"New_File003.txt\" should be returned based on Rule 2.");

            // Step6.
            // Call HTTP PROPFIND request with following settings:
            //  - Set "Depth" header to "1";
            //  - Set the time stamp to the critical time stamp for resource "[Sub-Folder]";
            //  - Set the "Request-URI" to "[Sub-Folder]".
            requestUri = Common.GetConfigurationPropertyValue("Server_SubFolderUri", this.Site);
            headersCollection = S03_PropFindExtension.ConstructHttpHeaders("1");
            body = S03_PropFindExtension.ConstructHttpBody(S03_PropFindExtension.GetUtcFormatString(criticalTimeStamp_SubFolder));
            response = this.Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            xmlDoc = response.BodyXmlData;

            // Get valid resource list from the response of Step 6.
            ArrayList resourceList_3 = this.GetValidResourceUnderMultistatusElement(xmlDoc);
            this.Site.Log.Add(LogEntryKind.Comment, "resourceList_3.Count = {0}", resourceList_3.Count);

            // Step7.
            // In the response of step 6, make sure resource "[Sub-Folder]", "New_File003.txt" and "New_File004.txt" are all not returned.
            enumratorResInfo = resourceList_3.GetEnumerator();
            bool doesNotReturnSubFolder = true;
            bool doesNotReturnNewFile003 = true;
            bool doesNotReturnNewFile004 = true;
            string expectedHref_New_File004 = destinationUri_File004;
            while (enumratorResInfo.MoveNext())
            {
                ResourceInfo resInfo = (ResourceInfo)enumratorResInfo.Current;

                // Try to find the resource "[Sub-Folder]".
                // In some protocol server, the return URI for collection resource end with "/"; 
                // In some protocol server, the return URI for collection resource end without "/";
                //  The example for such URI can be: 
                //              http://SUT01/sites/WDVMODUU/shared20%documents/SubFolder/
                //  or          http://SUT01/sites/WDVMODUU/shared20%documents/SubFolder
                //  The two URI identify the same collection resource.
                // So following codes use two compare results to check if the resource "[Sub-Folder]" is existed in the resource list.
                if ((string.Compare(resInfo.Href, expectedHref_SubFolder, true) == 0)
                    || (string.Compare(resInfo.Href + "/", expectedHref_SubFolder, true) == 0))
                {
                    doesNotReturnSubFolder = false;
                }

                // Try to find the resource "New_File003.txt".
                if (string.Compare(resInfo.Href, expectedHref_New_File003, true) == 0)
                {
                    doesNotReturnNewFile003 = false;
                }

                // Try to find the resource "New_File004.txt".
                if (string.Compare(resInfo.Href, expectedHref_New_File004, true) == 0)
                {
                    doesNotReturnNewFile004 = false;
                }
            }

            this.Site.Assert.IsTrue(doesNotReturnSubFolder, "The resource \"[Sub-Folder]\" should not be return.");
            this.Site.Assert.IsTrue(doesNotReturnNewFile003, "The resource \"New_File003.txt\" should not be return.");
            this.Site.Assert.IsTrue(doesNotReturnNewFile004, "The resource \"New_File004.txt\" should not be return.");

            // Step8.
            // Call HTTP DELETE method to delete the test file "New_File004.txt" that uploaded in the step 1.
            isSuccessful = this.DeleteTheFileInTheServer(destinationUri_File004);
            this.Site.Assert.IsTrue(isSuccessful, "Delete the test file \"New_File004.txt\" should be successful. ");
            this.RemoveFileUriFromDeleteList(destinationUri_File004);

            // Step9.
            // If the response "New_File003.txt" is return in the response of step 4, 
            // and resource "[Sub-Folder]", "New_File003.txt" and "New_File004.txt" are all not return in the response of step 6 ,
            // then capture MS-WDVMODUU_R105.
            if (findNewFile003DueToRule2 && doesNotReturnSubFolder && doesNotReturnNewFile003 && doesNotReturnNewFile004)
            {
                doesCaptureRequirement_105 = true;
            }

            this.Site.CaptureRequirementIfIsTrue(
                doesCaptureRequirement_105,
                105,
                @"[In Repl:collblob and Repl:repl] [When the server receives a PROPFIND request with the Repl:collblob element set to a timestamp, it includes a response element for each resource in the multistatus element that is a descendant of the Request-URI (limited by the Depth header specified in [RFC2518]) and that has changed according to the rule:] The resource is a descendant of a resource that has changed.");
        }

        /// <summary>
        /// This test case is used to verify the relevant requirements about the element collection "Repl:repl" and the element "Repl:collblob".
        /// </summary>
        [TestCategory("MSWDVMODUU"), TestMethod()]
        public void MSWDVMODUU_S03_TC05_PropFindExtension_Repl()
        {
            // Call HTTP PROPFIND request with following settings:
            //   Set "Depth" header to "1";
            //   Set the time stamp to the default time stamp " 1969-01-01T12:00:00Z";
            string requestUri = Common.GetConfigurationPropertyValue("Server_DefaultDocLibUri", this.Site);
            NameValueCollection headersCollection = S03_PropFindExtension.ConstructHttpHeaders("1");
            string body = S03_PropFindExtension.ConstructHttpBody(DefaultTimeStamp);
            WDVMODUUResponse response = this.Adapter.PropFind(requestUri, body, headersCollection);

            // Get XML data from the response of above HTTP PROFIND request.
            this.Site.Assert.IsNotNull(response.BodyXmlData, "The response object 'response.BodyXmlData' should not be null!");
            this.Site.Assert.IsNotNull(response.BodyXmlData.DocumentElement, "The response object 'response.BodyXmlData.DocumentElement' should not be null!");
            this.Site.Assert.IsNotNull(response.BodyXmlData.DocumentElement.LocalName, "The response object 'response.BodyXmlData.DocumentElement.LocalName' should not be null!");
            XmlDocument xmlDoc = response.BodyXmlData;

            this.Site.Assert.IsTrue(
                xmlDoc.DocumentElement.LocalName == "multistatus",
                "The element 'multistatus' should be return in the response.");

            string valueOfRepl_collblob = string.Empty;
            bool findRepl_replElementUnderDav_multistatus = false;
            bool findRepl_collblobUnderRepl_repl = false;
            bool findRepl_collblobUnderOtherElement = false;
            bool correctFormat_TimeStamp = false;
            bool onlyOneRepl_collblobUnderRepl_repl = false;
            bool findSpecialSchemaAlias = false;
            if (xmlDoc.DocumentElement.LocalName == "multistatus")
            {
                foreach (XmlNode node1 in xmlDoc.DocumentElement.ChildNodes)
                {
                    // Try to find "Repl:repl" under "DAV:multistatus".
                    if (node1.Name == "Repl:repl")
                    {
                        findRepl_replElementUnderDav_multistatus = true;

                        XmlNode nodeRepl_repl = node1;
                        int countOfRepl_collblob = 0;
                        foreach (XmlNode node2 in nodeRepl_repl.ChildNodes)
                        {
                            // Try to find "Repl:collblob" under "Repl:repl".
                            if (node2.Name == "Repl:collblob")
                            {
                                countOfRepl_collblob++;
                                findRepl_collblobUnderRepl_repl = true;

                                XmlNode nodeRepl_collblob = node2;
                                valueOfRepl_collblob = nodeRepl_collblob.InnerText;

                                // Define a regular expression pattern for the UTC timestamp syntax, for example: "2009-02-18T08:29:05Z"
                                string patternUTC = @"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z";
                                correctFormat_TimeStamp = Regex.IsMatch(valueOfRepl_collblob, patternUTC);
                            }
                        }

                        if (countOfRepl_collblob == 1)
                        {
                            onlyOneRepl_collblobUnderRepl_repl = true;
                        }
                        else
                        {
                            onlyOneRepl_collblobUnderRepl_repl = false;
                            this.Site.Log.Add(LogEntryKind.TestFailed, "Test Failed: There are {0} 'Repl:collblob' under 'Repl:repl' element!", countOfRepl_collblob);
                        }
                    }

                    if (node1.Name == "Repl:collblob")
                    {
                        // If we find "Repl:collblob" element out of the element "Repl:repl", 
                        // we can't capture requirement MS-WDVMODUU_R63, and should report error.
                        findRepl_collblobUnderOtherElement = true;
                    }
                }

                this.Site.Assert.IsTrue(findRepl_replElementUnderDav_multistatus, "There should be a 'Repl:repl' element under 'DAV:multistatus' element.");
                this.Site.Assert.IsTrue(findRepl_collblobUnderRepl_repl, "There should be a 'Repl:collblob' element under 'Repl:repl' element.");
                this.Site.Assert.IsTrue(correctFormat_TimeStamp, "The value of 'Repl:collblob' should a time stamp with correct format.");
                this.Site.Assert.IsFalse(findRepl_collblobUnderOtherElement, "The element 'Repl:collblob' should not under other element.");
                this.Site.Assert.IsTrue(onlyOneRepl_collblobUnderRepl_repl, "There should be only one 'Repl:collblob' under 'Repl_repl' element.");

                // Try to find special schema alias xmlns:Repl="http://schemas.microsoft.com/repl/"
                XmlAttributeCollection attCollection = xmlDoc.DocumentElement.Attributes;
                foreach (XmlAttribute xmlAttribute in attCollection)
                {
                    if ((xmlAttribute.Name == "xmlns:Repl") && (xmlAttribute.InnerText == "http://schemas.microsoft.com/repl/"))
                    {
                        findSpecialSchemaAlias = true;
                        break;
                    }
                }

                this.Site.Assert.IsTrue(findSpecialSchemaAlias, "The special schema alias xmlns:Repl=\"http://schemas.microsoft.com/repl/\" should be in the response.");
            }

            // In the response, if element collection "Repl:repl" is returned under "DAV: multistatus" element, 
            // then capture MS-WDVMODUU_R2, WDVMODUU_R67 and MS-WDVMODUU_R106.
            this.Site.CaptureRequirementIfIsTrue(
                findRepl_replElementUnderDav_multistatus,
                2,
                @"[In Common Data Types] A new XML element [<!ELEMENT multistatus (repl?, response+, responsedescription?) >] is added to the DAV:multistatus element collection, as defined in [RFC2518].");
            this.Site.CaptureRequirementIfIsTrue(
                findRepl_replElementUnderDav_multistatus,
                106,
                @"[In Repl:collblob and Repl:repl] [When the server receives a PROPFIND request with the Repl:collblob element set to a timestamp] In addition, the server includes the Repl:repl element collection in the response as specified.
<!ELEMENT multistatus (repl, response+, responsedescription?) >");
            this.Site.CaptureRequirementIfIsTrue(
                findRepl_replElementUnderDav_multistatus,
                67,
                @"[In Repl:repl Element Collection] [This Repl:repl Element collection appears] within the multistatus element collection (section 2.2.2). <!ELEMENT repl (collblob) >");

            // In the response, if the element "Repl:collblob" is only returned under "Repl:repl", than capture MS-WDVMODUU_R63.
            this.Site.CaptureRequirementIfIsTrue(
                findRepl_collblobUnderRepl_repl && !findRepl_collblobUnderOtherElement,
                63,
                @"[In Repl:collblob Element] The Repl:collblob element MUST NOT appear except within the Repl:repl XML element collection.");

            // In the response, if the value of element "Repl:collblob" conforms to the [ISO-8601] standard, then capture MS-WDVMODUU_R62.
            this.Site.CaptureRequirementIfIsTrue(
                correctFormat_TimeStamp,
                62,
                @"[In Repl:collblob Element] The Repl:collblob XML element MUST contain a UTC timestamp that conforms to the [ISO-8601] standard.
<!ELEMENT collblob (#PCDATA) >");

            // In the response, if only one "Repl:collblob" element under "Repl:repl", then capture MS-WDVMODUU_R65, and MS-WDVMODUU_R67.
            this.Site.CaptureRequirementIfIsTrue(
                onlyOneRepl_collblobUnderRepl_repl,
                65,
                @"[In Repl:repl Element Collection] The Repl:repl XML element collection MUST contain a single Repl:collblob element, as specified in section 2.2.2.1).");

            // In the response, if the attribute of element "DAV:multistatus"includes the schema alias xmlns:Repl="http://schemas.microsoft.com/repl/", 
            // then capture MS-WDVMODUU_R61. 
            this.Site.CaptureRequirementIfIsTrue(
                findRepl_replElementUnderDav_multistatus && findRepl_collblobUnderRepl_repl && findSpecialSchemaAlias,
                61,
                @"[In MODUU Extensions Property] When the Repl:collblob and Repl:repl elements appear in a response to a WebDAV client request, the response MUST also include this schema alias.
xmlns:Repl=""http://schemas.microsoft.com/repl/""");
        }

        #endregion PropFindExtension test cases

        #region Help Methods in Scenario 03

        /// <summary>
        /// Construct the HTTP XML Body that is used in HTTP PROPFIND method request, 
        /// the HTTP XML body includes Repl:repl Element Collection and Repl:collblob Element, 
        /// the value of Repl:collblob Element is the time stamp in the input parameter "timeStamp".
        /// </summary>
        /// <param name="timeStamp">The time stamp that is used in the HTTP body</param>
        /// <returns>Return the whole HTTP body that includes the time stamp.</returns>
        protected static string ConstructHttpBody(string timeStamp)
        {
            string httpBody = "<?xml version=\"1.0\"?>" +
                   "<D:propfind xmlns:D='DAV:' xmlns:Repl=\"http://schemas.microsoft.com/repl/\">" +
                      "<Repl:repl><Repl:collblob>" + timeStamp + "</Repl:collblob></Repl:repl>" +
                      "<D:allprop/>" +
                   "</D:propfind>";
            return httpBody;
        }

        /// <summary>
        /// Construct the HTTP Headers and values that is used in HTTP PROPFIND method request, 
        /// set the value of "Depth" header as the input parameter "depthValue".
        /// </summary>
        /// <param name="depthValue">The value of depth header</param>
        /// <returns>Return the header collections that includes whole HTTP headers and values.</returns>
        protected static NameValueCollection ConstructHttpHeaders(string depthValue)
        {
            NameValueCollection headersCollection = new NameValueCollection();
            headersCollection.Clear();
            headersCollection.Add("Depth", depthValue);
            headersCollection.Add("Cache-Control", "no-cache");
            headersCollection.Add("ContentType", "text/xml");
            headersCollection.Add("Pragma", "no-cache");
            headersCollection.Add("ProtocolVersion", "HTTP/1.1");
            return headersCollection;
        }

        /// <summary>
        /// Construct a string that conforms to the [ISO-8601] standard based on the value of input date time object.
        ///    Such as: 1969-01-01T12:00:00Z
        /// </summary>
        /// <param name="dateTime">The object that includes the date time information</param>
        /// <returns>Return the string that conforms to the [ISO-8601] standard based on the value of input parameter "dataTime".</returns>
        protected static string GetUtcFormatString(DateTime dateTime)
        {
            string year = dateTime.Year.ToString();
            string month = dateTime.Month.ToString();
            if (dateTime.Month < 10)
            {
                month = "0" + month;
            }

            string day = dateTime.Day.ToString();
            if (dateTime.Day < 10)
            {
                day = "0" + day;
            }

            string hour = dateTime.Hour.ToString();
            if (dateTime.Hour < 10)
            {
                hour = "0" + hour;
            }

            string minute = dateTime.Minute.ToString();
            if (dateTime.Minute < 10)
            {
                minute = "0" + minute;
            }

            string second = dateTime.Second.ToString();
            if (dateTime.Second < 10)
            {
                second = "0" + second;
            }

            string dateTimeString = string.Format("{0}-{1}-{2}T{3}:{4}:{5}Z", year, month, day, hour, minute, second);
            return dateTimeString;
        }

        /// <summary>
        /// This method is used to get the valid resource under the "DAV:multistatus" element collection in the HTTP body of response, 
        /// and record these valid resource in the "resourceList".
        /// A valid resource must include following two elements with none-empty value.
        ///     DAV:href
        ///     DAV:status
        /// And a valid resource can also include the element DAV:getlastmodified.
        /// </summary>
        /// <param name="xmlMultistatus">The whole XML data for "DAV:multistatus" element collection in the HTTP body of response</param>
        /// <returns>Return array list that contains all valid resource under the "DAV:multistatus" element.</returns>
        private ArrayList GetValidResourceUnderMultistatusElement(XmlDocument xmlMultistatus)
        {
            this.Site.Assert.IsNotNull(xmlMultistatus, "The object 'xmlMultistatus' should not be null!");
            this.Site.Assert.IsNotNull(xmlMultistatus.DocumentElement, "The object 'xmlMultistatus.DocumentElement' should not be null!");
            this.Site.Assert.IsNotNull(xmlMultistatus.DocumentElement.ChildNodes, "The object 'xmlMultistatus.DocumentElement.ChildNodes' should not be null!");

            ArrayList resourceList = new System.Collections.ArrayList();
            this.Site.Assert.IsTrue(xmlMultistatus.DocumentElement.LocalName == "multistatus", "The Xml Document object should include the element 'DAV:multistatus'!");

            if (xmlMultistatus.DocumentElement.LocalName == "multistatus")
            {
                foreach (XmlNode node_1 in xmlMultistatus.DocumentElement.ChildNodes)
                {
                    // Check each resource to find the valid resource.
                    if (node_1.LocalName == "response")
                    {
                        ResourceInfo resouceInfo = new ResourceInfo();
                        XmlNode nodeResponse = node_1;
                        bool findHrefElement = false;
                        foreach (XmlNode node_2 in nodeResponse.ChildNodes)
                        {
                            // In each resource, try to find "DAV:href" element under the resource.
                            if (node_2.LocalName == "href")
                            {
                                // When the server receives a PROPFIND request with the Repl:collblob element set to a timestamp, it includes a response element for each resource in the multistatus element, so there should be only one "href" element under each "response" element. 
                                // So if following assert is failed then we can't capture MS-WDVMODUU_R100.
                                this.Site.Assert.IsFalse(findHrefElement, "There should be only one \"href\" element under each \"response\" element.");
                                findHrefElement = true;

                                XmlNode nodeHref = node_2;
                                this.Site.Assert.IsNotNull(nodeHref.InnerText, "The value of element DAV:href should not be null.");
                                this.Site.Assert.IsTrue(nodeHref.InnerText != string.Empty, "The value of element DAV:href should not be empty string.");

                                // Record the value of "DAV:href"element. 
                                resouceInfo.Href = nodeHref.InnerText;
                                resouceInfo.Href = resouceInfo.Href.Replace("%20", " ");
                            }
                            #region Search in "DAV:propstat" element
                            if (node_2.LocalName == "propstat")
                            {
                                XmlNode nodePropstat = node_2;
                                foreach (XmlNode node_3 in nodePropstat)
                                {
                                    // Try to find "DAV:status"element under "DAV:propstat" element in the resource.
                                    if (node_3.LocalName == "status")
                                    {
                                        XmlNode nodeStatus = node_3;
                                        this.Site.Assert.IsNotNull(nodeStatus.InnerText, "The value of element DAV:status should not be null.");
                                        this.Site.Assert.IsTrue(nodeStatus.InnerText != string.Empty, "The value of element DAV:status should not be empty string.");

                                        // Record the value of "DAV:status"element. 
                                        resouceInfo.Status = nodeStatus.InnerText;
                                    }

                                    if (node_3.LocalName == "prop")
                                    {
                                        XmlNode nodeProp = node_3;
                                        foreach (XmlNode node_4 in nodeProp)
                                        {
                                            // Try to find "DAV:getlastmodified"element under "DAV:prop" element in the resource.
                                            if (node_4.LocalName == "getlastmodified")
                                            {
                                                XmlNode nodeGetLastModified = node_4;
                                                this.Site.Assert.IsNotNull(nodeGetLastModified.InnerText, "The value of element DAV:getlastmodified should not be null.");
                                                this.Site.Assert.IsTrue(nodeGetLastModified.InnerText != string.Empty, "The value of element DAV:getlastmodified should not be empty string.");
                                                string lastModified = nodeGetLastModified.InnerText;

                                                // Record the value of "DAV:getlastmodified"element. 
                                                resouceInfo.LastModifiedDateTime = DateTime.Parse(lastModified).ToUniversalTime();
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion
                        }

                        // Append the valid resource in the return resource list.
                        if ((resouceInfo.Href != string.Empty) && (resouceInfo.Status != string.Empty))
                        {
                            resourceList.Add(resouceInfo);
                        }
                    }
                }
            }

            return resourceList;
        }

        /// <summary>
        /// Try to find the special resource in the resource list.
        /// </summary>
        /// <param name="resourceList">The resource list includes all resource information</param>
        /// <param name="specialResourceUri">The URI of the special resource. If the URI is for collection resource, it must end with "/".</param>
        /// <returns>Return true if the special resource is found in the resource list, else return false.</returns>
        private bool FindSpecialResource(ArrayList resourceList, string specialResourceUri)
        {
            if ((resourceList == null) || string.IsNullOrEmpty(specialResourceUri))
            {
                return false;
            }

            bool findTheResource = false;
            int count = 0;
            IEnumerator enumratorResInfo = resourceList.GetEnumerator();
            while (enumratorResInfo.MoveNext())
            {
                ResourceInfo resInfo = (ResourceInfo)enumratorResInfo.Current;

                // In some protocol server, the return URI for collection resource end with "/"; 
                // In some protocol server, the return URI for collection resource end without "/";
                //  The example for such URI can be: 
                //              http://SUT01/sites/WDVMODUU/shared20%documents/
                //  or          http://SUT01/sites/WDVMODUU/shared20%documents
                //  The two URI identify the same collection resource.
                // So following codes use two compare results to make sure if the special resource is existed in the resource list.
                int compareResult1 = string.Compare(resInfo.Href, specialResourceUri, true);
                int compareResult2 = string.Compare(resInfo.Href + "/", specialResourceUri, true);
                if ((compareResult1 == 0) || (compareResult2 == 0))
                {
                    count++;
                    break;
                }
            }

            if (count == 1)
            {
                findTheResource = true;
            }
            else
            {
                findTheResource = false;
                this.Site.Log.Add(LogEntryKind.Comment, "The special resource URI {0} can't be found in the valid resource list.", specialResourceUri);
                enumratorResInfo = resourceList.GetEnumerator();
                while (enumratorResInfo.MoveNext())
                {
                    ResourceInfo resInfo = (ResourceInfo)enumratorResInfo.Current;
                    this.Site.Log.Add(LogEntryKind.Comment, "The valid resource {0} is in the valid resource list.", resInfo.Href);
                }
            }

            return findTheResource;
        }

        /// <summary>
        /// Call HTTP PUT method to upload a file into the server.
        /// </summary>
        /// <param name="destinationUri">The destination URI for the file to upload</param>
        /// <param name="filePath">The file path in the client</param>
        /// <returns>Return true if the file is uploaded successfully, else return false.</returns>
        private bool PutNewFileIntoServer(string destinationUri, string filePath)
        {
            if (string.IsNullOrEmpty(destinationUri))
            {
                return false;
            }

            if (string.IsNullOrEmpty(filePath))
            {
                return false;
            }

            bool isSuccessful = false;

            // Get the file content based on the file path.
            byte[] bytes = GetLocalFileContent(filePath);

            // Construct the request headers.
            NameValueCollection headersCollection = new NameValueCollection();
            headersCollection.Clear();
            headersCollection.Add("Cache-Control", "no-cache");
            headersCollection.Add("ContentLength", bytes.Length.ToString());
            headersCollection.Add("moss-cbfile", bytes.Length.ToString());
            headersCollection.Add("ProtocolVersion", "HTTP/1.1");

            // Call HTTP PUT method to upload the file into the destination URI.
            WDVMODUUResponse response = this.Adapter.Put(destinationUri, bytes, headersCollection);

            // Assert the response is successful.
            Site.Assume.IsTrue(
                response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Created,
                string.Format("Failed to PUT file {0} to the server under the path {1}! The return status code is {2}.", filePath, destinationUri, response.StatusCode));
            if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Created)
            {
                isSuccessful = true;
            }
            else
            {
                isSuccessful = false;
            }

            return isSuccessful;
        }

        #endregion Help Methods in Scenario 03

        #region Private structure

        /// <summary>
        /// The structure "ResourceInfo" is used to contain the key data of a resource in the response.
        /// </summary>
        private struct ResourceInfo
        {
            /// <summary>
            /// The URI of the resource. 
            /// </summary>
            public string Href;

            /// <summary>
            /// The status of the resource.
            /// </summary>
            public string Status;

            /// <summary>
            /// The last modified date time of the resource. 
            /// </summary>
            public DateTime LastModifiedDateTime;
        }

        #endregion Private structure
    }
}