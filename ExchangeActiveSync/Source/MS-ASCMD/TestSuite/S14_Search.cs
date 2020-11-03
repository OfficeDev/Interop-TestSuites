namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is used to test the Search command.
    /// </summary>
    [TestClass]
    public class S14_Search : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is used to verify the success status of Search command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC01_Search_Success()
        {
            #region Create a search request
            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4327");

            // Test Case verify requirement: MS-ASCMD_R4327
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.ResponseData.Response.Store.Status,
                4327,
                @"[In Status(Search)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when the Search command request has more than one And elements, the server should return the response with Status 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC02_Search_MultipleAndElements()
        {
            #region Create a search request with multiple And elements.
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { string.Empty },
                    ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.DeepTraversal }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And, Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType(), new Request.queryType() }
                }
            };

            // Create search request with multiple And elements.
            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Class, Request.ItemsChoiceType2.CollectionId, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)store.Query.Items[0]).Items = new object[] { "Email", this.User1Information.InboxCollectionId, "FreeText" };

            SearchRequest invalidSearchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            string sendStringRequest = invalidSearchRequest.GetRequestDataSerializedXML();
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, sendStringRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, sendStringRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string status = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R772");

            // Verify MS-ASCMD requirement: MS-ASCMD_R772
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                status,
                772,
                @"[In And] If multiple And elements are included in the request, the server responds with a Status element (section 2.2.3.162.12) value of 8 (SearchTooComplex).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when the Class element of Search command isn't a child element of And element, the server should return the response with status value 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC03_Search_TooComplex_Status8_Class()
        {
            #region Create a search request in which a class element is not in And element.

            // Insert Class element before And element.
            string invalidSearchRequest = this.CreateInvalidSearchRequest("<Class xmlns=\"AirSync\">Email</Class>", "<And>");
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string status = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R923");

            // Verify MS-ASCMD requirement: MS-ASCMD_R923
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                status,
                923,
                @"[In Class(Search)] If the airsync:Class element is included as a child of any element other than the And element, the server responds with a Status element (section 2.2.3.162.12) value of 8 (SearchTooComplex).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when the CollectionId element of Search command isn't a child element for And element, the server should return the response with Status 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC04_Search_TooComplex_Status8_CollectionId()
        {
            #region Create a search request in which a CollectionId element is not in And element.
            string collectionIdElement = "<CollectionId xmlns=\"AirSync\">" + this.User1Information.InboxCollectionId + "</CollectionId>";

            // Insert CollectionId element before And element.
            string invalidSearchRequest = this.CreateInvalidSearchRequest(collectionIdElement, "<And>");
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string status = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2032");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2032
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                status,
                2032,
                @"[In CollectionId(Search)] If the airsync:CollectionId element is included as a child of any element other than And, the server responds with a Status element (section 2.2.3.162.12) value of 8 (SearchTooComplex).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when the FreeText element of Search command isn't a child element for And element, the server should return the response with Status 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC05_Search_TooComplex_Status8_FreeText()
        {
            #region Create a search request in which a FreeText element is not in And element.
            string freeTextElement = "<FreeText>FreeText</FreeText>";

            // Insert FreeText element before And element.
            string invalidSearchRequest = this.CreateInvalidSearchRequest(freeTextElement, "<And>");
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string status = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3105");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3105
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                status,
                3105,
                @"[In FreeText (Search)] If the FreeText element is included as a child element of any element other than the And element, the server responds with a Status element (section 2.2.3.162.12) value of 8 (SearchTooComplex).");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search command with more than one DeepTraversal elements in And element, and the returned status value is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC06_Search_MultipleDeepTraversal()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2143, this.Site), "[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one DeepTraversal element in a Search command request]. (Exchange 2007 and above follow this behavior.)");
            #region Create a search request with multiple DeepTraversal elements.
            Request.Options1 searchOption = new Request.Options1
            {
                ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.DeepTraversal, Request.ItemsChoiceType6.DeepTraversal },
                Items = new object[] { string.Empty, string.Empty }
            };
            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Options = searchOption;
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2143");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2143
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.ResponseData.Response.Store.Status,
                2143,
                @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one DeepTraversal element in a Search command request]. (Exchange 2007 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search command with more than one airSync:MIMESupport elements in And element, and the returned status value is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC07_Search_MoreThanOneMIMESupport()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3374, this.Site), "[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one airsync:MIMESupport element in a Search command request]. (Exchange 2007 and above follow this behavior.)");
            #region Create a search request with multiple airsync:MIMESupport elements
            Request.Options1 searchOption = new Request.Options1
            {
                ItemsElementName = new Request.ItemsChoiceType6[]
                {
                    Request.ItemsChoiceType6.DeepTraversal,
                    Request.ItemsChoiceType6.MIMESupport,
                    Request.ItemsChoiceType6.MIMESupport
                }
            };
            byte mimeSupportValue = 2;
            searchOption.Items = new object[] { string.Empty, mimeSupportValue, mimeSupportValue };
            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Options = searchOption;
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3374");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3374
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.ResponseData.Response.Store.Status,
                3374,
                @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one airsync:MIMESupport element in a Search command request]. (Exchange 2007 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search command with more than one Range elements in And element, and the returned status value is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC08_Search_MultipleRange()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3723, this.Site), "[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one Range element in a Search command request is undefined]. (Exchange 2007 and above follow this behavior.)");
            #region Create a search request with multiple Range elements.
            Request.Options1 searchOption = new Request.Options1
            {
                ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.DeepTraversal, Request.ItemsChoiceType6.Range, Request.ItemsChoiceType6.Range },
                Items = new object[] { string.Empty, "0-9", "0-9" }
            };
            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Options = searchOption;
            #endregion

            #region Call Search command.
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3723");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3723
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.ResponseData.Response.Store.Status,
                3723,
                @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one Range element in a Search command request is undefined]. (Exchange 2007 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search command with more than one RebuildResults elements in And element, and the returned status value is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC09_Search_MoreThanOneRebuildResults()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3747, this.Site), "[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one RebuildResults element in a Search command request is undefined]. (Exchange 2007 and above follow this behavior.)");
            #region Create a search request with multiple RebuildResults elements.
            Request.Options1 searchOption = new Request.Options1
            {
                ItemsElementName = new Request.ItemsChoiceType6[]
                {
                    Request.ItemsChoiceType6.RebuildResults, 
                    Request.ItemsChoiceType6.RebuildResults
                },
                Items = new object[] { string.Empty, string.Empty }
            };
            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Options = searchOption;
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3747");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3747
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.ResponseData.Response.Store.Status,
                3747,
                @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one RebuildResults element in a Search command request is undefined]. (Exchange 2007 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search command with more than one Store elements in And element, and the returned status value is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC10_Search_MoreThanOneStores()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(4535, this.Site), "[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [[more than one Store element in a Search command request is undefined]. (Exchange 2007 and above follow this behavior.)");
            #region Create a search request with multiple Store elements.
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { string.Empty },
                    ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.DeepTraversal }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Class, Request.ItemsChoiceType2.CollectionId, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)store.Query.Items[0]).Items = new object[] { "Email", this.User1Information.InboxCollectionId, "FreeText" };

            // Create search request with two store elements
            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store, store });
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4535");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4535
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.ResponseData.Response.Store.Status,
                4535,
                @"[In Appendix A: Product Behavior] The implementation does not return a protocol status error in response to such a command request [more than one Store element in a Search command request is undefined]. (Exchange 2007 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search command with invalid formatted request to get status 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC11_Search_Invalid_Status2()
        {
            #region Create an invalid search request with email class element included in And element.
            string user1EmailAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string emailClassElement = "<To xmlns=\"Email\">" + user1EmailAddress + "</To>";

            // Search request should not contain any element belong to Email class, insert Email class's To element in the default search request to create an invalid search request.
            string invalidSearchRequest = this.CreateInvalidSearchRequest(emailClassElement, "</And>");
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string statusCode = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4329");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4329
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                statusCode,
                4329,
                @"[In Status(Search)] [When the scope is Item], [the cause of the status value 2 is] One or more of the client's search parameters was invalid.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify there is no matched item for the Search request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC12_Search_Result_NoItemMatched()
        {
            #region User1 calls SendMail command to send an email message to user2.
            string keyword = Guid.NewGuid().ToString().Substring(0, 5);
            string emailSubject = Common.GenerateResourceName(Site, keyword);
            SendMailResponse responseSendMail = this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            Site.Assert.AreEqual(string.Empty, responseSendMail.ResponseDataXML, "If SendMail command executes successfully, server should return empty xml data");
            #endregion

            #region Sync user2 mailbox changes
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Create a search request for searching mail.
            // Create one search query with GUID as search keyword, search should not get result with that keyword.
            string searchKeyword = Common.GenerateResourceName(Site, "search");
            Request.queryType searchQuery = new Request.queryType
            {
                ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                Items = new Request.queryType[] { new Request.queryType() }
            };
            ((Request.queryType)searchQuery.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Class, Request.ItemsChoiceType2.CollectionId, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)searchQuery.Items[0]).Items = new object[] { "Email", this.User2Information.InboxCollectionId, searchKeyword };

            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Query = searchQuery;
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            string statusCode = searchResponse.ResponseData.Status;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(searchResponse.ResponseDataXML);
            XmlNodeList results = doc.GetElementsByTagName("Result");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3847");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3847
            Site.CaptureRequirementIfIsTrue(
                statusCode.Equals("1") && results.Count == 1 && results[0].ChildNodes.Count == 0,
                3847,
                @"[In Result(Search)] If no matches are found, an empty Result element is present in the Store container element of the response XML.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3850");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3850
            Site.CaptureRequirementIfIsTrue(
                statusCode.Equals("1") && results.Count == 1 && results[0].ChildNodes.Count == 0,
                3850,
                @"[In Result(Search)] If no matches are found, an empty Result element is present in the Store container element of the response XML.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search response when the searched items have the multiple matched items.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC13_Search_Result_MultipleMatchedItems()
        {
            #region User1 calls SendMail command to send 2 email messages to user2.
            string keyWord = Guid.NewGuid().ToString().Substring(0, 5);
            uint mailIndex = 1;
            string emailSubject = keyWord + Common.GenerateResourceName(Site, "search", mailIndex);
            SendMailResponse responseSendMail = this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            Site.Assert.AreEqual(string.Empty, responseSendMail.ResponseDataXML, "If SendMail command executes successfully, server should return empty xml data");
            mailIndex++;
            string emailSubject2 = keyWord + Common.GenerateResourceName(Site, "search", mailIndex);
            SendMailResponse responseSendMail2 = this.SendPlainTextEmail(null, emailSubject2, this.User1Information.UserName, this.User2Information.UserName, null);
            Site.Assert.AreEqual(string.Empty, responseSendMail2.ResponseDataXML, "If SendMail command executes successfully, server should return empty xml data");
            #endregion

            #region Sync user2 mailbox changes
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject2);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject, emailSubject2);
            #endregion

            #region Create a search request for searching mail.
            Request.Options1 searchOption = new Request.Options1
            {
                ItemsElementName = new Request.ItemsChoiceType6[]
                {
                    Request.ItemsChoiceType6.DeepTraversal, Request.ItemsChoiceType6.RebuildResults,
                    Request.ItemsChoiceType6.Range
                },
                Items = new object[] { string.Empty, string.Empty, "0-2" }
            };

            Request.queryType searchQuery = new Request.queryType
            {
                ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                Items = new Request.queryType[] { new Request.queryType() }
            };
            ((Request.queryType)searchQuery.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Class, Request.ItemsChoiceType2.CollectionId, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)searchQuery.Items[0]).Items = new object[] { "Email", this.User2Information.InboxCollectionId, keyWord };

            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Options = searchOption;
            searchRequest.RequestData.Items[0].Query = searchQuery;
            #endregion

            #region Call Search command
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int sendMailCount = 2;
            int resultsCount;
            SearchResponse searchResponse;

            // Loop search to get correct results.
            do
            {
                Thread.Sleep(waitTime);
                searchResponse = this.CMDAdapter.Search(searchRequest);
                Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
                resultsCount = searchResponse.ResponseData.Response.Store.Result.Length;
                counter++;
            }
            while (resultsCount != sendMailCount && counter < retryCount);

            Site.Assert.AreEqual<int>(2, resultsCount, "Search response should contain two search results");
            Site.Log.Add(LogEntryKind.Debug, "Loop {0} times to get the search item", counter);
            #endregion

            #region Verify Requirements MS-ASCMD_R3846, MS-ASCMD_R3849, MS-ASCMD_R3862, MS-ASCMD_R3851
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3846");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3846
            // User1 sends 2 emails to user2 with same keyWord, user2 searches Inbox folder with specified keyword and gets 2 results, then MS-ASCMD_R3846 is verified
            Site.CaptureRequirementIfIsTrue(
                searchResponse.ResponseData.Response.Store.Status.Equals("1") && resultsCount == 2,
                3846,
                @"[In Result(Search)] One Result element is present for each match that is found.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3849");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3849
            Site.CaptureRequirementIfIsTrue(
                searchResponse.ResponseData.Response.Store.Status.Equals("1") && resultsCount == 2 && searchResponse.ResponseData.Response.Store.Result[0].CollectionId.Equals(this.User2Information.InboxCollectionId) && searchResponse.ResponseData.Response.Store.Result[1].CollectionId.Equals(this.User2Information.InboxCollectionId),
                3849,
                @"[In Result(Search)] When the store that is being searched is the mailbox, there is one Result element for each match that is found in the mailbox.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3862");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3862
            Site.CaptureRequirementIfIsTrue(
                searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName.Length > 0,
                3862,
                @"[In Result(Search)] Inside the Result element, the Properties element contains a list of requested properties for the mailbox item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3851");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3851
            Site.CaptureRequirementIfIsTrue(
                searchResponse.ResponseData.Response.Store.Result[0].CollectionId.Equals(this.User2Information.InboxCollectionId) && searchResponse.ResponseData.Response.Store.Result[0].Properties.ItemsElementName.Length > 0,
                3851,
                @"[In Result(Search)] When the store that is being searched is the mailbox, inside the Result element, the Properties element contains a list of requested properties for the mailbox item.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify search global address list success
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC14_Search_GAL_Success()
        {
            #region Create Search request with options
            int rangeEndIndex = 2;
            string requestRange = "0-" + rangeEndIndex.ToString();

            // Set search query.
            string keyWord = "MSASCMD_User";
            SearchRequest searchRequest = this.CreateSearchGALRequest(0, 0, requestRange, keyWord);
            #endregion

            #region Call search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            #endregion

            #region Verify Requirements MS-ASCMD_R5323, MS-ASCMD_R5368, MS-ASCMD_R5391, MS-ASCMD_R5360, MS-ASCMD_R5337, MSASCMD_R5330, MSASCMD_R5323, MSASCMD_R5286, MSASCMD_R5287, MSASCMD_R5293, MSASCMD_R5300, MSASCMD_R5316, MS-ASCMD_R419, MS-ASCMD_R420, MS-ASCMD_R4638, MS-ASCMD_R5711, MS-ASCMD_R421, MS-ASCMD_R431, MS-ASCMD_R426
            int searchResultCount = searchResponse.ResponseData.Response.Store.Result.Length;

            bool hasMobilePhoneElement = false;
            bool hasPhoneElement = false;
            bool hasTitleElement = false;
            bool hasOfficeElement = false;
            bool hasLastNameElement = false;
            bool hasFirstNameElement = false;
            bool hasHomePhoneElement = false;
            bool hasAliasElement = false;
            bool hasCompanyElement = false;
            bool hasDisplayNameElement = false;
            bool hasEmailAddressElement = false;

            foreach (Response.SearchResponseStoreResult result in searchResponse.ResponseData.Response.Store.Result)
            {
                if (result.Properties != null)
                {
                    // Verify optional property values.
                    foreach (Response.ItemsChoiceType6 itemElementName in result.Properties.ItemsElementName)
                    {
                        switch (itemElementName)
                        {
                            case Response.ItemsChoiceType6.Alias1:
                                hasAliasElement = true;
                                break;
                            case Response.ItemsChoiceType6.Company:
                                hasCompanyElement = true;
                                break;
                            case Response.ItemsChoiceType6.EmailAddress:
                                hasEmailAddressElement = true;
                                break;
                            case Response.ItemsChoiceType6.FirstName1:
                                hasFirstNameElement = true;
                                break;
                            case Response.ItemsChoiceType6.LastName1:
                                hasLastNameElement = true;
                                break;
                            case Response.ItemsChoiceType6.DisplayName1:
                                hasDisplayNameElement = true;
                                break;
                            case Response.ItemsChoiceType6.MobilePhone:
                                hasMobilePhoneElement = true;
                                break;
                            case Response.ItemsChoiceType6.HomePhone:
                                hasHomePhoneElement = true;
                                break;
                            case Response.ItemsChoiceType6.Office:
                                hasOfficeElement = true;
                                break;
                            case Response.ItemsChoiceType6.Phone:
                                hasPhoneElement = true;
                                break;
                            case Response.ItemsChoiceType6.Title1:
                                hasTitleElement = true;
                                break;
                        }
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5353");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5353
            Site.CaptureRequirementIfIsTrue(
                hasMobilePhoneElement,
                5353,
                @"[In MobilePhone] The MobilePhone element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5368");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5368
            Site.CaptureRequirementIfIsTrue(
                hasPhoneElement,
                5368,
                @"[In Phone] The Phone element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5391");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5391
            Site.CaptureRequirementIfIsTrue(
                hasTitleElement,
                5391,
                @"[In Title] The Title element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5360");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5360
            Site.CaptureRequirementIfIsTrue(
                hasOfficeElement,
                5360,
                @"[In Office (Search)] The Office element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5337");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5337
            Site.CaptureRequirementIfIsTrue(
                hasLastNameElement,
                5337,
                @"[In LastName (Search)] The LastName element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5330");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5330
            Site.CaptureRequirementIfIsTrue(
                hasHomePhoneElement,
                5330,
                @"[In HomePhone] The HomePhone element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5323");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5323
            Site.CaptureRequirementIfIsTrue(
                hasFirstNameElement,
                5323,
                @"[In FirstName (Search)] The FirstName element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5286");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5286
            Site.CaptureRequirementIfIsTrue(
                hasAliasElement,
                5286,
                @"[In Alias (Search)] The Alias element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5287");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5287
            Site.CaptureRequirementIfIsTrue(
                hasAliasElement,
                5287,
                @"[In Alias (Search)] It [Alias element] contains the alias of a recipient in the GAL that matched the search criteria from the corresponding Search command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5293");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5293
            Site.CaptureRequirementIfIsTrue(
                hasCompanyElement,
                5293,
                @"[In Company (Search)] The Company element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5300");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5300
            Site.CaptureRequirementIfIsTrue(
                hasDisplayNameElement,
                5300,
                @"[In DisplayName(Search)] The DisplayName element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5316");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5316
            Site.CaptureRequirementIfIsTrue(
                hasEmailAddressElement,
                5316,
                @"[In EmailAddress(Search)] The EmailAddress element is only present in response to Search command requests that contain a search:Name element with a value of ""GAL"", as specified in section 2.2.3.110.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R419");

            // Verify MS-ASCMD requirement: MS-ASCMD_R419
            Site.CaptureRequirementIfIsTrue(
                searchResultCount > rangeEndIndex,
                419,
                @"[In Search] The server MUST return entries up to the number that is requested.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R420");

            // Verify MS-ASCMD requirement: MS-ASCMD_R420
            Site.CaptureRequirementIfIsTrue(
                Convert.ToInt32(searchResponse.ResponseData.Response.Store.Total) >= searchResultCount,
                420,
                @"[In Search] [The server] MUST also indicate the total number of entries that are found.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4638");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4638
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                searchResultCount - rangeEndIndex,
                4638,
                @"[In Total(Search)] In cases where all the results are returned in the response XML, the value of the Total element is one more than the end-index value that is provided in the Range element.");

            List<object> displayName = GetElementsFromSearchResponse(searchResponse, Response.ItemsChoiceType6.DisplayName1);
            List<object> alias = GetElementsFromSearchResponse(searchResponse, Response.ItemsChoiceType6.Alias1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5711");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5711
            Site.CaptureRequirementIfIsTrue(
                displayName.Count == searchResultCount && alias.Count == searchResultCount,
                5711,
                @"[In Search] For each GAL entry that is found, the Search command returns email alias and display name.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R421");

            // Verify MS-ASCMD requirement: MS-ASCMD_R421
            Site.CaptureRequirementIfIsTrue(
                this.User1Information.UserName.StartsWith(keyWord, StringComparison.OrdinalIgnoreCase),
                421,
                @"[In Search] The text query string that is provided to the Search command is used in a prefix-string match.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R431");

            // Verify MS-ASCMD requirement: MS-ASCMD_R431
            Site.CaptureRequirementIfIsTrue(
                searchResponse.ResponseData.Response.Store.Range.Equals(requestRange),
                431,
                @"[In Search] The Range option is a zero-based index specifier in the form of ""m-n"".");

            // Verify display name in search result is in order.
            int compareResult = string.Compare((string)displayName[0], (string)displayName[1], StringComparison.CurrentCulture);
            bool displayNameInOrder = true;
            for (int displayNameIndex = 1; displayNameIndex < rangeEndIndex; displayNameIndex++)
            {
                int otherCompareResult = string.Compare((string)displayName[displayNameIndex], (string)displayName[displayNameIndex + 1], StringComparison.CurrentCulture);
                if (compareResult != otherCompareResult)
                {
                    displayNameInOrder = false;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R426");

            // Verify MS-ASCMD requirement: MS-ASCMD_R426
            Site.CaptureRequirementIfIsTrue(
                displayNameInOrder,
                426,
                @"[In Search] The Search command results are sorted by the server according to their ordering in the GAL (that is, by the display name property).");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify search with Tasks, Email, Calendar, Contacts, Notes, SMS class elements.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC15_Search_WithAllClass()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0", "[In Appendix A: Product Behavior] <14> Section 2.2.3.27.4: The SMS and Notes classes are only available if the MS-ASProtocolVersion header is set to 14.0 [or 14.1].");

            #region Create one item in User2 Inbox folder, Calendar folder, Contacts subfolder
            string searchPrefix = "keyWord" + TestSuiteBase.ClientId;

            // Respectively create one item in User2 Inbox folder, Calendar folder and Contacts sub-folder
            this.CreateItemsWithKeyword(searchPrefix);
            #endregion

            #region Create Search request with class element
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    ItemsElementName = new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.DeepTraversal,
                        Request.ItemsChoiceType6.RebuildResults
                    },
                    Items = new object[] { string.Empty, string.Empty }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.FreeText
            };

            ((Request.queryType)store.Query.Items[0]).Items = new object[] { this.User2Information.InboxCollectionId, this.User2Information.CalendarCollectionId, this.User2Information.ContactsCollectionId, "Tasks", "Email", "Calendar", "Contacts", "Notes", "SMS", searchPrefix };
            #endregion

            #region Call search command
            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            if (searchResponse.ResponseData.Response.Store.Status != "1" || searchResponse.ResponseData.Response.Store.Result.Length != 3)
            {
                searchResponse = this.LoopSearch(searchRequest, 3);
            }
            #endregion

            #region Verify requirements MS-ASCMD_R1322, MS-ASCMD_R920, MS-ASCMD_R5146, MS-ASCMD_R5883
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R920");

            // Verify MS-ASCMD requirement: MS-ASCMD_R920
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                searchResponse.ResponseData.Response.Store.Status,
                920,
                @"[In Class(Search)] The valid airsync:Class element values are: Tasks, Email, Calendar, Contacts, Notes, SMS<14>.");

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0")
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5146");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5146
                Site.CaptureRequirementIfAreEqual<string>(
                    "1",
                    searchResponse.ResponseData.Response.Store.Status,
                    5146,
                    @"[In Appendix A: Product Behavior] <14> Section 2.2.3.27.4: The SMS and Notes classes are only available if the MS-ASProtocolVersion header is set to 14.0 [or 14.1].");
            }

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1")
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5883");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5883
                Site.CaptureRequirementIfAreEqual<string>(
                    "1",
                    searchResponse.ResponseData.Response.Store.Status,
                    5883,
                    @"[In Appendix A: Product Behavior] <14> Section 2.2.3.27.4: The SMS and Notes classes are only available if the MS-ASProtocolVersion header is set to [14.0 or] 14.1.");
            }

            bool isAllResultContainClass = false;
            foreach (Response.SearchResponseStoreResult result in searchResponse.ResponseData.Response.Store.Result)
            {
                // Check if every search result contains Class element.
                isAllResultContainClass = result.Class != null && IsClassSupported(result.Class);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1322");

            // Verify MS-ASCMD requirement: MS-ASCMD_R1322
            Site.CaptureRequirementIfIsTrue(
                isAllResultContainClass,
                1322,
                @"[In Class(Search)] Element Class (Search) in Search command response, the number allowed is 1? (required, 1 per Result element).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify search without Notes and SMS class element when ProtocolVersion set to 12.1
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC16_Search_WithoutNoteAndSmsClass()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Notes, SMS class is not supported in ProtocolVersion 12.1");
            #region Create one item in User2 Inbox folder, Calendar folder, contacts sub folder
            string searchPrefix = "keyWord" + TestSuiteBase.ClientId;

            // Respectively create one item in User2 Inbox folder, Calendar folder  and Contacts subfolder.
            this.CreateItemsWithKeyword(searchPrefix);
            #endregion

            #region Create Search request with class element
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    ItemsElementName = new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.DeepTraversal,
                        Request.ItemsChoiceType6.RebuildResults
                    },
                    Items = new object[] { string.Empty, string.Empty }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.Class,
                Request.ItemsChoiceType2.FreeText
            };

            ((Request.queryType)store.Query.Items[0]).Items = new object[] { this.User2Information.InboxCollectionId, this.User2Information.CalendarCollectionId, this.User2Information.ContactsCollectionId, "Tasks", "Email", "Calendar", "Contacts", searchPrefix };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Call search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            if (searchResponse.ResponseData.Response.Store.Status != "1" || searchResponse.ResponseData.Response.Store.Result.Length != 3)
            {
                searchResponse = this.LoopSearch(searchRequest);
            }
            #endregion
        }

        /// <summary>
        /// This test case is used to verify search without class element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC17_Search_WithoutClass()
        {
            #region Create one item in User2 Inbox folder, Calendar folder  and Contacts subfolder
            string searchPrefix = "keyWord" + TestSuiteBase.ClientId;

            // Respectively create one item in User2 Inbox folder, Calendar folder  and Contacts subfolder.
            this.CreateItemsWithKeyword(searchPrefix);
            #endregion

            #region Create Search request without class element
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { string.Empty, string.Empty },
                    ItemsElementName = new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.DeepTraversal,
                        Request.ItemsChoiceType6.RebuildResults
                    }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.FreeText
            };

            ((Request.queryType)store.Query.Items[0]).Items = new object[] { this.User2Information.InboxCollectionId, this.User2Information.CalendarCollectionId, this.User2Information.ContactsCollectionId, searchPrefix };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Call search command
            SearchResponse searchResponse = this.LoopSearch(searchRequest);
            #endregion

            bool isClassInResultIsServerSupported = false;

            // Check if every search result contains Class element and the Class element is server supported.
            foreach (Response.SearchResponseStoreResult result in searchResponse.ResponseData.Response.Store.Result)
            {
                // Check if every search result contains Class element.
                if (result.Class != null)
                {
                    isClassInResultIsServerSupported = IsClassSupported(result.Class);
                }
                else
                {
                    isClassInResultIsServerSupported = false;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R922");

            // Verify MS-ASCMD requirement: MS-ASCMD_R922
            Site.CaptureRequirementIfIsTrue(
                isClassInResultIsServerSupported,
                922,
                @"[In Class(Search)] If one or more airsync:Class elements are not included in the Search request, the server will return all supported classes.");
        }

        /// <summary>
        /// This test case is used to verify if search point to one folder the first search result contains folder metadata and the second search result contains folder items metadata information.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC18_Search_DocumentLibraryFolder()
        {
            #region Create Search request for document library root folder
            string rootFolder = Common.GetConfigurationPropertyValue("SharedFolder", this.Site);
            SearchRequest searchRequest = CreateSearchDocumentLibraryRequest(rootFolder, this.User1Information.UserName, this.User1Information.UserPassword);
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            #endregion

            int resultCount = searchResponse.ResponseData.Response.Store.Result.Length;
            bool firstResultContainsFolderInfo = false;
            bool secondResultContainsContentInfo = false;
            if (resultCount >= 2)
            {
                Response.SearchResponseStoreResult firstResult = searchResponse.ResponseData.Response.Store.Result[0];
                byte isFolder = (byte)GetItemFromSearchResult(firstResult, Response.ItemsChoiceType6.IsFolder);
                if (isFolder == 1)
                {
                    // Means the first result contains root folder information.
                    firstResultContainsFolderInfo = true;
                }

                Response.SearchResponseStoreResult secondResult = searchResponse.ResponseData.Response.Store.Result[1];
                isFolder = (byte)GetItemFromSearchResult(secondResult, Response.ItemsChoiceType6.IsFolder);
                string linkID = (string)GetItemFromSearchResult(secondResult, Response.ItemsChoiceType6.LinkId);
                string contentType = (string)GetItemFromSearchResult(secondResult, Response.ItemsChoiceType6.ContentType);

                if (isFolder == 0 && contentType != null && linkID.StartsWith(rootFolder, true, null))
                {
                    // Means the secondResult contains contents in root folder.
                    secondResultContainsContentInfo = true;
                }
            }

            #region Verify Requirements MS-ASCMD_R3853, MS-ASCMD_R3855
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3853");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3853
            Site.CaptureRequirementIfIsTrue(
                firstResultContainsFolderInfo,
                3853,
                @"[In Result(Search)] When the store that is being searched is the document library, the first result that is returned in the Search response is the metadata for the Root folder or item to which the documentlibrary:LinkId element (section 2.2.3.48) value is pointing.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3855");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3855
            Site.CaptureRequirementIfIsTrue(
                firstResultContainsFolderInfo && secondResultContainsContentInfo,
                3855,
                @"[In Result(Search)] If the documentlibrary:LinkId element value in the request points to a folder, the metadata properties of the folder are returned as the first item, and the contents of the folder are returned as subsequent results.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify that if search document library point to specified folder item, only one search result will be returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC19_Search_DocumentLibraryItem()
        {
            #region Create search request
            string sharedItem = Common.GetConfigurationPropertyValue("SharedDocument1", this.Site);
            SearchRequest searchRequest = CreateSearchDocumentLibraryRequest(sharedItem, this.User1Information.UserName, this.User1Information.UserPassword);
            #endregion

            #region Calls Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Response.Store.Status, "If search command executes successfully, server should return status 1");
            #endregion

            int resultCount = searchResponse.ResponseData.Response.Store.Result.Length;
            string linkID = (string)GetItemFromSearchResult(searchResponse.ResponseData.Response.Store.Result[0], Response.ItemsChoiceType6.LinkId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3858");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3858
            Site.CaptureRequirementIfIsTrue(
                sharedItem.Equals(linkID, StringComparison.OrdinalIgnoreCase) && resultCount == 1,
                3858,
                @"[In Result(Search)] If the documentlibrary:LinkId element value in the request points to an item, only one result is returned: the metadata for the item.");
        }

        /// <summary>
        /// This test case is used to verify that if the contact photo exceeded the size limit set by MaxSize element , the search result should contain status 174.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC20_Search_GAL_Status174()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "12.1" && Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "14.0", "Status value 174 is not returned when the MS-ASProtocolVersion header value is set to 12.1 or 14.0.");
            #region Create Search request with options

            // Set max return 3 contacts with picture.
            uint maxPictures = 3;

            // Set the max contact picture returned won't exceed 1024 bytes.
            uint maxSize = 1 * 1024;
            int rangeEndIndex = 2;
            string requestRange = "0-" + rangeEndIndex.ToString();

            // Set search query.
            string keyWord = "MSASCMD_User";
            SearchRequest searchRequest = this.CreateSearchGALRequest(maxPictures, maxSize, requestRange, keyWord);
            #endregion

            #region Call search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            #endregion

            // Check if the search result 's picture element contains status 174.
            bool isStatus174Contains = false;
            List<object> picturesFromResponse = GetElementsFromSearchResponse(searchResponse, Response.ItemsChoiceType6.Picture1);
            foreach (object item in picturesFromResponse)
            {
                Response.SearchResponseStoreResultPropertiesPicture itemPicture = (Response.SearchResponseStoreResultPropertiesPicture)item;
                if (itemPicture.Data == null)
                {
                    // If contact picture size exceeds the size limit, server returns status value is 174.
                    if (itemPicture.Status.Equals("174"))
                    {
                        isStatus174Contains = true;
                    }
                }
            }

            #region Verify Requirements MS-ASCMD_R4374, MS-ASCMD_R4972
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4374");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4374
            Site.CaptureRequirementIfIsTrue(
                isStatus174Contains,
                4374,
                @"[In Status(Search)] [The meaning of the status value] 174 [is] The contact photo exceeded the size limit set by the MaxSize element (section 2.2.3.95.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4972");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4972
            Site.CaptureRequirementIfIsTrue(
                isStatus174Contains,
                4972,
                @"[In Common Status Codes] [The meaning of the status value 174 is] The contact photo exceeds the size limit set by the MaxSize element (section 2.2.3.95).<108>");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify that server returns 175 if the number of contact photos returned exceeded the size limit set by the MaxPictures element
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC21_Search_Status175()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "12.1" && Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "14.0", "Status value 175 is not returned when the MS-ASProtocolVersion header value is set to 12.1 or 14.0.");
            #region Create Search request with options

            // Set max return one contact with picture.
            uint maxPictures = 1;

            // Set max picture returned from server should not exceeds 10240 bytes.
            uint maxSize = 10 * 1024;
            int rangeEndIndex = 2;
            string requestRange = "0-" + rangeEndIndex.ToString();

            // Set search query.
            string keyWord = "MSASCMD_User";
            SearchRequest searchRequest = this.CreateSearchGALRequest(maxPictures, maxSize, requestRange, keyWord);
            #endregion

            #region Call search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            #endregion

            // Check if the search result 's picture element contains status 175.
            bool containsStatus175 = false;
            bool validStatusValue = true;
            List<object> picturesFromResponse = GetElementsFromSearchResponse(searchResponse, Response.ItemsChoiceType6.Picture1);
            foreach (object item in picturesFromResponse)
            {
                Response.SearchResponseStoreResultPropertiesPicture itemPicture = (Response.SearchResponseStoreResultPropertiesPicture)item;
                if (itemPicture.Data == null)
                {
                    if (!itemPicture.Status.Equals("1") && !itemPicture.Status.Equals("173") && !itemPicture.Status.Equals("174") && !itemPicture.Status.Equals("175"))
                    {
                        validStatusValue = false;
                    }

                    // If contact picture returned exceeds the number limit of MaxPictures, server returns status value is 175.
                    if (itemPicture.Status.Equals("175"))
                    {
                        containsStatus175 = true;
                    }
                }
            }

            #region Verify Requirements MS-ASCMD_R4375, MS-ASCMD_R4974, MS-ASCMD_R4371
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4375");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4375
            Site.CaptureRequirementIfIsTrue(
                containsStatus175,
                4375,
                @"[In Status(Search)] [The meaning of the status value] 175 [is] The number of contact photos returned exceeded the size limit set by the MaxPictures element (section 2.2.3.94.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4974");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4974
            Site.CaptureRequirementIfIsTrue(
                containsStatus175,
                4974,
                @"[In Common Status Codes] [The meaning of the status value 175 is] The number of contact photos returned exceeds the size limit set by the MaxPictures element (section 2.2.3.94).<109>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4371");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4371
            Site.CaptureRequirementIfIsTrue(
                validStatusValue,
                4371,
                @"[In Status(Search)] The following table shows valid values [1,173,174,175] for the Status element when it is returned as a child of the gal:Picture element.<75>");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify if search with DeepTraversal element, server should search subfolders.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC22_Search_WithDeepTraversal()
        {
            #region User1 calls SendMail command to send two emails to user2
            string searchPrefix = "keyWord" + TestSuiteBase.ClientId;
            string emailSubject1 = searchPrefix + Common.GenerateResourceName(Site, "subject1");
            string emailSubject2 = searchPrefix + Common.GenerateResourceName(Site, "subject2");
            this.SendPlainTextEmail(null, emailSubject1, this.User1Information.UserName, this.User2Information.UserName, null);
            this.SendPlainTextEmail(null, emailSubject2, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region User2 move one of emails to new subfolder in Inbox folder.
            this.SwitchUser(this.User2Information);
            string emailItemOneServerID = this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, emailSubject1);
            this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, emailSubject2);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1, emailSubject2);

            // User2 creates one subfolder in Inbox folder.
            string folderID = this.CreateFolder((byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), this.User2Information.InboxCollectionId);
            TestSuiteBase.RecordCaseRelativeFolders(this.User2Information, folderID);

            // User2 moves the email with emailSubject1 into new subfolder.
            MoveItemsRequest moveItemRequest = TestSuiteBase.CreateMoveItemsRequest(emailItemOneServerID, this.User2Information.InboxCollectionId, folderID);
            MoveItemsResponse moveItemResponse = this.CMDAdapter.MoveItems(moveItemRequest);
            Site.Assert.AreEqual(3, int.Parse(moveItemResponse.ResponseData.Response[0].Status), " If MoveItems command executes successful, server should return status 3");
            this.GetMailItem(folderID, emailSubject1);
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1);
            #endregion

            #region Creates Search request with DeepTraversale element
            Request.queryType query = new Request.queryType();
            string storeName = SearchName.Mailbox.ToString();
            Request.Options1 optionWithDeepTraversalAndRebuild = new Request.Options1();

            // Create search request query.
            query.ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And };
            query.Items = new Request.queryType[] { new Request.queryType() };

            ((Request.queryType)query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.FreeText
            };

            ((Request.queryType)query.Items[0]).Items = new object[] { this.User2Information.InboxCollectionId, searchPrefix };

            // Create Search option with DeepTraversal and Rebuild.
            optionWithDeepTraversalAndRebuild.Items = new object[] { string.Empty, string.Empty };
            optionWithDeepTraversalAndRebuild.ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.DeepTraversal, Request.ItemsChoiceType6.RebuildResults };

            Request.SearchStore store = new Request.SearchStore
            {
                Name = storeName,
                Query = query,
                Options = optionWithDeepTraversalAndRebuild
            };
            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });

            #endregion

            #region Call search command with DeepTraversal element
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            if (searchResponse.ResponseData.Response.Store.Status != "1" || searchResponse.ResponseData.Response.Store.Result.Length != 2)
            {
                searchResponse = this.LoopSearch(searchRequest, 2);
            }
            #endregion

            List<object> subjectInSearchResult = GetElementsFromSearchResponse(searchResponse, Response.ItemsChoiceType6.Subject1);
            bool resultContainsAllEmail = false;
            int findEmailCount = 0;
            for (int listIndex = 0; listIndex < subjectInSearchResult.Count; listIndex++)
            {
                string subject = (string)subjectInSearchResult[listIndex];
                if (subject.Equals(emailSubject1) || subject.Equals(emailSubject2))
                {
                    findEmailCount++;
                }
            }

            if (findEmailCount == 2)
            {
                resultContainsAllEmail = true;
            }
            #region Verify Requirements MS-ASCMD_R2031, MS-ASCMD_R5856
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2031");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2031
            Site.CaptureRequirementIfIsTrue(
                resultContainsAllEmail,
                2031,
                @"[In CollectionId(Search)] If the DeepTraversal element (section 2.2.3.41) is present, it applies to all folders under each airsync:CollectionId element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5856");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5856
            Site.CaptureRequirementIfIsTrue(
                resultContainsAllEmail,
                5856,
                @"[In DeepTraversal (Search)] [The DeepTraversal element] indicates that the client wants the server to search all subfolders for the folders that are specified in the query.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if search without DeepTraversal element, server should not search subfolders.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC23_Search_WithoutDeepTraversal()
        {
            #region User1 calls SendMail command to send two emails to user2
            string searchPrefix = "keyWord" + TestSuiteBase.ClientId;
            string emailSubject1 = searchPrefix + Common.GenerateResourceName(Site, "subject1");
            string emailSubject2 = searchPrefix + Common.GenerateResourceName(Site, "subject2");
            this.SendPlainTextEmail(null, emailSubject1, this.User1Information.UserName, this.User2Information.UserName, null);
            this.SendPlainTextEmail(null, emailSubject2, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region User2 moves one of the emails to new subfolder in Inbox folder.
            this.SwitchUser(this.User2Information);
            string emailItemOneServerID = this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, emailSubject1);
            this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, emailSubject2);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1, emailSubject2);

            // User2 creates one subfolder in Inbox folder.
            string folderID = this.CreateFolder((byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), this.User2Information.InboxCollectionId);
            this.SyncChanges(folderID);
            TestSuiteBase.RecordCaseRelativeFolders(this.User2Information, folderID);

            // User2 moves the email with emailSubject1 into new subfolder.
            MoveItemsRequest moveItemRequest = TestSuiteBase.CreateMoveItemsRequest(emailItemOneServerID, this.User2Information.InboxCollectionId, folderID);
            MoveItemsResponse moveItemResponse = this.CMDAdapter.MoveItems(moveItemRequest);
            Site.Assert.AreEqual(3, int.Parse(moveItemResponse.ResponseData.Response[0].Status), " If MoveItems command executes successfully, server should return status 3");
            this.GetMailItem(folderID, emailSubject1);
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1);
            #endregion

            #region Creates Search request without DeepTraversale element
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { string.Empty },
                    ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.RebuildResults }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.FreeText
            };

            ((Request.queryType)store.Query.Items[0]).Items = new object[] { this.User2Information.InboxCollectionId, searchPrefix };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Call search command without DeepTraversal element
            SearchResponse searchResponseWithOutDeepTraversal = this.LoopSearch(searchRequest);
            #endregion

            List<object> subjectInSearchResult = GetElementsFromSearchResponse(searchResponseWithOutDeepTraversal, Response.ItemsChoiceType6.Subject1);
            bool resultCotainsEmail1 = false;
            bool resultCotainsEmail2 = false;

            foreach (object result in subjectInSearchResult)
            {
                string subject = (string)result;
                if (subject.Equals(emailSubject1))
                {
                    resultCotainsEmail1 = true;
                }
                else if (subject.Equals(emailSubject2))
                {
                    resultCotainsEmail2 = true;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2146");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2146
            Site.CaptureRequirementIfIsTrue(
                !resultCotainsEmail1 && resultCotainsEmail2,
                2146,
                @"[In DeepTraversal] If the DeepTraversal element is not present, the subfolders are not searched.");
        }

        /// <summary>
        /// This test case is used to verify that if user access to the resource is denied, server return status 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC24_Search_DocumentLibraryStatus5()
        {
            #region Create Search request for document library folder with unauthorized user
            string rootFolder = Common.GetConfigurationPropertyValue("SharedFolder", this.Site);

            // Set user2 who has no permission to visit the specified folder.
            SearchRequest searchRequest = CreateSearchDocumentLibraryRequest(rootFolder, this.User2Information.UserName, this.User2Information.UserPassword);
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4339");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4339
            Site.CaptureRequirementIfAreEqual<string>(
                "5",
                searchResponse.ResponseData.Response.Store.Status,
                4339,
                @"[In Status(Search)] [When the scope is Global], [the cause of the status value 5 is] Access was denied to the resource");
        }

        /// <summary>
        /// This test case is used to verify that if user access the resource is not found, server return status 6.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC25_Search_DocumentLibraryStatus6()
        {
            #region Create Search request pointing to invalid document library folder
            string invalidRootFolder = "\\\\" + Common.GetConfigurationPropertyValue("SutComputerName", this.Site) + "\\" + "InvalidFolder";
            SearchRequest searchRequest = CreateSearchDocumentLibraryRequest(invalidRootFolder, this.User1Information.UserName, this.User1Information.UserPassword);
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4342");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4342
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                searchResponse.ResponseData.Response.Store.Status,
                4342,
                @"[In Status(Search)] [When the scope is Global], [the cause of the status value 6 is] Resource was not found.");
        }

        /// <summary>
        /// This test case is used to verify search GAL contacts with photo success.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC26_Search_GAL_Status173()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "12.1" && Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "14.0", "Status value 173 is not returned when the MS-ASProtocolVersion header value is set to 12.1 or 14.0.");

            #region Create Search request with options
            uint maxPictures = 3;
            uint maxSize = 10 * 1024;
            int rangeEndIndex = 2;
            string requestRange = "0-" + rangeEndIndex.ToString();

            // Set search query
            string keyWord = "MSASCMD_User";
            SearchRequest searchRequest = this.CreateSearchGALRequest(maxPictures, maxSize, requestRange, keyWord);
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
            #endregion

            #region Verify Requirements MS-ASCMD_R3674, MS-ASCMD_R3311, MS-ASCMD_R3310, MS-ASCMD_R4372, MS-ASCMD_R4373, MS-ASCMD_R3691, MS-ASCMD_R3692
            bool isContactPictureElementContainsData = false;
            bool isContactPhotoLessThanMaxSize = false;
            bool isPictureContainDataWithStatus1 = false;
            bool isContactWithoutPhotoWithStatus173 = false;
            int contactWithPhotoCount = 0;

            List<object> picturesfromResponse = GetElementsFromSearchResponse(searchResponse, Response.ItemsChoiceType6.Picture1);
            foreach (object item in picturesfromResponse)
            {
                Response.SearchResponseStoreResultPropertiesPicture itemPicture = (Response.SearchResponseStoreResultPropertiesPicture)item;
                if (itemPicture.Data != null)
                {
                    isContactPictureElementContainsData = true;
                    contactWithPhotoCount++;
                    int pictureSize = Convert.ToInt32(itemPicture.Data);

                    // If all contact picture size is less than MaxSize in request ,then isR3310Verified is true.
                    isContactPhotoLessThanMaxSize = pictureSize < maxSize;

                    // If contact photo was retrieved, status value is 1.
                    isPictureContainDataWithStatus1 = itemPicture.Status.Equals("1");
                }
                else
                {
                    // If contact doesn't have picture , status value is 173.
                    if (itemPicture.Status.Equals("173"))
                    {
                        isContactWithoutPhotoWithStatus173 = true;
                    }
                }
            }

            // If user calls Search GAL response contains Picture element and the Picture element contains non empty Data element, then MS-ASCMD_R3667, MS-ASCMD_R2135 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3667");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3667
            Site.CaptureRequirementIfIsTrue(
                isContactPictureElementContainsData,
                3667,
                @"[In Picture] The Picture element<63> is used in the following command requests and responses: ResolveRecipients command requests and responses (section 2.2.2.13), Search command requests and responses (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2135");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2135
            Site.CaptureRequirementIfIsTrue(
                isContactPictureElementContainsData,
                2135,
                @"[In Data(Search)] The gal:Data element<28> is an optional child element of the gal:Picture element in Search command responses that contains the binary data of the contact photo.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3674");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3674
            Site.CaptureRequirementIfIsTrue(
                isContactPictureElementContainsData,
                3674,
                @"[In Picture(Search)] The gal:Picture element<66> is an optional child element of the Properties element in Search command responses that contains the data related to the contact photos.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3311");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3311
            Site.CaptureRequirementIfIsTrue(
                contactWithPhotoCount <= maxPictures,
                3311,
                @"[In MaxSize(Search)] The MaxPictures element (section 2.2.3.94) specifies the maximum number of contact photos to return in the server response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3310");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3310
            Site.CaptureRequirementIfIsTrue(
                isContactPhotoLessThanMaxSize,
                3310,
                @"[In MaxSize(Search)] The MaxSize element specifies the maximum size of an individual contact photo that is returned in the response, in bytes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4372");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4372
            Site.CaptureRequirementIfIsTrue(
                isPictureContainDataWithStatus1,
                4372,
                @"[In Status(Search)] [the meaning of the status value] 1 [is] The contact photo was retrieved successfully.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4373");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4373
            Site.CaptureRequirementIfIsTrue(
                isContactWithoutPhotoWithStatus173,
                4373,
                @"[In Status(Search)] [The meaning of the status value] 173 is] The user does not have a contact photo.");

            bool isAllResultHasProperty = false;
            bool isNotAllResultHasSamePropertyLength = false;

            int item0ProperyElementCount = searchResponse.ResponseData.Response.Store.Result[0].Properties.Items.Length;
            foreach (Response.SearchResponseStoreResult result in searchResponse.ResponseData.Response.Store.Result)
            {
                // If all search result contains properties , then isR3691Verified is true.
                isAllResultHasProperty = result.Properties != null;

                // If search result has different properties item length, means different set of properties can be returned.
                if (result.Properties.Items.Length != item0ProperyElementCount)
                {
                    isNotAllResultHasSamePropertyLength = true;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3691");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3691
            Site.CaptureRequirementIfIsTrue(
                isAllResultHasProperty,
                3691,
                @"[In Properties(Search)] Only those properties that are attached to the specific GAL entry are returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3692");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3692
            Site.CaptureRequirementIfIsTrue(
                isNotAllResultHasSamePropertyLength,
                3692,
                @"[In Properties(Search)] Therefore different sets of properties can be returned in the response XML for different matching GAL entries.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if ConversationId element is included outside of And element, server will return status 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC27_Search_TooComplex_Status8_ConversationId()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1") || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"), "The ConversationId element is not supported when the MS-ASProtocolVersion header is set to 12.1");

            #region Create a default search request
            // Create default search request.
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { "0-0", string.Empty },
                    ItemsElementName = new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.Range, Request.ItemsChoiceType6.RebuildResults
                    }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.CollectionId
            };

            ((Request.queryType)store.Query.Items[0]).Items = new object[] { this.User1Information.InboxCollectionId };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Create invalid search request with conversationId element outside of And element
            string conversationIdElement = "<ConversationId><![CDATA[ BBA4726D4399D44C83297D4BD904ED2D]]></ConversationId>";

            // Insert conversationID element before And element.
            string originalXmlRequest = searchRequest.GetRequestDataSerializedXML();

            // Insert element before tag
            string invalidSearchRequest = originalXmlRequest.Insert(originalXmlRequest.IndexOf("<And", StringComparison.OrdinalIgnoreCase), conversationIdElement);
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string status = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2086");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2086
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                status,
                2086,
                @"[In ConversationId(Search)] If the ConversationId element is included as a child of any element other than the And element, the server responds with a Status element (section 2.2.3.162.12) value of 8 (SearchTooComplex).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify Search comparisons are performed by using case-insensitive matching.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC28_Search_CaseInsensitiveMatch()
        {
            #region User1 calls SendMail command to send 2 email messages to user2.
            string upperCaseSearchKeyword = "KEYWORD" + TestSuiteBase.ClientId;
            uint mailIndex = 1;
            string emailSubject1 = upperCaseSearchKeyword + Common.GenerateResourceName(Site, "search", mailIndex);
            SendMailResponse responseSendMail = this.SendPlainTextEmail(null, emailSubject1, this.User1Information.UserName, this.User2Information.UserName, null);
            Site.Assert.AreEqual(string.Empty, responseSendMail.ResponseDataXML, "If SendMail command executes successfully, server should return empty xml data");
            mailIndex++;
            string emailSubject2 = upperCaseSearchKeyword + Common.GenerateResourceName(Site, "search", mailIndex);
            SendMailResponse responseSendMail2 = this.SendPlainTextEmail(null, emailSubject2, this.User1Information.UserName, this.User2Information.UserName, null);
            Site.Assert.AreEqual(string.Empty, responseSendMail2.ResponseDataXML, "If SendMail command executes successfully, server should return empty xml data");
            #endregion

            #region Sync user2 mailbox changes
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject1);
            this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject2);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject1, emailSubject2);
            #endregion

            #region Create search request with low case search prefix
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { "0-0", string.Empty },
                    ItemsElementName = new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.Range, Request.ItemsChoiceType6.RebuildResults
                    }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.CollectionId,
                Request.ItemsChoiceType2.FreeText
            };

            // Search with low case search key word.
            ((Request.queryType)store.Query.Items[0]).Items = new object[] { this.User2Information.InboxCollectionId, upperCaseSearchKeyword.ToLower(System.Globalization.CultureInfo.InvariantCulture) };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Call search command
            SearchResponse searchResponse = this.LoopSearch(searchRequest, 2);
            #endregion

            #region Verify Requirements MS-ASCMD_R3729, MS-ASCMD_R3699
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3729");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3729
            Site.CaptureRequirementIfIsTrue(
                searchResponse.ResponseData.Response.Store.Result.Length == 1 && Convert.ToInt32(searchResponse.ResponseData.Response.Store.Total) >= 2,
                3729,
                @"[In Range(Search)] A Range element value of 0-0 indicates 1 item.");

            string subjectInSearchResult = (string)GetItemFromSearchResult(searchResponse.ResponseData.Response.Store.Result[0], Response.ItemsChoiceType6.Subject1);

            // If search use lower case search keyword get the correct result, then MS-ASCMD_R3699 is verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3699");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3699
            Site.CaptureRequirementIfIsTrue(
                subjectInSearchResult.Equals(emailSubject1) || subjectInSearchResult.Equals(emailSubject2),
                3699,
                @"[In Query] Search comparisons are performed by using case-insensitive matching.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the Search command with more than one ConversationId elements in And element, and the returned status value is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC29_Search_MoreThanOneConversationId()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0", "[In Appendix A: Product Behavior] <23> Section 2.2.3.35.2: The ConversationId element is not supported when the MS-ASProtocolVersion header is set to 12.1.");
            #region Create a search request with multiple ConversationId elements.
            string conversationId = Guid.NewGuid().ToString("N");
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[]
            {
                Request.ItemsChoiceType2.ConversationId,
                Request.ItemsChoiceType2.ConversationId,
                Request.ItemsChoiceType2.FreeText
            };

            ((Request.queryType)store.Query.Items[0]).Items = new object[] { conversationId, conversationId, "KeyWord" };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Call Search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            #endregion

            if (Common.IsRequirementEnabled(5815, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5815");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5815
                Site.CaptureRequirementIfAreNotEqual<string>(
                    "1",
                    searchResponse.ResponseData.Response.Store.Status,
                    5815,
                    @"[In Appendix A: Product Behavior] The implementation does return a protocol status error in response to such a command request [more than one ConversationId element in a Search command request]. (Exchange 2007 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify if Search request includes Or element, server will return status 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC30_Search_TooComplex_Status8_Or()
        {
            #region Creates search request with Or element
            // Create default search request.
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Or },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Class, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)store.Query.Items[0]).Items = new object[] { "Email", "KeyWord" };
            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            #endregion

            #region Calls search command
            SearchResponse searchResponse = this.CMDAdapter.Search(searchRequest);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when the GreaterThan element of Search command is not a child element for And element, the server should return the response with Status 8
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC31_Search_TooComplex_Status8_GreaterThan()
        {
            #region Create a search request in which a GreaterThan element is not in And element.
            string greaterThanElement = "<GreaterThan><DateReceived  xmlns=\"Email\"/><Value>2012-03-28T07:31:09.314Z</Value></GreaterThan>";

            // Insert GreaterThan element before And element
            string invalidSearchRequest = this.CreateInvalidSearchRequest(greaterThanElement, "<And>");
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string status = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3143");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3143
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                status,
                3143,
                @"[In GreaterThan] If the GreaterThan element is included as a child element of any element other than the And element (section 2.2.3.10), the server responds with a Status element (section 2.2.3.162.12) value of 8 (SearchTooComplex).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify when the LessThan element of Search command is not a child element for And element, the server should return the response with Status 8.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC32_Search_TooComplex_Status8_LessThan()
        {
            #region Create a request with LessThan element not in And element.
            string lessThanElement = "<LessThan><DateReceived  xmlns=\"Email\"/><Value>50</Value></LessThan>";

            // Insert LessThan element before And element
            string invalidSearchRequest = this.CreateInvalidSearchRequest(lessThanElement, "<And>");
            #endregion

            #region Call method SendStringRequest to send a plain text request.
            SendStringResponse response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);
            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("e", "Search");

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int counter = 1;
            XmlNode searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);

            while (counter < retryCount && searchStatus != null && searchStatus.InnerXml.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.CMDAdapter.SendStringRequest(CommandName.Search, null, invalidSearchRequest);
                doc.LoadXml(response.ResponseDataXML);
                xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("e", "Search");
                searchStatus = doc.SelectSingleNode("/e:Search/e:Status", xnm);
                counter++;
            }

            string status = Common.GetSearchStatusCode(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3220");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3220
            Site.CaptureRequirementIfAreEqual<string>(
                "8",
                status,
                3220,
                @"[In LessThan] If the LessThan element is included as a child element of any element other than the And element (section 2.2.3.10), the server responds with a Status element (section 2.2.3.162.12) value of 8 (SearchTooComplex).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the response of searching mail operation with MIMESupport setting that contains the Body element.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S14_TC33_Search_MIMESupport()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Calls Search command with MIMESupport set to 1.
            byte mimeSupportValue = 1;
            SearchRequest searchRequest = this.CreateSearchRequestWithMimeSupport(subject, mimeSupportValue);
            SearchResponse searchResponse = this.LoopSearch(searchRequest);

            bool containBodyElement = FindBodyElementInSearchResponse(searchResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3383");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3383
            Site.CaptureRequirementIfIsTrue(
                searchResponse.ResponseData.Response.Store.Status.Equals("1") && containBodyElement,
                3383,
                @"[In MIMESupport(Search)][If the airsync:MIMESupport element is set to 1 in the Search request:] The response from the server MUST include the airsyncbase:Body element ([MS-ASAIRS] section 2.2.2.4), which is a child element of the Properties element (section 2.2.3.128.2).");
            #endregion

            #region Calls Search command with MIMESupport set to 2.
            mimeSupportValue = 2;
            SearchRequest newSearchRequest = this.CreateSearchRequestWithMimeSupport(subject, mimeSupportValue);
            SearchResponse newSearchResponse = this.LoopSearch(newSearchRequest);

            containBodyElement = FindBodyElementInSearchResponse(newSearchResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3384");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3384
            Site.CaptureRequirementIfIsTrue(
                newSearchResponse.ResponseData.Status.Equals("1") && containBodyElement,
                3384,
                @"[In MIMESupport(Search)][If the airsync:MIMESupport element is set to 2 in the Search request:] The response from the server MUST include the airsyncbase:Body element ([MS-ASAIRS] section 2.2.2.4), which is a child element of the Properties element (section 2.2.3.128.2).");

            Response.Body bodyElement = (Response.Body)GetItemFromSearchResult(newSearchResponse.ResponseData.Response.Store.Result[0], Response.ItemsChoiceType6.Body);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3388");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3388
            Site.CaptureRequirementIfAreEqual<byte>(
                4,
                bodyElement.Type,
                3388,
                @"[In MIMESupport(Search)] The airsyncbase:Type element ([MS-ASAIRS] section 2.2.2.22.1) with a value of 4 to inform the device that the data is a MIME BLOB.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3387");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3387
            Site.CaptureRequirementIfIsTrue(
                bodyElement.Type == (byte)4 && bodyElement.EstimatedDataSizeSpecified == true && bodyElement.TruncatedSpecified == true && bodyElement.Data != null,
                3387,
                @"[In MIMESupport(Search)] [The airsyncbase:Body element] MUST contain the following child elements [the airsyncbase:Type element, the airsyncbase:EstimatedDataSize element, the airsyncbase:Truncated element, the airsyncbase:Data element] in an S/MIME Search response.");
            #endregion
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Verify if the classElement is supported.
        /// </summary>
        /// <param name="classElement">The class element.</param>
        /// <returns>Return true if class element is supported, return false if class element is not supported.</returns>
        private static bool IsClassSupported(string classElement)
        {
            switch (classElement)
            {
                case "Email":
                    return true;
                case "Tasks":
                    return true;
                case "Calendar":
                    return true;
                case "Contacts":
                    return true;
                case "Notes":
                    return true;
                case "SMS":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Create search document library request
        /// </summary>
        /// <param name="pointTo">Point to specified folder or item.</param>
        /// <param name="userName">The user name need to access the resource.</param>
        /// <param name="userPassword">The user password.</param>
        /// <returns>The search request.</returns>
        private static SearchRequest CreateSearchDocumentLibraryRequest(string pointTo, string userName, string userPassword)
        {
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.DocumentLibrary.ToString(),
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.EqualTo },
                    Items = new Request.queryTypeEqualTo[]
                    {
                        new Request.queryTypeEqualTo
                        {
                            LinkId = string.Empty,
                            Value = pointTo
                        }
                    }
                },
                Options = new Request.Options1
                {
                    ItemsElementName =
                        new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.UserName, Request.ItemsChoiceType6.Password },
                    Items = new object[] { userName, userPassword }
                }
            };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            return searchRequest;
        }

        /// <summary>
        /// Get item value from single Result element of Search command.
        /// </summary>
        /// <param name="searchResult">The single Result element of Search command.</param>
        /// <param name="itemType">The item type.</param>
        /// <returns>The item value.</returns>
        private static object GetItemFromSearchResult(Response.SearchResponseStoreResult searchResult, Response.ItemsChoiceType6 itemType)
        {
            for (int index = 0; index < searchResult.Properties.ItemsElementName.Length; index++)
            {
                if (searchResult.Properties.ItemsElementName[index] == itemType)
                {
                    return searchResult.Properties.Items[index];
                }
            }

            return null;
        }

        /// <summary>
        /// Check if all search results contain Body element
        /// </summary>
        /// <param name="searchResponse">The searchResponse</param>
        /// <returns>If all search result contains Body element return true, else return false</returns>
        private static bool FindBodyElementInSearchResponse(SearchResponse searchResponse)
        {
            // Check search result's properties contains Body element
            foreach (Response.SearchResponseStoreResult result in searchResponse.ResponseData.Response.Store.Result)
            {
                bool containBodyElement = false;
                for (int index = 0; index < result.Properties.ItemsElementName.Length; index++)
                {
                    if (result.Properties.ItemsElementName[index] == Response.ItemsChoiceType6.Body && result.Properties.Items[index] != null)
                    {
                        containBodyElement = true;
                        break;
                    }
                }

                if (containBodyElement == false)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Get elements from Search command response.
        /// </summary>
        /// <param name="searchResponse">The Search command response.</param>
        /// <param name="elementType">The element type need to get.</param>
        /// <returns>The element object list.</returns>
        private static List<object> GetElementsFromSearchResponse(SearchResponse searchResponse, Response.ItemsChoiceType6 elementType)
        {
            List<object> element = new List<object>();
            foreach (Response.SearchResponseStoreResult result in searchResponse.ResponseData.Response.Store.Result)
            {
                for (int itemIndex = 0; itemIndex < result.Properties.ItemsElementName.Length; itemIndex++)
                {
                    if (result.Properties.ItemsElementName[itemIndex] == elementType)
                    {
                        element.Add(result.Properties.Items[itemIndex]);
                    }
                }
            }

            return element;
        }

        /// <summary>
        /// Loop calls search command to search special items.
        /// </summary>
        /// <param name="searchRequest">The search request.</param>
        /// <returns>The search response contains results.</returns>
        private SearchResponse LoopSearch(SearchRequest searchRequest)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int resultsCount;
            SearchResponse searchResponse;

            // Loop search to get correct results.
            do
            {
                Thread.Sleep(waitTime);
                searchResponse = this.CMDAdapter.Search(searchRequest);
                Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
                resultsCount = searchResponse.ResponseData.Response.Store.Result.Length;
                counter++;
            }
            while (resultsCount <= 1 && counter < retryCount && searchResponse.ResponseData.Response.Store.Result[0].Properties == null);
            Site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store.Result[0].Properties, "If search get result, the result should contains properties");
            Site.Log.Add(LogEntryKind.Debug, "Loop {0} times to get the search item", counter);
            return searchResponse;
        }

        /// <summary>
        /// Loop calls search command to get specified results.
        /// </summary>
        /// <param name="searchRequest">The search request.</param>
        /// <param name="totalResult">The result count.</param>
        /// <returns>The search response contains required results.</returns>
        private SearchResponse LoopSearch(SearchRequest searchRequest, int totalResult)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int resultsCount;
            SearchResponse searchResponse;

            // Loop search to get correct results.
            do
            {
                Thread.Sleep(waitTime);
                searchResponse = this.CMDAdapter.Search(searchRequest);
                Site.Assert.AreEqual("1", searchResponse.ResponseData.Response.Store.Status, "If server successfully completed command, server should return status 1");
                resultsCount = searchResponse.ResponseData.Response.Store.Result.Length;
                counter++;
            }
            while (resultsCount <= 1 && counter < retryCount && (searchResponse.ResponseData.Response.Store.Result[0].Properties == null || Convert.ToInt32(searchResponse.ResponseData.Response.Store.Total) < totalResult));
            Site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store.Result[0].Properties, "If search get result, the result should contains properties");
            Site.Log.Add(LogEntryKind.Debug, "Loop {0} times to get the search item", counter);
            return searchResponse;
        }

        /// <summary>
        /// Respectively create one item in User2 Inbox folder, Calendar folder and Contacts subfolder.
        /// </summary>
        /// <param name="searchPrefix">The prefix in created item.</param>
        private void CreateItemsWithKeyword(string searchPrefix)
        {
            // User1 sends one meeting request mail to use2.
            string meetingRequestSubject = searchPrefix + Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2.
            this.SendMeetingRequest(meetingRequestSubject, calendar);

            // Switch to user2 mail box.
            this.SwitchUser(this.User2Information);

            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);

            // Create one contact in Contacts subfolder.
            string contactSubfolderID = this.CreateFolder((byte)FolderType.UserCreatedContacts, Common.GenerateResourceName(Site, "FolderCreate"), this.User2Information.ContactsCollectionId);
            this.FolderSync();
            TestSuiteBase.RecordCaseRelativeFolders(this.User2Information, contactSubfolderID);

            string contactFileAS = searchPrefix + Common.GenerateResourceName(Site, "Contact");
            Request.SyncCollectionAdd addData = this.CreateAddContactCommand("FirstName", "MiddleName", "LastName", contactFileAS, null);

            this.GetInitialSyncResponse(contactSubfolderID);
            SyncRequest syncRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, contactSubfolderID, addData);
            this.Sync(syncRequest);
            this.FolderSync();
            this.SyncChanges(contactSubfolderID);
        }

        /// <summary>
        /// Create one search request with MIMESupport and related element
        /// </summary>
        /// <param name="keyWord">Search keyword</param>
        /// <param name="mimeSupportValue">MIMESupport element value</param>
        /// <returns>The search request</returns>
        private SearchRequest CreateSearchRequestWithMimeSupport(string keyWord, byte mimeSupportValue)
        {
            // Create search Option element
            Request.Options1 searchOption = new Request.Options1();
            searchOption.ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.RebuildResults, Request.ItemsChoiceType6.DeepTraversal, Request.ItemsChoiceType6.MIMESupport, Request.ItemsChoiceType6.BodyPreference };

            // Set bodyPrference element value
            Request.BodyPreference bodyPreference = new Request.BodyPreference();
            bodyPreference.Type = 4;
            bodyPreference.TruncationSize = 100;
            bodyPreference.TruncationSizeSpecified = true;

            // Set search Option element value
            searchOption.Items = new object[] { string.Empty, string.Empty, mimeSupportValue, bodyPreference };

            // Create search Query element
            Request.queryType searchQuery = new Request.queryType();
            searchQuery.ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And };
            searchQuery.Items = new Request.queryType[] { new Request.queryType() };
            ((Request.queryType)searchQuery.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Class, Request.ItemsChoiceType2.CollectionId, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)searchQuery.Items[0]).Items = new object[] { "Email", this.User1Information.InboxCollectionId, keyWord };

            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Options = searchOption;
            searchRequest.RequestData.Items[0].Query = searchQuery;
            return searchRequest;
        }

        /// <summary>
        /// Create invalid Search request.
        /// </summary>
        /// <param name="element">Additional element insert into default search request.</param>
        /// <param name="insertTag">Insert tag shows where the additional element should be inserted.</param>
        /// <returns>String search quest.</returns>
        private string CreateInvalidSearchRequest(string element, string insertTag)
        {
            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            string originalXmlRequest = searchRequest.GetRequestDataSerializedXML();

            // Remove original element from request.
            if (searchRequest.GetRequestDataSerializedXML().Contains(element))
            {
                originalXmlRequest = searchRequest.GetRequestDataSerializedXML().Replace(element, string.Empty);
            }

            // Insert element before insertTag.
            string invalidSearchRequest = originalXmlRequest.Insert(originalXmlRequest.IndexOf(insertTag, StringComparison.OrdinalIgnoreCase), element);
            return invalidSearchRequest;
        }

        /// <summary>
        /// Create Search GAL request.
        /// </summary>
        /// <param name="maxPictures">The maxPictures value.</param>
        /// <param name="maxSize">The maxSize value.</param>
        /// <param name="requestRange">The range value.</param>
        /// <param name="keyWord">The search key word.</param>
        /// <returns>The GAL search request.</returns>
        private SearchRequest CreateSearchGALRequest(uint maxPictures, uint maxSize, string requestRange, string keyWord)
        {
            // Create search request with range, maxSize, maxPictures options.
            Request.Options1 searchOption = new Request.Options1
            {
                ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.Range, Request.ItemsChoiceType6.Picture }
            };

            if (maxPictures > 0 && maxSize > 0)
            {
                Request.OptionsPicture picture = new Request.OptionsPicture
                {
                    MaxPictures = maxPictures,
                    MaxPicturesSpecified = true,
                    MaxSize = maxSize,
                    MaxSizeSpecified = true
                };
                searchOption.Items = new object[] { requestRange, picture };
            }
            else
            {
                searchOption.Items = new object[] { requestRange };
            }

            Request.queryType searchQuery = new Request.queryType { Text = new string[] { keyWord } };

            // Set Name element, option element, query element in default search request.
            SearchRequest searchRequest = this.CreateDefaultSearchRequest();
            searchRequest.RequestData.Items[0].Name = SearchName.GAL.ToString();
            searchRequest.RequestData.Items[0].Options = searchOption;
            searchRequest.RequestData.Items[0].Query = searchQuery;

            return searchRequest;
        }

        /// <summary>
        /// Create one search request with default value.
        /// </summary>
        /// <returns>Return a SearchRequest instance.</returns>
        private SearchRequest CreateDefaultSearchRequest()
        {
            Request.SearchStore store = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { string.Empty },
                    ItemsElementName = new Request.ItemsChoiceType6[] { Request.ItemsChoiceType6.DeepTraversal }
                },
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)store.Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.Class, Request.ItemsChoiceType2.CollectionId, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)store.Query.Items[0]).Items = new object[] { "Email", this.User1Information.InboxCollectionId, "FreeText" };

            SearchRequest searchRequest = Common.CreateSearchRequest(new Request.SearchStore[] { store });
            return searchRequest;
        }
        #endregion
    }
}