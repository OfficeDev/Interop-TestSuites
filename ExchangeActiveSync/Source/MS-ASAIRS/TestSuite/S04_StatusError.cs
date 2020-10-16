namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the status error which is returned by the Sync command, Search command and ItemOperations command when the XML elements in AirSyncBase namespace don't comply with the requirements regarding data type, number of instance, order and placement in the XML hierarchy.
    /// </summary>
    [TestClass]
    public class S04_StatusError : TestSuiteBase
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
        public static void ClassCleanUp()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region MSASAIRS_S04_TC01_Sync_Status4
        /// <summary>
        /// This case is designed to test if the child elements of BodyPreference are not in the correct order, the AllOrNone element is not of type Boolean, or multiple AllOrNone elements are in a single BodyPreference element, the server will return protocol status error 4 for a Sync command.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S04_TC01_Sync_Status4()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);

            // Make sure the email has reached the inbox folder of the recipient
            this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Call Sync command with incorrect order of child elements of BodyPreference
            SyncRequest request = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User2Information.InboxCollectionId), this.User2Information.InboxCollectionId, null, bodyPreference, null);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(request.GetRequestDataSerializedXML());
            XmlNode bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");

            // Put the first node to the end.
            XmlNode temp = bodyPreferenceNode.ChildNodes[0];
            bodyPreferenceNode.RemoveChild(temp);
            bodyPreferenceNode.AppendChild(temp);

            SendStringResponse response = this.ASAIRSAdapter.Sync(doc.OuterXml);

            string status = this.GetStatusCodeFromXPath(response, "/a:Sync/a:Status");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R348");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R348
            Site.CaptureRequirementIfAreEqual(
                "4",
                status,
                348,
                @"[In Validating XML] If the child elements of BodyPreference (section 2.2.2.12) are not in the correct order, the server returns protocol status error 4 (for a Sync command).");
            #endregion

            #region Call Sync command with a non-boolean value of AllOrNone element
            doc.LoadXml(request.GetRequestDataSerializedXML());
            XmlNode allOrNoneNode = doc.SelectSingleNode("//*[name()='AllOrNone']");

            // Set the AllOrNone element value to non-boolean.
            allOrNoneNode.InnerText = "a";

            response = this.ASAIRSAdapter.Sync(doc.OuterXml);

            status = this.GetStatusCodeFromXPath(response, "/a:Sync/a:Status");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R350");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R350
            Site.CaptureRequirementIfAreEqual(
                "4",
                status,
                350,
                @"[In Validating XML] If the AllOrNone (section 2.2.2.3.2) element is not of type boolean, the server returns protocol status error 4 (for Sync command).");
            #endregion

            #region Call Sync command with multiple AllOrNone elements in a single BodyPreference element
            doc.LoadXml(request.GetRequestDataSerializedXML());
            bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");
            allOrNoneNode = doc.SelectSingleNode("//*[name()='AllOrNone']");
            temp = allOrNoneNode.Clone();
            bodyPreferenceNode.AppendChild(temp);

            response = this.ASAIRSAdapter.Sync(doc.OuterXml);

            status = this.GetStatusCodeFromXPath(response, "/a:Sync/a:Status");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R351");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R351
            Site.CaptureRequirementIfAreEqual(
                "4",
                status,
                351,
                @"[In Validating XML] If multiple AllOrNone elements are in a single BodyPreference element, the server returns protocol status error 4 (for Sync command).");
            #endregion
        }
        #endregion

        #region MSASAIRS_S04_TC02_IncorrectDataType
        /// <summary>
        /// This case is designed to test server will return protocol status error 2 for an ItemOperations command or Search command, protocol status error 6 for a Sync command, if an element doesn't meet the specified data type.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S04_TC02_IncorrectDataType()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Call ItemOperations command with the Type element value in incorrect data type
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            if (Common.IsRequirementEnabled(346, this.Site))
            {
                ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsRequest(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(itemOperationRequest.GetRequestDataSerializedXML());
                XmlNode typeNode = doc.SelectSingleNode("//*[name()='Type']");
                typeNode.InnerText = "a";

                SendStringResponse itemOperationResponse = this.ASAIRSAdapter.ItemOperations(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(itemOperationResponse, "/i:ItemOperations/i:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R346");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R346
                Site.CaptureRequirementIfAreEqual(
                    "2",
                    status,
                    346,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for an ItemOperations command (as specified in [MS-ASCMD] section 2.2.2.8), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding data type] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Call Search command with the Type element value in incorrect data type
                if (Common.IsRequirementEnabled(10033, this.Site))
                {
                    SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(searchRequest.GetRequestDataSerializedXML());
                    XmlNode typeNode = doc.SelectSingleNode("//*[name()='Type']");
                    typeNode.InnerText = "a";

                    SendStringResponse searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                    string searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                    int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                    int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                    int counter = 1;

                    while (counter < retryCount && searchStatus.Equals("10"))
                    {
                        Thread.Sleep(waitTime);
                        searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                        searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                        counter++;
                    }

                    string status = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Response/s:Store/s:Status");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10033");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R10033
                    Site.CaptureRequirementIfAreEqual(
                        "2",
                        status,
                        10033,
                        @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for a Search command (as specified in [MS-ASCMD] section 2.2.2.14), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding data type] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
                }
                #endregion
            }

            #region Call Sync command with the Type element value in incorrect data type
            if (Common.IsRequirementEnabled(10014, this.Site))
            {
                SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User2Information.InboxCollectionId), this.User2Information.InboxCollectionId, CreateSyncAddCommands(), null, null);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(syncAddRequest.GetRequestDataSerializedXML());
                XmlNode typeNode = doc.SelectSingleNode("//*[name()='Type']");
                typeNode.InnerText = "a";

                SendStringResponse syncAddResponse = this.ASAIRSAdapter.Sync(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(syncAddResponse, "/a:Sync/a:Collections/a:Collection/a:Responses/a:Add/a:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10014");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10014
                Site.CaptureRequirementIfAreEqual(
                    "6",
                    status,
                    10014,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 6 for a Sync command (as specified in [MS-ASCMD] section 2.2.2.19), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding data type] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASAIRS_S04_TC03_IncorrectNumberOfInstances
        /// <summary>
        /// This case is designed to test server will return protocol status error 2 for an ItemOperations command or Search command, protocol status error 6 for a Sync command, if an element doesn't meet the number of specified instances.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S04_TC03_IncorrectNumberOfInstances()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Call ItemOperations command with multiple AllOrNone elements
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            if (Common.IsRequirementEnabled(10030, this.Site))
            {
                ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsRequest(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null);

                // Add another AllOrNone element in BodyPreference element
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(itemOperationRequest.GetRequestDataSerializedXML());
                XmlNode bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");
                XmlNode allOrNoneNode = doc.SelectSingleNode("//*[name()='AllOrNone']");
                XmlNode temp = allOrNoneNode.Clone();
                bodyPreferenceNode.AppendChild(temp);

                SendStringResponse itemOperationResponse = this.ASAIRSAdapter.ItemOperations(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(itemOperationResponse, "/i:ItemOperations/i:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10030");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10030
                Site.CaptureRequirementIfAreEqual(
                    "2",
                    status,
                    10030,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for an ItemOperations command (as specified in [MS-ASCMD] section 2.2.2.8), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding number of instances] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Call Search command with multiple AllOrNone elements
                if (Common.IsRequirementEnabled(10034, this.Site))
                {
                    SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

                    // Add another AllOrNone element in BodyPreference element
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(searchRequest.GetRequestDataSerializedXML());
                    XmlNode bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");
                    XmlNode allOrNoneNode = doc.SelectSingleNode("//*[name()='AllOrNone']");
                    XmlNode temp = allOrNoneNode.Clone();
                    bodyPreferenceNode.AppendChild(temp);

                    SendStringResponse searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                    string searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                    int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                    int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                    int counter = 1;

                    while (counter < retryCount && searchStatus.Equals("10"))
                    {
                        Thread.Sleep(waitTime);
                        searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                        searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                        counter++;
                    }

                    string status = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Response/s:Store/s:Status");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10034");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R10034
                    Site.CaptureRequirementIfAreEqual(
                        "2",
                        status,
                        10034,
                        @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for a Search command (as specified in [MS-ASCMD] section 2.2.2.14), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding number of instances] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
                }
                #endregion
            }

            #region Call Sync add command with multiple Type elements in Body
            if (Common.IsRequirementEnabled(10037, this.Site))
            {
                SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User2Information.InboxCollectionId), this.User2Information.InboxCollectionId, CreateSyncAddCommands(), null, null);

                // Add another Type element in Body element
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(syncAddRequest.GetRequestDataSerializedXML());
                XmlNode bodyNode = doc.SelectSingleNode("//*[name()='Body']");
                XmlNode typeNode = bodyNode.SelectSingleNode("//*[name()='Type']");
                XmlNode temp = typeNode.Clone();
                bodyNode.InsertBefore(temp, typeNode);

                SendStringResponse syncAddResponse = this.ASAIRSAdapter.Sync(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(syncAddResponse, "/a:Sync/a:Collections/a:Collection/a:Responses/a:Add/a:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10037");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10037
                Site.CaptureRequirementIfAreEqual(
                    "6",
                    status,
                    10037,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 6 for a Sync command (as specified in [MS-ASCMD] section 2.2.2.19), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding number of instances] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASAIRS_S04_TC04_IncorrectOrder
        /// <summary>
        /// This case is designed to test server will return protocol status error 2 for an ItemOperations command or Search command, protocol status error 6 for a Sync command, if elements don't meet the specified order.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S04_TC04_IncorrectOrder()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Call ItemOperations command with incorrect BodyPreference element order
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            if (Common.IsRequirementEnabled(10031, this.Site))
            {
                ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsRequest(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(itemOperationRequest.GetRequestDataSerializedXML());
                XmlNode bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");

                // Put the first node to the end.
                XmlNode temp = bodyPreferenceNode.ChildNodes[0];
                bodyPreferenceNode.RemoveChild(temp);
                bodyPreferenceNode.AppendChild(temp);

                SendStringResponse itemOperationResponse = this.ASAIRSAdapter.ItemOperations(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(itemOperationResponse, "/i:ItemOperations/i:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10031");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10031
                Site.CaptureRequirementIfAreEqual(
                    "2",
                    status,
                    10031,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for an ItemOperations command (as specified in [MS-ASCMD] section 2.2.2.8), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding order] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Call Search command with incorrect BodyPreference element order
                if (Common.IsRequirementEnabled(10035, this.Site))
                {
                    SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(searchRequest.GetRequestDataSerializedXML());
                    XmlNode bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");

                    // Put the first node to the end.
                    XmlNode temp = bodyPreferenceNode.ChildNodes[0];
                    bodyPreferenceNode.RemoveChild(temp);
                    bodyPreferenceNode.AppendChild(temp);

                    SendStringResponse searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                    string searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                    int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                    int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                    int counter = 1;

                    while (counter < retryCount && searchStatus.Equals("10"))
                    {
                        Thread.Sleep(waitTime);
                        searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                        searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                        counter++;
                    }

                    string status = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Response/s:Store/s:Status");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10035");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R10035
                    Site.CaptureRequirementIfAreEqual(
                        "2",
                        status,
                        10035,
                        @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for a Search command (as specified in [MS-ASCMD] section 2.2.2.14), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding order] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
                }
                #endregion
            }

            #region Call Sync add command with incorrect body element order
            if (Common.IsRequirementEnabled(10038, this.Site))
            {
                SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User2Information.InboxCollectionId), this.User2Information.InboxCollectionId, CreateSyncAddCommands(), null, null);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(syncAddRequest.GetRequestDataSerializedXML());
                XmlNode bodyNode = doc.SelectSingleNode("//*[name()='Body']");

                // Put the first node to the end.
                XmlNode temp = bodyNode.ChildNodes[0];
                bodyNode.RemoveChild(temp);
                bodyNode.AppendChild(temp);

                SendStringResponse syncAddResponse = this.ASAIRSAdapter.Sync(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(syncAddResponse, "/a:Sync/a:Collections/a:Collection/a:Responses/a:Add/a:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10038");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10038
                Site.CaptureRequirementIfAreEqual(
                    "6",
                    status,
                    10038,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 6 for a Sync command (as specified in [MS-ASCMD] section 2.2.2.19), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding order] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASAIRS_S04_TC05_IncorrectPlacement
        /// <summary>
        /// This case is designed to test server will return protocol status error 2 for an ItemOperations command or Search command, protocol status error 6 for a Sync command, if an element doesn't meet the specified placement.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S04_TC05_IncorrectPlacement()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);
            #endregion

            #region Set BodyPreference element
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Call ItemOperations command with incorrect placement of BodyPreference element.
            DataStructures.Sync syncItem = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);

            if (Common.IsRequirementEnabled(10032, this.Site))
            {
                ItemOperationsRequest itemOperationRequest = TestSuiteHelper.CreateItemOperationsRequest(this.User2Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference, null);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(itemOperationRequest.GetRequestDataSerializedXML());
                XmlNode bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");

                // Add another BodyPreference element in the BodyPreference element, the placement is invalid.
                XmlNode temp = bodyPreferenceNode.Clone();
                temp.SelectSingleNode("//*[name()='Type']").InnerText = "2";
                bodyPreferenceNode.AppendChild(temp);

                SendStringResponse itemOperationResponse = this.ASAIRSAdapter.ItemOperations(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(itemOperationResponse, "/i:ItemOperations/i:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10032");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10032
                Site.CaptureRequirementIfAreEqual(
                    "2",
                    status,
                    10032,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for an ItemOperations command (as specified in [MS-ASCMD] section 2.2.2.8), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding placement] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion

            if (Common.IsRequirementEnabled(53, this.Site))
            {
                #region Call Search command with incorrect placement of BodyPreference element.
                if (Common.IsRequirementEnabled(10036, this.Site))
                {
                    SearchRequest searchRequest = TestSuiteHelper.CreateSearchRequest(subject, this.User2Information.InboxCollectionId, null, bodyPreference, null);
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(searchRequest.GetRequestDataSerializedXML());
                    XmlNode bodyPreferenceNode = doc.SelectSingleNode("//*[name()='BodyPreference']");

                    // Add another BodyPreference element in the BodyPreference element, the placement is invalid.
                    XmlNode temp = bodyPreferenceNode.Clone();
                    temp.SelectSingleNode("//*[name()='Type']").InnerText = "2";
                    bodyPreferenceNode.AppendChild(temp);

                    SendStringResponse searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                    string searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                    int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                    int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
                    int counter = 1;

                    while (counter < retryCount && searchStatus.Equals("10"))
                    {
                        Thread.Sleep(waitTime);
                        searchResponse = this.ASAIRSAdapter.Search(doc.OuterXml);
                        searchStatus = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Status");
                        counter++;
                    }

                    string status = this.GetStatusCodeFromXPath(searchResponse, "/s:Search/s:Response/s:Store/s:Status");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10036");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R10036
                    Site.CaptureRequirementIfAreEqual(
                        "2",
                        status,
                        10036,
                        @"[In Appendix B: Product Behavior] Implementation does return protocol status error 2 for a Search command (as specified in [MS-ASCMD] section 2.2.2.14), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding placement] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
                }
                #endregion
            }

            #region Call Sync add command with incorrect placement of Type element.
            if (Common.IsRequirementEnabled(10039, this.Site))
            {
                SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User2Information.InboxCollectionId), this.User2Information.InboxCollectionId, CreateSyncAddCommands(), null, null);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(syncAddRequest.GetRequestDataSerializedXML());
                XmlNode bodyNode = doc.SelectSingleNode("//*[name()='Body']");

                // Add another body element in the body element, the placement is invalid.
                XmlNode temp = bodyNode.Clone();
                temp.SelectSingleNode("//*[name()='Type']").InnerText = "2";
                bodyNode.AppendChild(temp);

                SendStringResponse syncAddResponse = this.ASAIRSAdapter.Sync(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(syncAddResponse, "/a:Sync/a:Collections/a:Collection/a:Responses/a:Add/a:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10039");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10039
                Site.CaptureRequirementIfAreEqual(
                    "6",
                    status,
                    10039,
                    @"[In Appendix B: Product Behavior] Implementation does return protocol status error 6 for a Sync command (as specified in [MS-ASCMD] section 2.2.2.19), if an element does not meet the requirements[any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding placement] specified for that element, unless specified in the following table[section 3.2.5.1]. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region MSASAIRS_S04_TC06_MultipleBodyPreferenceHaveSameTypeValue
        /// <summary>
        /// This case is designed to test the error will be returned if multiple BodyPreference elements are present with the same value in the Type child element.
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S04_TC06_MultipleBodyPreferenceHaveSameTypeValue()
        {
            #region Send a plain text email
            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            this.SendEmail(EmailType.Plaintext, subject, body);

            // Make sure the email has reached the inbox folder of the recipient
            this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null, null, null);
            #endregion

            #region Set two BodyPreference elements with same type value
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                },

                new Request.BodyPreference()
                {
                    Type = 1,
                    TruncationSize = 100,
                    TruncationSizeSpecified = true,
                    AllOrNone = true,
                    AllOrNoneSpecified = true
                }
            };
            #endregion

            #region Verify multiple BodyPreference elements with same type value in Sync for related requirements
            SyncRequest request = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(this.User2Information.InboxCollectionId), this.User2Information.InboxCollectionId, null, bodyPreference, null);

            if (Common.IsRequirementEnabled(10015, this.Site))
            {
                try
                {
                    this.ASAIRSAdapter.Sync(request);

                    Site.Assert.Fail("The server should return an HTTP error 500 if multiple BodyPreference elements are present with the same value in the Type child element.");
                }
                catch (WebException exception)
                {
                    int errorCode = ((HttpWebResponse)exception.Response).StatusCode.GetHashCode();

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10015");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R10015
                    Site.CaptureRequirementIfAreEqual<int>(
                        500,
                        errorCode,
                        10015,
                        @"[In Appendix B: Product Behavior] Implementation does return an HTTP error 500 instead of a Status value of 4 when multiple BodyPreference elements are present with the same value in the Type child element. (<1> Section 3.2.5.1:  Exchange 2007 SP1 returns an HTTP error 500 instead of a Status value of 4 when multiple BodyPreference elements are present with the same value in the Type child element.)");
                }
            }

            if (Common.IsRequirementEnabled(10016, this.Site))
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(request.GetRequestDataSerializedXML());

                SendStringResponse syncAddResponse = this.ASAIRSAdapter.Sync(doc.OuterXml);

                string status = this.GetStatusCodeFromXPath(syncAddResponse, "/a:Sync/a:Status");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R10016");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R10016
                Site.CaptureRequirementIfAreEqual(
                    "4",
                    status,
                    10016,
                    @"[In Appendix B: Product Behavior] Implementation does return 4 (for Sync command) if multiple BodyPreference elements are present with the same value in the Type child element. (Exchange Server 2010 and above follow this behavior.)");
            }
            #endregion
        }
        #endregion

        #region private methods
        /// <summary>
        /// Generate a Sync Add command with body element.
        /// </summary>
        /// <returns>Returns a command list instance.</returns>
        private static object[] CreateSyncAddCommands()
        {
            Request.Body addBody = new Request.Body { Type = 1, Data = "Test sync add" };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType8> itemsElementName = new List<Request.ItemsChoiceType8>();

            items.Add(addBody);
            itemsElementName.Add(Request.ItemsChoiceType8.Body);

            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData
            {
                Items = items.ToArray(),
                ItemsElementName = itemsElementName.ToArray()
            };

            Request.SyncCollectionAdd syncAdd = new Request.SyncCollectionAdd
            {
                ClientId = Guid.NewGuid().ToString("N"),
                ApplicationData = applicationData
            };

            List<object> commandList = new List<object> { syncAdd };
            return commandList.ToArray();
        }

        /// <summary>
        /// Get the status code from the response.
        /// </summary>
        /// <param name="response">The string format response.</param>
        /// <param name="xpath">The XPath to get the status code.</param>
        /// <returns>Returns the status code</returns>
        private string GetStatusCodeFromXPath(SendStringResponse response, string xpath)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response.ResponseDataXML);

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("a", "AirSync");
            nsmgr.AddNamespace("i", "ItemOperations");
            nsmgr.AddNamespace("s", "Search");

            XmlNode status = doc.SelectSingleNode(xpath, nsmgr);
            Site.Assert.IsNotNull(status, "The Status element should be returned in the response.");

            return status.InnerText;
        }
        #endregion
    }
}