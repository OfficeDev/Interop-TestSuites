namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that contains the common methods used by test suites.
    /// </summary>
    public static class Common
    {
        /// <summary>
        /// Get search response status which returned by the SendStringRequest operation.
        /// </summary>
        /// <param name="response">The SendStringResponse instance</param>
        /// <returns>The Status code</returns>
        public static string GetSearchStatusCode(SendStringResponse response)
        {
            string xmlResponse = response.ResponseDataXML;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlResponse);
            XmlNamespaceManager xmlNameSpaceManager = new XmlNamespaceManager(doc.NameTable);
            xmlNameSpaceManager.AddNamespace("e", "Search");
            XmlNode status = doc.SelectSingleNode("/e:Search/e:Response/e:Store/e:Status", xmlNameSpaceManager);

            if (status != null)
            {
                return status.InnerText;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Create a meeting request mime
        /// </summary>
        /// <param name="from">The from address of mail</param>
        /// <param name="to">The to address of the mail</param>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="body">The body content of the mail</param>
        /// <param name="icalendarContent">The content of iCalendar required by this meeting</param>
        /// <returns>Returns the corresponding sample meeting mime</returns>
        public static string CreateMeetingRequestMime(string from, string to, string subject, string body, string icalendarContent)
        {
            string meetingRequestMime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: multipart/alternative;
    boundary=""---Next Part---""
MIME-Version: 1.0

-----Next Part---
Content-Type: text/plain; charset=""us-ascii""

{3}

-----Next Part---
Content-Type: text/calendar; charset=""us-ascii""; method=REQUEST

{4}

";
            return FormatString(meetingRequestMime, from, to, subject, body, icalendarContent);
        }

        /// <summary>
        /// Extension of string.Format to use the string as the object to be formatted.
        /// When the string argument is null, this extension will treat it as string.Empty
        /// </summary>
        /// <param name="format">A composite format string</param>
        /// <param name="args">A string array that contains zero or more string to format.</param>
        /// <returns>A copy of format in which the format items have been replaced</returns>
        public static string FormatString(string format, params string[] args)
        {
            string[] tmpargs = new string[args.Length];

            for (int i = 0; i < args.Length; i++)
            {
                tmpargs[i] = args[i] ?? string.Empty;
            }

            return string.Format(format, tmpargs);
        }

        /// <summary>
        /// Get a specified property value from ptfconfig file.
        /// </summary>
        /// <param name="propertyName">The name of property.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>The value of the specified property.</returns>
        public static string GetConfigurationPropertyValue(string propertyName, ITestSite site)
        {
            string propertyValue = site.Properties[propertyName];
            if (propertyValue != null)
            {
                string propertyRegex = @"\[(?<property>[^\[]+?)\]";

                if (Regex.IsMatch(propertyValue, propertyRegex, RegexOptions.IgnoreCase))
                {
                    propertyValue = Regex.Replace(
                        propertyValue,
                        propertyRegex,
                        (m) =>
                        {
                            string matchedPropertyName = m.Groups["property"].Value;
                            if (site.Properties[matchedPropertyName] != null)
                            {
                                return GetConfigurationPropertyValue(matchedPropertyName, site);
                            }
                            else
                            {
                                return m.Value;
                            }
                        },
                        RegexOptions.IgnoreCase);
                }
            }
            else if (string.Compare(propertyName, "CommonConfigurationFileName", StringComparison.CurrentCultureIgnoreCase) != 0)
            {
                // 'CommonConfigurationFileName' property can be set to null when the common properties were moved from the common ptfconfig file to the local ptfconfig file.
                site.Assert.Fail("Property '{0}' was not found in the ptfconfig file. Note: When processing property values, string in square brackets ([...]) will be replaced with the property value whose name is the same string.", propertyName);
            }

            return propertyValue;
        }

        /// <summary>
        /// Create a sample plain text mime
        /// </summary>
        /// <param name="from">The from address of mail</param>
        /// <param name="to">The to address of the mail</param>
        /// <param name="cc">The cc address of the mail</param>
        /// <param name="bcc">The bcc address of the mail</param>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="body">The body content of the mail</param>
        /// <param name="sender">The sender of the mail</param>
        /// <param name="replyTo">The replyTo of the mail</param>
        /// <returns>Returns the corresponding sample plain text mime</returns>
        public static string CreatePlainTextMime(string from, string to, string cc, string bcc, string subject, string body, string sender = null, string replyTo = null)
        {
            cc = string.IsNullOrEmpty(cc) ? string.Empty : string.Format("Cc: {0}\r\n", cc);
            bcc = string.IsNullOrEmpty(bcc) ? string.Empty : string.Format("Bcc: {0}\r\n", bcc);
            sender = string.IsNullOrEmpty(sender) ? string.Empty : string.Format("Sender: {0}\r\n", sender);
            replyTo = string.IsNullOrEmpty(replyTo) ? string.Empty : string.Format("Reply-To: {0}\r\n", replyTo);

            string plainTextMime =
@"From: {0}
To: {1}
"
+ sender + cc + bcc + replyTo + @"Subject: {2}
Content-Type: text/plain; charset=""us-ascii""
MIME-Version: 1.0

{3}
";
            return FormatString(plainTextMime, from, to, subject, body);
        }

        /// <summary>
        /// Builds a SendMail request by using the specified client Id, copyToSentItems option and mail mime content.
        /// </summary>
        /// <param name="clientId">Specify the client Id</param>
        /// <param name="copyToSentItems">Specify whether needs to store a mail copy to sent items</param>
        /// <param name="mime">Specify the mail mime</param>
        /// <returns>Returns the SendMailRequest instance</returns>
        public static SendMailRequest CreateSendMailRequest(string clientId, bool copyToSentItems, string mime)
        {
            Request.SendMail sendMail = new Request.SendMail
            {
                SaveInSentItems = copyToSentItems ? string.Empty : null,
                ClientId = clientId,
                Mime = mime
            };

            SendMailRequest sendMailRequest = Common.CreateSendMailRequest();
            sendMailRequest.RequestData = sendMail;
            return sendMailRequest;
        }

        /// <summary>
        /// Merge the properties from the global ptfconfig file.
        /// </summary>
        /// <param name="globalConfigFileName">Global ptfconfig filename.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeGlobalConfig(string globalConfigFileName, ITestSite site)
        {
            if (string.IsNullOrEmpty(globalConfigFileName))
            {
                site.Log.Add(
                    LogEntryKind.Warning,
                    string.Format(
                    "The common ptfconfig file '{0}' was not loaded since the 'CommonConfigurationFileName' property or its value is not available at the local ptfconfig file.", 
                    globalConfigFileName));
            }
            else
            {
                MergeConfigurationFile(globalConfigFileName, site);
            }
        }

        /// <summary>
        /// Merge the properties from the SHOULD/MAY ptfconfig file.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeSHOULDMAYConfig(ITestSite site)
        {
            string shouldMayConfigFileName = string.Format("{0}_{1}_SHOULDMAY.deployment.ptfconfig", site.DefaultProtocolDocShortName, Common.GetSutVersion(site).ToString());

            MergeConfigurationFile(shouldMayConfigFileName, site);

            site.Log.Add(LogEntryKind.Comment, "Use {0} file for optional requirements configuration", shouldMayConfigFileName);
        }

        /// <summary>
        /// Merge the properties from the specified ptfconfig file.
        /// </summary>
        /// <param name="configFileName">Name of the ptfconfig file.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeConfigurationFile(string configFileName, ITestSite site)
        {
            if (!File.Exists(configFileName))
            {
                throw new FileNotFoundException(string.Format("The ptfconfig file '{0}' could not be found.", configFileName));
            }

            XmlNodeList properties = null;

            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(configFileName);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("tc", "http://schemas.microsoft.com/windows/ProtocolsTest/2007/07/TestConfig");

                if (doc.DocumentElement != null)
                {
                    properties = doc.DocumentElement.SelectNodes("//tc:Property", nsmgr);
                }

                if (properties == null)
                {
                    return;
                }
            }
            catch (XmlException exception)
            {
                throw new PtfConfigLoadException(string.Format("Merging the ptfconfig file '{0}' failed. It is an invalid XML file. " + exception.Message, configFileName));
            }

            foreach (XmlNode property in properties)
            {
                string propertyName;
                string propertyValue;

                if (property.Attributes == null || property.Attributes["name"] == null || string.IsNullOrEmpty(property.Attributes["name"].Value))
                {
                    throw new PtfConfigLoadException(string.Format("A property defined in the ptfconfig file '{0}' has a missing or a empty 'name' attribute.", configFileName));
                }
                else
                {
                    propertyName = property.Attributes["name"].Value;
                }

                if (property.Attributes == null || property.Attributes["value"] == null)
                {
                    throw new PtfConfigLoadException(string.Format("Property '{0}' defined in the ptfconfig file '{1}' has a missing 'value' attribute.", propertyName, configFileName));
                }
                else
                {
                    propertyValue = property.Attributes["value"].Value;
                }

                if (site.Properties[propertyName] == null)
                {
                    site.Properties.Add(propertyName, propertyValue);
                }
                else
                {
                    // Since the test suite specific ptfconfig file should take precedence over the global ptfconfig file, when the same property exists in both, the global ptfconfig property is ignored.
                    site.Log.Add(LogEntryKind.Warning, string.Format("Same property '{0}' exists in both the local ptfconfig file and the global ptfconfig file '{1}'. Test suite is using the one from the local ptfconfig file.", propertyName, configFileName));

                    continue;
                }
            }
        }

        /// <summary>
        /// Check whether the specified requirement is enabled to run or not.
        /// </summary>
        /// <param name="requirementId">The unique requirement number.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>True if the specified requirement is enabled to run, otherwise false.</returns>
        public static bool IsRequirementEnabled(int requirementId, ITestSite site)
        {
            string requirementPropertyName = string.Format("R{0}Enabled", requirementId);
            string requirementPropertyValue = GetConfigurationPropertyValue(requirementPropertyName, site);

            if (string.Compare("true", requirementPropertyValue, StringComparison.CurrentCultureIgnoreCase) != 0 && string.Compare("false", requirementPropertyValue, StringComparison.CurrentCultureIgnoreCase) != 0)
            {
                site.Assert.Fail("The property {0} value must be true or false in the SHOULD/MAY ptfconfig file.", requirementPropertyName);
            }

            return string.Compare("true", requirementPropertyValue, StringComparison.CurrentCultureIgnoreCase) == 0;
        }

        /// <summary>
        /// A method used to generate a unique name with protocol short name(without dash "-"), resource name, index and time stamp, when creating multiple resources of the same type.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <param name="resourceName">A parameter that represents the resource name which is used to compose the unique name.</param>
        /// <param name="index">A parameter that represents the index of the resources of the same type, which is used to compose the unique name.</param>
        /// <returns>A return value that represents the unique name composed of protocol short name, test case number, resource name, index and time stamp.</returns>
        public static string GenerateResourceName(ITestSite site, string resourceName, uint index)
        {
            string newPrefixOfResourceName = GeneratePrefixOfResourceName(site);
            return string.Format(@"{0}_{1}{2}_{3}", newPrefixOfResourceName, resourceName, index, FormatCurrentDateTime());
        }

        /// <summary>
        /// A method used to generate a unique name with protocol short name(without dash "-"), resource name and time stamp.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <param name="resourceName">A parameter that represents the resource name which is used to compose the unique name.</param>
        /// <returns>A return value that represents the unique name composed of protocol short name, test case number, resource name and time stamp.</returns>
        public static string GenerateResourceName(ITestSite site, string resourceName)
        {
            string newPrefixOfResourceName = GeneratePrefixOfResourceName(site);
            return string.Format(@"{0}_{1}_{2}", newPrefixOfResourceName, resourceName, FormatCurrentDateTime());
        }

        /// <summary>
        /// A method used to generate the prefix of a resource name based on the current test case name.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>A return value that represents the prefix of a resource name.</returns>
        public static string GeneratePrefixOfResourceName(ITestSite site)
        {
            string newPrefixOfResourceName = string.Empty;
            if (site != null)
            {
                site.Assume.IsNotNull(site.TestProperties, "The dictionary object 'site.TestProperties' should NOT be null! ");
                site.Assume.IsTrue(site.TestProperties.ContainsKey("CurrentTestCaseName"), "The dictionary object 'site.TestProperties' should contain the key 'CurrentTestCaseName'!");
                site.Assume.IsNotNull(site.DefaultProtocolDocShortName, "The 'site.DefaultProtocolDocShortName' should NOT be null! ");
                string currentTestCaseName = site.TestProperties["CurrentTestCaseName"].ToString();
                string currentProtocolShortName = string.Empty;
                if (site.DefaultProtocolDocShortName.IndexOf("-", StringComparison.CurrentCultureIgnoreCase) >= 0)
                {
                    foreach (string partName in site.DefaultProtocolDocShortName.Split(new char[1] { '-' }))
                    {
                        currentProtocolShortName += partName;
                    }
                }
                else
                {
                    currentProtocolShortName = site.DefaultProtocolDocShortName;
                }

                int startPos = currentTestCaseName.IndexOf(currentProtocolShortName, StringComparison.CurrentCultureIgnoreCase);
                site.Assume.IsTrue(startPos >= 0, "The '{0}' should contain '{1}'!", currentTestCaseName, currentProtocolShortName);
                if (startPos >= 0)
                {
                    currentTestCaseName = currentTestCaseName.Substring(startPos);
                }

                string currentTestScenarioNumber = currentTestCaseName.Split(new char[1] { '_' })[1];
                string currentTestCaseNumber = currentTestCaseName.Split(new char[1] { '_' })[2];
                newPrefixOfResourceName = string.Format(@"{0}_{1}_{2}", currentProtocolShortName, currentTestScenarioNumber, currentTestCaseNumber);
            }

            return newPrefixOfResourceName;
        }

        /// <summary>
        /// Try to parse the no separator time string to DateTime
        /// </summary>
        /// <param name="time">The specified DateTime string</param>
        /// <returns>Return the DateTime</returns>
        public static DateTime GetNoSeparatorDateTime(string time)
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(time.Substring(0, 4));
            stringBuilder.Append("-");
            stringBuilder.Append(time.Substring(4, 2));
            stringBuilder.Append("-");
            stringBuilder.Append(time.Substring(6, 5));
            stringBuilder.Append(":");
            stringBuilder.Append(time.Substring(11, 2));
            stringBuilder.Append(":");
            stringBuilder.Append(time.Substring(13));

            DateTime dateTime = DateTime.Parse(stringBuilder.ToString());
            return dateTime;
        }

        /// <summary>
        /// Get SUT version.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>Returns the SUT version.</returns>
        public static SutVersion GetSutVersion(ITestSite site)
        {
            string configuredSutVersion = GetConfigurationPropertyValue("SutVersion", site);
            SutVersion sutVersion = (SutVersion)Enum.Parse(typeof(SutVersion), configuredSutVersion, true);
            return sutVersion;
        }

        /// <summary>
        /// Merge common configuration and SHOULD/MAY configuration files.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeConfiguration(ITestSite site)
        {
            // Get the name of common configuration file.
            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", site);

            // Merge the common configuration
            Common.MergeGlobalConfig(commonConfigFileName, site);

            // Merge the SHOULD/MAY configuration
            Common.MergeSHOULDMAYConfig(site);
        }

        /// <summary>
        /// Creates an Autodiscover request
        /// </summary>
        /// <param name="emailAddress">An email address</param>
        /// <param name="acceptableResponseSchema">Acceptable response schema</param>
        /// <returns>An AutodiscoverRequest instance</returns>
        public static AutodiscoverRequest CreateAutodiscoverRequest(string emailAddress, string acceptableResponseSchema)
        {
            AutodiscoverRequest requestTemp = new AutodiscoverRequest();
            requestTemp.SetCommandParameters(null);
            Request.Autodiscover requestdata = new Request.Autodiscover
            {
                Request =
                    new Request.RequestType
                    {
                        AcceptableResponseSchema = acceptableResponseSchema,
                        EMailAddress = emailAddress
                    }
            };
            requestTemp.RequestData = requestdata;
            return requestTemp;
        }

        /// <summary>
        /// Creates an empty Sync request
        /// </summary>
        /// <returns>A SyncRequest instance</returns>
        public static SyncRequest CreateSyncRequest()
        {
            SyncRequest requestTemp = new SyncRequest { RequestData = new Request.Sync() };
            return requestTemp;
        }

        /// <summary>
        ///  Creates a Sync request using the specified SyncCollection array
        /// </summary>
        /// <param name="syncCollections">Specified SyncCollection array</param>
        /// <returns>A SyncRequest instance</returns>
        public static SyncRequest CreateSyncRequest(Request.SyncCollection[] syncCollections)
        {
            SyncRequest requestTemp = new SyncRequest();
            Request.Sync requestData = new Request.Sync();
            requestTemp.RequestData = requestData;
            requestTemp.RequestData.Collections = syncCollections;
            return requestTemp;
        }

        /// <summary>
        /// Create an initial Sync request.
        /// </summary>
        /// <param name="collectionId">The server id of the folder to be synchronized.</param>
        /// <returns>The SyncRequest instance.</returns>
        public static SyncRequest CreateInitialSyncRequest(string collectionId)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                CollectionId = collectionId,
                SyncKey = "0"
            };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Creates a folder Sync request using specified syncKey
        /// </summary>
        /// <param name="syncKey">Specified sync key</param>
        /// <returns>A FolderSyncRequest instance</returns>
        public static FolderSyncRequest CreateFolderSyncRequest(string syncKey)
        {
            FolderSyncRequest request = new FolderSyncRequest();
            Request.FolderSync requestData = new Request.FolderSync { SyncKey = syncKey };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        ///  Creates an empty FolderCreateRequest instance
        /// </summary>
        /// <returns>A FolderCreateRequest instance</returns>
        public static FolderCreateRequest CreateFolderCreateRequest()
        {
            FolderCreateRequest request = new FolderCreateRequest();
            Request.FolderCreate requestData = new Request.FolderCreate();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates a FolderCreateRequest instance using the specified information
        /// </summary>
        /// <param name="syncKey">The syncKey of this request</param>
        /// <param name="type">The type of created folder</param>
        /// <param name="displayName">The display name of the folder</param>
        /// <param name="parentId">The parentId of the created folder</param>
        /// <returns>A FolderCreateRequest instance</returns>
        public static FolderCreateRequest CreateFolderCreateRequest(string syncKey, byte type, string displayName, string parentId)
        {
            FolderCreateRequest request = new FolderCreateRequest();
            Request.FolderCreate requestData = new Request.FolderCreate
            {
                DisplayName = displayName,
                SyncKey = syncKey,
                ParentId = parentId,
                Type = type
            };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates an empty FolderDeleteRequest instance
        /// </summary>
        /// <returns>A FolderDeleteRequest instance</returns>
        public static FolderDeleteRequest CreateFolderDeleteRequest()
        {
            FolderDeleteRequest request = new FolderDeleteRequest();
            Request.FolderDelete requestData = new Request.FolderDelete();

            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates a FolderDeleteRequest instance using the specified information
        /// </summary>
        /// <param name="syncKey">The syncKey of this request</param>
        /// <param name="serverId">Specified the deleted folder's server id</param>
        /// <returns>A FolderDeleteRequest instance</returns>
        public static FolderDeleteRequest CreateFolderDeleteRequest(string syncKey, string serverId)
        {
            FolderDeleteRequest request = new FolderDeleteRequest();
            Request.FolderDelete requestData = new Request.FolderDelete { SyncKey = syncKey, ServerId = serverId };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates an empty FolderUpdateRequest instance
        /// </summary>
        /// <returns>A FolderUpdateRequest instance</returns>
        public static FolderUpdateRequest CreateFolderUpdateRequest()
        {
            FolderUpdateRequest request = new FolderUpdateRequest();
            Request.FolderUpdate requestData = new Request.FolderUpdate();

            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates an empty FolderUpdateRequest instance
        /// </summary>
        /// <param name="syncKey">The syncKey of this request</param>
        /// <param name="serverId">Specified the updated folder's server id</param>
        /// <param name="displayName">Specified new display name of the folder</param>
        /// <param name="parentId">Specified new parentId of the created folder</param>
        /// <returns>A FolderUpdateRequest instance</returns>
        public static FolderUpdateRequest CreateFolderUpdateRequest(string syncKey, string serverId, string displayName, string parentId)
        {
            FolderUpdateRequest request = new FolderUpdateRequest();
            Request.FolderUpdate requestData = new Request.FolderUpdate
            {
                DisplayName = displayName,
                SyncKey = syncKey,
                ParentId = parentId,
                ServerId = serverId
            };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create GetAttachmentRequest instance, because there is XML body for the GetAttachment command, then the RequestData always will be null
        /// </summary>
        /// <returns>A GetAttachmentRequest instance</returns>
        public static GetAttachmentRequest CreateGetAttachmentRequest()
        {
            GetAttachmentRequest request = new GetAttachmentRequest { RequestData = null };
            return request;
        }

        /// <summary>
        /// Creates an empty GetItemEstimateRequest instance
        /// </summary>
        /// <returns>An empty GetItemEstimateRequest instance</returns>
        public static GetItemEstimateRequest CreateGetItemEstimateRequest()
        {
            GetItemEstimateRequest request = new GetItemEstimateRequest();
            Request.GetItemEstimate requestData = new Request.GetItemEstimate();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates a GetItemEstimateRequest using the specified GetItemEstimateCollection array
        /// </summary>
        /// <param name="collections">Specified GetItemEstimateCollection array</param>
        /// <returns>A GetItemEstimateRequest instance</returns>
        public static GetItemEstimateRequest CreateGetItemEstimateRequest(Request.GetItemEstimateCollection[] collections)
        {
            GetItemEstimateRequest request = new GetItemEstimateRequest();
            Request.GetItemEstimate requestData = new Request.GetItemEstimate { Collections = collections };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates an empty ItemOperationsRequest instance
        /// </summary>
        /// <returns>An ItemOperationsRequest instance</returns>
        public static ItemOperationsRequest CreateItemOperationsRequest()
        {
            ItemOperationsRequest request = new ItemOperationsRequest();
            Request.ItemOperations requestData = new Request.ItemOperations();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Creates an ItemOperationsRequest instance using specified items
        /// </summary>
        /// <param name="items">Specified object items array</param>
        /// <returns>An ItemOperationsRequest instance</returns>
        public static ItemOperationsRequest CreateItemOperationsRequest(object[] items)
        {
            ItemOperationsRequest request = new ItemOperationsRequest();
            Request.ItemOperations requestData = new Request.ItemOperations { Items = items };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty MeetingResponseRequest instance
        /// </summary>
        /// <returns>A MeetingResponseRequest instance</returns>
        public static MeetingResponseRequest CreateMeetingResponseRequest()
        {
            MeetingResponseRequest request = new MeetingResponseRequest();
            Request.MeetingResponse requestData = new Request.MeetingResponse();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a MeetingResponseRequest instance using specified MeetingResponseRequest array
        /// </summary>
        /// <param name="requests">Specified the MeetingResponseRequest array</param>
        /// <returns>A MeetingResponseRequest instance</returns>
        public static MeetingResponseRequest CreateMeetingResponseRequest(Request.MeetingResponseRequest[] requests)
        {
            MeetingResponseRequest request = new MeetingResponseRequest();
            Request.MeetingResponse requestData = new Request.MeetingResponse { Request = requests };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty MoveItemsRequest instance
        /// </summary>
        /// <returns>A MoveItemsRequest instance</returns>
        public static MoveItemsRequest CreateMoveItemsRequest()
        {
            MoveItemsRequest request = new MoveItemsRequest();
            Request.MoveItems requestData = new Request.MoveItems();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a MoveItemsRequest instance using specified MoveItemsMove array
        /// </summary>
        /// <param name="items">Specified the MoveItemsMove array</param>
        /// <returns>A MoveItemsRequest instance</returns>
        public static MoveItemsRequest CreateMoveItemsRequest(Request.MoveItemsMove[] items)
        {
            MoveItemsRequest request = new MoveItemsRequest();
            Request.MoveItems requestData = new Request.MoveItems { Move = items };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty PingRequest instance
        /// </summary>
        /// <returns>A PingRequest instance</returns>
        public static PingRequest CreatePingRequest()
        {
            PingRequest request = new PingRequest();
            Request.Ping requestData = new Request.Ping();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a PingRequest instance using specified PingFolder array and heartbeatInterval
        /// </summary>
        /// <param name="folders">Specified PingFolder array</param>
        /// <param name="heartbeatInterval">Specified heartbeatInterval</param>
        /// <returns>A PingRequest instance</returns>
        public static PingRequest CreatePingRequest(Request.PingFolder[] folders, string heartbeatInterval)
        {
            PingRequest request = new PingRequest();
            Request.Ping requestData = new Request.Ping { Folders = folders, HeartbeatInterval = heartbeatInterval };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty ProvisionRequest instance
        /// </summary>
        /// <returns>An empty ProvisionRequest instance</returns>
        public static ProvisionRequest CreateProvisionRequest()
        {
            ProvisionRequest request = new ProvisionRequest();
            Request.Provision requestData = new Request.Provision();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a ProvisionRequest instance using specified information
        /// </summary>
        /// <param name="deviceInformation">Specified the device information</param>
        /// <param name="policies">Specified the policies</param>
        /// <param name="remoteWipe">Specified the remoteWipe</param>
        /// <returns>A ProvisionRequest instance</returns>
        public static ProvisionRequest CreateProvisionRequest(Request.DeviceInformation deviceInformation, Request.ProvisionPolicies policies, Request.ProvisionRemoteWipe remoteWipe)
        {
            ProvisionRequest request = new ProvisionRequest();
            Request.Provision requestData = new Request.Provision
            {
                DeviceInformation = deviceInformation,
                Policies = policies,
                RemoteWipe = remoteWipe
            };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty ResolveRecipientsRequest instance
        /// </summary>
        /// <returns>An empty ResolveRecipientsRequest instance</returns>
        public static ResolveRecipientsRequest CreateResolveRecipientsRequest()
        {
            ResolveRecipientsRequest request = new ResolveRecipientsRequest();
            Request.ResolveRecipients requestData = new Request.ResolveRecipients();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a ResolveRecipientsRequest instance using specified information
        /// </summary>
        /// <param name="items">Specified the resolve recipients items</param>
        /// <returns>A ResolveRecipientsRequest instance</returns>
        public static ResolveRecipientsRequest CreateResolveRecipientsRequest(object[] items)
        {
            ResolveRecipientsRequest request = new ResolveRecipientsRequest();
            Request.ResolveRecipients requestData = new Request.ResolveRecipients { Items = items };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty SearchRequest instance
        /// </summary>
        /// <returns>An empty SearchRequest instance</returns>
        public static SearchRequest CreateSearchRequest()
        {
            SearchRequest request = new SearchRequest();
            Request.Search requestData = new Request.Search();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a SearchRequest using specified information
        /// </summary>
        /// <param name="searchStores">Specified SearchStore arrays</param>
        /// <returns>A SearchRequest instance</returns>
        public static SearchRequest CreateSearchRequest(Request.SearchStore[] searchStores)
        {
            SearchRequest request = new SearchRequest();
            Request.Search requestData = new Request.Search { Items = searchStores };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a Search request using the specified keyword and folder collection ID
        /// </summary>
        /// <param name="keyword">Specify a string value for which to search.</param>
        /// <param name="folderCollectionId">Specify the folder in which to search.</param>
        /// <returns>A SearchRequest instance</returns>
        public static SearchRequest CreateSearchRequest(string keyword, string folderCollectionId)
        {
            Request.SearchStore searchStore = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Options = new Request.Options1
                {
                    Items = new object[] { string.Empty, string.Empty },

                    ItemsElementName = new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.RebuildResults,
                        Request.ItemsChoiceType6.DeepTraversal
                    }
                }
            };

            // Build up query condition by using the keyword and folder CollectionID
            Request.queryType queryItem = new Request.queryType
            {
                Items = new object[] { folderCollectionId, keyword },

                ItemsElementName = new Request.ItemsChoiceType2[] 
                {
                    Request.ItemsChoiceType2.CollectionId,
                    Request.ItemsChoiceType2.FreeText
                }
            };

            searchStore.Query = new Request.queryType
            {
                Items = new object[] { queryItem },
                ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And }
            };

            return Common.CreateSearchRequest(new Request.SearchStore[] { searchStore });
        }

        /// <summary>
        /// Create an empty SendMailRequest instance
        /// </summary>
        /// <returns>An empty SendMailRequest instance</returns>
        public static SendMailRequest CreateSendMailRequest()
        {
            SendMailRequest request = new SendMailRequest();
            Request.SendMail requestData = new Request.SendMail();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a SendMailRequest instance using specified information
        /// </summary>
        /// <param name="accountId">Specified the account Id</param>
        /// <param name="clientId">Specified the client Id</param>
        /// <param name="mime">Specified the mime</param>
        /// <returns>A SendMailRequest instance</returns>
        public static SendMailRequest CreateSendMailRequest(string accountId, string clientId, string mime)
        {
            SendMailRequest request = new SendMailRequest();
            Request.SendMail requestData = new Request.SendMail
            {
                AccountId = accountId,
                ClientId = clientId,
                Mime = mime
            };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty SettingsRequest instance
        /// </summary>
        /// <returns>An empty SettingsRequest instance</returns>
        public static SettingsRequest CreateSettingsRequest()
        {
            SettingsRequest request = new SettingsRequest();
            Request.Settings requestData = new Request.Settings();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty SmartForwardRequest instance
        /// </summary>
        /// <returns>An empty SmartForwardRequest instance</returns>
        public static SmartForwardRequest CreateSmartForwardRequest()
        {
            SmartForwardRequest request = new SmartForwardRequest();
            Request.SmartForward requestData = new Request.SmartForward();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a SmartForwardRequest instance using specified information
        /// </summary>
        /// <param name="accountId">Specified the account Id</param>
        /// <param name="clientId">Specified the client Id</param>
        /// <param name="mime">Specified the mime</param>
        /// <param name="source">Specified the SmartForwardSource</param>
        /// <returns>A SmartForwardRequest instance</returns>
        public static SmartForwardRequest CreateSmartForwardRequest(string accountId, string clientId, string mime, Request.Source source)
        {
            SmartForwardRequest request = new SmartForwardRequest();
            Request.SmartForward requestData = new Request.SmartForward
            {
                AccountId = accountId,
                ClientId = clientId,
                Mime = mime,
                Source = source
            };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty SmartReplyRequest instance
        /// </summary>
        /// <returns>An empty SmartReplyRequest instance</returns>
        public static SmartReplyRequest CreateSmartReplyRequest()
        {
            SmartReplyRequest request = new SmartReplyRequest();
            Request.SmartReply requestData = new Request.SmartReply();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a SmartReplyRequest instance using specified information
        /// </summary>
        /// <param name="accountId">Specified the account Id</param>
        /// <param name="clientId">Specified the client Id</param>
        /// <param name="mime">Specified the mime</param>
        /// <param name="source">Specified the SmartReplySource</param>
        /// <returns>A SmartReplyRequest instance</returns>
        public static SmartReplyRequest CreateSmartReplyRequest(string accountId, string clientId, string mime, Request.Source source)
        {
            SmartReplyRequest request = new SmartReplyRequest();
            Request.SmartReply requestData = new Request.SmartReply
            {
                AccountId = accountId,
                ClientId = clientId,
                Mime = mime,
                Source = source
            };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create an empty ValidateCertRequest instance
        /// </summary>
        /// <returns>An empty ValidateCertRequest instance</returns>
        public static ValidateCertRequest CreateValidateCertRequest()
        {
            ValidateCertRequest request = new ValidateCertRequest();
            Request.ValidateCert requestData = new Request.ValidateCert();
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Create a ValidateCertRequest by specified the certificates in the byte[][] format
        /// </summary>
        /// <param name="certificates">Specified the certificates</param>
        /// <returns>A ValidateCertRequest instance</returns>
        public static ValidateCertRequest CreateValidateCertRequest(byte[][] certificates)
        {
            ValidateCertRequest request = new ValidateCertRequest();
            Request.ValidateCert requestData = new Request.ValidateCert { Certificates = certificates };
            request.RequestData = requestData;
            return request;
        }

        /// <summary>
        /// Get the server ID for the given default folder type in the actual FolderSyncResponse.
        /// </summary>
        /// <param name="folderSyncResponse">A FolderSyncResponse returned by previous ActiveSync FolderSync operation.</param>
        /// <param name="folderType">Specify the default folder type.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>Returns the server ID associated with the specified default folder type.</returns>
        public static string GetDefaultFolderServerId(FolderSyncResponse folderSyncResponse, FolderType folderType, ITestSite site)
        {
            site.Assert.IsNotNull(folderSyncResponse, "The FolderSyncResponse should not be null.");
            site.Assert.IsNotNull(folderSyncResponse.ResponseData, "The FolderSyncResponse.ResponseData should not be null.");
            site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes, "The FolderSyncResponse.ResponseData.Changes should not be null.");
            site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes.Add, "The FolderSyncResponse.ResponseData.Changes.Add should not be null.");

            string serverId = null;

            foreach (FolderSyncChangesAdd addResponse in folderSyncResponse.ResponseData.Changes.Add)
            {
                int type = (int)folderType;
                if (string.Equals(addResponse.Type, type.ToString(), StringComparison.CurrentCultureIgnoreCase))
                {
                    serverId = addResponse.ServerId;
                    break;
                }
            }

            return serverId;
        }

        /// <summary>
        /// Get mailbox by user name and domain.
        /// </summary>
        /// <param name="userName">The user name.</param>
        /// <param name="domain">The domain in which the user is.</param>
        /// <returns>The mailbox of the user.</returns>
        public static string GetMailAddress(string userName, string domain)
        {
            return string.Format(@"{0}@{1}", userName, domain);
        }

        /// <summary>
        /// Get tag name and token.
        /// </summary>
        /// <param name="combinedTagAndToken">A string contain tag name and matched token following the format"identifier|CodePageName|TagName|Token"</param>
        /// <param name="matchedToken">Output the matched Token</param>
        /// <returns>Tag Name in MS-ASWBXML processing</returns>
        public static string GetTagName(string combinedTagAndToken, out byte matchedToken)
        {
            byte tokenTemp = 0;
            if (string.IsNullOrEmpty(combinedTagAndToken))
            {
                matchedToken = tokenTemp;
                return string.Empty;
            }

            string[] dataTemp = combinedTagAndToken.Split('|');

            // A string contain tag name and matched token following the format "identifier|CodePageName|TagName|Token" a total of 4 properties.
            if (dataTemp.Length < 4)
            {
                matchedToken = tokenTemp;
                return string.Empty;
            }

            // Convert the 4th information : Token in format "identifier|CodePageName|TagName|Token".
            if (!byte.TryParse(dataTemp[3], out tokenTemp))
            {
                matchedToken = tokenTemp;
                return string.Empty;
            }

            matchedToken = tokenTemp;

            // Convert the third information : TagName in format "identifier|CodePageName|TagName|Token".
            return dataTemp[2];
        }

        /// <summary>
        /// Get the CodePageName from the CombinedString
        /// </summary>
        /// <param name="combinedString">a string contain tag name and matched token following the format"identifier|CodePageName|TagName|Token"</param>
        /// <returns>CodePage in MS-ASWBXML processing</returns>
        public static string GetCodePageName(string combinedString)
        {
            if (string.IsNullOrEmpty(combinedString))
            {
                return string.Empty;
            }

            string[] dataTemp = combinedString.Split('|');

            // A string contain tag name and matched token following the format "identifier|CodePageName|TagName|Token" a total of 4 properties.
            if (dataTemp.Length < 4)
            {
                return string.Empty;
            }

            // Convert the second information : CodePageName in format "identifier|CodePageName|TagName|Token".
            return dataTemp[1];
        }

        /// <summary>
        /// Convert the ActiveSyncProtocolVersion gotten from ptfconfig file to the string could be used in ActiveSyncClient.
        /// </summary>
        /// <param name="originalVersion">The original ActiveSyncProtocolVersion gotten from ptfconfig file.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>The converted string used in ActiveSyncClient.</returns>
        public static string ConvertActiveSyncProtocolVersion(string originalVersion, ITestSite site)
        {
            string convertedVersion = null;
            if (string.Equals(originalVersion, "12.1", StringComparison.CurrentCultureIgnoreCase))
            {
                convertedVersion = "121";
            }
            else if (string.Equals(originalVersion, "14.0", StringComparison.CurrentCultureIgnoreCase))
            {
                convertedVersion = "140";
            }
            else if (string.Equals(originalVersion, "14.1", StringComparison.CurrentCultureIgnoreCase))
            {
                convertedVersion = "141";
            }
            else if (string.Equals(originalVersion, "16.0", StringComparison.CurrentCultureIgnoreCase))
            {
                convertedVersion = "160";
            }
            else if (string.Equals(originalVersion, "16.1", StringComparison.CurrentCultureIgnoreCase))
            {
                convertedVersion = "161";
            }
            else
            {
                site.Assert.Fail(originalVersion + " is not a valid value of ActiveSyncProtocolVersion property, the value should be 12.1, 14.0, 14.1, 16.0 or 16.1.");
            }

            return convertedVersion;
        }

        /// <summary>
        /// Get the TimeZone string which is base64 encoded according the specified name and bias. For detailed algorithm refer to [MS-ASDTYPE] 2.6.4 TimeZone
        /// </summary>
        /// <param name="name">The StandardName or DaylightName for TimeZone</param>
        /// <param name="bias">The offset from UTC</param>
        /// <returns>A base64 encoded TimeZone string which represents the [MS-ASTDTYPE] TimeZone Structure</returns>
        public static string GetTimeZone(string name, int bias)
        {
            List<byte> container = new List<byte>();

            byte[] biasBytes = BitConverter.GetBytes(bias);
            container.AddRange(biasBytes);

            byte[] standardNameBytes = Encoding.Unicode.GetBytes(name);
            container.AddRange(standardNameBytes);

            byte[] standardDateBytes = new byte[16];
            Array.Clear(standardDateBytes, 0x0, standardDateBytes.Length);
            container.AddRange(standardDateBytes);

            byte[] standardBiasBytes = new byte[4];
            Array.Clear(standardBiasBytes, 0x0, standardBiasBytes.Length);
            container.AddRange(standardBiasBytes);

            byte[] daylightNameBytes = Encoding.Unicode.GetBytes(name);
            container.AddRange(daylightNameBytes);

            byte[] daylightDateBytes = new byte[16];
            Array.Clear(daylightDateBytes, 0x0, daylightDateBytes.Length);
            container.AddRange(daylightDateBytes);

            byte[] daylightBiasBytes = new byte[4];
            Array.Clear(daylightBiasBytes, 0x0, daylightBiasBytes.Length);
            container.AddRange(daylightBiasBytes);

            return Convert.ToBase64String(container.ToArray());
        }

        /// <summary>
        /// Extend XmlNodeList can convert to an XmlNode array
        /// </summary>
        /// <param name="nodes">The XmlNodeList array</param>
        /// <returns>The XmlNode array</returns>
        public static XmlNode[] ToArray(this XmlNodeList nodes)
        {
            List<XmlNode> lists = new List<XmlNode>();

            for (int i = 0; i < nodes.Count; i++)
            {
                lists.Add(nodes.Item(i));
            }

            return lists.ToArray();
        }

        /// <summary>
        /// This method is used to verify whether the data element is Base64Encoding
        /// </summary>
        /// <param name="stringData">The data value</param>
        /// <returns>Whether the data element is Base64Encoding</returns>
        public static bool IsStringBase64Encoded(string stringData)
        {
            try
            {
                Convert.FromBase64String(stringData);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        /// <summary>
        /// Get status code from web exception which will be returned by IIS.
        /// </summary>
        /// <param name="webException">Web exception</param>
        /// <returns>Status code</returns>
        public static string GetErrorCodeFromException(WebException webException)
        {
            if (null == webException)
            {
                return string.Empty;
            }

            string exceptionMessage = webException.Message;
            string statusCode = string.Empty;
            if (exceptionMessage.Contains("(") && exceptionMessage.Contains(")"))
            {
                int leftParenthesis = exceptionMessage.IndexOf("(", StringComparison.CurrentCultureIgnoreCase);
                int rightParenthesis = exceptionMessage.IndexOf(")", StringComparison.CurrentCultureIgnoreCase);
                statusCode = exceptionMessage.Substring(leftParenthesis + 1, rightParenthesis - leftParenthesis - 1);
            }

            return statusCode;
        }

        /// <summary>
        /// Get the Content-Type header.
        /// </summary>
        /// <param name="headerValue">The string included charset and type in the Content-Type header.</param>
        /// <returns>The type of the content.</returns>
        public static string GetContentTypeHeader(string headerValue)
        {
            string[] headervalues = headerValue.Split(';');
            if (headervalues.Length > 0)
            {
                return headervalues[0];
            }

            return string.Empty;
        }

        /// <summary>
        /// Set a value in the target object using the specified property name
        /// </summary>
        /// <param name="targetObject">The target object</param>
        /// <param name="propertyName">The property name</param>
        /// <param name="value">The property value</param>
        public static void SetSpecifiedPropertyValueByName(object targetObject, string propertyName, object value)
        {
            if (string.IsNullOrEmpty(propertyName) || null == value || null == targetObject)
            {
                return;
            }

            PropertyInfo matchedProperty = targetObject.GetType().GetProperty(propertyName);

            if (matchedProperty != null)
            {
                if (matchedProperty.PropertyType == typeof(DateTime?))
                {
                    value = DateTime.Parse(value.ToString());
                }
                else if (matchedProperty.PropertyType == typeof(byte) || matchedProperty.PropertyType == typeof(byte?))
                {
                    value = byte.Parse(value.ToString());
                }
                else if (matchedProperty.PropertyType == typeof(bool) || matchedProperty.PropertyType == typeof(bool?))
                {
                    if (value.ToString() == "0")
                    {
                        value = false;
                    }
                    else if (value.ToString() == "1")
                    {
                        value = true;
                    }
                    else
                    {
                        value = bool.Parse(value.ToString());
                    }
                }
                else if (matchedProperty.PropertyType == typeof(uint) || matchedProperty.PropertyType == typeof(uint?))
                {
                    value = uint.Parse(value.ToString());
                }
                else if (matchedProperty.PropertyType == typeof(int) || matchedProperty.PropertyType == typeof(int?))
                {
                    value = int.Parse(value.ToString());
                }
                else if (matchedProperty.PropertyType == typeof(ushort) || matchedProperty.PropertyType == typeof(ushort?))
                {
                    value = ushort.Parse(value.ToString());
                }

                matchedProperty.SetValue(targetObject, value, null);
            }
        }

        /// <summary>
        /// Get all the valid values of one element
        /// </summary>
        /// <param name="elementName">The name of the element.</param>
        /// <param name="expectedValues">The array of all expected values.</param>
        /// <param name="actualValue">The actual value of one element.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void VerifyActualValues(string elementName, string[] expectedValues, string actualValue, ITestSite site)
        {
            if (!expectedValues.Contains(actualValue))
            {
                site.Assert.Fail("The value of '{0}' element should be one of the following values: {1}. Actually it is: {2}.", elementName, string.Join(", ", expectedValues), actualValue);
            }
        }

        /// <summary>
        /// Load SyncResponse to SyncStore.
        /// </summary>
        /// <param name="response">A SyncResponse instance.</param>
        /// <returns>The SyncStore instance.</returns>
        public static DataStructures.SyncStore LoadSyncResponse(SyncResponse response)
        {
            DataStructures.SyncStore result = new DataStructures.SyncStore();
            if (string.IsNullOrEmpty(response.ResponseData.Status))
            {
                result.StatusSpecified = false;
            }
            else
            {
                result.StatusSpecified = true;
            }

            if (result.StatusSpecified)
            {
                result.Status = Convert.ToInt32(response.ResponseData.Status);
            }

            if (response.ResponseData.Item == null)
            {
                return result;
            }

            Response.SyncCollectionsCollection collection = ((Response.SyncCollections)response.ResponseData.Item).Collection[0];
            for (int i = 0; i < collection.ItemsElementName.Length; i++)
            {
                switch (collection.ItemsElementName[i])
                {
                    case Response.ItemsChoiceType10.CollectionId:
                        result.CollectionId = collection.Items[i].ToString();
                        break;
                    case Response.ItemsChoiceType10.SyncKey:
                        result.SyncKey = collection.Items[i].ToString();
                        break;
                    case Response.ItemsChoiceType10.Status:
                        result.CollectionStatus = Convert.ToByte(collection.Items[i]);
                        break;
                    case Response.ItemsChoiceType10.Commands:
                        Response.SyncCollectionsCollectionCommands commands = collection.Items[i] as Response.SyncCollectionsCollectionCommands;
                        if (commands != null)
                        {
                            LoadAddElements(commands, result.AddElements);
                            LoadChangeElements(commands, result.ChangeElements);
                            LoadDeleteElements(commands, result.DeleteElements);
                        }

                        break;
                    case Response.ItemsChoiceType10.Responses:
                        Response.SyncCollectionsCollectionResponses responses = collection.Items[i] as Response.SyncCollectionsCollectionResponses;
                        if (responses != null)
                        {
                            if (responses.Add != null)
                            {
                                foreach (Response.SyncCollectionsCollectionResponsesAdd add in responses.Add)
                                {
                                    result.AddResponses.Add(add);
                                }
                            }

                            if (responses.Change != null)
                            {
                                foreach (Response.SyncCollectionsCollectionResponsesChange change in responses.Change)
                                {
                                    result.ChangeResponses.Add(change);
                                }
                            }
                        }

                        break;
                }
            }

            return result;
        }

        /// <summary>
        /// Load sync collection Add elements which are not sync until this sync
        /// </summary>
        /// <param name="collectionCommands">The add operation response</param>
        /// <param name="addElements">The collection to be loaded in</param>
        public static void LoadAddElements(Response.SyncCollectionsCollectionCommands collectionCommands, Collection<DataStructures.Sync> addElements)
        {
            if (collectionCommands.Add == null)
            {
                return;
            }

            foreach (Response.SyncCollectionsCollectionCommandsAdd addCommand in collectionCommands.Add)
            {
                DataStructures.Sync syncItem = new DataStructures.Sync { ServerId = addCommand.ServerId };
                if (addCommand.ApplicationData != null && addCommand.ApplicationData.ItemsElementName.Length > 0)
                {
                    syncItem.Email = DataStructures.Email.DeserializeFromAddApplicationData<DataStructures.Email>(addCommand.ApplicationData);
                    syncItem.Calendar = DataStructures.Calendar.DeserializeFromAddApplicationData<DataStructures.Calendar>(addCommand.ApplicationData);
                    syncItem.Note = DataStructures.Note.DeserializeFromAddApplicationData<DataStructures.Note>(addCommand.ApplicationData);
                    syncItem.Contact = DataStructures.Contact.DeserializeFromAddApplicationData<DataStructures.Contact>(addCommand.ApplicationData);
                    syncItem.Task = DataStructures.Task.DeserializeFromAddApplicationData<DataStructures.Task>(addCommand.ApplicationData);
                }

                addElements.Add(syncItem);
            }
        }

        /// <summary>
        /// Load Change elements in Sync response
        /// </summary>
        /// <param name="collectionCommands">The change operation response</param>
        /// <param name="changeElements">The collection to be loaded in</param>
        public static void LoadChangeElements(Response.SyncCollectionsCollectionCommands collectionCommands, Collection<DataStructures.Sync> changeElements)
        {
            if (collectionCommands.Change == null)
            {
                return;
            }

            foreach (Response.SyncCollectionsCollectionCommandsChange changeCommand in collectionCommands.Change)
            {
                DataStructures.Sync syncItem = new DataStructures.Sync { ServerId = changeCommand.ServerId };
                if (changeCommand.ApplicationData != null && changeCommand.ApplicationData.ItemsElementName.Length > 0)
                {
                    syncItem.Email = DataStructures.Email.DeserializeFromChangeApplicationData<DataStructures.Email>(changeCommand.ApplicationData);
                    syncItem.Calendar = DataStructures.Calendar.DeserializeFromChangeApplicationData<DataStructures.Calendar>(changeCommand.ApplicationData);
                    syncItem.Note = DataStructures.Note.DeserializeFromChangeApplicationData<DataStructures.Note>(changeCommand.ApplicationData);
                    syncItem.Contact = DataStructures.Contact.DeserializeFromChangeApplicationData<DataStructures.Contact>(changeCommand.ApplicationData);
                    syncItem.Task = DataStructures.Task.DeserializeFromChangeApplicationData<DataStructures.Task>(changeCommand.ApplicationData);
                }

                changeElements.Add(syncItem);
            }
        }

        /// <summary>
        /// Load Delete elements in Sync response
        /// </summary>
        /// <param name="collectionCommands">The delete operation response</param>
        /// <param name="deleteElements">The collection to be loaded in</param>
        public static void LoadDeleteElements(Response.SyncCollectionsCollectionCommands collectionCommands, Collection<string> deleteElements)
        {
            if (collectionCommands.Delete == null)
            {
                return;
            }

            foreach (Response.SyncCollectionsCollectionCommandsDelete deleteCommand in collectionCommands.Delete)
            {
                deleteElements.Add(deleteCommand.ServerId);
            }
        }

        /// <summary>
        /// Load the SearchStore from the SearchResponse
        /// </summary>
        /// <param name="search">The returned Search response.</param>
        /// <returns>A SearchStore instance</returns>
        public static DataStructures.SearchStore LoadSearchResponse(SearchResponse search)
        {
            DataStructures.SearchStore searchStore = new DataStructures.SearchStore();

            if (search.ResponseData.Status != null)
            {
                searchStore.Status = search.ResponseData.Status;

                if (!searchStore.Status.Equals("1"))
                {
                    return searchStore;
                }
            }

            searchStore.StoreStatus = search.ResponseData.Response.Store.Status;

            if (!searchStore.StoreStatus.Equals("1"))
            {
                return searchStore;
            }

            searchStore.Range = search.ResponseData.Response.Store.Range;

            if (!string.IsNullOrEmpty(search.ResponseData.Response.Store.Total))
            {
                searchStore.Total = Convert.ToInt32(search.ResponseData.Response.Store.Total, CultureInfo.InvariantCulture);
            }

            if (search.ResponseData.Response.Store.Result == null)
            {
                return searchStore;
            }

            if (search.ResponseData.Response.Store.Result.Length == 1 && search.ResponseData.Response.Store.Result[0].CollectionId == null)
            {
                return searchStore;
            }

            foreach (Response.SearchResponseStoreResult result in search.ResponseData.Response.Store.Result)
            {
                DataStructures.Search searchResult = new DataStructures.Search
                {
                    LongId = result.LongId,
                    Class = result.Class,
                    CollectionId = result.CollectionId,
                    Note = DataStructures.Note.DeserializeFromSearchProperties<DataStructures.Note>(result.Properties),
                    Calendar = DataStructures.Calendar.DeserializeFromSearchProperties<DataStructures.Calendar>(result.Properties),
                    Contact = DataStructures.Contact.DeserializeFromSearchProperties<DataStructures.Contact>(result.Properties),
                    Email = DataStructures.Email.DeserializeFromSearchProperties<DataStructures.Email>(result.Properties),
                    Task = DataStructures.Task.DeserializeFromSearchProperties<DataStructures.Task>(result.Properties)
                };

                searchStore.Results.Add(searchResult);
            }

            return searchStore;
        }

        /// <summary>
        /// Convert the ItemOperationsResponse to ItemOperationsStore
        /// </summary>
        /// <param name="itemOperations">The returned ItemOperations response.</param>
        /// <returns>An ItemOperationsStore instance</returns>
        public static DataStructures.ItemOperationsStore LoadItemOperationsResponse(ItemOperationsResponse itemOperations)
        {
            DataStructures.ItemOperationsStore store = new DataStructures.ItemOperationsStore
            {
                Status = itemOperations.ResponseData.Status
            };

            if (itemOperations.ResponseData.Response != null && itemOperations.ResponseData.Response.Fetch != null)
            {
                foreach (Response.ItemOperationsResponseFetch item in itemOperations.ResponseData.Response.Fetch)
                {
                    DataStructures.ItemOperations fetchItem = new DataStructures.ItemOperations
                    {
                        Status = item.Status,
                        Class = item.Class,
                        ServerId = item.ServerId,
                        CollectionId = item.CollectionId
                    };

                    if (item.Properties != null)
                    {
                        fetchItem.Calendar = DataStructures.Calendar.DeserializeFromFetchProperties<DataStructures.Calendar>(item.Properties);
                        fetchItem.Email = DataStructures.Email.DeserializeFromFetchProperties<DataStructures.Email>(item.Properties);
                        fetchItem.Note = DataStructures.Note.DeserializeFromFetchProperties<DataStructures.Note>(item.Properties);
                        fetchItem.Contact = DataStructures.Contact.DeserializeFromFetchProperties<DataStructures.Contact>(item.Properties);
                        fetchItem.Task = DataStructures.Task.DeserializeFromFetchProperties<DataStructures.Task>(item.Properties);
                    }

                    store.Items.Add(fetchItem);
                }
            }

            return store;
        }

        /// <summary>
        /// Record items impacted by current test case.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder which the item is placed.</param>
        /// <param name="itemSubject">The subject of the item to delete.</param>
        /// <returns>The items impacted by current test case.</returns>
        public static CreatedItems RecordCreatedItem(string collectionId, string itemSubject)
        {
            CreatedItems createdItems = new CreatedItems { CollectionId = collectionId };
            createdItems.ItemSubject.Add(itemSubject);

            return createdItems;
        }

        /// <summary>
        /// Format the current timestamp with the format "HHmmss_ffffff".
        /// </summary>
        /// <returns>The formatted current timestamp string.</returns>
        private static string FormatCurrentDateTime()
        {
            return DateTime.Now.ToString("HHmmss_ffffff");
        }
    }
}