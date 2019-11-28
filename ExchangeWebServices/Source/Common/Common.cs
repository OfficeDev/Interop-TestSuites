namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Security;
    using System.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Policy;
    using System.Text.RegularExpressions;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that contains the common methods used by test suites.
    /// </summary>
    public static class Common
    {
        /// <summary>
        /// If the protocol uses SOAP over HTTPS as transport, the adapter uses this function to avoid closing base connection. Local client will accept any valid server certificate after executing this function.
        /// </summary>
        public static void AcceptServerCertificate()
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
        }

        /// <summary>
        /// Verify the remote Secure Sockets Layer (SSL) certificate used for authentication.
        /// </summary>
        /// <param name="sender">An object that contains state information for this validation.</param>
        /// <param name="certificate">The certificate used to authenticate the remote party.</param>
        /// <param name="chain">The chain of certificate authorities associated with the remote certificate.</param>
        /// <param name="sslPolicyErrors">One or more errors associated with the remote certificate.</param>
        /// <returns>A Boolean value that determines whether the specified certificate is accepted for authentication.</returns>
        public static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            SslPolicyErrors errors = sslPolicyErrors;

            if ((errors & SslPolicyErrors.RemoteCertificateNameMismatch) == SslPolicyErrors.RemoteCertificateNameMismatch)
            {
                Zone zone = Zone.CreateFromUrl(((HttpWebRequest)sender).RequestUri.ToString());
                if (zone.SecurityZone == SecurityZone.Intranet || zone.SecurityZone == SecurityZone.MyComputer)
                {
                    errors -= SslPolicyErrors.RemoteCertificateNameMismatch;
                }
            }

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == SslPolicyErrors.RemoteCertificateChainErrors)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        // Self-signed certificates have the issuer in the subject field. 
                        if ((certificate.Subject == certificate.Issuer) && (status.Status == X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else if (status.Status != X509ChainStatusFlags.NoError)
                        {
                            // If there are any other errors in the certificate chain, the certificate is invalid, the method returns false.
                            return false;
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are untrusted root errors for self-signed certificates. 
                // These certificates are valid.
                errors -= SslPolicyErrors.RemoteCertificateChainErrors;
            }

            return errors == SslPolicyErrors.None;
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
        /// Merge the properties from the global ptfconfig file.
        /// </summary>
        /// <param name="globalConfigFilename">Global ptfconfig filename.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeGlobalConfig(string globalConfigFilename, ITestSite site)
        {
            if (string.IsNullOrEmpty(globalConfigFilename))
            {
                site.Log.Add(
                    LogEntryKind.Warning,
                    string.Format(
                    "The common ptfconfig file '{0}' was not loaded since the 'CommonConfigurationFileName' property or its value is not available at the local ptfconfig file.",
                    globalConfigFilename));
            }
            else
            {
                MergeConfigurationFile(globalConfigFilename, site);
            }
        }

        /// <summary>
        /// Merge the properties from the SHOULD/MAY ptfconfig file.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeSHOULDMAYConfig(ITestSite site)
        {
            string shouldMayConfigFileName = string.Format("{0}_{1}_SHOULDMAY.deployment.ptfconfig", site.DefaultProtocolDocShortName, GetConfigurationPropertyValue("SutVersion", site));

            MergeConfigurationFile(shouldMayConfigFileName, site);

            site.Log.Add(LogEntryKind.Comment, "Use {0} file for optional requirements configuration", shouldMayConfigFileName);
        }

        /// <summary>
        /// Merge the properties from the specified ptfconfig file.
        /// </summary>
        /// <param name="configFilename">ptfconfig filename.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeConfigurationFile(string configFilename, ITestSite site)
        {
            if (!File.Exists(configFilename))
            {
                throw new FileNotFoundException(string.Format("The ptfconfig file '{0}' could not be found.", configFilename));
            }

            XmlNodeList properties = null;

            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(configFilename);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("tc", "http://schemas.microsoft.com/windows/ProtocolsTest/2007/07/TestConfig");
                properties = doc.DocumentElement.SelectNodes("//tc:Property", nsmgr);
                if (properties == null)
                {
                    return;
                }
            }
            catch (XmlException exception)
            {
                throw new PtfConfigLoadException(
                    string.Format("Merging the ptfconfig file '{0}' failed. It is an invalid XML file. " + exception.Message, configFilename));
            }

            foreach (XmlNode property in properties)
            {
                string propertyName;
                string propertyValue;

                if (property.Attributes["name"] == null || string.IsNullOrEmpty(property.Attributes["name"].Value))
                {
                    throw new PtfConfigLoadException(
                        string.Format(
                            "A property defined in the ptfconfig file '{0}' has a missing or a empty 'name' attribute.",
                            configFilename));
                }
                else
                {
                    propertyName = property.Attributes["name"].Value;
                }

                if (property.Attributes["value"] == null)
                {
                    throw new PtfConfigLoadException(
                        string.Format(
                            "Property '{0}' defined in the ptfconfig file '{1}' has a missing 'value' attribute.",
                            propertyName,
                            configFilename));
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
                    site.Log.Add(LogEntryKind.Warning, string.Format("Same property '{0}' exists in both the local ptfconfig file and the global ptfconfig file '{1}'. Test suite is using the one from the local ptfconfig file.", propertyName, configFilename));

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
                site.Assume.Fail("The property {0} value must be true or false in the SHOULD/MAY ptfconfig file.", requirementPropertyName);
            }

            return string.Compare("true", requirementPropertyValue, StringComparison.CurrentCultureIgnoreCase) == 0;
        }

        /// <summary>
        /// A method used to generate a unique name with protocol short name (without dash "-"), resource name, index and time stamp when creating multiple resources of the same type.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <param name="resourceName">A parameter that represents the resource name which is used to combine the unique name</param>
        /// <param name="index">A parameter that represents the index of the resources of the same type, which is used to combine the unique name </param>
        /// <returns>A return value that represents the unique name combined with test case ID, resource name, index and time stamp </returns>
        public static string GenerateResourceName(ITestSite site, string resourceName, uint index)
        {
            string newPrefixOfResourceName = GeneratePrefixOfResourceName(site);
            return string.Format(@"{0}_{1}{2}_{3}", newPrefixOfResourceName, resourceName, index, FormatCurrentDateTime());
        }

        /// <summary>
        /// A method used to generate a unique name with protocol short name (without dash "-"), resource name and time stamp.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <param name="resourceName">A parameter that represents the resource name which is used to combine the unique name</param>
        /// <returns>A return value that represents the unique name combined with test case ID, resource name and time stamp </returns>
        public static string GenerateResourceName(ITestSite site, string resourceName)
        {
            string newPrefixOfResourceName = GeneratePrefixOfResourceName(site);
            return string.Format(@"{0}_{1}_{2}", newPrefixOfResourceName, resourceName, FormatCurrentDateTime());
        }

        /// <summary>
        /// A method used to generate the prefix of a resource name based on the current test case name.
        /// For example, if the current test case name is "MSOXCRPC_S01_TC01_TestEcDoConnectEx", 
        /// the prefix will be "MSOXCRPC_S01_TC01".
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
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
                if (site.DefaultProtocolDocShortName.IndexOf("-") >= 0)
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

                int startPos = currentTestCaseName.IndexOf(currentProtocolShortName);
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
        /// Initialize the ExchangeServiceBinding. 
        /// </summary>
        /// <param name="exchangeServiceBinding">An ExchangeServiceBinding instance.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void InitializeServiceBinding(ExchangeServiceBinding exchangeServiceBinding, ITestSite site)
        {
            TransportProtocol transportType = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", site), true);

            // Avoid closing base connection
            if (transportType == TransportProtocol.HTTPS)
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                Common.AcceptServerCertificate();
            }

            exchangeServiceBinding.SoapVersion = SoapProtocolVersion.Soap11;
            SutVersion sutVersion = GetSutVersion(site);
            exchangeServiceBinding.RequestServerVersionValue = new RequestServerVersion();
            switch (sutVersion)
            {
                case SutVersion.ExchangeServer2007:
                    exchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2007_SP1;
                    break;
                case SutVersion.ExchangeServer2010:
                    exchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2010_SP2;
                    break;
                case SutVersion.ExchangeServer2013:
                    exchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2013;
                    break;
                case SutVersion.ExchangeServer2016:
                    exchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2016;
                    break;
                case SutVersion.ExchangeServer2019:
                    exchangeServiceBinding.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2016;
                    break;
                default:
                    site.Assert.Fail(@"The SUT version ""{0}"" is invalid, the expected value is ""ExchangeServer2007"", ""ExchangeServer2010"", ""ExchangeServer2013"", ""ExchangeServer2016"" or ""ExchangeServer2019"".", sutVersion);
                    break;
            }
        }

        /// <summary>
        /// Get the items' id and change key from successful response. 
        /// </summary>
        /// <param name="response">Response message from server.</param> 
        /// <returns>Item ids contained in response.</returns> 
        public static ItemIdType[] GetItemIdsFromInfoResponse(BaseResponseMessageType response)
        {
            ArrayList arrayOfItems = new ArrayList();
            ArrayOfResponseMessagesType responseMessages = response.ResponseMessages;
            foreach (ResponseMessageType responseMessage in responseMessages.Items)
            {
                if (responseMessage.ResponseCode != ResponseCodeType.NoError ||
                    responseMessage.ResponseClass != ResponseClassType.Success)
                {
                    continue;
                }

                ItemInfoResponseMessageType itemInfoResponseMessage = responseMessage as ItemInfoResponseMessageType;
                ArrayOfRealItemsType arrayOfRealItemsType = itemInfoResponseMessage.Items;
                if (arrayOfRealItemsType.Items == null)
                {
                    continue;
                }

                foreach (ItemType tempItem in arrayOfRealItemsType.Items)
                {
                    arrayOfItems.Add(tempItem.ItemId);
                }
            }

            return (ItemIdType[])arrayOfItems.ToArray(typeof(ItemIdType));
        }

        /// <summary>
        /// Get the items from successful response.
        /// </summary>
        /// <typeparam name="T">The type of item</typeparam>
        /// <param name="response">Response message from server.</param>
        /// <returns>Items contained in response.</returns>
        public static T[] GetItemsFromInfoResponse<T>(BaseResponseMessageType response)
        where T : ItemType, new()
        {
            T itemReturnedByServer = null;
            ArrayList arrayOfItems = new ArrayList();
            ArrayOfResponseMessagesType responseMessages = response.ResponseMessages;
            foreach (ResponseMessageType responseMessage in responseMessages.Items)
            {
                if (responseMessage.ResponseCode != ResponseCodeType.NoError ||
                    responseMessage.ResponseClass != ResponseClassType.Success)
                {
                    continue;
                }

                ItemInfoResponseMessageType itemRet = responseMessage as ItemInfoResponseMessageType;
                ArrayOfRealItemsType aorit = itemRet.Items;
                if (aorit.Items == null)
                {
                    continue;
                }

                foreach (ItemType tempItem in aorit.Items)
                {
                    itemReturnedByServer = tempItem as T;
                    arrayOfItems.Add(itemReturnedByServer);
                }
            }

            return (T[])arrayOfItems.ToArray(typeof(T));
        }

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        /// <param name="headerValues">Specify the values for SOAP header.</param>
        /// <param name="exchangeServiceBinding">An ExchangeServiceBinding instance.</param>
        public static void ConfigureSOAPHeader(Dictionary<string, object> headerValues, ExchangeServiceBinding exchangeServiceBinding)
        {
            foreach (KeyValuePair<string, object> soapHeader in headerValues)
            {
                exchangeServiceBinding.GetType().GetProperty(soapHeader.Key).SetValue(
                    exchangeServiceBinding,
                    Convert.ChangeType(soapHeader.Value, exchangeServiceBinding.GetType().GetProperty(soapHeader.Key).PropertyType),
                    null);
            }
        }

        /// <summary>
        /// Check if the response message only contains Id elements.
        /// </summary>
        /// <param name="rawResponseXml">The raw xml of the response returned by SUT</param>
        /// <param name="tagName">The name of an xml tag.</param>
        /// <param name="elementName">The element name that the raw xml should contain.</param>
        /// <returns>If the response only contains Id elements, return true; otherwise, false.</returns>
        public static bool IsIdOnly(XmlElement rawResponseXml, string tagName, string elementName)
        {
            bool isIdOnly = false;
            if (rawResponseXml == null)
            {
                isIdOnly = false;
            }
            else
            {
                XmlNodeList nodes = rawResponseXml.GetElementsByTagName(tagName);
                foreach (XmlNode node in nodes)
                {
                    if (node.HasChildNodes)
                    {
                        XmlNodeList children = node.ChildNodes;
                        if (children != null && children.Count > 0)
                        {
                            foreach (XmlNode child in children)
                            {
                                if (string.Equals(child.Name, elementName))
                                {
                                    isIdOnly = true;
                                }
                                else
                                {
                                    if (string.Equals(child.Name, "t:PersonalNotes"))
                                    {
                                        isIdOnly = true;
                                    }
                                    else
                                    {
                                        isIdOnly = false;
                                        break;
                                    }                                    
                                }
                            }
                        }
                        else
                        {
                            isIdOnly = false;
                        }
                    }
                    else
                    {
                        isIdOnly = false;
                    }
                }
            }

            return isIdOnly;
        }

        /// <summary>
        /// Check the response and the count of items in ResponseMessages.
        /// </summary>
        /// <param name="response">The response of a operation.</param>
        /// <param name="expectedItemCount">The count of expected items in ResponseMessages.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void CheckOperationSuccess(BaseResponseMessageType response, int expectedItemCount, ITestSite site)
        {
            site.Assert.AreEqual<int>(
                 expectedItemCount,
                 response.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 expectedItemCount,
                 response.ResponseMessages.Items.GetLength(0));

            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        responseMessage.ResponseClass,
                        string.Format(
                            "The operation should be successful! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            responseMessage.ResponseCode));
            }
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