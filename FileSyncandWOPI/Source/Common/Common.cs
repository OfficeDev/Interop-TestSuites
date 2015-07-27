//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Security;
    using System.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Policy;
    using System.Text.RegularExpressions;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    ///  A class that contains the common methods used by test suites.
    /// </summary>
    public static class Common
    {
        /// <summary>
        /// A field represents the all registered Should/May properties.
        /// </summary>
        private static Dictionary<string, List<Tuple<int, bool>>> shouldMayProperties = new Dictionary<string, List<Tuple<int, bool>>>();

        /// <summary>
        /// If the protocol uses SOAP over HTTPS as transport, the adapter uses this function to avoid closing base connection.
        /// Local client will accept any valid server certificate after executing this function.
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
                if (zone.SecurityZone == SecurityZone.Intranet || zone.SecurityZone == SecurityZone.Internet || zone.SecurityZone == SecurityZone.MyComputer)
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
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
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
            else if (string.Compare(propertyName, "CommonConfigurationFileName", StringComparison.OrdinalIgnoreCase) != 0)
            {
                // 'CommonConfigurationFileName' property can be set to null when the common properties were moved from the common ptfconfig file to the local ptfconfig file.
                site.Assert.Fail("Property '{0}' was not found in the ptfconfig file. Note: When processing property values, string in square brackets ([...]) will be replaced with the property value whose name is the same string.", propertyName);
            }

            return propertyValue;
        }

        /// <summary>
        /// Get a specified property value from ptfconfig file and covert to value of Type T.
        /// </summary>
        /// <typeparam name="T">The Type to which to convert value.</typeparam>
        /// <param name="propertyName">The name of property.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <returns>The value of the specified property</returns>
        public static T GetConfigurationPropertyValue<T>(string propertyName, ITestSite site) where T : struct
        {
            Type expectedType = typeof(T);
            T propertyValue;
            string propertyValueString = GetConfigurationPropertyValue(propertyName, site);
            string errorFormat = @"The PTF configuration property '{0}''s value should represent a [{1}] type value, actual value: '{2}'";
            Exception exceptionInConvertProcess = null;

            try
            {
                if (expectedType.IsEnum)
                {
                    propertyValue = (T)Enum.Parse(expectedType, propertyValueString, true);
                }
                else
                {
                    propertyValue = (T)Convert.ChangeType(propertyValueString, expectedType);
                }

                return propertyValue;
            }
            catch (FormatException formatEx)
            {
                exceptionInConvertProcess = formatEx;
            }
            catch (ArgumentException argumentEx)
            {
                exceptionInConvertProcess = argumentEx;
            }
            catch (InvalidCastException invalidCastEx)
            {
                exceptionInConvertProcess = invalidCastEx;
            }
            catch (OverflowException overflowEx)
            {
                exceptionInConvertProcess = overflowEx;
            }

            throw new InvalidCastException(string.Format(errorFormat, propertyName, expectedType.Name, propertyValueString), exceptionInConvertProcess);
        }

        /// <summary>
        /// Extract error code from soap fault.
        /// </summary>
        /// <param name="exception">The soap fault containing the error code.</param>
        /// <returns>The value of ErrorCode element in soap fault.</returns>
        public static string ExtractErrorCodeFromSoapFault(SoapException exception)
        {
            string errorCode = null;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(exception.Detail.OuterXml);

            XmlNode errorCodeNode = doc.DocumentElement["errorcode"];
            if (errorCodeNode != null)
            {
                errorCode = errorCodeNode.InnerText;
            }

            return errorCode;
        }

        /// <summary>
        /// Merge the properties from the global ptfconfig file.
        /// </summary>
        /// <param name="globalConfigFileName">Global ptfconfig filename.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        public static void MergeGlobalConfig(string globalConfigFileName, ITestSite site)
        {
            if (string.IsNullOrEmpty(globalConfigFileName))
            {
                site.Log.Add(
                    LogEntryKind.Warning,
                    string.Format("The common ptfconfig file '{0}' was not loaded since the 'CommonConfigurationFileName' property or its value is not available at the local ptfconfig file.", globalConfigFileName));
            }
            else
            {
                MergeConfigurationFile(globalConfigFileName, site);
                CheckCommonProperties(site);
            }
        }

        /// <summary>
        /// Merge the properties from the SHOULD/MAY ptfconfig file.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        public static void MergeSHOULDMAYConfig(ITestSite site)
        {
            string shouldMayConfigFilename = string.Format("{0}_{1}_SHOULDMAY.deployment.ptfconfig", site.DefaultProtocolDocShortName, GetConfigurationPropertyValue("SutVersion", site));
            MergeConfigurationFile(shouldMayConfigFilename, site);
            site.Log.Add(LogEntryKind.Comment, "Use {0} file for optional requirements configuration", shouldMayConfigFilename);
        }

        /// <summary>
        /// Merge the properties from the specified ptfconfig file.
        /// </summary>
        /// <param name="configFileName">ptfconfig filename.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
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
                properties = doc.DocumentElement.SelectNodes("//tc:Property", nsmgr);
                if (properties == null)
                {
                    return;
                }
            }
            catch (XmlException exception)
            {
                throw new PtfConfigLoadException(
                    string.Format("Merging the ptfconfig file '{0}' failed. It is an invalid XML file. " + exception.Message, configFileName),
                    exception);
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
                            configFileName));
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
                            configFileName));
                }
                else
                {
                    propertyValue = property.Attributes["value"].Value;
                }

                if (site.Properties[propertyName] == null)
                {
                    site.Properties.Add(propertyName, propertyValue);
                }
                else if (configFileName.Contains("SHOULDMAY"))
                {
                    // Same property should not exist in both the test suite specific ptfconfig file and the SHOULD/MAY ptfconfig file.
                    throw new PtfConfigLoadException(
                        string.Format(
                            "Same property '{0}' was found in both the local ptfconfig file and the SHOULD/MAY ptfconfig file '{1}'. " +
                            "It cannot exist in the local ptfconfig file.",
                            propertyName,
                            configFileName));
                }
                else
                {
                    // Since the test suite specific ptfconfig file should take precedence over the global ptfconfig file, 
                    // when the same property exists in both, the global ptfconfig property is ignored.
                    site.Log.Add(
                        LogEntryKind.Warning,
                        string.Format(
                            "Same property '{0}' exists in both the local ptfconfig file and the global ptfconfig file '{1}'. " +
                            "Test suite is using the one from the local ptfconfig file.",
                            propertyName,
                            configFileName));

                    continue;
                }
            }
        }

        /// <summary>
        /// Check whether the specified requirement is enabled to run or not.
        /// </summary>
        /// <param name="requirementId">The unique requirement number.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <returns>True if the specified requirement is enabled to run, otherwise false.</returns>
        public static bool IsRequirementEnabled(int requirementId, ITestSite site)
        {
            return IsRequirementEnabled(site.DefaultProtocolDocShortName, requirementId, site);
        }

        /// <summary>
        /// Check whether the specified requirement is enabled to run or not.
        /// </summary>
        /// <param name="shortName">The protocol short name.</param>
        /// <param name="requirementId">The unique requirement number.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <returns>True if the specified requirement is enabled to run, otherwise false.</returns>
        public static bool IsRequirementEnabled(string shortName, int requirementId, ITestSite site)
        {
            string requirementPropertyName = string.Format("R{0}Enabled_{1}", requirementId, shortName);
            string requirementPropertyValue = GetConfigurationPropertyValue(requirementPropertyName, site);

            if (string.Compare("true", requirementPropertyValue, StringComparison.OrdinalIgnoreCase) != 0 && string.Compare("false", requirementPropertyValue, StringComparison.OrdinalIgnoreCase) != 0)
            {
                site.Assert.Fail("The property {0} value must be true or false in the SHOULD/MAY ptfconfig file.", requirementPropertyName);
            }

            if (site.TestProperties["CurrentTestCaseName"] != null)
            {
                string fullCurrentTestCaseName = site.TestProperties["CurrentTestCaseName"].ToString();
                string testCaseName = fullCurrentTestCaseName.Split(new char[1] { '.' }).Last();
                RegisterShouldMayProperty(testCaseName, requirementId, Convert.ToBoolean(requirementPropertyValue));
            }

            return string.Compare("true", requirementPropertyValue, StringComparison.OrdinalIgnoreCase) == 0;
        }

        /// <summary>
        /// A method used to generate a unique name with protocol short name(without dash "-"), resource name, index and time stamp, when create multiple resources of the same type.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <param name="resourceName">A parameter that represents the resource name which is used to combine the unique name</param>
        /// <param name="index">A parameter that represents the index of the resources of the same type, which is used to combine the unique name </param>
        /// <returns>A return value that represents the unique name combined with protocol short name, resource name, index and time stamp </returns>
        public static string GenerateResourceName(ITestSite site, string resourceName, uint index)
        {
            string newPrefixOfResourceName = GeneratePrefixOfResourceName(site);
            return string.Format(@"{0}_{1}{2}_{3}", newPrefixOfResourceName, resourceName, index, FormatCurrentDateTime());
        }

        /// <summary>
        /// A method used to generate a unique name with protocol short name(without dash "-"), resource name and time stamp.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <param name="resourceName">A parameter that represents the resource name which is used to combine the unique name</param>
        /// <returns>A return value that represents the unique name combined with protocol short name, resource name and time stamp </returns>
        public static string GenerateResourceName(ITestSite site, string resourceName)
        {
            string newPrefixOfResourceName = GeneratePrefixOfResourceName(site);
            return string.Format(@"{0}_{1}_{2}", newPrefixOfResourceName, resourceName, FormatCurrentDateTime());
        }

        /// <summary>
        /// A method used to generate the prefix of a resource name based on the current test case name.
        /// For example, if the current test case name is "TestCase_S01_TC22_UploadContents_Success_LockID", 
        /// the prefix will be "TestCase_S01_TC22".
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>A return value that represents the prefix of a resource name.</returns>
        public static string GeneratePrefixOfResourceName(ITestSite site)
        {
            string newPrefixOfResourceName = string.Empty;
            if (site != null)
            {
                site.Assume.IsNotNull(site.TestProperties, "The dictionary object 'site.TestProperities' should NOT be null! ");
                site.Assume.IsTrue(site.TestProperties.ContainsKey("CurrentTestCaseName"), "The dictionary object 'site.TestProperities' should contain the key 'CurrentTestCaseName'!");
                site.Assume.IsNotNull(site.DefaultProtocolDocShortName, "The 'site.DefaultProtocolDocShortName' should NOT be null! ");
                string fullCurrentTestCaseName = site.TestProperties["CurrentTestCaseName"].ToString();
                string testCaseName = fullCurrentTestCaseName.Split(new char[1] { '.' }).Last();

                string regex = @"([a-zA-Z_]+)_(S\d+)_(TC\d+)_";
                Match m = Regex.Match(testCaseName, regex);
                if (m.Success)
                {
                    string shortName = m.Groups[1].Value;
                    string currentTestScenarioNumber = m.Groups[2].Value;
                    string currentTestCaseNumber = m.Groups[3].Value;
                    newPrefixOfResourceName = string.Format(@"{0}_{1}_{2}", shortName, currentTestScenarioNumber, currentTestCaseNumber);
                }
                else
                {
                    site.Assert.Fail(string.Format("The test case name does not match the pattern {0}", "protocolShortName_S[scenarioNumber]_TC[TestCaseNumber]"));
                }
            }

            return newPrefixOfResourceName;
        }

        /// <summary>
        /// Merge common configuration and SHOULD/MAY configuration files.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeConfiguration(ITestSite site)
        {
            // Get the name of common configuration file.
            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", site);

            // Merge the common configuration.
            Common.MergeGlobalConfig(commonConfigFileName, site);

            // Merge the SHOULD/MAY configuration.
            Common.MergeSHOULDMAYConfig(site);
        }

        /// <summary>
        /// Format the current timestamp to this format "HHmmss_fff".
        /// </summary>
        /// <returns>The formatted current timestamp string.</returns>
        public static string FormatCurrentDateTime()
        {
            return DateTime.Now.ToString("HHmmss_ffffff");
        }

        /// <summary>
        /// Generate invalid password by replacing the first character of valid password if it is not empty, otherwise an "a" will be returned.
        /// </summary>
        /// <param name="validPassword">The valid password.</param>
        /// <returns>The invalid password.</returns>
        public static string GenerateInvalidPassword(string validPassword)
        {
            if (string.IsNullOrEmpty(validPassword))
            {
                return "a";
            }
            else if (validPassword[0].Equals('a'))
            {
                return string.Format("{0}{1}", "b", validPassword.Substring(1, validPassword.Length - 1));
            }
            else
            {
                return string.Format("{0}{1}", "a", validPassword.Substring(1, validPassword.Length - 1));
            }
        }

        /// <summary>
        /// Get attribute value by specified attribute name and index from a z:row collection.
        /// </summary>
        /// <param name="zrowDatas">A parameter represents the z:row collection where the method will target.</param>
        /// <param name="index">A parameter represents the index of a z:row item in a z:row collection.It starts on "Zero".</param>
        /// <param name="expectedAttributeName">A parameter represents the attributeName of which value will be returned.</param>
        /// <returns>A return value represents the attribute value.</returns>
        public static string GetZrowAttributeValue(XmlNode[] zrowDatas, int index, string expectedAttributeName)
        {
            if (string.IsNullOrEmpty(expectedAttributeName))
            {
                throw new ArgumentException("String type parameter should not be empty or null.", "expectedAttributeName");
            }

            if (null == zrowDatas)
            {
                throw new ArgumentNullException("zrowDatas");
            }

            if (index > zrowDatas.Count() - 1)
            {
                string errMsg = string.Format(
                    "The index[{0}] should be less than zrowDatas array length [{1}].",
                    index,
                    zrowDatas.Count() - 1);
                throw new ArgumentException(errMsg, "index");
            }

            XmlNode expectedItem = zrowDatas[index];
            var expectedValues = from XmlAttribute attributeItem in expectedItem.Attributes
                                 where attributeItem.Name.Equals(expectedAttributeName, StringComparison.OrdinalIgnoreCase)
                                 select attributeItem.Value;

            if (expectedValues.Count() > 0)
            {
                return expectedValues.ElementAt(0);
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        ///  A method used to process unsupported Product behavior, it perform a directly capture for a requirement which describe an implementation does support  products.
        /// </summary>
        /// <param name="testDocumentShortName">The value of the test document short name.</param>
        /// <param name="requirementId">The value of the test requirement id of which requirement describe an implementation does support  products. </param>
        /// <param name="requirementDescription">The value of the test requirement description.</param>
        /// <param name="currentTestSite">An instance of the TestSite.</param>
        public static void ProcessingUnsupportProducts(string testDocumentShortName, int requirementId, string requirementDescription, ITestSite currentTestSite)
        {
            if (null == currentTestSite)
            {
                throw new ArgumentNullException("currentTestSite");
            }

            if (Common.IsRequirementEnabled(testDocumentShortName, requirementId, currentTestSite))
            {
                string documentShortName = string.Empty;

                // If the testDocumentShortName parameter do not have value, try to get from the DefaultProtocolDocShortName property of ITestSite instance.
                if (string.IsNullOrEmpty(testDocumentShortName))
                {
                    documentShortName = currentTestSite.DefaultProtocolDocShortName;
                }
                else
                {
                    documentShortName = testDocumentShortName;
                }

                if (string.IsNullOrEmpty(documentShortName))
                {
                    currentTestSite.CaptureRequirement(requirementId, requirementDescription);
                }
                else
                {
                    currentTestSite.CaptureRequirement(documentShortName, requirementId, requirementDescription);
                }
            }
            else
            {
                SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", currentTestSite);
                currentTestSite.Assert.Inconclusive("This test suite does not support current SUT version[{0}].", currentSutVersion);
            }
        }

        /// <summary>
        /// Merge the properties from the SHOULD/MAY ptfconfig file according to the supported products.
        /// </summary>
        /// <param name="supportedProducts">A parameter represents the all supported products.</param>
        /// <param name="supportedProductsRelatedRequirementId">A parameter represents the id of the requirement which specifies the supported products.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeSHOULDMAYConfig(string[] supportedProducts, int supportedProductsRelatedRequirementId, ITestSite site)
        {
            MergeSHOULDMAYConfig(supportedProducts, supportedProductsRelatedRequirementId, site, null);
        }

        /// <summary>
        /// Merge the properties from the SHOULD/MAY ptfconfig file according to the supported products and specified protocol short name.
        /// </summary>
        /// <param name="supportedProducts">A parameter represents the all supported products.</param>
        /// <param name="supportedProductsRelatedRequirementId">A parameter represents the id of the requirement which specifies the supported products.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <param name="shortName">A parameter represents the current protocol short name. If this parameter is string.empty or null, this method will not use the parameter.</param>
        public static void MergeSHOULDMAYConfig(string[] supportedProducts, int supportedProductsRelatedRequirementId, ITestSite site, string shortName)
        {
            if (null == supportedProducts)
            {
                throw new ArgumentNullException("supportedProducts");
            }

            if (supportedProductsRelatedRequirementId <= 0)
            {
                throw new ArgumentException("The value should be greater than Zero.", "supportedProductsRelatedRequirementId");
            }

            SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", site);
            string currentSutVersionValue = currentSutVersion.ToString();
            bool isCurrentSutSupported = supportedProducts.Any(supportedProduct => supportedProduct.Equals(currentSutVersionValue, StringComparison.OrdinalIgnoreCase));
            string requirementPropertyName = string.Empty;
            string actualShortNameValue = string.Empty;
            if (string.IsNullOrEmpty(shortName))
            {
                requirementPropertyName = string.Format("R{0}Enabled", supportedProductsRelatedRequirementId);
                actualShortNameValue = site.DefaultProtocolDocShortName;
            }
            else
            {
                requirementPropertyName = string.Format("R{0}Enabled_{1}", supportedProductsRelatedRequirementId, shortName);
                actualShortNameValue = shortName;
            }

            if (isCurrentSutSupported)
            {
                string shouldMayConfigFilename = string.Format("{0}_{1}_SHOULDMAY.deployment.ptfconfig", actualShortNameValue, currentSutVersionValue);
                Common.MergeConfigurationFile(shouldMayConfigFilename, site);
                site.Log.Add(LogEntryKind.Debug, "Use {0} file for optional requirements configuration", shouldMayConfigFilename);
                bool isExpectedPropertyExist = site.Properties.AllKeys.Any(property => property.Equals(requirementPropertyName, StringComparison.OrdinalIgnoreCase));
                if (!isExpectedPropertyExist)
                {
                    site.Assert.Fail("There should be a property [{0}].", requirementPropertyName);
                }
            }
            else
            {
                // If the current SUT version does not support the protocol, this method add a R***enable property into the properties collection.
                bool isExpectedPropertyExist = site.Properties.AllKeys.Any(property => property.Equals(requirementPropertyName, StringComparison.OrdinalIgnoreCase));
                if (isExpectedPropertyExist)
                {
                    site.Properties[requirementPropertyName] = bool.FalseString;
                }
                else
                {
                    site.Properties.Add(requirementPropertyName, bool.FalseString);
                }
            }
        }

        /// <summary>
        /// Check the values of common properties of "SutVersion", "TransportType", Throw exception if check fails.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void CheckCommonProperties(ITestSite site)
        {
            string exceptionInfo = string.Empty;

            // Check the value of common property "SutVersion".
            string propSutVersion = GetConfigurationPropertyValue("SutVersion", site);
            if ((string.Compare(propSutVersion, "SharePointFoundation2010", StringComparison.OrdinalIgnoreCase) != 0)
                && (string.Compare(propSutVersion, "SharePointFoundation2013", StringComparison.OrdinalIgnoreCase) != 0)
                && (string.Compare(propSutVersion, "SharePointServer2010", StringComparison.OrdinalIgnoreCase) != 0)
                && (string.Compare(propSutVersion, "SharePointServer2013", StringComparison.OrdinalIgnoreCase) != 0))
            {
                exceptionInfo += "\r\n" + @"Error: The current value of common property ""SutVersion"" is """ + propSutVersion + @""", the value is not valid. " + "\r\n";
                exceptionInfo += @"Please use one of following valid value for common property ""SutVersion"": " + "\r\n";
                exceptionInfo += @" ""SharePointFoundation2010"", the SUT is Microsoft® SharePoint® Foundation 2010 Service Pack 2 (SP2);" + "\r\n";
                exceptionInfo += @" ""SharePointFoundation2013"", the SUT is Microsoft® SharePoint® Foundation 2013 Service Pack 1 (SP1);" + "\r\n";
                exceptionInfo += @" ""SharePointServer2010"", the SUT is Microsoft® SharePoint® Server 2010 Service Pack 2 (SP2);" + "\r\n";
                exceptionInfo += @" ""SharePointServer2013"", the SUT is Microsoft® SharePoint® Server 2013 Service Pack 1 (SP1). " + "\r\n";
            }

            // Check the value of common property "TransportType".
            string propTransportType = GetConfigurationPropertyValue("TransportType", site);
            if ((string.Compare(propTransportType, "HTTP", StringComparison.OrdinalIgnoreCase) != 0)
                && (string.Compare(propTransportType, "HTTPS", StringComparison.OrdinalIgnoreCase) != 0))
            {
                exceptionInfo += "\r\n" + @"Error: The current value of common property ""TransportType"" is """ + propTransportType + @""", the value is not valid. " + "\r\n";
                exceptionInfo += @"Please use one of following valid value for common property ""TransportType"": " + "\r\n";
                exceptionInfo += @" ""HTTP"", the transport is HTTP;" + "\r\n";
                exceptionInfo += @" ""HTTPS"", the transport is HTTPS." + "\r\n";
            }

            // Throw the AssertInconclusiveException if check fails. 
            if (!string.IsNullOrEmpty(exceptionInfo))
            {
                site.Assert.Fail(exceptionInfo);
            }
        }

        /// <summary>
        /// Throw inconclusive when the current test case passes but all the Should/May configure properties within it are false. 
        /// </summary>
        /// <param name="context">A parameter represents the test context.</param>
        /// <param name="site">A parameter represents the test site.</param>
        public static void AssumeInconclusiveIfAllShouldMayPropertiesFalse(TestContext context, ITestSite site)
        {
            if (context.CurrentTestOutcome != UnitTestOutcome.Passed)
            {
                return;
            }

            string fullCurrentTestCaseName = site.TestProperties["CurrentTestCaseName"].ToString();
            string testCaseName = fullCurrentTestCaseName.Split(new char[1] { '.' }).Last();

            if (shouldMayProperties.ContainsKey(testCaseName))
            {
                List<Tuple<int, bool>> properties = shouldMayProperties[testCaseName];
                if (properties != null)
                {
                    if (properties.All(tuple => tuple.Item2 == false))
                    {
                        string description = string.Join(",", properties.Select(tuple => string.Format("R{0}", tuple.Item1)).ToArray());

                        site.Assume.Inconclusive(string.Format("At least one of the requirements {0} should be enabled", description));
                    }

                    shouldMayProperties.Remove(testCaseName);
                }
            }
        }

        /// <summary>
        /// Register the Should/May configure property and its value to facilitate the determination of inconclusive the current test case when 
        /// it passes and all the Should/May configure properties within it are false. 
        /// </summary>
        /// <param name="testCaseName">A parameter represents the current test case name.</param>
        /// <param name="requirementId">A parameter represents Should/May configure requirement id.</param>
        /// <param name="value">A parameter represents the value of Should/May configure requirement id.</param>
        private static void RegisterShouldMayProperty(string testCaseName, int requirementId, bool value)
        {
            if (shouldMayProperties.ContainsKey(testCaseName))
            {
                List<Tuple<int, bool>> properties = shouldMayProperties[testCaseName];
                Tuple<int, bool> property = properties.FirstOrDefault(tuple => tuple.Item1 == requirementId);
                if (property != null)
                {
                    if (property.Item2 != value)
                    {
                        properties.Remove(property);
                        properties.Add(new Tuple<int, bool>(requirementId, value));
                    }
                }
                else
                {
                    properties.Add(new Tuple<int, bool>(requirementId, value));
                }
            }
            else
            {
                List<Tuple<int, bool>> properties = new List<Tuple<int, bool>>();
                properties.Add(new Tuple<int, bool>(requirementId, value));
                shouldMayProperties.Add(testCaseName, properties);
            }
        }
    }
}