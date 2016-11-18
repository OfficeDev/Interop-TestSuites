namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
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

    /// <summary>
    /// A class that contains the common methods used by test suites.
    /// </summary>
    public static class Common
    {
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
            else if (string.Compare(propertyName, "CommonConfigurationFileName", true) != 0)
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
            string errorFormat = @"The PTF config property '{0}''s value should represent a [{1}] type value, actual value: '{2}'";
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
                    throw new PtfConfigLoadException(string.Format("Same property '{0}' was found in both the local ptfconfig file and the SHOULD/MAY ptfconfig file '{1}'. It cannot exist in the local ptfconfig file.", propertyName, configFileName));
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
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <returns>True if the specified requirement is enabled to run, otherwise false.</returns>
        public static bool IsRequirementEnabled(int requirementId, ITestSite site)
        {
            string requirementPropertyName = string.Format("R{0}Enabled", requirementId);
            string requirementPropertyValue = GetConfigurationPropertyValue(requirementPropertyName, site);

            if (string.Compare("true", requirementPropertyValue, true) != 0 && string.Compare("false", requirementPropertyValue, true) != 0)
            {
                site.Assert.Fail("The property {0} value must be true or false in the SHOULD/MAY ptfconfig file.", requirementPropertyName);
            }

            return string.Compare("true", requirementPropertyValue, true) == 0;
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
        /// For example, if the current test case name is "MSADMINS_S01_TC20_CreateSiteSuccessfully_OwnerEmailLessThanMaxLength", 
        /// the prefix will be "MSADMINS_S01_TC20".
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
            return DateTime.Now.ToString("HHmmss_fff");
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
        /// A method used to get a field definition from response of GetList operation which is described in [MS-LISTSWS] by specified field name.
        /// </summary>
        /// <param name="getListResponse">A parameter represents the response instance of GetList operation which includes fields definition.</param>
        /// <param name="expectedFieldName">A parameter represents the field name which is expected.</param>
        /// <param name="currentTestSite">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>A return value represents the field definition which name is equal to specified field name.</returns>
        public static FieldDefinition GetFieldItemByName(GetListResponseGetListResult getListResponse, string expectedFieldName, ITestSite currentTestSite)
        {
            if (null == getListResponse)
            {
                throw new ArgumentNullException("getListResponse");
            }

            if (string.IsNullOrEmpty(expectedFieldName))
            {
                throw new ArgumentNullException("expectedFieldName");
            }

            if (null == getListResponse.List || null == getListResponse.List.Fields || null == getListResponse.List.Fields.Field
                || 0 == getListResponse.List.Fields.Field.Length)
            {
                currentTestSite.Assert.Fail("The response of GetList operation should contain the fields data of a list.");
            }

            var expectedField = from FieldDefinition fieldItem in getListResponse.List.Fields.Field
                                where fieldItem.Name.Equals(expectedFieldName, StringComparison.OrdinalIgnoreCase)
                                select fieldItem;

            int matchFieldCount = expectedField.Count();
            FieldDefinition fiedDefinition = null;
            if (0 == matchFieldCount)
            {
                currentTestSite.Assert.Fail("The response of GetList operation should contain expected field definition[{0}].", expectedFieldName);
            }
            else if (matchFieldCount > 1)
            {
                currentTestSite.Assert.Fail("The response of GetList operation contain fields which have duplicate name[{0}].", expectedFieldName);
            }
            else
            {
                fiedDefinition = expectedField.ElementAt(0);
            }

            return fiedDefinition;
        }

        /// <summary>
        /// A method used to verify a field definition which is described in [MS-LISTSWS] whether has expected field type.
        /// </summary>
        /// <param name="fieldDefinition">A parameter represents the field definition which is used to verify.</param>
        /// <param name="fieldType">A parameter represents the expected field type value which is used to check the field definition.</param>
        /// <param name="currentTestSite">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>Returns True indicating the verification of specified fieldDefinition is successful.</returns>
        public static bool VerifyFieldType(FieldDefinition fieldDefinition, string fieldType, ITestSite currentTestSite)
        {
            if (string.IsNullOrEmpty(fieldType))
            {
                throw new ArgumentNullException("fieldType");
            }

            currentTestSite.Log.Add(
                        LogEntryKind.Debug,
                        @"The fieldDefinition ""[{0}]"" current type:[{1}], expected type:[{2}].",
                        fieldDefinition.Name,
                        fieldDefinition.Type,
                        fieldType);

            return fieldType.Equals(fieldDefinition.Type, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// A method used to verify a field definition which is described in [MS-LISTSWS] whether has expected field ID.
        /// </summary>
        /// <param name="fieldDefinition">A parameter represents the field definition which is used to verify.</param>
        /// <param name="fieldId">>A parameter represents the expected field ID value which is used to check the field definition.</param>
        /// <param name="currentTestSite">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>Returns True indicating the verification of specified fieldDefinition was successful.</returns>
        public static bool VerifyFieldId(FieldDefinition fieldDefinition, string fieldId, ITestSite currentTestSite)
        {
            if (string.IsNullOrEmpty(fieldId))
            {
                throw new ArgumentNullException("fieldId");
            }

            Guid expectedValue;
            if (!Guid.TryParse(fieldId, out expectedValue))
            {
                throw new ArgumentException("The fieldId must be a valid GUID format value.", "fieldId");
            }

            currentTestSite.Log.Add(
                        LogEntryKind.Debug,
                        @"The fieldDefinition ""[{0}]"" current ID:[{1}], expected ID:[{2}].",
                        fieldDefinition.Name,
                        fieldDefinition.ID,
                        fieldId);

            Guid actualValue = Guid.Parse(fieldDefinition.ID);
            return expectedValue.Equals(actualValue);
        }

        /// <summary>
        /// Check the values of common properties of "SutVersion", "TransportType", and "SoapVersion". Throw exception if check fails.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <param name="checkSoapVersion">Set to true if check value of common property "SoapVersion", else set to false.</param>
        public static void CheckCommonProperties(ITestSite site, bool checkSoapVersion)
        {
            string exceptionInfo = string.Empty;

            // Check the value of common property "SutVersion".
            string propSutVersion = GetConfigurationPropertyValue("SutVersion", site);
            if ((string.Compare(propSutVersion, "WindowsSharePointServices3", true) != 0)
                && (string.Compare(propSutVersion, "SharePointFoundation2010", true) != 0)
                && (string.Compare(propSutVersion, "SharePointFoundation2013", true) != 0)
                && (string.Compare(propSutVersion, "SharePointServer2007", true) != 0)
                && (string.Compare(propSutVersion, "SharePointServer2010", true) != 0)
                && (string.Compare(propSutVersion, "SharePointServer2013", true) != 0)
                && (string.Compare(propSutVersion, "SharePointServer2016", true) != 0))
            {
                exceptionInfo += "\r\n" + @"Error: The current value of common property ""SutVersion"" is """ + propSutVersion + @""", the value is not valid. " + "\r\n";
                exceptionInfo += @"Please use one of following valid value for common property ""SutVersion"": " + "\r\n";
                exceptionInfo += @" ""WindowsSharePointServices3"", the SUT is Windows® SharePoint® Services 3.0 Service Pack 3 (SP3);" + "\r\n";
                exceptionInfo += @" ""SharePointFoundation2010"", the SUT is Microsoft® SharePoint® Foundation 2010 Service Pack 2 (SP2);" + "\r\n";
                exceptionInfo += @" ""SharePointFoundation2013"", the SUT is Microsoft® SharePoint® Foundation 2013 Service Pack 1 (SP1);" + "\r\n";
                exceptionInfo += @" ""SharePointServer2007"", the SUT is Microsoft® Office SharePoint® Server 2007 Service Pack 3 (SP3);" + "\r\n";
                exceptionInfo += @" ""SharePointServer2010"", the SUT is Microsoft® SharePoint® Server 2010 Service Pack 2 (SP2);" + "\r\n";
                exceptionInfo += @" ""SharePointServer2013"", the SUT is Microsoft® SharePoint® Server 2013 Service Pack 1 (SP1). " + "\r\n";
                exceptionInfo += @" ""SharePointServer2016"", the SUT is Microsoft® SharePoint® Server 2016. " + "\r\n";
            }

            // Check the value of common property "TransportType".
            string propTransportType = GetConfigurationPropertyValue("TransportType", site);
            if ((string.Compare(propTransportType, "HTTP", true) != 0)
                && (string.Compare(propTransportType, "HTTPS", true) != 0))
            {
                exceptionInfo += "\r\n" + @"Error: The current value of common property ""TransportType"" is """ + propTransportType + @""", the value is not valid. " + "\r\n";
                exceptionInfo += @"Please use one of following valid value for common property ""TransportType"": " + "\r\n";
                exceptionInfo += @" ""HTTP"", the transport is HTTP;" + "\r\n";
                exceptionInfo += @" ""HTTPS"", the transport is HTTPS." + "\r\n";
            }

            if (checkSoapVersion)
            {
                // Check the value of common property "SoapVersion".
                string propSoapVersion = GetConfigurationPropertyValue("SoapVersion", site);
                if ((string.Compare(propSoapVersion, "SOAP11", true) != 0)
                    && (string.Compare(propSoapVersion, "SOAP12", true) != 0))
                {
                    exceptionInfo += "\r\n" + @"Error: The current value of common property ""SoapVersion"" is """ + propSoapVersion + @""", the value is not valid." + " \r\n";
                    exceptionInfo += @"Please use one of following valid value for common property ""SoapVersion"": " + "\r\n";
                    exceptionInfo += @" ""SOAP11"", the messages are formatted with SOAP 1.1;" + "\r\n";
                    exceptionInfo += @" ""SOAP12"", the messages are formatted with SOAP 1.2." + "\r\n";
                }
            }

            // Throw the AssertInconclusiveException if check fails. 
            if (exceptionInfo != string.Empty)
            {
                site.Assert.Fail(exceptionInfo);
            }
        }
    }
}