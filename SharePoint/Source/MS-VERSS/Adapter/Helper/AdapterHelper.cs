namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System;
    using System.Collections.ObjectModel;
    using System.Text.RegularExpressions;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to assist MS-VERSSAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// Transfer ITestSite into adapter, make adapter can use ITestSite's function.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// Initialize the AdapterHelper class.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite.</param>
        public static void Initialize(ITestSite testSite)
        {
            site = testSite;
        }

        /// <summary>
        /// Generate a doc full URL.
        /// </summary>
        /// <param name="requestUrl">A Uri indicates the absolute URL for the site collection.</param>
        /// <param name="docLibName">The name of library in which a document will be created.</param>
        /// <param name="docName">The name of document that will be created.</param>
        /// <returns>The full URL of the document that will be created.</returns>
        public static Uri ConstructDocFileFullUrl(Uri requestUrl, string docLibName, string docName)
        {
            Uri docAbsoluteUrl = null;
            if (requestUrl.AbsoluteUri.EndsWith("/", StringComparison.CurrentCulture) == false)
            {
                docAbsoluteUrl = new Uri(string.Format("{0}/{1}/{2}", requestUrl.AbsoluteUri, docLibName, docName));
            }
            else
            {
                docAbsoluteUrl = new Uri(string.Format("{0}{1}/{2}", requestUrl.AbsoluteUri, docLibName, docName));
            }

            return docAbsoluteUrl;
        }

        /// <summary>
        /// Verify whether the expected versions from GetFileVersions are equal to the actual operation response.
        /// </summary>
        /// <param name="expect">The expected version from GetFileVersions method.</param>
        /// <param name="actual">The actual version from operation response.</param>
        /// <returns>The Boolean value that indicates whether the expected versions are equal to 
        /// the actual operation response.</returns>
        public static bool AreVersionsResultEqual(string expect, VersionData[] actual)
        {
            string[] expectVersions = expect.Split(new string[] { "^" }, StringSplitOptions.RemoveEmptyEntries);

            if (expectVersions.Length != actual.Length)
            {
                return false;
            }

            Collection<string> expectVersionsCollection = new Collection<string>();
            foreach (string expectVersion in expectVersions)
            {
                expectVersionsCollection.Add(expectVersion);
            }

            site.Log.Add(
                LogEntryKind.Debug,
                string.Format("The expected version number contains {0}.", expect.Replace("^", ",")));

            Collection<string> actualVersionsCollection =
                ConvertVersionDataArrayToVersionNumberCollection(actual);

            site.Log.Add(
              LogEntryKind.Debug,
              string.Format("The actual version number contains {0}.", ConvertVersionDataToString(actual)));

            return AreVersionNumberEqual(expectVersionsCollection, actualVersionsCollection);
        }

        /// <summary>
        /// Verify whether the expected versions from GetVersions operation are equal to the responses of other operations.
        /// </summary>
        /// <param name="expect">The expected versions from GetVersions operation response.</param>
        /// <param name="actual">The actual versions from other operations response. </param>
        /// <returns>The Boolean value that indicates whether the expected versions 
        /// from GetVersions operation are equal to the responses of other operations.</returns>
        public static bool AreVersionsResultEqual(VersionData[] expect, VersionData[] actual)
        {
            if (expect.Length != actual.Length)
            {
                return false;
            }

            Collection<string> expectVersions =
                ConvertVersionDataArrayToVersionNumberCollection(expect);
             site.Log.Add(
             LogEntryKind.Debug,
             string.Format("The expected version number contains {0}.", ConvertVersionDataToString(expect)));

            Collection<string> actualVersions =
                ConvertVersionDataArrayToVersionNumberCollection(actual);
            site.Log.Add(
             LogEntryKind.Debug,
             string.Format("The actual version number contains {0}.", ConvertVersionDataToString(actual)));

            return AreVersionNumberEqual(expectVersions, actualVersions);
        }

        /// <summary>
        /// Get the latest version that is less than the current version. 
        /// </summary>
        /// <param name="versions">The versions of specified file.</param>
        /// <returns>A string indicates a version that only less than current version.</returns>
        public static string GetPreviousVersion(VersionData[] versions)
        {
            #region Sort of descending order of version number
            // The version numbers collection that sort of descending order.
            Collection<string> versionsSortCollection = new Collection<string>();

            // Sort of descending order of version number
            foreach (VersionData versionData in versions)
            {
                // A Boolean value indicates whether this version number is smallest in the versionsSortCollection.
                bool isSmallest = true;
                for (int i = 0; i < versionsSortCollection.Count; i++)
                {
                    if (CompareVersionNumber(versionsSortCollection[i], versionData.version) == 1)
                    {
                        // If find a version in versionsSortCollection that less than this version,
                        // and then insert this version prior to the version that less than this version in the collection.
                        versionsSortCollection.Insert(i, versionData.version);
                        isSmallest = false;
                        break;
                    }
                }

                // If this version is the smallest one in the versionsSortCollection, add it to the end of 
                // the versionsSortCollection.
                if (isSmallest == true)
                {
                    versionsSortCollection.Add(versionData.version);
                }
            }
            #endregion

            // According to MS-OFCGLOS,the current version is the latest version of a document.
            // Then current version is the most recent version of the file.
            // So the first item in versionsSortCollection is current version and the second item is the version 
            // that less than current version.
            return versionsSortCollection[1];
        }

        /// <summary>
        /// Verify whether the specified version exists in all the versions of the file.
        /// </summary>
        /// <param name="versions">All the versions of file.</param>
        /// <param name="version">The specified version number which will be checked.</param>
        /// <returns>A Boolean value indicates whether the specified version exists in all the version of the file.</returns>
        public static bool IsVersionExist(VersionData[] versions, string version)
        {
            bool isExist = false;

            foreach (VersionData versionData in versions)
            {
                if (CompareVersionNumber(versionData.version, version) == 0)
                {
                    isExist = true;
                    break;
                }
            }

            return isExist;
        }

        /// <summary>
        /// Get the current version.
        /// </summary>
        /// <param name="versions">All the versions of file.</param>
        /// <returns>A string value indicates the current version of file.</returns>
        public static string GetCurrentVersion(VersionData[] versions)
        {
            // A string indicates the current version.
            // According to MS-OFCGLOS,the current version is the latest version of a document.
            // Then current version is the most recent version of the file.
            string currentVersion = versions[0].version;

            for (int i = 1; i < versions.Length; i++)
            {
                // Compare the version number with the value of currentVersion variable.
                if (CompareVersionNumber(currentVersion, versions[i].version) == 1)
                {
                    // If the CompareVersionNumber method returns 1, then the version number is greater than the value of
                    // currentVersion variable.
                    // The version number will assigned to currentVersion variable.
                    currentVersion = versions[i].version;
                }
            }

            return currentVersion;
        }

        /// <summary>
        /// Check whether the current version was increased.
        /// </summary>
        /// <param name="oldVersion">The old version.</param>
        /// <param name="currentVersion">The current version.</param>
        /// <returns>A Boolean value indicates whether the current version was increased.</returns>
        public static bool IsCurrentVersionIncreased(string oldVersion, string currentVersion)
        {
            decimal oldVersionNumber = decimal.Parse(oldVersion.Replace("@", string.Empty));
            decimal currentVersionNumber = decimal.Parse(currentVersion.Replace("@", string.Empty));

            if (currentVersionNumber > oldVersionNumber)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Verify whether the specified element exists in SOAP body.
        /// </summary>
        /// <param name="soapBody">A string value indicates the SOAP body.</param>
        /// <param name="elementName">A string value indicates the element name.</param>
        /// <returns>A Boolean value indicates whether the specified element exists in SOAP body.</returns>
        public static bool IsExistElementInSoapBody(string soapBody, string elementName)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(soapBody);
            XmlNodeList elements = doc.GetElementsByTagName(elementName);

            if (elements.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// This method is used to extract error string from soap fault.
        /// </summary>
        /// <param name="exception">The soap fault that returned from server.</param>
        /// <returns>The value of ErrorString element in soap fault.</returns>
        public static string ExtractErrorStringFromSoapFault(SoapException exception)
        {
            string errorString = string.Empty;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(exception.Detail.OuterXml);

            // Get the ErrorCode element from response.
            XmlNode errorStringNode = doc.DocumentElement["errorstring"];
            if (errorStringNode != null)
            {
                errorString = errorStringNode.InnerText;
            }

            return errorString;
        }

        /// <summary>
        /// A method used to get the schema from XSD file. It is designed to read schema definition from XSD file.
        /// </summary>
        /// <param name="schemaFilePath">A parameter represents a XSD file path where the schema definitions should be loaded.</param>
        /// <returns>A return value represents the schema definitions.</returns>
        public static string GetSchemaStringFromXsdFile(string schemaFilePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(schemaFilePath);
            return xmlDoc.OuterXml;
        }

        /// <summary>
        /// This method is used to verify URL format.
        /// </summary>
        /// <param name="url">The URL to be verified</param>
        /// <returns>If it's URL format, true will be returned. Otherwise, false will be returned.</returns>
        public static bool ValidateAbsoluteUrlFormat(Uri url)
        {
            // <digit> ::= any one of the ten digits 0 through 9
            string digit = @"[0-9]";

            // <letter> ::= any one of the 52 alphabetic characters A through Z in upper case and a through z in lower case
            string letter = @"[A-Za-z]";

            // <let-dig> ::= <letter> | <digit>
            string let_dig = letter + "|" + digit;

            // <let-dig-hyp> ::= <let-dig> | "-"
            string let_dig_hyp = let_dig + "|" + "[-]";

            // <ldh-str> ::= <let-dig-hyp> | <let-dig-hyp> <ldh-str>
            string ldh_str = let_dig_hyp + "+";

            // <label> ::= <letter> [ [ <ldh-str> ] <let-dig> ]
            string label = letter + "((" + ldh_str + ")?" + let_dig + ")?";

            // <SubDomain> ::= <label> | <SubDomain> "." <label>
            string subdomain = label + @"(\." + label + ")*";

            // <domain> ::= <SubDomain> | " "
            string ip = @"\d+\.\d+\.\d+\.\d+";

            // a host name that is either in a domain format or in a IP address format
            string host = subdomain + "|" + ip;

            string port = @":\d+";

            bool isSatisfyGuidFormat = false;
            Regex reg = new Regex(@"^[A-Za-z]{3,}://" + host + "(" + port + "){0,1}" + @"?/[^\s]*/?[^\s]+\.[^\s]+$", RegexOptions.Compiled);
            isSatisfyGuidFormat = reg.IsMatch(url.AbsoluteUri);
            return isSatisfyGuidFormat;
        }

        /// <summary>
        /// Compare the version numbers.
        /// </summary>
        /// <param name="versionA">A string indicates the first version number to be compared.</param>
        /// <param name="versionB">A string indicates the second version number to be compared.</param>
        /// <returns>An integer indicates the result of the comparison of the version numbers
        /// (-1: versionA is greater than versionB; 0: versionA equals to versionB; 
        /// 1: versionA is less than versionB).</returns>
        private static int CompareVersionNumber(string versionA, string versionB)
        {
            // When the version number preceded with an at sign (@), remove the at sign.
            versionA = versionA.Replace("@", string.Empty);
            versionB = versionB.Replace("@", string.Empty);

            // Get the major version and the minor version of first version number.
            int majorA = int.Parse(versionA.Substring(0, versionA.IndexOf(".", StringComparison.CurrentCulture)));
            int minorA = int.Parse(versionA.Substring(versionA.IndexOf(".", StringComparison.CurrentCulture) + 1));

            // Get the major version and the minor version of second version number.
            int majorB = int.Parse(versionB.Substring(0, versionB.IndexOf(".", StringComparison.CurrentCulture)));
            int minorB = int.Parse(versionB.Substring(versionB.IndexOf(".", StringComparison.CurrentCulture) + 1));

            // Compare the version numbers.
            // If versionA greater than versionB then return -1.
            // If versionA equals to versionB then return 0.
            // If versionA less than versionB then return 1.
            if (majorA > majorB)
            {
                return -1;
            }
            else if (majorA == majorB)
            {
                if (minorA > minorB)
                {
                    return -1;
                }
                else if (minorA == minorB)
                {
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            else
            {
                return 1;
            }
        }

        /// <summary>
        /// Verify whether the two version numbers are equal.
        /// </summary>
        /// <param name="versionsA">A string array indicates the first version number.</param>
        /// <param name="versionsB">A string array indicates the second version number.</param>
        /// <returns>A Boolean value indicates whether the two version numbers are equal.</returns>
        private static bool AreVersionNumberEqual(Collection<string> versionsA, Collection<string> versionsB)
        {
            foreach (string versionA in versionsA)
            {
                for (int i = 0; i < versionsB.Count; i++)
                {
                    if (string.Compare(versionA, versionsB[i], true, System.Globalization.CultureInfo.CurrentCulture) == 0)
                    {
                        versionsB.RemoveAt(i);
                        break;
                    }
                }
            }

            if (versionsB.Count == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Convert an array of VersionData object to a version numbers collection.
        /// </summary>
        /// <param name="versionDataArray">An array of VersionData object that will be converted.</param>
        /// <returns>A Collection object indicates a collection of version number.</returns>
        private static Collection<string> ConvertVersionDataArrayToVersionNumberCollection(VersionData[] versionDataArray)
        {
            Collection<string> versions = new Collection<string>();

            foreach (VersionData versionData in versionDataArray)
            {
                versions.Add(versionData.version);
            }

            return versions;
        }

        /// <summary>
        /// Convert an array of VersionData to string.
        /// </summary>
        /// <param name="versionDataArray">An array of VersionData object that will be converted.</param>
        /// <returns>A string indicates version information.</returns>
        private static string ConvertVersionDataToString(VersionData[] versionDataArray)
        {
            string versionsString = string.Empty;

            foreach (VersionData versionData in versionDataArray)
            {
                versionsString = versionsString + versionData.version + ",";
            }

            if (versionsString.Length > 0)
            {
                versionsString = versionsString.Substring(0, versionsString.Length - 1);
            }
            
            return versionsString;
        }
    }
}