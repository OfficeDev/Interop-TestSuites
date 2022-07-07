namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Security;
    using System.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Policy;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that contains the common methods used by test suites.
    /// </summary>
    public class Common
    {
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
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
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
        /// <param name="configFilename">ptfconfig filename.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
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
                throw new InvalidOperationException(
                    string.Format("Merging the ptfconfig file '{0}' failed. It is an invalid XML file." + exception.Message, configFilename),
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
                            "A property defined in the ptfconfig file '{0}' has a missing or an empty 'name' attribute.",
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
                else if (configFilename.Contains("SHOULDMAY"))
                {
                    // Same property should not exist in both the test suite specific ptfconfig file and the SHOULD/MAY ptfconfig file.
                    throw new PtfConfigLoadException(
                        string.Format(
                            "Same property '{0}' was found in both the local ptfconfig file and the SHOULD/MAY ptfconfig file '{1}'. " +
                            "It cannot exist in the local ptfconfig file.",
                            propertyName,
                            configFilename));
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
                            configFilename));

                    continue;
                }
            }
        }

        /// <summary>
        /// Merge common configuration and SHOULD/MAY configuration files.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        public static void MergeConfiguration(ITestSite site)
        {
            // Merge the common configuration into local configuration
            string commonConfigFileName = GetConfigurationPropertyValue("CommonConfigurationFileName", site);

            // Merge the common configuration
            MergeGlobalConfig(commonConfigFileName, site);

            // Merge the SHOULD/MAY configuration
            MergeSHOULDMAYConfig(site);
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
                site.Assume.Fail("The property {0} value must be true or false in the SHOULD/MAY ptfconfig file.", requirementPropertyName);
            }

            return string.Compare("true", requirementPropertyValue, true) == 0;
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
        /// Compresses stream using LZ77 algorithm if CompressRpcRequest is set to true in configuration, and obfuscates stream if XorRpcRequest is set to true in configuration. If none of the properties is true, the original stream is returned.
        /// </summary>
        /// <param name="inputStream">The stream to be compressed and obfuscated. The stream should contain header and payload.</param>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
        /// <returns>Returns the stream that is compressed and obfuscated.</returns>
        public static byte[] CompressAndObfuscateRequest(byte[] inputStream, ITestSite site)
        {
            bool xorRpcRequest = bool.Parse(GetConfigurationPropertyValue("XorRpcRequest", site));
            bool compressRpcRequest = bool.Parse(GetConfigurationPropertyValue("CompressRpcRequest", site));
            byte[] processedStream = inputStream;
            if (compressRpcRequest)
            {
                processedStream = CompressStream(processedStream);
            }

            if (xorRpcRequest)
            {
                processedStream = XOR(processedStream);
            }

            return processedStream;
        }

        /// <summary>
        /// Compresses payload in the stream with LZ77 algorithm and updates the header.
        /// </summary>
        /// <param name="inputStream">The stream to be compressed. The stream should contain header and payload.</param>
        /// <returns>Returns the compressed stream. The original stream will be returned if Compressed flag is set in header or compressed stream is larger.</returns>
        public static byte[] CompressStream(byte[] inputStream)
        {
            #region Consts
            // Size of Version field in RPC_HEADER_EXT
            const int RpcHeaderExtVersionByteSize = 2;

            // Size of Flags field in RPC_HEADER_EXT
            const int RpcHeaderExtFlagsByteSize = 2;

            // Size of Size field in RPC_HEADER_EXT
            const int RpcHeaderExtSizeByteSize = 2;

            // The length of RPC_HEADER_EXT.
            const int RPCHeaderExtLength = 8;
            #endregion

            if (inputStream == null)
            {
                throw new ArgumentNullException("inputStream");
            }

            if (inputStream.Length < RPCHeaderExtLength)
            {
                throw new ArgumentException("Input must contain valid RPC_HEADER_EXT.", "inputStream");
            }

            // The third and fourth byte represents Flags field in RPC_HEADER_EXT
            RpcHeaderExtFlags flag = (RpcHeaderExtFlags)BitConverter.ToInt16(inputStream, RpcHeaderExtVersionByteSize);
            if ((flag & RpcHeaderExtFlags.Compressed) == RpcHeaderExtFlags.Compressed)
            {
                // Data stream is already compressed.
                return inputStream;
            }

            ushort originalPayloadSize = BitConverter.ToUInt16(inputStream, RpcHeaderExtVersionByteSize + RpcHeaderExtFlagsByteSize);

            // Compress payload
            byte[] originalPayload = new byte[originalPayloadSize];
            Array.Copy(inputStream, RPCHeaderExtLength, originalPayload, 0, originalPayloadSize);
            byte[] compressedPayload = LZ77Compress(originalPayload);

            // Output the original stream if compressed payload is bigger.
            if (originalPayloadSize <= compressedPayload.Length)
            {
                return inputStream;
            }

            byte[] outStream = new byte[RPCHeaderExtLength + compressedPayload.Length];
            Array.Copy(inputStream, 0, outStream, 0, RPCHeaderExtLength);

            // Update Compressed flag.
            flag |= RpcHeaderExtFlags.Compressed;
            Array.Copy(BitConverter.GetBytes((short)flag), 0, outStream, RpcHeaderExtVersionByteSize, RpcHeaderExtFlagsByteSize);

            // Update Size field.
            Array.Copy(BitConverter.GetBytes((ushort)compressedPayload.Length), 0, outStream, RpcHeaderExtVersionByteSize + RpcHeaderExtFlagsByteSize, RpcHeaderExtSizeByteSize);

            // Fill compressed payload.
            Array.Copy(compressedPayload, 0, outStream, RPCHeaderExtLength, compressedPayload.Length);

            return outStream;
        }

        /// <summary>
        /// Decompresses payload in the stream with LZ77 algorithm and updates the header.
        /// </summary>
        /// <param name="inputStream">The stream to be decompressed. The stream should contain header and payload.</param>
        /// <returns>Returns the compressed stream. The original stream will be returned if Compressed flag is not set in header.</returns>
        public static byte[] DecompressStream(byte[] inputStream)
        {
            #region Consts

            // Size of Version field in RPC_HEADER_EXT
            const int RpcHeaderExtVersionByteSize = 2;

            // Size of Flags field in RPC_HEADER_EXT
            const int RpcHeaderExtFlagsByteSize = 2;

            // Size of Size field in RPC_HEADER_EXT
            const int RpcHeaderExtSizeByteSize = 2;

            // Size of SizeActual field in RPC_HEADER_EXT
            const int RpcHeaderExtSizeActualByteSize = 2;

            // The length of RPC_HEADER_EXT.
            const int RPCHeaderExtLength = 8;

            #endregion

            if (inputStream == null)
            {
                throw new ArgumentNullException("inputStream");
            }

            if (inputStream.Length < RPCHeaderExtLength)
            {
                throw new ArgumentException("Input must contain valid RPC_HEADER_EXT.", "inputStream");
            }

            // The third and fourth byte represents Flags field in RPC_HEADER_EXT
            RpcHeaderExtFlags flag = (RpcHeaderExtFlags)BitConverter.ToInt16(inputStream, RpcHeaderExtVersionByteSize);
            if ((flag & RpcHeaderExtFlags.Compressed) == RpcHeaderExtFlags.None)
            {
                // Data is not compressed, no need to decompress
                return inputStream;
            }

            ushort compressedSize = BitConverter.ToUInt16(inputStream, RpcHeaderExtVersionByteSize + RpcHeaderExtFlagsByteSize);
            ushort actualSize = BitConverter.ToUInt16(inputStream, RpcHeaderExtVersionByteSize + RpcHeaderExtFlagsByteSize + RpcHeaderExtSizeByteSize);

            // Decompress payload
            byte[] originalPayload = new byte[compressedSize];
            Array.Copy(inputStream, RPCHeaderExtLength, originalPayload, 0, compressedSize);
            byte[] decompressedPayload = LZ77Decompress(originalPayload, actualSize);
            byte[] outStream = new byte[RPCHeaderExtLength + actualSize];
            Array.Copy(inputStream, 0, outStream, 0, RPCHeaderExtLength);
            Array.Copy(decompressedPayload, 0, outStream, RPCHeaderExtLength, actualSize);

            // Change Flag field to uncompressed
            flag &= ~RpcHeaderExtFlags.Compressed;
            Array.Copy(BitConverter.GetBytes((short)flag), 0, outStream, RpcHeaderExtVersionByteSize, RpcHeaderExtFlagsByteSize);

            // Set Size value to ActualSize value.
            Array.Copy(outStream, RpcHeaderExtVersionByteSize + RpcHeaderExtFlagsByteSize + RpcHeaderExtSizeActualByteSize, outStream, RpcHeaderExtVersionByteSize + RpcHeaderExtFlagsByteSize, RpcHeaderExtSizeByteSize);

            return outStream;
        }

        /// <summary>
        /// Obfuscates or reverts payload in the stream by applying XOR to each byte of the data with the value 0xA5 and updates the header.
        /// </summary>
        /// <param name="inputStream">The stream to be obfuscated or reverted. The stream should contain header and payload.</param>
        /// <returns>Returns the obfuscated or reverted stream.</returns>
        public static byte[] XOR(byte[] inputStream)
        {
            #region Consts
            // Size of Version field in RPC_HEADER_EXT
            const int RpcHeaderExtVersionByteSize = 2;

            // Size of Flags field in RPC_HEADER_EXT
            const int RpcHeaderExtFlagsByteSize = 2;

            // The length of RPC_HEADER_EXT.
            const int RPCHeaderExtLength = 8;

            // The value that each byte to be obfuscated has XOR applied with, section 3.1.4.11.1.3.
            const byte XorMask = 0xA5;
            #endregion

            if (inputStream == null)
            {
                throw new ArgumentNullException("inputStream");
            }

            if (inputStream.Length < RPCHeaderExtLength)
            {
                throw new ArgumentException("Input must contain valid RPC_HEADER_EXT.", "inputStream");
            }

            ushort size = BitConverter.ToUInt16(inputStream, RpcHeaderExtVersionByteSize + RpcHeaderExtFlagsByteSize);
            byte[] outStream = new byte[size + RPCHeaderExtLength];

            // Copy RPC_HEADER_EXT header info
            Array.Copy(inputStream, 0, outStream, 0, RPCHeaderExtLength);

            // Update XorMagic flag.
            RpcHeaderExtFlags flag = (RpcHeaderExtFlags)BitConverter.ToInt16(inputStream, RpcHeaderExtVersionByteSize);
            flag ^= RpcHeaderExtFlags.XorMagic;
            Array.Copy(BitConverter.GetBytes((short)flag), 0, outStream, RpcHeaderExtVersionByteSize, RpcHeaderExtFlagsByteSize);

            // Obfuscate data by applying XOR to each byte with the value 0xA5
            for (uint counter = RPCHeaderExtLength; counter < outStream.Length; counter++)
            {
                outStream[counter] = (byte)(inputStream[counter] ^ XorMask);
            }

            return outStream;
        }

        /// <summary>
        /// Add bytes length to the first two bytes of the new array, and append bytes to the new array
        /// </summary>
        /// <param name="bytes">Original byte array</param>
        /// <returns>Returns the new byte array contains the length of original array</returns>
        public static byte[] AddInt16LengthBeforeBinaryArray(byte[] bytes)
        {
            int len = 0;
            if (bytes != null)
            {
                len = bytes.Length;
            }

            byte[] retValue = new byte[len + 2];
            retValue[0] = (byte)len;
            retValue[1] = (byte)(len >> 8);
            Array.Copy(bytes, 0, retValue, 2, len);

            return retValue;
        }

        /// <summary>
        /// Compare two byte arrays
        /// </summary>
        /// <param name="bytes1">The first byte array</param>
        /// <param name="bytes2">The second byte array</param>
        /// <returns>Return true if bytes1 is same with bytes2, otherwise return false</returns>
        public static bool CompareByteArray(byte[] bytes1, byte[] bytes2)
        {
            if (bytes1 == null || bytes2 == null)
            {
                if (bytes1 == null && bytes2 == null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            if (bytes1.Length == bytes2.Length)
            {
                for (int i = 0; i < bytes1.Length; i++)
                {
                    if (bytes1[i] != bytes2[i])
                    {
                        return false;
                    }
                }
            }
            else
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Convert a byte array to an unsigned integer. If the length of byte array is greater than 4, just convert the first 4 bytes.
        /// </summary>
        /// <param name="bytes">The byte array to be converted</param>
        /// <returns>Return the unsigned integer value</returns>
        public static uint ConvertByteArrayToUint(byte[] bytes)
        {
            uint uintValue = 0;
            int position = bytes.Length;

            if (position > 4)
            {
                position = 4;
            }

            for (int i = position - 1; i >= 0; i--)
            {
                uintValue <<= 8;
                uintValue += System.Convert.ToUInt32(bytes[i]);
            }

            return uintValue;
        }

        /// <summary>
        /// Check whether the byte array is null terminated ASCII string.
        /// </summary>
        /// <param name="bytes">The byte array to be checked</param>
        /// <returns>Return true if byte array contains a null-terminated ASCII string</returns>
        public static bool IsNullTerminatedASCIIStr(byte[] bytes)
        {
            if (bytes == null)
            {
                return false;
            }

            int len = bytes.Length;

            // Check null terminate.
            if (!(bytes[len - 1] == 0x00))
            {
                return false;
            }

            for (int i = 0; i < bytes.Length; i++)
            {
                // ASCII between 0x00 and 0x7F
                if (!(bytes[i] >= 0x00 && bytes[i] <= 0x7F))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Get string from the byte array. If the byte array contains length, remove the first four byte
        /// </summary>
        /// <param name="bytes">The byte array contains string</param>
        /// <param name="isFirstTwoByteStoreLength">Determine if the byte array contains length</param>
        /// <returns>Return string got from bytes</returns>
        public static string GetStringFromBinary(byte[] bytes, bool isFirstTwoByteStoreLength)
        {
            StringBuilder stringByteValue = new StringBuilder();
            foreach (byte by in bytes)
            {
                stringByteValue.Append(by.ToString("X2"));
            }

            if (isFirstTwoByteStoreLength)
            {
                // Del first four byte which store the byte array length
                stringByteValue.Remove(0, 4);
            }

            return stringByteValue.ToString();
        }

        /// <summary>
        /// Convert a string to byte array.
        /// </summary>
        /// <param name="str">The string to be converted</param>
        /// <returns>Returns a byte array from the string</returns>
        public static byte[] GetBytesFromUnicodeString(string str)
        {
            List<byte> lstBytes = new List<byte>();
            lstBytes.AddRange(new System.Text.UnicodeEncoding().GetBytes(str));
            lstBytes.Add(0x00);
            lstBytes.Add(0x00);
            return lstBytes.ToArray();
        }

        /// <summary>
        /// Convert a string array to byte array
        /// </summary>
        /// <param name="strs">The string array to be converted</param>
        /// <returns>Return a byte array from the list of string</returns>
        public static byte[] GetBytesFromMutiUnicodeString(string[] strs)
        {
            List<byte> lstResult = new List<byte>();
            lstResult.AddRange(BitConverter.GetBytes(strs.Length));

            foreach (string str in strs)
            {
                lstResult.AddRange(GetBytesFromUnicodeString(str));
            }

            return lstResult.ToArray();
        }

        /// <summary>
        /// Check whether the byte array is GUID or not
        /// </summary>
        /// <param name="bytes">The byte array to be checked</param>
        /// <returns>Return true if the byte array is GUID, otherwise return false.</returns>
        public static bool IsGUID(byte[] bytes)
        {
            bool isGUID = false;

            // Check GUID length with 16.
            if (bytes != null && bytes.Length == 16)
            {
                Guid guid = new Guid(bytes);

                // GUID format check regExpression.
                string regexPatten = @"^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\"
                    + @"-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$";
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(regexPatten);
                string guidStr = guid.ToString();
                isGUID = regex.IsMatch(guidStr);
            }

            return isGUID;
        }

        /// <summary>
        /// Convert a hex16 string to byte array
        /// </summary>
        /// <param name="str">The original hex16 string</param>
        /// <returns>Return byte array</returns>
        public static byte[] GetBytesFromBinaryHexString(string str)
        {
            char[] charArray = str.ToLower().ToCharArray();
            if (str.Length % 2 != 0)
            {
                throw new Exception("The hex16 string is invalid");
            }

            byte[] byteArray = new byte[str.Length / 2];
            int index = 0;
            int tempHex = 0;
            for (int i = 0; i < charArray.Length;)
            {
                char a = charArray[i++];

                if (a >= 'a' && a <= 'f')
                {
                    switch (a)
                    {
                        case 'a':
                            tempHex = 10;
                            break;
                        case 'b':
                            tempHex = 11;
                            break;
                        case 'c':
                            tempHex = 12;
                            break;
                        case 'd':
                            tempHex = 13;
                            break;
                        case 'e':
                            tempHex = 14;
                            break;
                        case 'f':
                            tempHex = 15;
                            break;
                    }
                }
                else if (a >= '0' && a <= '9')
                {
                    tempHex = Convert.ToInt32(a) - 48;
                }
                else
                {
                    throw new Exception("The hex16 string is invalid");
                }

                a = charArray[i++];
                if (a >= 'a' && a <= 'f')
                {
                    switch (a)
                    {
                        case 'a':
                            tempHex = (tempHex * 16) + 10;
                            break;
                        case 'b':
                            tempHex = (tempHex * 16) + 11;
                            break;
                        case 'c':
                            tempHex = (tempHex * 16) + 12;
                            break;
                        case 'd':
                            tempHex = (tempHex * 16) + 13;
                            break;
                        case 'e':
                            tempHex = (tempHex * 16) + 14;
                            break;
                        case 'f':
                            tempHex = (tempHex * 16) + 15;
                            break;
                    }
                }
                else if (a >= '0' && a <= '9')
                {
                    tempHex = (tempHex * 16) + Convert.ToInt32(a) - 48;
                }
                else
                {
                    throw new Exception("The hex16 string is invalid");
                }

                byteArray[index++] = Convert.ToByte(tempHex);
            }

            return byteArray;
        }

        /// <summary>
        /// Check whether a byte array is a valid utf16 encoding or not
        /// </summary>
        /// <param name="bytes">The byte array to be checked</param>
        /// <returns>Return true if it is utf16 encoding, otherwise return false</returns>
        public static bool IsUtf16LEString(byte[] bytes)
        {
            if (bytes == null)
            {
                return false;
            }

            int len = bytes.Length;
            if (len % 2 != 0)
            {
                return false;
            }

            for (int i = 1; i < len; i += 2)
            {
                if (bytes[i] != '\0')
                {
                    return false;
                }
            }

            if (bytes[len - 2] == '\0' && bytes[len - 1] == '\0')
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// To determine whether the DN matches regular expression or not.
        /// The regular expression is generated according to the ABNF in [MS-OXOABK] section 2.2.1.1. 
        /// </summary>
        /// <param name="dn">The distinguished name value to be matched.</param>
        /// <param name="format">The DN format needs to be matched.</param>
        /// <returns>If the DN matches the ABNF definition, return true, else false.</returns>     
        public static bool IsDNMatchABNF(string dn, DNFormat format)
        {
            // The regular expression specifies the non-space-teletex ABNF definition "!" / DQUOTE / "%" / "&" / "\" / "(" / ")" / 
            // "*" / "+" / "," / "-" / "." / "0" / "1" / 
            // "2" / "3" / "4" / "5" / "6" / "7" / "8" /
            // "9" / ":" / "<" / "=" / ">" / "?" / "@" /
            // "A" / "B" / "C" / "D" / "E" / "F" / "G" / 
            // "H" / "I" / "J" / "K" / "L" / "M" / "N" / 
            // "O" / "P" / "Q" / "R" / "S" / "T" / "U" / 
            // "V" / "W" / "X" / "Y" / "Z" / "[" / "]" /
            // "_" / "a" / "b" / "c" / "d" / "e" / "f" /
            // "g" / "h" / "i" / "j" / "k" / "l" / "m" /
            // "n" / "o" / "p" / "q" / "r" / "s" / "t" /
            // "u" / "v" / "w" / "x" / "y" / "z" / "|"
            string nonSpaceTeletex = "(" + @"[A-Za-z0-9!%&\()*+,-.:<=>?\[\]@_|]|" + "\"" + ")";

            string teletextChar = "(" + nonSpaceTeletex + @"|\s" + ")";

            string rdn = "(" + "(" + nonSpaceTeletex
                             + "(" + teletextChar + ")" + "{0,62}"
                             + nonSpaceTeletex + ")"
                             + "|"
                             + "(" + nonSpaceTeletex + ")"
                             + ")";

            string orgRdn = "(" + "/(o|O)=" + rdn + ")";

            string orgUnitRdn = "(" + "/(o|O)(u|U)=" + rdn + ")";

            string containerRdn = "(" + "/(c|C)(n|N)=" + rdn + ")";

            string objectRdn = "(" + "/(c|C)(n|N)=" + rdn + ")";

            string regexX500ContainerDn = "(" + orgRdn + orgUnitRdn + containerRdn + "{0,13}" + ")";

            string regexX500Dn = "(" + regexX500ContainerDn + objectRdn + ")";

            string regexX500DnWithNoContainerRdn = "(" + objectRdn + ")";

            string organizationDn = "(" + orgRdn + ")";

            string galAddrlistDn = "/";

            string addresslistDn = "(" + "/guid=" + "[a-fA-F0-9]{32}" + "|/" + ")";

            string[] regexStrs = null;

            switch (format)
            {
                case DNFormat.AddressListDn:
                    regexStrs = new string[] { addresslistDn };
                    break;

                case DNFormat.X500Dn:
                    regexStrs = new string[] { regexX500Dn };
                    break;

                case DNFormat.GalAddrlistDn:
                    regexStrs = new string[] { galAddrlistDn };
                    break;

                case DNFormat.X500DnWithNoContainerRdn:
                    regexStrs = new string[] { regexX500DnWithNoContainerRdn };
                    break;

                case DNFormat.Dn:
                    regexStrs = new string[] { regexX500Dn, addresslistDn, organizationDn };
                    break;

                default:
                    break;
            }

            foreach (string regexStr in regexStrs)
            {
                Regex regex = new Regex(regexStr);
                MatchCollection matchResult = regex.Matches(dn);

                if (matchResult.Count == 1)
                {
                    if (matchResult[0].Value.Equals(dn))
                    {
                        return true;
                    }
                }
            }

            return false;
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
        /// Identify whether the Rop request contains more than one server object handle. Refers to [MS-OXCROPS] for more details.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <returns>Return true if the Rop request contains more than one server object handle, otherwise return false.</returns>
        public static bool IsOutputHandleInRopRequest(ISerializable ropRequest)
        {
            byte ropId = (byte)BitConverter.ToInt16(ropRequest.Serialize(), 0);

            switch (ropId)
            {
                case 0x01: // RopRelease ROP
                case 0x07: // RopGetPropertiesSpecific ROP
                case 0x08: // RopGetPropertiesAll ROP
                case 0x09: // RopGetPropertiesList ROP
                case 0x0A: // RopSetProperties ROP
                case 0x0B: // RopDeleteProperties ROP
                case 0x0D: // RopRemoveAllRecipients ROP
                case 0x0E: // RopModifyRecipients ROP
                case 0x0F: // RopReadRecipients ROP
                case 0x10: // RopReloadCachedInformation ROP
                case 0x12: // RopSetColumns ROP
                case 0x13: // RopSortTable ROP
                case 0x14: // RopRestrict ROP
                case 0x15: // RopQueryRows ROP
                case 0x16: // RopGetStatus ROP
                case 0x17: // RopQueryPosition ROP
                case 0x18: // RopSeekRow ROP
                case 0x19: // RopSeekRowBookmark ROP
                case 0x1A: // RopSeekRowFractional ROP
                case 0x1B: // RopCreateBookmark ROP
                case 0x1D: // RopDeleteFolder ROP
                case 0x1E: // RopDeleteMessages ROP
                case 0x1F: // RopGetMessageStatus ROP
                case 0x20: // RopSetMessageStatus ROP
                case 0x24: // RopDeleteAttachment ROP
                case 0x26: // RopSetReceiveFolder ROP
                case 0x27: // RopGetReceiveFolder ROP
                case 0x2A: // RopNotify ROP
                case 0x2C: // RopReadStream ROP
                case 0x2D: // RopWriteStream ROP
                case 0x2E: // RopSeekStream ROP
                case 0x2F: // RopSetStreamSize ROP
                case 0x30: // RopSetSearchCriteria ROP
                case 0x31: // RopGetSearchCriteria ROP
                case 0x32: // RopSubmitMessage ROP
                case 0x34: // RopAbortSubmit ROP
                case 0x37: // RopQueryColumnsAll ROP
                case 0x38: // RopAbort ROP
                case 0x40: // RopModifyPermissions ROP
                case 0x41: // RopModifyRules ROP
                case 0x42: // RopGetOwningServers ROP
                case 0x43: // RopLongTermIdFromId ROP
                case 0x44: // RopIdFromLongTermId ROP
                case 0x45: // RopPublicFolderIsGhosted ROP
                case 0x47: // RopSetSpooler ROP
                case 0x48: // RopSpoolerLockMessage ROP
                case 0x49: // RopGetAddressTypes ROP
                case 0x4A: // RopTransportSend ROP
                case 0x4E: // RopFastTransferSourceGetBuffer ROP
                case 0x4F: // RopFindRow ROP
                case 0x50: // RopProgress ROP
                case 0x51: // RopTransportNewMail ROP
                case 0x52: // RopGetValidAttachments ROP
                case 0x54: // RopFastTransferDestinationPutBuffer ROP
                case 0x55: // RopGetNamesFromPropertyIds ROP
                case 0x56: // RopGetPropertyIdsFromNames ROP
                case 0x57: // RopUpdateDeferredActionMessages ROP
                case 0x58: // RopEmptyFolder ROP
                case 0x59: // RopExpandRow ROP
                case 0x5A: // RopCollapseRow ROP
                case 0x5B: // RopLockRegionStream ROP
                case 0x5C: // RopUnlockRegionStream ROP
                case 0x5D: // RopCommitStream ROP
                case 0x5E: // RopGetStreamSize ROP
                case 0x5F: // RopQueryNamedProperties ROP
                case 0x60: // RopGetPerUserLongTermIds ROP
                case 0x61: // RopGetPerUserGuid ROP
                case 0x63: // RopReadPerUserInformation ROP
                case 0x64: // RopWritePerUserInformation ROP
                case 0x66: // RopSetReadFlags ROP
                case 0x68: // RopGetReceiveFolderTable ROP
                case 0x6B: // RopGetCollapseState ROP
                case 0x6C: // RopSetCollapseState ROP
                case 0x6D: // RopGetTransportFolder ROP
                case 0x6E: // RopPending ROP
                case 0x6F: // RopOptionsData ROP
                case 0x73: // RopSynchronizationImportHierarchyChange ROP
                case 0x74: // RopSynchronizationImportDeletes ROP
                case 0x75: // RopSynchronizationUploadStateStreamBegin ROP
                case 0x76: // RopSynchronizationUploadStateStreamContinue ROP
                case 0x77: // RopSynchronizationUploadStateStreamEnd ROP
                case 0x78: // RopSynchronizationImportMessageMove ROP
                case 0x79: // RopSetPropertiesNoReplicate ROP
                case 0x7A: // RopDeletePropertiesNoReplicate ROP
                case 0x7B: // RopGetStoreState ROP
                case 0x7F: // RopGetLocalReplicaIds ROP
                case 0x80: // RopSynchronizationImportReadStateChanges ROP
                case 0x81: // RopResetTable ROP
                case 0x86: // RopTellVersion ROP
                case 0x89: // RopFreeBookmark ROP
                case 0x90: // RopWriteAndCommitStream ROP
                case 0x91: // RopHardDeleteMessages ROP
                case 0x92: // RopHardDeleteMessagesAndSubfolders ROP
                case 0x93: // RopSetLocalReplicaMidsetDeleted ROP
                case 0xF9: // RopBackoff ROP
                case 0xFE: // RopLogon ROP
                    return false;
                
                case 0x02: // RopOpenFolder ROP
                case 0x03: // RopOpenMessage ROP
                case 0x04: // RopGetHierarchyTable ROP
                case 0x05: // RopGetContentsTable ROP
                case 0x06: // RopCreateMessage ROP
                case 0x0C: // RopSaveChangesMessage ROP
                case 0x11: // RopSetMessageReadFlag ROP
                case 0x1C: // RopCreateFolder ROP
                case 0x21: // RopGetAttachmentTable ROP
                case 0x22: // RopOpenAttachment ROP
                case 0x23: // RopCreateAttachment ROP
                case 0x25: // RopSaveChangesAttachment ROP
                case 0x29: // RopRegisterNotification ROP
                case 0x2B: // RopOpenStream ROP
                case 0x33: // RopMoveCopyMessages ROP
                case 0x35: // RopMoveFolder ROP
                case 0x36: // RopCopyFolder ROP
                case 0x39: // RopCopyTo ROP
                case 0x3A: // RopCopyToStream ROP
                case 0x3B: // RopCloneStream ROP
                case 0x3E: // RopGetPermissionsTable ROP
                case 0x3F: // RopGetRulesTable ROP
                case 0x46: // RopOpenEmbeddedMessage ROP
                case 0x4B: // RopFastTransferSourceCopyMessages ROP
                case 0x4C: // RopFastTransferSourceCopyFolder ROP
                case 0x4D: // RopFastTransferSourceCopyTo ROP
                case 0x53: // RopFastTransferDestinationConfigure ROP
                case 0x67: // RopCopyProperties ROP
                case 0x69: // RopFastTransferSourceCopyProperties ROP
                case 0x70: // RopSynchronizationConfigure ROP
                case 0x72: // RopSynchronizationImportMessageChange ROP
                case 0x7E: // RopSynchronizationOpenCollector ROP
                case 0x82: // RopSynchronizationGetTransferState ROP
                    return true;

                default:
                    return true;
            }
        }

        /// <summary>
        /// Format the raw data of Rop request and response.
        /// </summary>
        /// <param name="content">Raw data of Rops</param>
        /// <param name="line">The line of the data to be displayed.</param>
        /// <param name="column">The column of the data to be displayed.</param>
        /// <returns>Return the formatted string.</returns>
        public static string FormatBinaryDate(byte[] content, int line = 16, int column = 16)
        {
            if (content == null)
            {
                return string.Empty;
            }

            string targetString = string.Empty;
            int col = 0;
            int lin = 0;
            int position = 0;

            while (position < line * column)
            {
                for (; lin < line; lin++)
                {
                    for (col = 0; col < column; col++)
                    {
                        targetString += content[position].ToString("x2");
                        position = position + 1;
                        if (col != column - 1)
                        {
                            targetString += " ";
                        }
                        else
                        {
                            targetString += "\r\n";
                        }

                        if (position >= content.Length)
                        {
                            return targetString;
                        }
                    }
                }

                if (col == column && lin == line)
                {
                    targetString += "...";
                }
            }

            return targetString;
        }

        /// <summary>
        /// Format the current timestamp to this format "HHmmss_fff".
        /// </summary>
        /// <returns>The formatted current timestamp string.</returns>
        private static string FormatCurrentDateTime()
        {
            return DateTime.Now.ToString("HHmmss_ffffff");
        }

        /// <summary>
        /// Compresses stream using LZ77 algorithm and encodes using Direct2 algorithm.
        /// </summary>
        /// <param name="inputStream">The input stream needed to be compressed.</param>
        /// <returns>Returns the compressed stream.</returns>
        private static byte[] LZ77Compress(byte[] inputStream)
        {
            if (inputStream == null)
            {
                throw new ArgumentNullException("inputStream");
            }

            #region Consts
            // The minimum match is 3 bytes. [MS-OXCRPC], section 3.1.4.11.1.2.1.4.
            const int SizeOfMinimumMatch = 3;

            // The maximum window size restricted by the metadata offset length (13 bytes). [MS-OXCRPC], section 3.1.4.11.1.2.2.3.
            const int MaximumWindowSize = 8193;

            // The size of bitmask. [MS-OXCRPC], section 3.1.4.11.1.2.2.1.
            const int SizeOfBitMask = sizeof(uint);

            // Indicates unused bits in bitmask are filled with 1.
            const uint DefaultBitMaskFilling = 0xFFFFFFFF;

            // Means the first 31 bytes are actual data. (1000 0000 0000 0000 0000 0000 0000 0000). Uses as the beginning of checking bitmask.
            const uint BitMaskOf31ActualData = 0x80000000;

            // The size of metadata. [MS-OXCRPC], section 3.1.4.11.1.2.2.3.
            const int SizeOfMetadata = sizeof(short);

            // The low-order three bits are the length. [MS-OXCRPC], section 3.1.4.11.1.2.2.4.
            const int MetadataLengthBitLength = 3;

            // The maximum value when all low-order three bits are "1". [MS-OXCRPC], section 3.1.4.11.1.2.2.4.
            const int MetadataLengthFullBitsValue = 7;

            // The shared byte with value 1111.
            const byte SharedByteSetLow4Bits = 0xF;

            // The next byte with value 11111111.
            const byte NextByteSetAllBits = 0xFF;

            // The size of final two bytes which is used to calculate the match length equal or greater than 280.
            const int SizeOfFinalTwoBytes = sizeof(short);

            // The increase of the length of output stream each time stream length is insufficient.
            const int OutStreamLengthIncrease = 128;

            // The maximum metadata length. It would be the first 2 bytes, 1 shared byte, 1 additional byte, and the final 2 bytes, section 3.1.4.11.1.2.2.4.
            const int MaximumMetadataLength = 6;
            #endregion

            #region Variables
            // The position of the byte in the input stream that is currently being coded (the beginning of the lookahead buffer).
            int codingPosition;

            // The starting position of the window in the input stream.
            int windowStartingPosition;

            // The position of the byte in the output stream where the data byte or metadata is being written.
            int outBytesPosition;

            // To distinguish data from metadata in the compressed byte stream. [MS-OXCRPC], section 3.1.4.11.1.2.2.1.
            uint bitMask;

            // Indicates the bit representing the next byte to be processed is "1".
            uint bitMaskPointer;

            // The position of the bitmask in the output stream.
            int outBitMaskPostion;

            // The position of the shared byte in the output stream. After the high-order nibble of the byte is used, this value is set to "-1" to indicate it needs to be set to a new position next time.
            int sharedBytePosition;
            #endregion

            int size = inputStream.Length;
            byte[] outStream = new byte[size];

            // Set the coding position to the beginning of the input stream.
            codingPosition = 0;
            windowStartingPosition = 0;
            outBytesPosition = 0;
            outBitMaskPostion = 0;
            bitMaskPointer = 0;
            sharedBytePosition = -1;
            while (codingPosition < size)
            {
                // Enlarge output stream length to ensure it's sufficient.
                if (outBytesPosition + MaximumMetadataLength + SizeOfBitMask > outStream.Length)
                {
                    Array.Resize<byte>(ref outStream, outStream.Length + OutStreamLengthIncrease);
                }

                // Move to the next bitmask if all bits in current bitmask are set.
                if (bitMaskPointer == 0)
                {
                    outBitMaskPostion = outBytesPosition;
                    Array.Copy(BitConverter.GetBytes(DefaultBitMaskFilling), 0, outStream, outBitMaskPostion, SizeOfBitMask);
                    outBytesPosition += SizeOfBitMask;
                    bitMaskPointer = BitMaskOf31ActualData;
                }

                // Find the longest match in the window for the lookahead buffer.
                int matchLength = 0;
                int matchOffset = 0;
                for (int matchOffsetPosition = windowStartingPosition; matchOffsetPosition < codingPosition - matchLength; matchOffsetPosition++)
                {
                    if (inputStream[codingPosition] == inputStream[matchOffsetPosition])
                    {
                        int currentMatchLength = 1;
                        while (currentMatchLength < size - codingPosition - 1)
                        {
                            if (inputStream[codingPosition + currentMatchLength] == inputStream[matchOffsetPosition + currentMatchLength])
                            {
                                currentMatchLength++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (currentMatchLength >= matchLength)
                        {
                            matchLength = currentMatchLength;
                            matchOffset = codingPosition - matchOffsetPosition;
                        }
                    }
                }

                // Output the data byte or metadata with DIRECT2 encoding.
                if (matchLength < SizeOfMinimumMatch)
                {
                    // Next one is not metadata. Set bitmask bit to '0'.
                    bitMask = BitConverter.ToUInt32(outStream, outBitMaskPostion);
                    bitMask &= ~bitMaskPointer;
                    Array.Copy(BitConverter.GetBytes(bitMask), 0, outStream, outBitMaskPostion, sizeof(uint));
                    bitMaskPointer >>= 1;

                    // Fill data byte in output stream.
                    outStream[outBytesPosition] = inputStream[codingPosition];
                    outBytesPosition++;
                    codingPosition++;
                }
                else
                {
                    // Next one is metadata. Set bitmask bit to '1'.
                    bitMask = BitConverter.ToUInt32(outStream, outBitMaskPostion);
                    bitMask |= bitMaskPointer;
                    Array.Copy(BitConverter.GetBytes(bitMask), 0, outStream, outBitMaskPostion, sizeof(uint));
                    bitMaskPointer >>= 1;

                    // Fill metadata offset and length in output stream. 
                    // Use the high-order 13 bits in metadata bytes to store metadata offset.
                    matchOffset--;
                    int remainningMatchLength = matchLength - SizeOfMinimumMatch;
                    if (remainningMatchLength < MetadataLengthFullBitsValue)
                    {
                        // Use the low-order bits in metadata bytes to represent metadata length.
                        short metadata = (short)((matchOffset << MetadataLengthBitLength) + remainningMatchLength);
                        Array.Copy(BitConverter.GetBytes(metadata), 0, outStream, outBytesPosition, SizeOfMetadata);
                        outBytesPosition += SizeOfMetadata;
                    }
                    else
                    {
                        short metadata = (short)((matchOffset << MetadataLengthBitLength) + MetadataLengthFullBitsValue);
                        Array.Copy(BitConverter.GetBytes(metadata), 0, outStream, outBytesPosition, SizeOfMetadata);
                        outBytesPosition += SizeOfMetadata;
                        remainningMatchLength -= MetadataLengthFullBitsValue;
                        if (remainningMatchLength < (int)SharedByteSetLow4Bits)
                        {
                            // Additionally use the low-order or high-order nibble in shared byte to represent metadata length.
                            if (sharedBytePosition < 0)
                            {
                                sharedBytePosition = outBytesPosition++;
                                outStream[sharedBytePosition] = (byte)remainningMatchLength;
                            }
                            else
                            {
                                outStream[sharedBytePosition] += (byte)(remainningMatchLength << 4);
                                sharedBytePosition = -1;
                            }
                        }
                        else
                        {
                            if (sharedBytePosition < 0)
                            {
                                sharedBytePosition = outBytesPosition++;
                                outStream[sharedBytePosition] = SharedByteSetLow4Bits;
                            }
                            else
                            {
                                outStream[sharedBytePosition] += SharedByteSetLow4Bits << 4;
                                sharedBytePosition = -1;
                            }

                            remainningMatchLength -= (int)SharedByteSetLow4Bits;

                            if (remainningMatchLength < (int)NextByteSetAllBits)
                            {
                                // Additionally use another byte to represent metadata length.
                                outStream[outBytesPosition++] = (byte)remainningMatchLength;
                            }
                            else
                            {
                                outStream[outBytesPosition++] = NextByteSetAllBits;

                                // Use the final two bytes to represent metadata length.
                                Array.Copy(BitConverter.GetBytes((short)(matchLength - SizeOfMinimumMatch)), 0, outStream, outBytesPosition, SizeOfFinalTwoBytes);
                                outBytesPosition += SizeOfFinalTwoBytes;
                            }
                        }
                    }

                    // If the lookahead buffer is not empty, move the coding position (and the window) L bytes forward.
                    codingPosition += matchLength;
                    if (codingPosition - windowStartingPosition > MaximumWindowSize)
                    {
                        windowStartingPosition = codingPosition - MaximumWindowSize;
                    }
                }
            }

            // Move to the next bitmask if all bits in current bitmask are set because additional bit "1" is needed as EOF.
            if (bitMaskPointer == 0)
            {
                outBitMaskPostion = outBytesPosition;
                Array.Copy(BitConverter.GetBytes(DefaultBitMaskFilling), 0, outStream, outBitMaskPostion, SizeOfBitMask);
                outBytesPosition += SizeOfBitMask;
            }

            // Resize output stream
            Array.Resize<byte>(ref outStream, outBytesPosition);
            return outStream;
        }

        /// <summary>
        /// Decodes stream using Direct2 algorithm and decompresses using LZ77 algorithm.
        /// </summary>
        /// <param name="inputStream">The input stream needed to be decompressed.</param>
        /// <param name="actualSize">The expected size of the decompressed output stream.</param>
        /// <returns>Returns the decompressed stream.</returns>
        private static byte[] LZ77Decompress(byte[] inputStream, int actualSize)
        {
            if (inputStream == null)
            {
                throw new ArgumentNullException("inputStream");
            }

            #region Variables
            // To distinguish data from metadata in the compressed byte stream. [MS-OXCRPC], section 3.1.7.2.2.1.
            int bitMask;

            // Indicates the bit representing the next byte to be processed is "1".
            uint bitMaskPointer;

            // The count of bitmask.
            uint bitMaskCount;

            // Metadata offset.
            int offset;

            // The length of metadata.
            int length;

            // The container of redundant information which is used to reduce the size of input data.
            int metadata;

            // The length of metadata. For more detail, refer to [MS-OXCRPC], section 3.1.7.2.2.4.
            int metadataLength;

            // The additive length contained by the nibble of shared byte.
            int lengthInSharedByte;

            // The byte follows the bitmask.
            byte nextByte;

            // The byte follows the initial 2-byte metadata whenever the match length is greater than nine.
            // The nibble of this byte is "reserved" for the next metadata instance when the length is greater than nine.
            // For more detail, refer to [MS-OXCRPC], section 3.1.7.2.2.4.
            byte sharedByte;

            // Indicates which nibble of shared byte to be used. True indicates high-order nibble, false indicates low-order nibble.
            bool useSharedByteHighOrderNibble;

            // The count of bytes in inStream.
            int inputBytesCount;

            // The count of bytes in outStream.
            int outCount;
            #endregion

            #region Consts
            // Means the first 31 bytes are actual data. (1000 0000 0000 0000 0000 0000 0000 0000)
            // Uses as the beginning of checking bitmask.
            const uint BitMaskOf31ActualData = 0x80000000;

            // The high-order 13 bits are a first complement of the offset. (1111 1111 1111 1000)
            const int BitMaskOfHigh13AreFirstComplementOfOffset = 0xFFF8;

            // The size of metadata. [MS-OXCRPC], section 3.1.7.2.2.3.
            const int SizeOfMetadata = sizeof(short);

            // The size of shared byte.
            const int SizeOfSharedByte = sizeof(byte);

            // The low-order three bits are the length. [MS-OXCRPC], section 3.1.7.2.2.3.
            const int OffsetOfLengthInMetadata = 3;

            // The three bits in the original two bytes of metadata with value b'111'
            const int BitSetLower3Bits = 0x7;

            // The size of nibble in shared byte.
            const int SizeOfNibble = 4;

            // The minimum match is 3 bytes. [MS-OXCRPC], section 3.1.7.2.2.4.
            const int SizeOfMinimumMatch = 3;

            // Three low-order bits of the 2-bytes metadata allow for the expression of lengths from 3 to 9.
            // Because 3 is the minimum match and b'111' is reserved.
            // So every time the match length is greater than 9, there will be an additional byte follows the initial 2-byte metadata.
            // Refer to [MS-OXCRPC], section 3.1.7.2.2.4.
            const int MatchLengthWithAdditionalByte = 10;

            // The shared byte with value 1111.
            const byte SharedByteSetLow4Bits = 0xF;

            // The next byte with value 11111111.
            const byte NextByteSetAllBits = 0xFF;

            // The size of final two bytes which is used to calculate the match length equal or greater than 280.
            const int SizeOfFinalTwoBytes = 4;

            // Each bit in bitmask (4 bytes) can distinguish data from metadata in the compressed byte stream.
            const int CountOfBitmask = 32;

            #endregion

            byte[] outStream = new byte[actualSize];
            int size = inputStream.Length;

            outCount = 0;
            inputBytesCount = 0;
            useSharedByteHighOrderNibble = false;
            sharedByte = 0;
            while (inputBytesCount < size)
            {
                bitMask = BitConverter.ToInt32(inputStream, inputBytesCount);

                // The size of bitmask is 4 bytes.
                inputBytesCount += sizeof(uint);
                bitMaskPointer = BitMaskOf31ActualData;
                bitMaskCount = 0;
                do
                {
                    // The size of RPC_HEADER_EXT.
                    if (inputBytesCount < size)
                    {
                        // If the next byte in inStream is not metadata
                        if ((bitMask & bitMaskPointer) == 0)
                        {
                            outStream[outCount] = inputStream[inputBytesCount];
                            outCount++;
                            inputBytesCount++;

                            // Move to the next bitmask.
                            bitMaskPointer >>= 1;
                            bitMaskCount++;
                        }
                        else
                        {
                            // If next set of bytes is metadata, count offset and length
                            // This protocol assumes the metadata is two bytes in length
                            metadata = (int)BitConverter.ToInt16(inputStream, inputBytesCount);

                            // The high-order 13 bits are a first complement of the offset
                            offset = (metadata & BitMaskOfHigh13AreFirstComplementOfOffset) >> OffsetOfLengthInMetadata;
                            offset++;

                            #region Count Length
                            // If three bits in the original two bytes of metadata is not b'111', length equals to bit value plus 3. (Less than 10)
                            if ((metadata & BitSetLower3Bits) != BitSetLower3Bits)
                            {
                                length = (metadata & BitSetLower3Bits) + SizeOfMinimumMatch;
                                metadataLength = SizeOfMetadata;
                            }
                            else
                            {
                                // If three bits in the original two bytes of metadata is b'111', need shared byte. (Larger than 9)
                                // First time use low-order nibble
                                if (!useSharedByteHighOrderNibble)
                                {
                                    sharedByte = inputStream[inputBytesCount + SizeOfMetadata];
                                    lengthInSharedByte = sharedByte & SharedByteSetLow4Bits;

                                    // Next time will use high-order nibble of shared byte.
                                    useSharedByteHighOrderNibble = true;
                                    metadataLength = SizeOfMetadata + SizeOfSharedByte;
                                }
                                else
                                {
                                    // Next time use high-order nibble
                                    lengthInSharedByte = sharedByte >> SizeOfNibble;

                                    // Next time will use low-order nibble of shared byte.
                                    useSharedByteHighOrderNibble = false;
                                    metadataLength = SizeOfMetadata;
                                }

                                // If length in shared byte is not b'1111', length equals to 3+7+lengthInSharedByte
                                if (lengthInSharedByte != SharedByteSetLow4Bits)
                                {
                                    length = MatchLengthWithAdditionalByte + lengthInSharedByte;
                                }
                                else
                                {
                                    // If length in shared byte is b'1111'(larger than 24), next byte will be use.
                                    if (useSharedByteHighOrderNibble)
                                    {
                                        nextByte = inputStream[inputBytesCount + SizeOfMetadata + 1];
                                    }
                                    else
                                    {
                                        nextByte = inputStream[inputBytesCount + SizeOfMetadata];
                                    }

                                    // If next byte is not b'11111111', length equals to 3+7+lengthInSharedByte + nextByte
                                    if (nextByte != NextByteSetAllBits)
                                    {
                                        length = MatchLengthWithAdditionalByte + lengthInSharedByte + nextByte;
                                        metadataLength++;
                                    }
                                    else
                                    {
                                        // Consider the presence of shared bytes
                                        int useSharedSizeOfTwoBytes = useSharedByteHighOrderNibble ? SizeOfFinalTwoBytes : SizeOfFinalTwoBytes - 1;

                                        // If next byte is b'11111111' (larger than 279), use the next two bytes to represent length
                                        // These two bytes represent a length of 277+3 (minimum match length)
                                        length = (int)BitConverter.ToInt16(inputStream, inputBytesCount + useSharedSizeOfTwoBytes) + SizeOfMinimumMatch;
                                        metadataLength += SizeOfMinimumMatch;
                                    }
                                }
                            }
                            #endregion

                            for (int counter = 0; counter < length; counter++)
                            {
                                outStream[outCount + counter] = outStream[outCount - offset + counter];
                            }

                            inputBytesCount += metadataLength;
                            outCount += length;

                            // Move to the next bitmask.
                            bitMaskPointer >>= 1;
                            bitMaskCount++;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                while (bitMaskCount != CountOfBitmask);
            }

            // If the output stream's length doesn't equal to the expected size, the decompression is failed.
            if (outCount != actualSize)
            {
                throw new InvalidOperationException(string.Format("Decompression failed because decompressed byte array length ({0}) doesn't equal to the expected length ({1}).", outCount, actualSize));
            }

            return outStream;
        }
    }
}