//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.IO;
    using System.Runtime.Serialization.Json;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;

    /// <summary>
    /// A class is used to perform serializer operations for MS-WOPI protocol.
    /// </summary>
    public class WOPISerializerHelper : HelperBase
    {
        /// <summary>
        /// The require elements of the JSON string from CheckFileInfo response.
        /// </summary>
        private static string[] jsonRequireItemsForCheckFileInfo = { "BaseFileName", "OwnerId", "SHA256", "Size", "Version" };

        /// <summary>
        /// The require elements of the JSON string from PutRelativeFile response.
        /// </summary>
        private static string[] jsonRequireItemsForPutRelativeFile = { "Name", "Url" };

        /// <summary>
        /// The require elements of the JSON string from SecureStore response.
        /// </summary>
        private static string[] jsonRequireItemsForReadSecureStore = { "UserName", "Password" };

        /// <summary>
        /// The require elements of the JSON string from CheckFolderInfo response.
        /// </summary>
        private static string[] jsonRequireItemsForCheckFolderInfo = { "FolderName", "OwnerId" };

        /// <summary>
        /// Prevents a default instance of the WOPISerializerHelper class from being created
        /// </summary>
        private WOPISerializerHelper()
        {
        }

        /// <summary>
        /// Convert the JSON string to the specified Object.
        /// </summary>
        /// <typeparam name="T">The type of the JSON object which is defined in MS-WOPI protocol.</typeparam>
        /// <param name="jsonValue">The value of the JSON strings.</param>
        /// <returns>A return value represents the object which is de-serialize from JSON string.</returns>
        public static T JsonToObject<T>(string jsonValue) where T : class
        {
            Type currentType = typeof(T);

            DataContractJsonSerializer serializer = new DataContractJsonSerializer(currentType);

            MemoryStream memoryStreamInstance = new MemoryStream(Encoding.Default.GetBytes(jsonValue));

            if (currentType.Name.Equals("CheckFileInfo"))
            {
                CheckRequiredJsonItem(jsonRequireItemsForCheckFileInfo, jsonValue);
            }
            else if (currentType.Name.Equals("PutRelativeFile"))
            {
                CheckRequiredJsonItem(jsonRequireItemsForPutRelativeFile, jsonValue);
            }
            else if (currentType.Name.Equals("ReadSecureStore"))
            {
                CheckRequiredJsonItem(jsonRequireItemsForReadSecureStore, jsonValue);
            }
            else if (currentType.Name.Equals("CheckFolderInfo"))
            {
                CheckRequiredJsonItem(jsonRequireItemsForCheckFolderInfo, jsonValue);
            }

            T expectedInstance = serializer.ReadObject(memoryStreamInstance) as T;
            memoryStreamInstance.Dispose();
            if (null == expectedInstance)
            {
                throw new InvalidOperationException(
                    string.Format(
                    "Convert the JSON string to [{0}] type failed, the JSON string:\r\n{1}",
                    currentType.Name,
                    jsonValue));
            }

            return expectedInstance;
        }

        /// <summary>
        /// This method is used to convert the XML date to the Discovery object.
        /// </summary>
        /// <param name="xmlValue">The value of the xml string.</param>
        /// <returns>The object value which is converted from the xml string.</returns>
        public static wopidiscovery DeserializeXmlToDiscoveryObject(string xmlValue)
        {
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<string>(xmlValue, "xmlString", "DeserializeXmlToDiscoveryObject");

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(wopidiscovery));
            wopidiscovery discovery = null;

            using (StringReader strReader = new StringReader(xmlValue))
            {
                discovery = xmlSerializer.Deserialize(strReader) as wopidiscovery;
                if (null == discovery)
                {
                    throw new ArgumentNullException("discovery", "Could not get the current xml string to the expected Discovery type.");
                }
            }

            return discovery;
        }

        /// <summary>
        /// A method is used to get a xml string from a WOPI Discovery type object. The xml string is used to response the discovery request.
        /// </summary>
        /// <param name="discoveryObject">A parameter represents the WOPI Discovery object which contain the discovery information.</param>
        /// <returns>A parameter represents the xml string which contains the discovery information.</returns>
        public static string GetDiscoveryXmlFromDiscoveryObject(wopidiscovery discoveryObject)
        {
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<wopidiscovery>(discoveryObject, "discoveryObject", "GetDiscoveryXmlFromDiscoveryObject");

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(wopidiscovery));
            string xmlString = string.Empty;

            MemoryStream memorySteam = null;
            try
            {
                memorySteam = new MemoryStream();
                StreamWriter streamWriter = null;
                try
                {
                    streamWriter = new StreamWriter(memorySteam, Encoding.UTF8);

                    // Remove w3c default namespace prefix in serialize process.
                    XmlSerializerNamespaces nameSpaceInstance = new XmlSerializerNamespaces();
                    nameSpaceInstance.Add(string.Empty, string.Empty);
                    xmlSerializer.Serialize(streamWriter, discoveryObject, nameSpaceInstance);

                    // Read the MemoryStream to output the xml string.
                    memorySteam.Position = 0;
                    using (StreamReader streamReader = new StreamReader(memorySteam))
                    {
                        xmlString = streamReader.ReadToEnd();
                    }
                }
                finally
                {
                    if (streamWriter != null)
                    {
                        streamWriter.Dispose();
                    }
                }
            }
            finally
            {
                if (memorySteam != null)
                {
                    memorySteam.Dispose();
                }
            }

            if (string.IsNullOrEmpty(xmlString))
            {
                throw new InvalidOperationException("Could not get the xml string.");
            }

            // Format the serialized xml string.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlString);
            return xmlDoc.OuterXml;
        }

        /// <summary>
        /// A method is used to check whether the item exists in the JSON strings.
        /// </summary>
        /// <param name="jsonValue">A parameter represents the JSON string </param>
        /// <param name="jsonItem">A parameter represents the JSON item which is expected to be contained.</param>
        /// <returns>Return 'true' indicating the item exists.</returns>
        public static bool CheckContainItem(string jsonValue, string jsonItem)
        {
            if (jsonValue.Contains("\"" + jsonItem + "\"" + ":"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// A method is used to check whether the required items exists in the JSON strings.
        /// </summary>
        /// <param name="jsonItems">The collection for the require items.</param>
        /// <param name="jsonString">The JSON string.</param>
        /// <returns>Return 'true' indicating all required item exists in the JSON string.</returns>
        private static bool CheckRequiredJsonItem(string[] jsonItems, string jsonString)
        {
            foreach (string item in jsonItems)
            {
                if (!jsonString.Contains("\"" + item + "\"" + ":"))
                {
                    throw new InvalidOperationException("The require item" + item + "doesn't exist in the" + jsonString + "Json string.");
                }
            }

            return true;
        }
    }
}
