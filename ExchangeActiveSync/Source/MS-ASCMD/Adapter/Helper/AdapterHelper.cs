namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to assist MS-ASCMDAdapter.
    /// </summary>
    public class AdapterHelper : ManagedAdapterBase
    {
        #region Public Methods

        /// <summary>
        /// Pick up the Status value from the XMLString of response.
        /// </summary>
        /// <param name="responsexmlString">Response xml String.</param>
        /// <returns>status value</returns>
        public static byte PickUpRootStatusValueFromXMLString(string responsexmlString)
        {
            // When the input response xml string is Null or Empty, return value "0" to track it.
            if (string.IsNullOrEmpty(responsexmlString))
            {
                return 0;
            }

            // Pick up root Status value if it exist in XMLString of the response , otherwise return "0" for root Status.
            int startPointOfFirstStatusTag = responsexmlString.IndexOf(@"<Status>", 0, StringComparison.OrdinalIgnoreCase);
            int startPointOfCloseStatus = responsexmlString.IndexOf(@"</Status>", 0, StringComparison.OrdinalIgnoreCase);

            if (startPointOfFirstStatusTag > 0 && startPointOfCloseStatus > startPointOfFirstStatusTag)
            {
                int lengOfStatusTag = @"<Status>".Length;
                int startPointOfStatusValue = startPointOfFirstStatusTag + lengOfStatusTag;
                int lengOfValue = startPointOfCloseStatus - startPointOfStatusValue;
                if (lengOfValue != 0)
                {
                    string statusValyeString = responsexmlString.Substring(startPointOfStatusValue, lengOfValue);
                    byte statusValue;
                    if (byte.TryParse(statusValyeString, out statusValue))
                    {
                        // Return the expected output.
                        return statusValue;
                    }
                }
            }

            return 0;
        }

        /// <summary>
        /// Get ValidateCert response status code returned by the ValidateCert operation.
        /// </summary>
        /// <param name="response">The data of ValidateCert response.</param>
        /// <returns>The status code</returns>
        public static string GetValidateCertStatusCode(ValidateCertResponse response)
        {
            string xmlResponse = response.ResponseDataXML;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlResponse);
            XmlNamespaceManager xmlNameSpaceManager = new XmlNamespaceManager(doc.NameTable);
            xmlNameSpaceManager.AddNamespace("e", "ValidateCert");
            XmlNodeList status = doc.SelectNodes("/e:ValidateCert/e:Certificate/e:Status", xmlNameSpaceManager);

            if (status != null && status.Count == 0)
            {
                status = doc.SelectNodes("/e:ValidateCert/e:Status", xmlNameSpaceManager);
            }

            return status[0].InnerText;
        }

        /// <summary>
        /// Get an array of the possible status codes for a command.
        /// </summary>
        /// <param name="statusValues">The status codes for a specific command.</param>
        /// <returns>The possible status codes for a command.</returns>
        public static string[] ValidStatus(string[] statusValues)
        {
            List<string> commonStatus = new List<string>(new string[] { "101", "102", "103", "104", "105", "106", "107", "108", "109", "110", "111", "112", "113", "114", "115", "116", "117", "118", "119", "120", "121", "122", "123", "124", "125", "126", "127", "128", "129", "130", "131", "132", "133", "134", "135", "136", "137", "138", "139", "140", "141", "142", "143", "144", "145", "146", "147", "148", "149", "150", "151", "152", "153", "154", "155", "156", "160", "161", "162", "163", "164", "165", "166", "167", "168", "169", "170", "171", "172", "173", "174", "175", "176", "177" });

            for (int i = 0; i < statusValues.Length; i++)
            {
                commonStatus.Add(statusValues[i]);
            }

            return commonStatus.ToArray();
        }

        /// <summary>
        /// Get the ASCMD Status from Response.
        /// </summary>
        /// <param name="response">ASCMD command responses.</param>
        /// <returns>status value</returns>
        public byte GetStatusFromResponses(object response)
        {
            byte statusValue = 0;
            if (null == response)
            {
                return statusValue;
            }

            Type responsesType = response.GetType();
            Type baseType = responsesType.BaseType;

            // Verify whether it is inherited from ActiveSyncDataStructure.ActiveSyncResponses<T>.
            if (null == baseType || 0 != baseType.FullName.IndexOf("ActiveSyncDataStructure.ActiveSyncResponses`1", StringComparison.OrdinalIgnoreCase))
            {
                return statusValue;
            }

            // Verify whether the responses is "SendStringResponses"
            if (0 == responsesType.FullName.IndexOf("ActiveSyncDataStructure.SendStringResponses", StringComparison.OrdinalIgnoreCase))
            {
                SendStringResponse stringResponses = response as SendStringResponse;
                if (stringResponses != null)
                {
                    return PickUpRootStatusValueFromXMLString(stringResponses.ResponseDataXML);
                }
            }

            // If it is not a SendStringResponses, get the status field under the Responses element.
            PropertyInfo lowLevelResponsesData = responsesType.GetProperty("ResponsesData");
            if (null == lowLevelResponsesData)
            {
                return statusValue;
            }

            object responsesDataInstance = lowLevelResponsesData.GetValue(response, null);
            if (null == responsesDataInstance)
            {
                return statusValue;
            }

            // Get the ResponsesData type
            Type responsesDataType = responsesDataInstance.GetType();
            PropertyInfo statusproperty = responsesDataType.GetProperty("Status", typeof(byte));

            // If it doesn't contain the "status" in serialized data type, try to pick up from "ResponsesDataXML"
            if (null == statusproperty)
            {
                PropertyInfo responsesXMLProperty = responsesType.GetProperty("ResponsesDataXML");
                if (responsesXMLProperty != null)
                {
                    string responsesDataXMLTemp = responsesXMLProperty.GetValue(response, null) as string;
                    return PickUpRootStatusValueFromXMLString(responsesDataXMLTemp);
                }

                return statusValue;
            }

            try
            {
                statusValue = (byte)statusproperty.GetValue(responsesDataInstance, null);
            }
            catch (FormatException ex)
            {
                Site.Log.Add(LogEntryKind.TestError, "Could not parse the ResponsesData.Status to byte type\r\n{0}", ex.Message);
            }

            return statusValue;
        }

        #endregion
    }
}