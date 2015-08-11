namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The class provides the methods to assist MS_ASPROVAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// Get policies from Provision command response.
        /// </summary>
        /// <param name="provisionResponse">The response of Provision command.</param>
        /// <returns>The dictionary of policies gotten from Provision command response.</returns>
        public static Dictionary<string, string> GetPoliciesFromProvisionResponse(ActiveSyncResponseBase<Response.Provision> provisionResponse)
        {
            Dictionary<string, string> policiesSetting = new Dictionary<string, string>();
            if (null == provisionResponse || string.IsNullOrEmpty(provisionResponse.ResponseDataXML))
            {
                return policiesSetting;
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(provisionResponse.ResponseDataXML);
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
            namespaceManager.AddNamespace("prov", "Provision");
            XmlNode provisionDocNode = xmlDoc.SelectSingleNode(@"//prov:EASProvisionDoc", namespaceManager);

            if (provisionDocNode != null && provisionDocNode.HasChildNodes)
            {
                foreach (XmlNode policySetting in provisionDocNode.ChildNodes)
                {
                    string policyValue = string.IsNullOrEmpty(policySetting.InnerText) ? string.Empty : policySetting.InnerText;
                    string policyName = string.IsNullOrEmpty(policySetting.LocalName) ? string.Empty : policySetting.LocalName;
                    policiesSetting.Add(policyName, policyValue);
                }
            }

            return policiesSetting;
        }
    }
}