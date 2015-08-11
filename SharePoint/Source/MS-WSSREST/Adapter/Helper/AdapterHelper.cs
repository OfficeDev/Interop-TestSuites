namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to assist MS-WSSRESTAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        #region Variables

        /// <summary>
        /// An object provides logging, assertions, and SUT adapters for test code onto its execution context.
        /// </summary>
        private static ITestSite site;

        #endregion Variables

        #region Methods

        /// <summary>
        /// Initialize object of "Site".
        /// </summary>
        /// <param name="testSite">A object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void Initialize(ITestSite testSite)
        {
            site = testSite;
        }

        /// <summary>
        /// Get xml content from http response.
        /// </summary>
        /// <param name="httpResponse">The http web response.</param>
        /// <returns>The xml content.</returns>
        public static XmlDocument GetXmlData(HttpWebResponse httpResponse)
        {
            string temp = GetResponseContent(httpResponse.GetResponseStream());

            if (string.IsNullOrEmpty(temp))
            {
                return null;
            }

            XmlDocument rawXmlDoc = new XmlDocument();
            rawXmlDoc.LoadXml(temp);

            return rawXmlDoc;
        }

        /// <summary>
        /// Get content from stream.
        /// </summary>
        /// <param name="stream">The response stream.</param>
        /// <returns>The content of response.</returns>
        public static string GetResponseContent(Stream stream)
        {
            string temp = string.Empty;
            using (StreamReader sr = new StreamReader(stream))
            {
                temp = sr.ReadToEnd();
            }

            return temp;
        }

        /// <summary>
        /// Analyze http response.
        /// </summary>
        /// <param name="response">The http response.</param>
        /// <returns>A list of Entry instance.</returns>
        public static List<Entry> AnalyseResponse(XmlDocument response)
        {
            List<Entry> results = new List<Entry>();

            if (response == null)
            {
                return null;
            }

            XmlNodeList nodes = response.GetElementsByTagName("entry");

            foreach (XmlNode entryNode in nodes)
            {
                Entry result = new Entry();

                if (null != entryNode.Attributes["m:etag"])
                {
                    result.Etag = entryNode.Attributes["m:etag"].Value;
                }

                foreach (XmlNode node in entryNode.ChildNodes)
                {
                    if (node.Name == "id")
                    {
                        result.ID = node.InnerText;
                    }
                    else if (node.Name == "title")
                    {
                        result.Title = node.InnerText;
                    }
                    else if (node.Name == "updated")
                    {
                        result.Updated = DateTime.Parse(node.InnerText);
                    }
                    else if (node.Name == "content")
                    {
                        if (node.ChildNodes.Count > 0)
                        {
                            result.Properties = new Dictionary<string, string>();
                            XmlNode pro = node.FirstChild;

                            foreach (XmlNode xn in pro.ChildNodes)
                            {
                                result.Properties.Add(xn.LocalName, xn.InnerText);
                            }
                        }
                    }
                    else if (node.LocalName == "properties")
                    {
                        result.Properties = new Dictionary<string, string>();
                        foreach (XmlNode xn in node.ChildNodes)
                        {
                            result.Properties.Add(xn.LocalName, xn.InnerText);
                        }
                    }
                }

                results.Add(result);
            }

            return results;
        }

        #endregion
    }
}