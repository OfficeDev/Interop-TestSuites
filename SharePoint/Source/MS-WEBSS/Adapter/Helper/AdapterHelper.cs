//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    ///  The class is used to define some common and helpful methods for adapter.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// A object provides logging, assertions, and SUT adapters for test code onto its execution context.
        /// </summary>
        private static ITestSite baseSite;

        /// <summary>
        /// Initialize "Site".
        /// </summary>
        /// <param name="site">A object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void Initialize(ITestSite site)
        {
            baseSite = site;
        }

        /// <summary>
        /// Retrieves the user credential based on authentication type.
        /// </summary>
        /// <param name="userAuthentication">The authentication information client used.</param>
        /// <returns>The credentials of the specified authentication type.</returns>
        public static NetworkCredential ConfigureCredential(UserAuthentication userAuthentication)
        {
            string userName, password, domain;

            switch (userAuthentication)
            {
                case UserAuthentication.Authenticated:
                    // use authenticated account info
                    domain = Common.GetConfigurationPropertyValue("Domain", baseSite);
                    userName = Common.GetConfigurationPropertyValue("UserName", baseSite);
                    password = Common.GetConfigurationPropertyValue("Password", baseSite);
                    return new NetworkCredential(userName, password, domain);
                default:
                    // use unauthenticated account info
                    domain = GenerateRandomString(5);
                    userName = GenerateRandomString(5);
                    password = GenerateRandomString(10);
                    return new NetworkCredential(userName, password, domain);
            }
        }

        /// <summary>
        /// This method is used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>Returns the random generated string.</returns>
        public static string GenerateRandomString(int size)
        {
            Random random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
                ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Combine network path.
        /// </summary>
        /// <param name="path1">Network path1.</param>
        /// <param name="path2">Network path2.</param>
        /// <returns>Combined network path.</returns>
        public static string CombineNetworkPath(string path1, string path2)
        {
            if (path1.EndsWith("/") && path2.StartsWith("/"))
            {
                return path1 + path2.Substring(1);
            }
            else if (!path1.EndsWith("/") && !path2.StartsWith("/"))
            {
                return path1 + "/" + path2;
            }

            return path1 + path2;
        }

        /// <summary>
        /// Verify XML Element Exists.
        /// </summary>
        /// <param name="xmlElement">A XmlElement object which is the raw XML response from server.</param>
        /// <param name="elementName">The Name of the response element.</param>
        /// <returns>Whether the element exists in the response XML.</returns>
        public static bool ElementExists(XmlElement xmlElement, string elementName)
        {
            string content = xmlElement.OuterXml;
            if (content.Contains(elementName))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Get SOAP exception result.
        /// </summary>
        /// <param name="exception">SOAP exception.</param>
        /// <returns>The error code of the SOAP exception.</returns>
        public static string GetSoapExceptionErrorcode(SoapException exception)
        {
            string errorCode = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(exception.Detail.OuterXml);
            XmlNodeList nodeList = xmlDoc.GetElementsByTagName("errorcode");
            if (nodeList.Count >= 1)
            {
                errorCode = nodeList[0].InnerText;
            }
            else
            {
                errorCode = string.Empty;
            }

            return errorCode;
        }

        /// <summary>
        /// get transport type
        /// </summary>
        /// <returns>TransportProtocol.HTTP or TransportProtocol.HTTPS</returns>
        public static TransportProtocol GetTransportType()
        {
            if (Common.GetConfigurationPropertyValue("TransportType", baseSite).ToUpper().Equals("HTTP"))
            {
                return TransportProtocol.HTTP;
            }
            else
            {
                return TransportProtocol.HTTPS;
            }
        }
    }
}