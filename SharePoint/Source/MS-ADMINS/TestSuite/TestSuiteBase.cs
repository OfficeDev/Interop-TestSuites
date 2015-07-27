//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ADMINS
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains all helper methods used in test case level.
    /// </summary>
    public static class TestSuiteBase
    {
        /// <summary>
        /// A random generator using current time seeds.
        /// </summary>
        private static Random random;

        /// <summary>
        /// A ITestSite instance
        /// </summary>
        private static ITestSite testSite;

        /// <summary>
        /// A method used to initialize the TestSuiteBase with specified ITestSite instance.
        /// </summary>
        /// <param name="site">A parameter represents ITestSite instance.</param>
        public static void Initialize(ITestSite site)
        {
            testSite = site;
        }

        /// <summary>
        /// A method is used to get a unique site title.
        /// </summary>
        /// <returns>A return value represents the unique site title that is combined with the object name and time stamp</returns>
        public static string GenerateUniqueSiteTitle()
        {
            return Common.GenerateResourceName(testSite, "SiteTitle");
        }

        /// <summary>
        /// A method is used to generate owner name.
        /// </summary>
        /// <returns>A return value represents the unique owner name that is combined with the object name and time stamp</returns>
        public static string GenerateUniqueOwnerName()
        {
            return Common.GenerateResourceName(testSite, "OwnerName");
        }

        /// <summary>
        /// A method is used to generate portal name.
        /// </summary>
        /// <returns>A return value represents the unique portal name that is combined with the Object name and time stamp</returns>
        public static string GenerateUniquePortalName()
        {
            return Common.GenerateResourceName(testSite, "PortalName");
        }

        /// <summary>
        /// This method is used to random generate an integer in the specified range.
        /// </summary>
        /// <param name="minValue">The inclusive lower bound of the random number returned.</param>
        /// <param name="maxValue">The exclusive upper bound of the random number returned.</param>
        /// <returns>A 32-bit signed integer greater than or equal to minValue and less than maxValue</returns>
        public static int GenerateRandomNumber(int minValue, int maxValue)
        {
            random = new Random((int)DateTime.Now.Ticks);
            return random.Next(minValue, maxValue);
        }

        /// <summary>
        /// This method is used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>Returns the random generated string.</returns>
        public static string GenerateRandomString(int size)
        {
            random = new Random((int)DateTime.Now.Ticks);
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
        /// This method is used to generate random e-mail with the specified size.
        /// </summary>
        /// <param name="size">A parameter represents the generated e-mail size.</param>
        /// <returns>Returns the random generated e-mail.</returns>
        public static string GenerateEmail(int size)
        {
            string returnValue = string.Empty;
            string suffix = "@contoso.com";
            if (size > suffix.Length)
            {
                returnValue = GenerateRandomString(size - suffix.Length) + suffix;
            }
            else
            {
                returnValue = "a" + suffix;
            }

            return returnValue;
        }

        /// <summary>
        /// This method is used to generate random portal url with the specified size.
        /// </summary>
        /// <param name="size">A parameter represents the generated portal url size.</param>
        /// <returns>Returns the random generated portal url.</returns>
        public static string GeneratePortalUrl(int size)
        {
            string returnValue = string.Empty;
            string suffix = ".com";
            string prefix = "www.";
            if (size > prefix.Length + suffix.Length)
            {
                returnValue = prefix + GenerateRandomString(size - prefix.Length - suffix.Length) + suffix;
            }
            else
            {
                returnValue = prefix + "a" + suffix;
            }

            return returnValue;
        }

        /// <summary>
        /// This method is used to generate random url without port number which includes [TransportType]://[SUTComputerName].
        /// </summary>
        /// <param name="size">A parameter represents the random string size.</param>
        /// <returns>Returns the random generated url.</returns>
        public static string GenerateUrlWithoutPort(int size)
        {
            string returnValue = string.Empty;
            string subSite = "/sites/";
            string prefix = Common.GetConfigurationPropertyValue("TransportType", testSite) + "://" + Common.GetConfigurationPropertyValue("SUTComputerName", testSite);
            if (size > subSite.Length)
            {
                returnValue = prefix + subSite + GenerateRandomString(size - subSite.Length);
            }
            else 
            {
                returnValue = prefix + subSite + "a";
            }

            return returnValue;
        }

        /// <summary>
        /// This method is used to generate a url prefix with port number. The format should be "HTTP://[SUTComputerName]:[HTTPPortNumber]/sites/" or "HTTPS://[SUTComputerName]:[HTTPSPortNumber]/sites/".
        /// </summary>
        /// <returns>Returns the generated url prefix.</returns>
        public static string GenerateUrlPrefixWithPortNumber()
        {
            string returnValue = string.Empty;
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", testSite);
            switch (transport)
            {
                case TransportProtocol.HTTP:
                    returnValue = Common.GetConfigurationPropertyValue("UrlWithHTTPPortNumber", testSite);
                    break;
                case TransportProtocol.HTTPS:
                    returnValue = Common.GetConfigurationPropertyValue("UrlWithHTTPSPortNumber", testSite);
                    break;
                default:
                    testSite.Assume.Fail("Unaccepted transport type: {0}.", transport);
                    break;
            }
     
            return returnValue;
        }

        /// <summary>
        /// This method is used to generate a url prefix with admin port number. The format should be "HTTP://[SUTComputerName]:[AdminHTTPPortNumber]/sites/" or "HTTPS://[SUTComputerName]:[AdminHTTPSPortNumber]/sites/".
        /// </summary>
        /// <returns>Returns the generated url prefix.</returns>
        public static string GenerateUrlPrefixWithAdminPort()
        {
            string returnValue = string.Empty;
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", testSite);
            switch (transport)
            {
                case TransportProtocol.HTTP:
                    returnValue = Common.GetConfigurationPropertyValue("UrlWithAdminHTTPPort", testSite);
                    break;
                case TransportProtocol.HTTPS:
                    returnValue = Common.GetConfigurationPropertyValue("UrlWithAdminHTTPSPort", testSite);
                    break;
                default:
                    testSite.Assume.Fail("Unaccepted transport type: {0}.", transport);
                    break;
            }

            return returnValue;
        }
    }
}