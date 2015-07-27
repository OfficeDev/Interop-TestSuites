//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class which is used to initialize the SharedContext <see cref="SharedContext"/> instance for MS-FSSHTTP, MS-FSSHTTPB and MS-FSSHTTPD.
    /// </summary>
    public static class SharedContextUtils
    {
        /// <summary>
        /// Initialize the SharedContext <see cref="SharedContext"/> based on the specified request file URL, user name, password and domain for the MS-FSSHTTP test purpose.
        /// </summary>
        /// <param name="userName">Specify the user name.</param>
        /// <param name="password">Specify the password.</param>
        /// <param name="domain">Specify the domain.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void InitializeSharedContextForFSSHTTP(string userName, string password, string domain, ITestSite site)
        {
            SharedContext context = SharedContext.Current;

            if (string.Equals("HTTP", Common.GetConfigurationPropertyValue("TransportType", site), System.StringComparison.OrdinalIgnoreCase))
            {
                context.TargetUrl = Common.GetConfigurationPropertyValue("HttpTargetServiceUrl", site);
                context.EndpointConfigurationName = Common.GetConfigurationPropertyValue("HttpEndPointName", site);
            }
            else
            {
                context.TargetUrl = Common.GetConfigurationPropertyValue("HttpsTargetServiceUrl", site);
                context.EndpointConfigurationName = Common.GetConfigurationPropertyValue("HttpsEndPointName", site);
            }

            context.Site = site;
            context.OperationType = OperationType.FSSHTTPCellStorageRequest;
            context.UserName = userName;
            context.Password = password;
            context.Domain = domain;
        }
    }
}
