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

    /// <summary>
    /// A class which is used to provide help methods for configuration files.
    /// </summary>
    public static class ConfigurationFileHelper
    {
        /// <summary>
        /// Merge the common configuration and should/may configuration file.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        public static void MergeConfigurationFile(TestTools.ITestSite site)
        {
            site.DefaultProtocolDocShortName = "MS-FSSHTTP-FSSHTTPB";

            // Get the name of common configuration file.
            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", site);

            // Merge the common configuration.
            Common.MergeGlobalConfig(commonConfigFileName, site);

            // Merge the MS_FSSHTTP-FSSHTTPB should/may configuration file.
            Common.MergeSHOULDMAYConfig(site);
        }
    }
}
