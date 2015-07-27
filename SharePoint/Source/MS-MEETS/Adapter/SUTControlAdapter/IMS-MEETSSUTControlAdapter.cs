//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Defines methods used by test suite to control the SUT
    /// </summary>
    public interface IMS_MEETSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Make sure there are no meeting workspaces under the specified site.
        /// </summary>
        /// <param name="siteUrl">The site Url</param>
        /// <returns>Returns if the method is succeed.</returns>
        [MethodHelp("Make sure there are no meeting workspaces under the specified site. Entering True in return field indicates that there are no meeting workspaces under the specified site. Entering False indicates that there are some meeting workspaces under the specified site.")]
        bool PrepareTestEnvironment(string siteUrl);
    }
}