//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    /// <summary>
    /// Parameters for InitializeService operation.
    /// </summary>
    public class InitialPara
    {
        /// <summary>
        /// Gets or sets the URL of official file service.
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Gets or sets the property name of user in configuration file.
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// Gets or sets the property name of password in configuration file.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets the property name of domain in configuration file.
        /// </summary>
        public string Domain { get; set; }
    }
}
