//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// The protocol transport type which is used to transfer messages between the client and SUT.
    /// </summary>
    public enum TransportProtocol
    {
        /// <summary>
        /// The transport is SOAP over HTTP.
        /// </summary>
        HTTP,

        /// <summary>
        /// The transport is SOAP over HTTPS.
        /// </summary>
        HTTPS
    }

    /// <summary>
    /// The version of SUT.
    /// </summary>
    public enum SutVersion
    {
        /// <summary>
        /// The SUT is Microsoft SharePoint Foundation 2010 SP2.
        /// </summary>
        SharePointFoundation2010,

        /// <summary>
        /// The SUT is Microsoft SharePoint Foundation 2013 SP1.
        /// </summary>
        SharePointFoundation2013,

        /// <summary>
        /// The SUT is Microsoft SharePoint Server 2010 SP2.
        /// </summary>
        SharePointServer2010,

        /// <summary>
        /// The SUT is Microsoft SharePoint Server 2013 SP1.
        /// </summary>
        SharePointServer2013
    }

    /// <summary>
    /// Represent Result of Validation
    /// </summary>
    public enum ValidationResult
    {
        /// <summary>
        /// Indicate the validation is success.
        /// </summary>
        Success,

        /// <summary>
        /// Indicate the validation is error.
        /// </summary>
        Error,

        /// <summary>
        /// Indicate the validation is warning.
        /// </summary>
        Warning
    }
}