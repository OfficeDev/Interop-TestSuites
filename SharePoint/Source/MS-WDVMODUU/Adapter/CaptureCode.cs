//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System.Net;

    /// <summary>
    /// Adapter requirements capture code for MS-WDVMODUU server role.
    /// </summary>
    public partial class MS_WDVMODUUAdapter
    {
        /// <summary>
        ///  Validate and capture the requirement about the transport, this protocol only support HTTP as transport now.
        /// </summary>
        /// <param name="httpResponse">The response for the HTTP request.</param>
        private void ValidateAndCaptureTransport(HttpWebResponse httpResponse)
        {
            this.Site.CaptureRequirementIfIsNotNull(httpResponse, 1, "[In Transport] Messages [in this protocol] are transported by using HTTP, as specified in [RFC2518] and [RFC2616].");
        }
    }
}