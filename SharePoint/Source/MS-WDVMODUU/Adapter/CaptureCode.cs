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