namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with WhoAmI Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with Versioning Sub-request.
        /// </summary>
        /// <param name="versioningSubResponse">Containing the VersioningSubResponse information</param>
        /// <param name="site">Instance of ITestSite</param>
        public static void ValidateFileOperationSubResponse(FileOperationSubResponseType fileOperationSubResponse, ITestSite site)
        {
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11124
            site.CaptureRequirement(
                "MS-FSSHTTP",
                11124,
                @"[In FileOperationSubResponseType]  
 <xs:complexType name=""FileOperationSubResponseType"">
   <xs:complexContent>
     <xs:extension base=""tns:SubResponseType"">
     </xs:extension>
   </xs:complexContent>
 </xs:complexType>");
        }
    }
}