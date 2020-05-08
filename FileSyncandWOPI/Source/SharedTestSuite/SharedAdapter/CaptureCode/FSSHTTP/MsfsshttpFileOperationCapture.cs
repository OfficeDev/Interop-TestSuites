namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with FileOperation Sub-request.
    /// </summary>
    public sealed partial class MsfsshttpAdapterCapture
    {
        /// <summary>
        /// Capture requirements related with FileOperation Sub-request.
        /// </summary>
        /// <param name="fileOperationSubResponse">Containing the FileOperationSubResponse information</param>
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
        <xs:sequence minOccurs=""0"" maxOccurs=""1"">
       < xs:element name = ""SubResponseData"" />
       </ xs:sequence >

     </ xs:extension>
   </xs:complexContent>
 </xs:complexType>");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2349
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(FileOperationSubResponseType),
                     fileOperationSubResponse.GetType(),
                     "MS-FSSHTTP",
                     2349,
                     @"[FileOperation SubRequest]The protocol server responds with a FileOperation SubResponse message, which is of type FileOperationSubResponseType as specified in section 2.3.1.35. ");

        }
    }
}