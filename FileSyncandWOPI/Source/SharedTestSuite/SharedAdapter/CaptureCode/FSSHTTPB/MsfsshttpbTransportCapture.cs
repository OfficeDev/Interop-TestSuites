namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Reflection;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This is the partial part of the class MsfsshttpbAdapterCapture for MS-FSSHTTPB transport.
    /// </summary>
    public partial class MsfsshttpbAdapterCapture
    {
        /// <summary>
        /// This method is used to test Transport related adapter requirements.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyTransport(ITestSite site)
        {
            // Directly capture requirement MS-FSSHTTPB_R7, if embedded text in the cell sub response in the MS-FSSHTTP can be parsed successfully. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     7,
                     @"[In Transport] This protocol[MS-FSSHTTPB] uses File Synchronization via SOAP over HTTP protocol as specified in [MS-FSSHTTP].");

            // Directly capture requirement MS-FSSHTTPB_R8, if embedded text in the cell sub response in the MS-FSSHTTP can be parsed successfully. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     8,
                     @"[In Common Data Types] Unless noted otherwise, the following statements apply to this specification[MS-FSSHTTPB]:
                     Fields that consist of more than a single byte are specified in little-endian byte order.");
        }

        /// <summary>
        /// This method is used to invoke the corresponding type capture codes.
        /// </summary>
        /// <param name="type">Specify the type which will be captured requirements. </param>
        /// <param name="instance">Specify he instance of the type.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void InvokeCaptureMethod(System.Type type, object instance, ITestSite site)
        {
            // Find the capture function 
            MethodInfo captureMethod = typeof(MsfsshttpbAdapterCapture).GetMethod("Verify" + type.Name);
            if (captureMethod == null)
            {
                throw new InvalidOperationException(string.Format("Cannot find the function Verify{0} function in the MsfsshttpbAdapterCapture type.", type.Name));
            }

            captureMethod.Invoke(this, new object[] { instance, site });
        }
    }
}