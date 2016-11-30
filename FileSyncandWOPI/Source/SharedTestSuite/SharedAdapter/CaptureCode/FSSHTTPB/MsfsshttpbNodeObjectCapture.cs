namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This is the partial part of the class MsfsshttpbAdapterCapture for MS-FSSHTTPB transport.
    /// </summary>
    public partial class MsfsshttpbAdapterCapture
    {
        /// <summary>
        /// This method is used to verify the signature object related requirements.
        /// </summary>
        /// <param name="instance">Specify the signature object instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifySignatureObject(SignatureObject instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the SignatureObject related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type SignatureObject is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to verify the data size related requirements.
        /// </summary>
        /// <param name="instance">Specify the data size  instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyDataSizeObject(DataSizeObject instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the DataSizeObject related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type DataSizeObject is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to verify the intermediate node related requirements.
        /// </summary>
        /// <param name="instance">Specify the intermediate node instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyIntermediateNodeObject(LeafNodeObjectData instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the LeafNodeObjectData related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type LeafNodeObjectData is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R55
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPD",
                     55,
                     @"[In Intermediate Node Object Data] Intermediate Node Start (2 bytes): A 16-bit stream object header, as specified in [MS-FSSHTTPB] section 2.2.1.5.1, with a Header Type of 0x00, Compound of 0x1, Type of 0x1F, and Length of 0x00.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R56
            site.CaptureRequirementIfAreEqual<ushort>(
                     0xFC,
                     LittleEndianBitConverter.ToUInt16(instance.StreamObjectHeaderStart.SerializeToByteList().ToArray(), 0),
                     "MS-FSSHTTPD",
                     56,
                     @"[In Intermediate Node Object Data] Intermediate Node Start (2 bytes): The value of this field[Intermediate Node Start] MUST be 0x00FC.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R57
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.Signature.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPD",
                     57,
                     @"[In Intermediate Node Object Data] Signature Header (2 bytes): A 16-bit stream object header, as specified in [MS-FSSHTTPB] section 2.2.1.5.1, with a Header Type of  0x00, Compound of 0x0, Type of 0x21, and Length equal to the size of Signature Data.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R61
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.DataSize.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPD",
                     61,
                     @"[In Intermediate Node Object Data] Data Size Header (2 bytes): A 16-bit stream object header, as specified in [MS-FSSHTTPB] section 2.2.1.5.1, with a Header Type of  0x00, Compound of 0x0, Type of 0x22, and Length of 0x08 (the size, in bytes, of Data Size).");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8013
            site.CaptureRequirementIfAreEqual<uint>(
                     0x1110,
                     LittleEndianBitConverter.ToUInt16(instance.DataSize.StreamObjectHeaderStart.SerializeToByteList().ToArray(), 0),
                     "MS-FSSHTTPD",
                     8013,
                     @"[In Intermediate Node Object Data] Data Size Header (2 bytes): The value of this field[Data Size Header] MUST be 0x1110.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R63
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPD",
                     63,
                     @"[In Intermediate Node Object Data] Intermediate Node End (1 byte): An 8-bit stream object header end, as specified in [MS-FSSHTTPB] section 2.2.1.5.3, that specifies a stream object of type 0x1F.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8014
            site.CaptureRequirementIfAreEqual<byte>(
                     0x7D,
                     instance.StreamObjectHeaderEnd.SerializeToByteList()[0],
                     "MS-FSSHTTPD",
                     8014,
                     @"[In Intermediate Node Object Data] Intermediate Node End (1 byte):The value of this field[Intermediate Node End] MUST be 0x7D.");
        }

        /// <summary>
        /// This method is used to verify the root node related requirements.
        /// </summary>
        /// <param name="instance">Specify the intermediate node instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyRootNodeObject(IntermediateNodeObject instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the IntermediateNodeObject related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type IntermediateNodeObject is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPD_R37, if stream object start type is StreamObjectHeaderStart32bit. 
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     37,
                     @"[In Root Node Object Data] Root Node Start (2 bytes): A 16-bit stream object header, as specified in [MS-FSSHTTPB] section 2.2.1.5.1, with a Header Type of  0x00, Compound of 0x1, Type of 0x20, and Length of 0x00.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R38
            site.CaptureRequirementIfAreEqual<ushort>(
                     0x104,
                     LittleEndianBitConverter.ToUInt16(instance.StreamObjectHeaderStart.SerializeToByteList().ToArray(), 0),
                     "MS-FSSHTTPD",
                     38,
                     @"[In Root Node Object Data] Root Node Start (2 bytes): The value of this field[Root Node Start] MUST be 0x0104.");

            // Directly capture requirement MS-FSSHTTPD_R38, if all above asserts pass. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.DataSize.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPD",
                     43,
                     @"[In Root Node Object Data] Data Size Header (2 bytes): A 16-bit stream object header, as specified in [MS-FSSHTTPB] section 2.2.1.5.1, with a Header Type of  0x00, Compound of 0x0, Type of 0x22, and Length of 0x08 (the size, in bytes, of Data Size).");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8010
            site.CaptureRequirementIfAreEqual<uint>(
                     0x1110,
                     LittleEndianBitConverter.ToUInt16(instance.DataSize.StreamObjectHeaderStart.SerializeToByteList().ToArray(), 0),
                     "MS-FSSHTTPD",
                     8010,
                     @"[In Root Node Object Data] Data Size Header (2 bytes): The value of this field[Data Size Header] MUST be 0x1110.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R39
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.Signature.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPD",
                     39,
                     @"[In Root Node Object Data] Signature Header (2 bytes): A 16-bit stream object header, as specified in [MS-FSSHTTPB] section 2.2.1.5.1, with a Header Type of 0x00, Compound of 0x0, Type of 0x21, and Length equal to the size of Signature Data.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R45
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPD",
                     45,
                     @"[In Root Node Object Data] Root Node End (1 byte): An 8-bit stream object header end, as specified in [MS-FSSHTTPB] section 2.2.1.5.3, that specifies a stream object of type 0x20.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R46
            site.CaptureRequirementIfAreEqual<byte>(
                     0x81,
                     instance.StreamObjectHeaderEnd.SerializeToByteList()[0],
                     "MS-FSSHTTPD",
                     46,
                     @"[In Root Node Object Data] Root Node End (1 byte): The value of this field[Root Node End] MUST be 0x81.");
        }
    }
}