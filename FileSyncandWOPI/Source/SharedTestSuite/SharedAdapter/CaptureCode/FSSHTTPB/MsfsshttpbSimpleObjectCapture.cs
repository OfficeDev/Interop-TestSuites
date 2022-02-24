namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This is the partial part of the class MsfsshttpbAdapterCapture for MS-FSSHTTPB simple object.
    /// </summary>
    public partial class MsfsshttpbAdapterCapture
    {
        /// <summary>
        /// This method is used to test Compact Unsigned 64-bit Integer related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCompact64bitInt(Compact64bitInt instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Compact Unsigned 64-bit Integer related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type Compact64bitInt is null due to parsing error or type casting error.");
            }

            bool isVerifyR10 = instance.DecodedValue <= (ulong)18446744073709551615;

            site.Log.Add(
                LogEntryKind.Debug,
                "The Compact Unsigned 64-bit Integer actual value is {0}, which should be less or equal to 18446744073709551615 for requirement MS-FSSHTTPB_R10.",
                instance.DecodedValue);

            // Directly capture requirement MS-FSSHTTPB_R10, if there are no parsing errors. 
            site.CaptureRequirementIfIsTrue(
                     isVerifyR10,
                     "MS-FSSHTTPB",
                     10,
                     @"[In Compact Unsigned 64-bit Integer] A variable-width encoding of unsigned integers less than 18446744073709551616.");

            // Based on the different compact unsigned 64-bit type to capture the corresponding requirements.
            switch (instance.Type)
            {
                case Compact64bitInt.CompactUintNullType:
                    this.VerifyCompactUintZero(instance, site);
                    break;

                case Compact64bitInt.CompactUint7bitType:
                    this.VerifyCompactUint7BitValues(instance, site);
                    break;

                case Compact64bitInt.CompactUint14bitType:
                    this.VerifyCompactUint14BitValues(instance, site);
                    break;

                case Compact64bitInt.CompactUint21bitType:
                    this.VerifyCompactUint21BitValues(instance, site);
                    break;

                case Compact64bitInt.CompactUint28bitType:
                    this.VerifyCompactUint28BitValues(instance, site);
                    break;

                case Compact64bitInt.CompactUint35bitType:
                    this.VerifyCompactUint35BitValues(instance, site);
                    break;

                case Compact64bitInt.CompactUint42bitType:
                    this.VerifyCompactUint42BitValues(instance, site);
                    break;

                case Compact64bitInt.CompactUint49bitType:
                    this.VerifyCompactUint49BitValues(instance, site);
                    break;

                case Compact64bitInt.CompactUint64bitType:
                    this.VerifyCompactUint64BitValues(instance, site);
                    break;

                default:
                    site.Assert.Fail("Unsupported Compact64bitInt type value " + (int)instance.Type);
                    break;
            }
        }

        /// <summary>
        /// This method is used to test File Chunk Reference related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyFileChunk(FileChunk instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the File Chunk Reference related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type FileChunk is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R48, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     48,
                     @"[In File Chunk Reference] Start (variable): A compact unsigned 64-bit integer (section 2.2.1.1 ) that specifies the byte-offset within the file of the beginning of the file chunk.");

            // Directly capture requirement MS-FSSHTTPB_R49, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     49,
                     @"[In File Chunk Reference] Length (variable): A compact unsigned 64-bit integer that specifies the count of bytes included in the file chunk.");
        }

        /// <summary>
        /// This method is used to test Binary Item related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyBinaryItem(BinaryItem instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Binary Item related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type BinaryItem is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R51, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     51,
                     @"[In Binary Item] Length (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the count of bytes of Content of the item.");

            // Directly capture requirement MS-FSSHTTPB_R52, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     52,
                     @"[In Binary Item] Content (variable): A byte stream that specifies the data for the item.");
        }

        /// <summary>
        /// This method is used to test String Item related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStringItem(StringItem instance, ITestSite site)
        {
            // There is not capture codes for this structure, this structure is not used by the protocol.
            // If the instance is not null and there are no parsing errors, then the String Item related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StringItem is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R54, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     54,
                     @"[In String Item] Count (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the count of characters in the string.");

            // Directly capture requirement MS-FSSHTTPB_R56, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     56,
                     @"[In String Item] Content (variable): It[Content (variable)] MUST NOT be null-terminated.");
        }

        /// <summary>
        /// This method is used to test String Item Array related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStringItemArray(StringItem instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the String Item Array related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StringItemArray is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R500, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     500,
                     @"[In String Item Array][String Item Array]The length and content of an array of String Items as specified in section 2.2.1.4.");

            // Directly capture requirement MS-FSSHTTPB_R502, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     502,
                     @"[In String Item Array] Count (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the count of String Items in the array.");

            // Directly capture requirement MS-FSSHTTPB_R503, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     503,
                     @"[In String Item Array] Content (variable): A String Item Array that specifies an array of items.");
        }

        /// <summary>
        /// This method is used to test 16-bit Stream Object Header Start related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStreamObjectHeaderStart16bit(StreamObjectHeaderStart16bit instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the 16-bit Stream Object Header Start related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StreamObjectHeaderStart16bit is null due to parsing error or type casting error.");
            }

            // Verify the common stream object header part related requirements.
            this.VerifyStreamObjectHeader(site);

            // Directly capture requirement MS-FSSHTTPB_R59, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     59,
                     @"[In 16-bit Stream Object Header Start] A - Header Type (2-bit): A flag that specifies a 16-bit stream object start.");

            // Directly capture requirement MS-FSSHTTPB_R60, if there are no parsing errors. 
            site.CaptureRequirementIfAreEqual<int>(
                     0x0,
                     instance.HeaderType,
                     "MS-FSSHTTPB",
                     60,
                     @"[In 16-bit Stream Object Header Start] This[A - Header Type(2-bit)] MUST be set to 0x0.");

            // Directly capture requirement MS-FSSHTTPB_R64, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     64,
                     @"[In 16-bit Stream Object Header Start] Type (6-bits): A 6-bit unsigned integer that specifies the stream object type [(see the following table for possible values)].");

            // Directly capture requirement MS-FSSHTTPB_R65, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     65,
                     @"[In 16-bit Stream Object Header Start] Length (7-bits): A 7-bit unsigned integer that specifies length in bytes for additional data (if any) before the next Stream Object Header start or Stream Object Header end.");
        }

        /// <summary>
        /// This method is used to test 32-bit Stream Object Header Start related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStreamObjectHeaderStart32bit(StreamObjectHeaderStart32bit instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the 32-bit Stream Object Header Start related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StreamObjectHeaderStart32bit is null due to parsing error or type casting error.");
            }

            // Verify the common stream object header part related requirements.
            this.VerifyStreamObjectHeader(site);

            // If the type value in the 16-bit stream object header, but actual the instance is 32-bit stream object header instance. And when the length is larger than 127 then capture the requirement MS-FSSHTTPB_R66.
            if ((int)instance.Type <= 0x3F && instance.Length > 127)
            {
                // Directly capture requirement MS-FSSHTTPB_R66.
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         66,
                         @"[In 16-bit Stream Object Header Start] If the length is more than 127 bytes, a 32-bit Stream Object Header start (section 2.2.1.5.2) MUST be used.");
            }

            // Directly capture requirement MS-FSSHTTPB_R97, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     97,
                     @"[In 32-bit Stream Object Header Start] A - Header Type (2-bit): A flag that specifies a 32-bit stream object start.");

            // Capture requirement MS-FSSHTTPB_R98, if there are no parsing errors. 
            site.CaptureRequirementIfAreEqual<int>(
                     0x2,
                     instance.HeaderType,
                     "MS-FSSHTTPB",
                     98,
                     @"[In 32-bit Stream Object Header Start] This[A - Header Type (2-bit)] MUST be set to 0x2.");

            // Directly capture requirement MS-FSSHTTPB_R102, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     102,
                     @"[In 32-bit Stream Object Header Start] Type (14-bits): A 14-bit unsigned integer that specifies the stream object type (see the following table for possible values).");

            // Directly capture requirement MS-FSSHTTPB_R103, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     103,
                     @"[In 32-bit Stream Object Header Start] Length (15-bits): A 15-bit unsigned integer that specifies the length in bytes for additional data (if any) before the next Stream Object Header start or Stream Object Header end.");

            if (instance.Length <= 32766)
            {
                site.Assert.IsNull(
                            instance.LargeLength,
                            "When the length less or equal to the value 32766, the large length MUST not be specified.");

                // Directly capture requirement MS-FSSHTTPB_R933, if the above assert validate.
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         933,
                         @"[In 32-bit Stream Object Header Start] Large Length (variable): MUST NOT be specified if the Length field contains any other value than 32767.");
            }

            if (instance.LargeLength != null)
            {
                // Directly capture requirement MS-FSSHTTPB_R931, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         931,
                         @"[In 32-bit Stream Object Header Start] Length (15-bits): If the length is more than 32766, a Large Length field MUST be specified.");

                // Directly capture requirement MS-FSSHTTPB_R2095, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         2095,
                         @"[In 32-bit Stream Object Header Start] Large Length (variable): An optional compact unsigned 64-bit integer (section 2.2.1.1) that specifies the length in bytes for additional data (if any).");

                // Capture requirement MS-FSSHTTPB_R932 and MS-FSSHTTPB_R930, if the length equals to 32767. 
                site.CaptureRequirementIfAreEqual<int>(
                         32767,
                         instance.Length,
                         "MS-FSSHTTPB",
                         932,
                         @"[In 32-bit Stream Object Header Start] Large Length (variable): This field[Large Length (variable)] MUST be specified if the Length field contains 32767.");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R930
                site.CaptureRequirementIfAreEqual<int>(
                         32767,
                         instance.Length,
                         "MS-FSSHTTPB",
                         930,
                         @"[In 32-bit Stream Object Header Start] Length (15-bits): If the length is more than 32766, this field [Length (15-bits)] MUST specify 32767.");
            }
        }

        /// <summary>
        /// This method is used to test 8-bit Stream Object Header End related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStreamObjectHeaderEnd8bit(StreamObjectHeaderEnd8bit instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the 8-bit Stream Object Header End related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StreamObjectHeaderEnd8bit is null due to parsing error or type casting error.");
            }

            // Verify the common stream object header part related requirements.
            this.VerifyStreamObjectHeader(site);

            // Directly capture requirement MS-FSSHTTPB_R147, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     147,
                     @"[In 8-bit Stream Object Header End] A – Header Type (2-bit): A flag that specifies an 8-bit stream object end.");

            // Directly capture requirement MS-FSSHTTPB_R148, if there are no parsing errors. 
            // The value equals to 0x01 can be guaranteed by the parsing progress.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     148,
                     @"[In 8-bit Stream Object Header End] This[A - Header Type (2-bit)] MUST be set to 0x1.");

            // Directly capture requirement MS-FSSHTTPB_R149, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     149,
                     @"[In 8-bit Stream Object Header End] Type (6-bits): A 6-bit unsigned integer that specifies the stream object type (see the following table for possible values).");

            int streamEndType = Convert.ToInt32(instance.Type);
            bool flag = streamEndType == 0x01 ||
                        streamEndType == 0x10 ||
                        streamEndType == 0x14 ||
                        streamEndType == 0x15 ||
                        streamEndType == 0x1D ||
                        streamEndType == 0x1E ||
                        streamEndType == 0x1F ||
                        streamEndType == 0x20 ||
                        streamEndType == 0x29 ||
                        streamEndType == 0x2D;

            site.Assert.IsTrue(
                         flag,
                         "Actual 8-bit stream object end type value is {0}, expect the value [0x01, 0x10, 0x14, 0x15, 0x1D, 0x1E, 0x29, 0x2D] for requirement MS-FSSHTTPB_R162.");

            // Directly capture requirement MS-FSSHTTPB_R162, if the above assertion was validated.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     162,
                     @"[In 8-bit Stream Object Header End] The following table lists the possible stream object types:[Its value must be one of 0x01, 0x10, 0x14, 0x15, 0x1D, 0x1E, 0x29, 0x2D].");
        }

        /// <summary>
        /// This method is used to test 16-bit Stream Object Header End related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStreamObjectHeaderEnd16bit(StreamObjectHeaderEnd16bit instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the 16-bit Stream Object Header End related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StreamObjectHeaderEnd16bit is null due to parsing error or type casting error.");
            }

            // Verify the common stream object header part related requirements.
            this.VerifyStreamObjectHeader(site);

            // Directly capture requirement MS-FSSHTTPB_R158, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     158,
                     @"[In 16-bit Stream Object Header End] A 16-bit header for a compound object that indicates the end of a stream object has the following format: [A - Header Type (2-bits), Type (14-bits)].");

            // Directly capture requirement MS-FSSHTTPB_R159, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     159,
                     @"[In 16-bit Stream Object Header End] A - Header Type (2-bit): A flag that specifies a 16-bit stream object end.");

            // Directly capture requirement MS-FSSHTTPB_R160, if there are no parsing errors. 
            // The value equals to 0x03 can be guaranteed by the parsing progress.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     160,
                     @"[In 16-bit Stream Object Header End] This[A - Header Type (2-bit)] MUST be set to 0x3.");

            // Directly capture requirement MS-FSSHTTPB_R161, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     161,
                     @"[In 16-bit Stream Object Header End] Type (14-bits): A 14-bit unsigned integer that specifies the stream object type (see the following table for possible values).");

            int streamEndType = Convert.ToInt32(instance.Type);
            bool flag = streamEndType == 0x40 ||
            streamEndType == 0x41 ||
            streamEndType == 0x42 ||
            streamEndType == 0x43 ||
            streamEndType == 0x44 ||
            streamEndType == 0x46 ||
            streamEndType == 0x47 ||
            streamEndType == 0x4D ||
            streamEndType == 0x51 ||
            streamEndType == 0x5D ||
            streamEndType == 0x62 ||
            streamEndType == 0x6B ||
            streamEndType == 0x083 ||
            streamEndType == 0x79 ||
            streamEndType == 0x7A;

            site.Assert.IsTrue(
                        flag,
                        "Actual 16-bit stream object end type value is {0}, expect the value [0x40, 0x41, 0x42, 0x43, 0x44, 0x46, 0x47, 0x4D, 0x51, 0x5D, 0x62, 0x6B, ox79, 0x083] for requirement MS-FSSHTTPB_R163.",
                        streamEndType);

            // Directly capture requirement MS-FSSHTTPB_R163, if the above assertion was validated.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     163,
                     @"[In 16-bit Stream Object Header End] The following table lists the possible stream object types:[Its value must be one of 0x40, 0x41, 0x42, 0x43, 0x44, 0x46, 0x47, 0x4D, 0x51, 0x5D, 0x62, 0x6B, 0x079].");
        }

        /// <summary>
        /// This method is used to test Extended GUID related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyExGuid(ExGuid instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Extended GUID related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ExGuid is null due to parsing error or type casting error.");
            }

            switch (instance.Type)
            {
                case ExGuid.ExtendedGUIDNullType:
                    this.VerifyExtendedGUIDNullValue(instance, site);
                    break;

                case ExGuid.ExtendedGUID5BitUintType:
                    this.VerifyExtendedGUID5BitUintValue(instance, site);
                    break;

                case ExGuid.ExtendedGUID10BitUintType:
                    this.VerifyExtendedGUID10BitUintValue(instance, site);
                    break;

                case ExGuid.ExtendedGUID17BitUintType:
                    this.VerifyExtendedGUID17BitUintValue(instance, site);
                    break;

                case ExGuid.ExtendedGUID32BitUintType:
                    this.VerifyExtendedGUID32BitUintValue(instance, site);
                    break;

                default:
                    site.Assert.Fail("Unsupported ExGuid type value " + instance.Type);
                    break;
            }
        }

        /// <summary>
        /// This method is used to test Extended GUID Array related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyExGUIDArray(ExGUIDArray instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Extended GUID Array related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type [$TypeName] is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R216, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     216,
                     @"[In Extended GUID Array] Count (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the count of Extended GUIDs in the array.");

            // Directly capture requirement MS-FSSHTTPB_R221, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     221,
                     @"[In Extended GUID Array] Content (variable): An Extended GUID Array that specifies an array of items.");
        }

        /// <summary>
        /// This method is used to test Serial Number related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifySerialNumber(SerialNumber instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Serial Number related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type SerialNumber is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R4006, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     4006,
                     @"[In Serial Number] The Serial Number of a data element can be created by the creator of the data element, but the server is authoritative and can replace a serial number that is created by the client.");

            switch (instance.Type)
            {
                case 0:
                    this.VerifySerialNumberNullValue(instance, site);
                    break;

                case 128:
                    this.VerifySerialNumber64BitUintValue(instance, site);
                    break;

                default:
                    site.Assert.Fail("Unsupported serial number type value " + (int)instance.Type);
                    break;
            }
        }

        /// <summary>
        /// This method is used to test Cell ID related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCellID(CellID instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Cell ID related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type CellID is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R234, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     234,
                     @"[In Cell ID] EXGUID1 (variable): An Extended GUID that specifies the first cell identifier.");

            // Directly capture requirement MS-FSSHTTPB_R235, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     235,
                     @"[In Cell ID] EXGUID2 (variable): An Extended GUID that specifies the second cell identifier.");
        }

        /// <summary>
        /// This method is used to test Cell ID Array related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCellIDArray(CellIDArray instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Cell ID Array related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type CellIDArray is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R236, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     236,
                     @"[In Cell ID Array][Cell ID Array is] The count and content of an array of Cell IDs.");

            // Directly capture requirement MS-FSSHTTPB_R237, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     237,
                     @"[In Cell ID Array] Count (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the count of Cell IDs in the array.");

            // Directly capture requirement MS-FSSHTTPB_R238, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     238,
                     @"[In Cell ID Array] Content (variable): A Cell ID Array that specifies an array of cells.");
        }

        #region Private method
        /// <summary>
        /// This method is used to test Serial Number Null Value related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifySerialNumberNullValue(SerialNumber instance, ITestSite site)
        {
            // All the serial number null values related requirements can be located here.
            site.Log.Add(LogEntryKind.Debug, "Runs for VerifySerialNumberNullValue operation with the value {0}.", instance.Value);
        }

        /// <summary>
        /// This method is used to test Serial Number 64 Bit Uint Value related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifySerialNumber64BitUintValue(SerialNumber instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Serial Number 64 Bit Uint Value related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type SerialNumber is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R227, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     227,
                     @"[In Serial Number 64 Bit Uint Value] A 25-byte encoding of the Serial Number.");

            // Directly capture requirement MS-FSSHTTPB_R228, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     228,
                     @"[In Serial Number 64 Bit Uint Value] Type (1 byte): An unsigned integer that specifies the type.");

            // Directly capture requirement MS-FSSHTTPB_R229, if there are no parsing errors. 
            site.CaptureRequirementIfAreEqual<uint>(
                     128,
                     instance.Type,
                     "MS-FSSHTTPB",
                     229,
                     @"[In Serial Number 64 Bit Uint Value] Type (1 byte): MUST be 128.");

            // Directly capture requirement MS-FSSHTTPB_R230, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     230,
                     @"[In Serial Number 64 Bit Uint Value] GUID (16 bytes): A GUID that specifies the item.");

            // Directly capture requirement MS-FSSHTTPB_R231, if there are no parsing errors. 
            site.CaptureRequirementIfAreNotEqual<System.Guid>(
                     System.Guid.Empty,
                     instance.GUID,
                     "MS-FSSHTTPB",
                     231,
                     @"[In Serial Number 64 Bit Uint Value] GUID (16 bytes): It[GUID (16 bytes)] MUST NOT be ""{00000000-0000-0000-0000-000000000000}"".");

            // Directly capture requirement MS-FSSHTTPB_R232, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     232,
                     @"[In Serial Number 64 Bit Uint Value] Value (8 bytes): An unsigned integer that specifies the value of the Serial Number.");
        }

        /// <summary>
        /// This method is used to test Extended GUID Null Value related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyExtendedGUIDNullValue(ExGuid instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Extended GUID Null Value related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ExGuid is null due to parsing error or type casting error.");
            }

            site.Assert.AreEqual<System.Guid>(
                System.Guid.Empty,
                instance.GUID,
                "The GUID part of the extended GUID null value MUST be {00000000-0000-0000-0000-000000000000}.");

            site.Assert.AreEqual<uint>(
                        0,
                        instance.Value,
                        "The unsigned integer part of the extended GUID null value MUST be zero.");

            // Directly capture requirement MS-FSSHTTPB_R188, if there are no parsing errors and the above two asserts are passed.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     188,
                     @"[In Extended GUID Null Value] A 1-byte encoding of the Extended GUID when the GUID part is {00000000-0000-0000-0000-000000000000} and the unsigned integer part is zero.");

            // Directly capture requirement MS-FSSHTTPB_R190, if the type is zero. 
            site.CaptureRequirementIfAreEqual<uint>(
                     0,
                     instance.Type,
                     "MS-FSSHTTPB",
                     190,
                     @"[In Extended GUID Null Value] Type (8 bits): MUST be zero.");

            // Directly capture requirement MS-FSSHTTPB_R189, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     189,
                     @"[In Extended GUID Null Value] Type (8 bits): An unsigned integer that specifies a null GUID.");
        }

        /// <summary>
        /// This method is used to test Extended GUID 5 Bit Uint Value related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyExtendedGUID5BitUintValue(ExGuid instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Extended GUID 5 Bit Uint Value related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type [$TypeName] is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R191, if the value between 0x0 and 0x1F.
            bool isVerifyR191 = instance.Value >= 0x0 && instance.Value <= 0x1F;
            site.Log.Add(
                    LogEntryKind.Debug,
                    "The Extended GUID actual unsigned value is {0}, which should be between 0x0 and 0x1F for requirement MS-FSSHTTPB_R191.",
                    instance.Value);
            site.CaptureRequirementIfIsTrue(
                     isVerifyR191,
                     "MS-FSSHTTPB",
                     191,
                     @"[In Extended GUID 5 Bit Uint Value] A 17-byte encoding of the Extended GUID when the integer part ranges from 0x0 through 0x1F.");

            // Directly capture requirement MS-FSSHTTPB_R192, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     192,
                     @"[In Extended GUID 5 Bit Uint Value] Type (3 bits): An unsigned integer that specifies the type.");

            // Capture requirement MS-FSSHTTPB_R193, if the type value equals to 4. 
            site.CaptureRequirementIfAreEqual<uint>(
                     4,
                     instance.Type,
                     "MS-FSSHTTPB",
                     193,
                     @"[In Extended GUID 5 Bit Uint Value] Type (3 bits): MUST be 4.");

            // Directly capture requirement MS-FSSHTTPB_R194, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     194,
                     @"[In Extended GUID 5 Bit Uint Value] Value (5 bits): An unsigned integer that specifies the value.");

            // Directly capture requirement MS-FSSHTTPB_R195, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     195,
                     @"[In Extended GUID 5 Bit Uint Value] GUID (16 bytes): A GUID that specifies the item.");

            // Capture requirement MS-FSSHTTPB_R196, if the GUID does not equal to {00000000-0000-0000-0000-000000000000}.
            site.CaptureRequirementIfAreNotEqual<System.Guid>(
                     System.Guid.Empty,
                     instance.GUID,
                     "MS-FSSHTTPB",
                     196,
                     @"[In Extended GUID 5 Bit Uint Value] GUID (16 bytes): This[GUID (16 bytes)] MUST NOT be ""{00000000-0000-0000-0000-000000000000}"".");
        }

        /// <summary>
        /// This method is used to test Extended GUID 10 Bit Uint Value related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyExtendedGUID10BitUintValue(ExGuid instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Extended GUID 10 Bit Uint Value related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ExGuid is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R198, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     198,
                     @"[In Extended GUID 10 Bit Uint Value] Type (6 bits): An unsigned integer that specifies the type.");

            // Directly capture requirement MS-FSSHTTPB_R200, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     200,
                     @"[In Extended GUID 10 Bit Uint Value] Value (10 bits): An unsigned integer that specifies the value.");

            // Directly capture requirement MS-FSSHTTPB_R201, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     201,
                     @"[In Extended GUID 10 Bit Uint Value] GUID (16 bytes): A GUID that specifies the item.");
        }

        /// <summary>
        /// This method is used to test Extended GUID 17 Bit Uint Value related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyExtendedGUID17BitUintValue(ExGuid instance, ITestSite site)
        {
            // All the extended GUID 17 bits values related requirements can be located here.
            site.Log.Add(LogEntryKind.Debug, "Runs for VerifyExtendedGUID17BitUintValue operation with the value {0}.", instance.Value);
        }

        /// <summary>
        /// This method is used to test Extended GUID 32 Bit Uint Value related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyExtendedGUID32BitUintValue(ExGuid instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Extended GUID 32 Bit Uint Value related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ExGuid is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R209, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     209,
                     @"[In Extended GUID 32 Bit Uint Value][Extended GUID 32 Bit Uint Value is] A 21-byte encoding of the Extended GUID when the integer part ranges from 0x20000 through 0xFFFFFFFF.");

            // Directly capture requirement MS-FSSHTTPB_R210, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     210,
                     @"[In Extended GUID 32 Bit Uint Value] Type (1 byte): An unsigned integer that specifies the type.");

            // Directly capture requirement MS-FSSHTTPB_R211, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     211,
                     @"[In Extended GUID 32 Bit Uint Value] Type (1 byte): MUST be 128.");

            // Directly capture requirement MS-FSSHTTPB_R212, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     212,
                     @"[In Extended GUID 32 Bit Uint Value] Value (4 bytes): An unsigned integer that specifies the value.");

            // Directly capture requirement MS-FSSHTTPB_R213, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     213,
                     @"[In Extended GUID 32 Bit Uint Value] GUID (16 bytes): A GUID that specifies the item.");

            // Directly capture requirement MS-FSSHTTPB_R214, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     214,
                     @"[In Extended GUID 32 Bit Uint Value] GUID (16 bytes): This[GUID (16 bytes)] MUST NOT be ""{00000000-0000-0000-0000-000000000000}"".");
        }

        /// <summary>
        /// This method is used to test Compact Uint Zero related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUintZero(Compact64bitInt instance, ITestSite site)
        {
            // Directly capture requirement MS-FSSHTTPB_R12, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     12,
                     @"[In Compact Uint Zero] A 1-byte encoding of the value zero.");

            // Capture requirement MS-FSSHTTPB_R14, if representing value 0. 
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     instance.DecodedValue,
                     "MS-FSSHTTPB",
                     14,
                     @"[In Compact Unit Zero] Uint(8 bits): MUST be zero.");
        }

        /// <summary>
        /// This method is used to test Compact Uint 7 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint7BitValues(Compact64bitInt instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Compact Uint 7 bit values related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type Compact64bitInt is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R16, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     16,
                     @"[In Compact Uint 7 bit values] A – Type (1 bit): A flag that specifies this format from all other formats of a compact unsigned 64-bit integer (section 2.2.1.1).");

            // Directly capture requirement MS-FSSHTTPB_R18, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     18,
                     @"[In Compact Uint 7 bit values] Uint (7 bits): An unsigned integer that specifies the value.");

            // Capture requirement MS-FSSHTTPB_R17, if the type value equals to 1. 
            site.CaptureRequirementIfAreEqual<uint>(
                     1,
                     instance.Type,
                     "MS-FSSHTTPB",
                     17,
                     @"[In Compact Uint 7 bit values] A – Type (1 bit): MUST be one.");

            // Directly capture requirement MS-FSSHTTPB_R15, if the value between 0x1F and 0x7F. 
            bool isVerifyR15 = instance.DecodedValue >= 0x1 && instance.DecodedValue <= 0x7F;
            site.Log.Add(
                    LogEntryKind.Debug,
                    "The Compact Unsigned 64-bit Integer actual value is {0}, which should be between 0x01 and 0x7F for requirement MS-FSSHTTPB_R15.",
                    instance.DecodedValue);
            site.CaptureRequirementIfIsTrue(
                     isVerifyR15,
                     "MS-FSSHTTPB",
                     15,
                     @"[In Compact Uint 7 bit values] A 1-byte encoding of values in the range 0x01 through 0x7F.");
        }

        /// <summary>
        /// This method is used to test Compact Uint 14 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint14BitValues(Compact64bitInt instance, ITestSite site)
        {
            // Directly capture requirement MS-FSSHTTPB_R20, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     20,
                     @"[In Compact Uint 14 bit values] A – Type (2 bits): A flag that specifies this format from all other formats of a compact unsigned 64-bit integer (section 2.2.1.1).");

            // Directly capture requirement MS-FSSHTTPB_R22, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     22,
                     @"[In Compact Uint 14 bit values] Uint (14 bits): An unsigned integer that specifies the value.");

            // Capture requirement MS-FSSHTTPB_R21, if the type value equals to 2. 
            site.CaptureRequirementIfAreEqual<uint>(
                     2,
                     instance.Type,
                     "MS-FSSHTTPB",
                     21,
                     @"[In Compact Uint 14 bit values] A - Type (2 bits): MUST be two.");

            // Directly capture requirement MS-FSSHTTPB_R19, if the value between 0x1F and 0x7F. 
            bool isVerifyR19 = instance.DecodedValue >= 0x0080 && instance.DecodedValue <= 0x3FFF;
            site.Log.Add(
                    LogEntryKind.Debug,
                    "The Compact Unsigned 64-bit Integer actual value is {0}, which should be between 0x0080 and 0x3FFF for requirement MS-FSSHTTPB_R19.",
                    instance.DecodedValue);
            site.CaptureRequirementIfIsTrue(
                     isVerifyR19,
                     "MS-FSSHTTPB",
                     19,
                     @"[In Compact Uint 14 bit values] A 2-byte encoding of values in the range 0x0080 through 0x3FFF.");
        }

        /// <summary>
        /// This method is used to test Compact Uint 21 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint21BitValues(Compact64bitInt instance, ITestSite site)
        {
            // Directly capture requirement MS-FSSHTTPB_R24, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     24,
                     @"[In Compact Uint 21 bit values] Type (3 bits): A flag that specifies this format from all other formats of a compact unsigned 64-bit integer (section 2.2.1.1).");

            // Directly capture requirement MS-FSSHTTPB_R26, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     26,
                     @"[In Compact Uint 21 bit values] Uint (21 bits): An unsigned integer that specifies the value.");

            // Capture requirement MS-FSSHTTPB_R25, if the type value equals to 4. 
            site.CaptureRequirementIfAreEqual<uint>(
                     4,
                     instance.Type,
                     "MS-FSSHTTPB",
                     25,
                     @"[In Compact Uint 21 bit values] Type (3 bits): MUST be four.");

            // Directly capture requirement MS-FSSHTTPB_R23, if the value between 0x004000 and 0x1FFFFF. 
            bool isVerifyR23 = instance.DecodedValue >= 0x004000 && instance.DecodedValue <= 0x1FFFFF;
            site.Log.Add(
                    LogEntryKind.Debug,
                    "The Compact Unsigned 64-bit Integer actual value is {0}, which should be between 0x004000 and 0x1FFFFF for requirement MS-FSSHTTPB_R23.",
                    instance.DecodedValue);
            site.CaptureRequirementIfIsTrue(
                     isVerifyR23,
                     "MS-FSSHTTPB",
                     23,
                     @"[In Compact Uint 21 bit values] A 3-byte encoding of values in the range 0x004000 through 0x1FFFFF.");
        }

        /// <summary>
        /// This method is used to test Compact Uint 28 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint28BitValues(Compact64bitInt instance, ITestSite site)
        {
            // All the unit35 bit values related requirements can be located here.
            site.Log.Add(LogEntryKind.Debug, "Runs for VerifyCompactUint28BitValues operation with the value {0}.", instance.DecodedValue);
        }

        /// <summary>
        /// This method is used to test Compact Uint 35 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint35BitValues(Compact64bitInt instance, ITestSite site)
        {
            // All the unit35 bit values related requirements can be located here.
            site.Log.Add(LogEntryKind.Debug, "Runs for VerifyCompactUint35BitValues operation with the value {0}.", instance.DecodedValue);
        }

        /// <summary>
        /// This method is used to test Compact Uint 42 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint42BitValues(Compact64bitInt instance, ITestSite site)
        {
            // All the unit42 bit values related requirements can be located here.
            site.Log.Add(LogEntryKind.Debug, "Runs for VerifyCompactUint42BitValues operation with the value {0}.", instance.DecodedValue);
        }

        /// <summary>
        /// This method is used to test Compact Uint 49 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint49BitValues(Compact64bitInt instance, ITestSite site)
        {
            // All the unit49 bit values related requirements can be located here.
            site.Log.Add(LogEntryKind.Debug, "Runs for VerifyCompactUint49BitValues operation with the value {0}.", instance.DecodedValue);
        }

        /// <summary>
        /// This method is used to test Compact Uint 64 bit values related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCompactUint64BitValues(Compact64bitInt instance, ITestSite site)
        {
            // Directly capture requirement MS-FSSHTTPB_R44, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     44,
                     @"[In Compact Uint 64 bit values] Type (8 bits): A flag that specifies this format from all other formats of a compact unsigned 64-bit integer (section 2.2.1.1).");

            // Directly capture requirement MS-FSSHTTPB_R46, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     46,
                     @"[In Compact Uint 64 bit values] Uint (64 bits): An unsigned integer that specifies the value.");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R43
            bool isVerify43 = instance.DecodedValue >= 0x0002000000000000 && instance.DecodedValue <= 0xFFFFFFFFFFFFFFFF;
            site.Log.Add(
                    LogEntryKind.Debug,
                    "The Compact Unsigned 64-bit Integer actual value is {0}, which should be between 0x0002000000000000 and 0x1FFFFFFFFFFFF for requirement MS-FSSHTTPB_R43.",
                    instance.DecodedValue);
            site.CaptureRequirementIfIsTrue(
                     isVerify43,
                     "MS-FSSHTTPB",
                     43,
                     @"[In Compact Uint 64 bit values] A 9-byte encoding of values in the range 0x0002000000000000 through 0xFFFFFFFFFFFFFFFF.");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R45
            site.CaptureRequirementIfAreEqual<uint>(
                     128,
                     instance.Type,
                     "MS-FSSHTTPB",
                     45,
                     @"[In Compact Uint 64 bit values] Type (8 bits): MUST be 128.");
        }

        /// <summary>
        /// This method is used to test Stream Object Header related adapter requirements.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyStreamObjectHeader(ITestSite site)
        {
            // Directly capture requirement MS-FSSHTTPB_R2089, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2089,
                     @"[In Stream Object Header] The length value does not include the size of the stream object headers.");
        }

        /// <summary>
        /// This method is used to test stream object header start type related requirements.
        /// </summary>
        /// <param name="header">Specify the stream object header start instance.</param>
        /// <param name="type">Specify the stream object header type corresponding type instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void ExpectStreamObjectHeaderStart(StreamObjectHeaderStart header, Type type, ITestSite site)
        {
            switch (header.Type)
            {
                case StreamObjectTypeHeaderStart.PutChangesResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(PutChangesResponse),
                                    type,
                                    "PutChangesResponse stream object header only represents IntermediateNodeObject instance.");
                    site.Assert.AreEqual<int>(
                                    0x087,
                                    (int)header.Type,
                                    "PutChangesResponse stream object header has header value 0x087.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "PutChangesResponse stream object header has compound value 0.");
                    
                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R99004
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             99004,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ] Put Changes Response, [the Type field is set to]0x87 [and the B-Compound field is set to] 0.");
                    break;

                case StreamObjectTypeHeaderStart.IntermediateNodeObject:
                    site.Assert.AreEqual<Type>(
                                    typeof(IntermediateNodeObject),
                                    type,
                                    "IntermediateNodeObject stream object header only represents IntermediateNodeObject instance.");
                    site.Assert.AreEqual<int>(
                                    0x20,
                                    (int)header.Type,
                                    "IntermediateNodeObject stream object header has header value 0x0104.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "IntermediateNodeObject stream object header has compound value 1.");
                    break;

                case StreamObjectTypeHeaderStart.LeafNodeObject:
                    site.Assert.AreEqual<Type>(
                                    typeof(LeafNodeObject),
                                    type,
                                    "LeafNodeObjectData stream object header only represents LeafNodeObjectData instance.");
                    site.Assert.AreEqual<int>(
                                    0x1F,
                                    (int)header.Type,
                                    "LeafNodeObjectData stream object header has header value 0x1F.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "LeafNodeObjectData stream object header has compound value 1.");
                    break;

                case StreamObjectTypeHeaderStart.SignatureObject:
                    site.Assert.AreEqual<Type>(
                                    typeof(SignatureObject),
                                    type,
                                    "SignatureObject stream object header only represents SignatureObject instance.");
                    site.Assert.AreEqual<int>(
                                    0x21,
                                    (int)header.Type,
                                    "SignatureObject stream object header has header value 0x21.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "SignatureObject stream object header has compound value 0.");
                    break;

                case StreamObjectTypeHeaderStart.DataSizeObject:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataSizeObject),
                                    type,
                                    "DataSizeObject stream object header only represents DataSizeObject instance.");
                    site.Assert.AreEqual<int>(
                                    0x22,
                                    (int)header.Type,
                                    "DataSizeObject stream object header has header value 0x22.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "DataSizeObject stream object header has compound value 0.");
                    break;

                case StreamObjectTypeHeaderStart.DataHashObject:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataHashObject),
                                    type,
                                    "DataHashObject stream object header only represents DataHashObject instance.");
                    site.Assert.AreEqual<int>(
                                    0x2F,
                                    (int)header.Type,
                                    "DataHashObject stream object header has header value 0x2F.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "DataHashObject stream object header has compound value 0.");
                    break;

                case StreamObjectTypeHeaderStart.DataElement:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataElement),
                                    type,
                                    "DataElement stream object header only represents DataElement instance.");
                    site.Assert.AreEqual<int>(
                                    0x1,
                                    (int)header.Type,
                                    "DataElement stream object header has header value 1.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "DataElement stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R68, if all above asserts pass. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             68,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of  structure] Data Element, [the Type field is set to]0x01 [and the B-Compound field is set to]1.");
                    break;

                case StreamObjectTypeHeaderStart.WaterlineKnowledgeEntry:
                    site.Assert.AreEqual<Type>(
                                    typeof(WaterlineKnowledgeEntry),
                                    type,
                                    "WaterlineKnowledgeEntry stream object header only represents WaterlineKnowledgeEntry instance.");
                    site.Assert.AreEqual<int>(
                                    0x4,
                                    (int)header.Type,
                                    "WaterlineKnowledgeEntry stream object header has header value 0x4.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "WaterlineKnowledgeEntry stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R70, if all above asserts pass.
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             70,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Waterline Knowledge Entry (section 2.2.1.13.4.1), [the Type field is set to]0x04 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.DataElementHash:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataElementHash),
                                    type,
                                    "DataElementHash stream object header only represents DataElementHash instance.");
                    site.Assert.AreEqual<int>(
                                    0x06,
                                    (int)header.Type,
                                    "DataElementHash stream object header has header value 0x06.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "DataElementHash stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R4000, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4000,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Data Element Hash, [the Type field is set to]0x06 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.StorageManifestRootDeclare:
                    site.Assert.AreEqual<Type>(
                                    typeof(StorageManifestRootDeclare),
                                    type,
                                    "StorageManifestRootDeclare stream object header only represents StorageManifestRootDeclare instance.");
                    site.Assert.AreEqual<int>(
                                    0x07,
                                    (int)header.Type,
                                    "StorageManifestRootDeclare stream object header has header value 0x07.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "StorageManifestRootDeclare stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R73, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             73,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Storage Manifest root declare, [the Type field is set to]0x07 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.RevisionManifestRootDeclare:
                    site.Assert.AreEqual<Type>(
                                    typeof(RevisionManifestRootDeclare),
                                    type,
                                    "RevisionManifestRootDeclare stream object header only represents RevisionManifestRootDeclare instance.");
                    site.Assert.AreEqual<int>(
                                    0x0A,
                                    (int)header.Type,
                                    "RevisionManifestRootDeclare stream object header has header value 0x0A.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "RevisionManifestRootDeclare stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R74, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             74,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Revision Manifest root declare, [the Type field is set to]0x0A [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.CellManifestCurrentRevision:
                    site.Assert.AreEqual<Type>(
                                    typeof(CellManifestCurrentRevision),
                                    type,
                                    "CellManifestCurrentRevision stream object header only represents CellManifestCurrentRevision instance.");
                    site.Assert.AreEqual<int>(
                                    0x0B,
                                    (int)header.Type,
                                    "CellManifestCurrentRevision stream object header has header value 0x0B.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "CellManifestCurrentRevision stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R75, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             75,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Cell Manifest current revision, [the Type field is set to]0x0B [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.StorageManifestSchemaGUID:
                    site.Assert.AreEqual<Type>(
                                    typeof(StorageManifestSchemaGUID),
                                    type,
                                    "StorageManifestSchemaGUID stream object header only represents StorageManifestSchemaGUID instance.");
                    site.Assert.AreEqual<int>(
                                    0x0C,
                                    (int)header.Type,
                                    "StorageManifestSchemaGUID stream object header has header value 0x0C.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "StorageManifestSchemaGUID stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R76, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             76,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Storage Manifest schema GUID, [the Type field is set to]0x0C [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.StorageIndexRevisionMapping:
                    site.Assert.AreEqual<Type>(
                                    typeof(StorageIndexRevisionMapping),
                                    type,
                                    "StorageIndexRevisionMapping stream object header only represents StorageIndexRevisionMapping instance.");
                    site.Assert.AreEqual<int>(
                                    0x0D,
                                    (int)header.Type,
                                    "StorageIndexRevisionMapping stream object header has header value 0x0D.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "StorageIndexRevisionMapping stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R77, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             77,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Storage Index Revision Mapping, [the Type field is set to]0x0D [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.StorageIndexCellMapping:
                    site.Assert.AreEqual<Type>(
                                    typeof(StorageIndexCellMapping),
                                    type,
                                    "StorageIndexCellMapping stream object header only represents StorageIndexCellMapping instance.");
                    site.Assert.AreEqual<int>(
                                    0x0E,
                                    (int)header.Type,
                                    "StorageIndexCellMapping stream object header has header value 0x0E.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "StorageIndexCellMapping stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R78, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             78,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Storage Index Cell Mapping, [the Type field is set to]0x0E [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.CellKnowledgeRange:
                    site.Assert.AreEqual<Type>(
                                    typeof(CellKnowledgeRange),
                                    type,
                                    "Cell knowledge range stream object header only represents CellKnowledgeRange instance.");
                    site.Assert.AreEqual<int>(
                                    0x0F,
                                    (int)header.Type,
                                    "Cell knowledge range stream object header has header value 0x0F.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "Cell knowledge range stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R79, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             79,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Cell Knowledge Range (section 2.2.1.13.2.1), [the Type field is set to]0x0F [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.Knowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(Knowledge),
                                    type,
                                    "Knowledge stream object header only represents Knowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x10,
                                    (int)header.Type,
                                    "Knowledge stream object header has header value 0x10.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "Knowledge stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R80, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             80,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Knowledge (section 2.2.1.13), [the Type field is set to]0x10 [and the B-Compound field is set to]1.");
                    break;

                case StreamObjectTypeHeaderStart.StorageIndexManifestMapping:
                    site.Assert.AreEqual<Type>(
                                    typeof(StorageIndexManifestMapping),
                                    type,
                                    "StorageIndexManifestMapping stream object header only represents StorageIndexManifestMapping instance.");
                    site.Assert.AreEqual<int>(
                                    0x11,
                                    (int)header.Type,
                                    "StorageIndexManifestMapping stream object header has header value 0x11.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "StorageIndexManifestMapping stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R81, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             81,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Storage Index Manifest Mapping, [the Type field is set to]0x11 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.CellKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(CellKnowledge),
                                    type,
                                    "CellKnowledge stream object header only represents CellKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x14,
                                    (int)header.Type,
                                    "CellKnowledge stream object header has header value 0x14.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "CellKnowledge stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R83, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             83,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Cell Knowledge (section 2.2.1.13.2), [the Type field is set to]0x14 [and the B-Compound field is set to]1.");
                    break;

                case StreamObjectTypeHeaderStart.ObjectGroupObjectData:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupObjectData),
                                    type,
                                    "ObjectGroupObjectData stream object header only represents ObjectGroupObjectData instance.");
                    site.Assert.AreEqual<int>(
                                    0x16,
                                    (int)header.Type,
                                    "ObjectGroupObjectData stream object header has header value 0x16.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "ObjectGroupObjectData stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R85, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             85,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Object Group Object Data, [the Type field is set to]0x16 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.CellKnowledgeEntry:
                    site.Assert.AreEqual<Type>(
                                    typeof(CellKnowledgeEntry),
                                    type,
                                    "CellKnowledgeEntry stream object header only represents CellKnowledgeEntry instance.");
                    site.Assert.AreEqual<int>(
                                    0x17,
                                    (int)header.Type,
                                    "CellKnowledgeEntry stream object header has header value 0x17.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "CellKnowledgeEntry stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R86, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             86,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Cell Knowledge Entry (section 2.2.1.13.2.2), [the Type field is set to]0x17 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.ObjectGroupObjectDeclare:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupObjectDeclare),
                                    type,
                                    "ObjectGroupObjectDeclare stream object header only represents ObjectGroupObjectDeclare instance.");
                    site.Assert.AreEqual<int>(
                                    0x18,
                                    (int)header.Type,
                                    "ObjectGroupObjectDeclare stream object header has header value 0x18.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "ObjectGroupObjectDeclare stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R87, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             87,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Object Group Object Declare, [the Type field is set to]0x18 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.RevisionManifestObjectGroupReferences:
                    site.Assert.AreEqual<Type>(
                                    typeof(RevisionManifestObjectGroupReferences),
                                    type,
                                    "RevisionManifestObjectGroupReferences stream object header only represents RevisionManifestObjectGroupReferences instance.");
                    site.Assert.AreEqual<int>(
                                    0x19,
                                    (int)header.Type,
                                    "RevisionManifestObjectGroupReferences stream object header has header value 0x19.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "RevisionManifestObjectGroupReferences stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R88, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             88,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Revision Manifest Object Group references, [the Type field is set to]0x19 [and the B-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.RevisionManifest:
                    site.Assert.AreEqual<Type>(
                                    typeof(RevisionManifest),
                                    type,
                                    "RevisionManifest stream object header only represents RevisionManifest instance.");
                    site.Assert.AreEqual<int>(
                                    0x1A,
                                    (int)header.Type,
                                    "RevisionManifest stream object header has header value 0x1A.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "RevisionManifest stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R89, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             89,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Revision Manifest, [the Type field is set to]0x1A [and the B-Compound field is set to]0.");

                    break;

                case StreamObjectTypeHeaderStart.ObjectGroupDeclarations:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupDeclarations),
                                    type,
                                    "ObjectGroupDeclarations stream object header only represents ObjectGroupDeclarations instance.");
                    site.Assert.AreEqual<int>(
                                    0x1D,
                                    (int)header.Type,
                                    "ObjectGroupDeclarations stream object header has header value 0x1D.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "ObjectGroupDeclarations stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R91, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             91,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Object Group Declarations, [the Type field is set to]0x1D [and the B-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.ObjectGroupData:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupData),
                                    type,
                                    "ObjectGroupData stream object header only represents ObjectGroupData instance.");
                    site.Assert.AreEqual<int>(
                                    0x1E,
                                    (int)header.Type,
                                    "ObjectGroupData stream object header has header value 0x1E.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "ObjectGroupData stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R92, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             92,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Object Group Data, [the Type field is set to]0x1E [and the B-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.WaterlineKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(WaterlineKnowledge),
                                    type,
                                    "WaterlineKnowledge stream object header only represents WaterlineKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x29,
                                    (int)header.Type,
                                    "WaterlineKnowledge stream object header has header value 0x29.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "WaterlineKnowledge stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R93, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             93,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Waterline Knowledge (section 2.2.1.13.4), [the Type field is set to]0x29 [and the B-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.ContentTagKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(ContentTagKnowledge),
                                    type,
                                    "ContentTagKnowledge stream object header only represents ContentTagKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x2D,
                                    (int)header.Type,
                                    "ContentTagKnowledge stream object header has header value 0x2D.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "ContentTagKnowledge stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R94, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             94,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Content Tag Knowledge (section 2.2.1.13.5), [the Type field is set to]0x2D [and the B-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.ContentTagKnowledgeEntry:
                    site.Assert.AreEqual<Type>(
                                    typeof(ContentTagKnowledgeEntry),
                                    type,
                                    "ContentTagKnowledgeEntry stream object header only represents ContentTagKnowledgeEntry instance.");
                    site.Assert.AreEqual<int>(
                                    0x2E,
                                    (int)header.Type,
                                    "ContentTagKnowledgeEntry stream object header has header value 0x2E.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "ContentTagKnowledgeEntry stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R95, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             95,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Content Tag Knowledge Entry, [the Type field is set to]0x2E [and the B-Compound field is set to]0.");

                    break;

                case StreamObjectTypeHeaderStart.ObjectGroupMetadataDeclarations:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupMetadataDeclarations),
                                    type,
                                    "ObjectGroupMetadataDeclarations stream object header only represents ObjectGroupMetadataDeclarations instance.");
                    site.Assert.AreEqual<int>(
                                    0x79,
                                    (int)header.Type,
                                    "ObjectGroupMetadataDeclarations stream object header has header value 0x79.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "ObjectGroupMetadataDeclarations stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R2090, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             2090,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ] Object Group metadata declarations, [the Type field is set to]0x79 [and the B-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.ObjectGroupMetadata:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupMetadata),
                                    type,
                                    "ObjectGroupMetadata stream object header only represents ObjectGroupMetadata instance.");
                    site.Assert.AreEqual<int>(
                                    0x78,
                                    (int)header.Type,
                                    "ObjectGroupMetadata stream object header has header value 0x78.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "ObjectGroupMetadata stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R2091, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             2091,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ] Object Group metadata, [the Type field is set to]0x78 [and the B-Compound field is set to]0.");

                    break;

                case StreamObjectTypeHeaderStart.AllocateExtendedGUIDRangeResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(AllocateExtendedGuidRangeSubResponseData),
                                    type,
                                    "Allocate ExtendedGUID range response stream object header only represents AllocateExtendedGuidRangeData instance.");
                    site.Assert.AreEqual<int>(
                                    0x081,
                                    (int)header.Type,
                                    "AllocateExtendedGuidRangeData stream object header has header value 0x081.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "AllocateExtendedGuidRangeData stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R2094, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             2094,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ] Allocate Extended GUID range response (section 2.2.3.1.4), [the Type field is set to]0x081 [and the B-Compound field is set to]0.");
                    break;  

                case StreamObjectTypeHeaderStart.FsshttpbSubResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(FsshttpbSubResponse),
                                    type,
                                    "sub response stream object header only represents FSSHTTPBSubResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x041,
                                    (int)header.Type,
                                    "FSSHTTPBSubResponse stream object header has header value 0x041.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "FSSHTTPBSubResponse stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R106, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             106,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Sub-response, [the Type field is set to]0x041 [and the C-Compound field is set to]1.");
                    break;

                case StreamObjectTypeHeaderStart.ReadAccessResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(ReadAccessResponse),
                                    type,
                                    "ReadAccessResponse stream object header only represents ReadAccessResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x043,
                                    (int)header.Type,
                                    "ReadAccessResponse stream object header has header value 0x043.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "ReadAccessResponse stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R108, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             108,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Read access response, [the Type field is set to]0x043 [and the C-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.SpecializedKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(SpecializedKnowledge),
                                    type,
                                    "SpecializedKnowledge stream object header only represents SpecializedKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x044,
                                    (int)header.Type,
                                    "SpecializedKnowledge stream object header has header value 0x044.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "SpecializedKnowledge stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R109, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             109,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Specialized Knowledge, [the Type field is set to]0x044 [and the C-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.WriteAccessResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(WriteAccessResponse),
                                    type,
                                    "WriteAccessResponse stream object header only represents WriteAccessResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x046,
                                    (int)header.Type,
                                    "WriteAccessResponse stream object header has header value 0x046.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "WriteAccessResponse stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R111, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             111,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Write access response, [the Type field is set to]0x046 [and the C-Compound field is set to]1.");
                    break;

                case StreamObjectTypeHeaderStart.ResponseError:
                    site.Assert.AreEqual<Type>(
                                    typeof(ResponseError),
                                    type,
                                    "Error stream object header only represents ResponseError instance.");
                    site.Assert.AreEqual<int>(
                                    0x04D,
                                    (int)header.Type,
                                    "ResponseError stream object header has header value 0x04D.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "ResponseError stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R116, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             116,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Error, [the Type field is set to]0x04D [and the C-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.QueryChangesResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(QueryChangesSubResponseData),
                                    type,
                                    "Query changes response stream object header only represents QueryChangesSubResponseData instance.");
                    site.Assert.AreEqual<int>(
                                    0x05F,
                                    (int)header.Type,
                                    "QueryChangesSubResponseData stream object header has header value 0x05F.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "QueryChangesSubResponseData stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R131, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             131,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Query Changes response, [the Type field is set to]0x05F [and the C-Compound field is set to]0.");

                    break;

                case StreamObjectTypeHeaderStart.FsshttpbResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(FsshttpbResponse),
                                    type,
                                    "Response stream object header only represents FsshttpbResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x062,
                                    (int)header.Type,
                                    "FsshttpbResponse stream object header has header value 0x062.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "FsshttpbResponse stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R133, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             133,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Response, [the Type field is set to]0x062 [and the C-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.CellError:
                    site.Assert.AreEqual<Type>(
                                    typeof(CellError),
                                    type,
                                    "Error cell stream object header only represents CellError instance.");
                    site.Assert.AreEqual<int>(
                                    0x066,
                                    (int)header.Type,
                                    "CellError stream object header has header value 0x066.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "CellError stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R136, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             136,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Error cell, [the Type field is set to]0x066 [and the C-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.DataElementFragment:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataElementFragment),
                                    type,
                                    "Data element fragment stream object header only represents DataElementFragment instance.");
                    site.Assert.AreEqual<int>(
                                    0x06A,
                                    (int)header.Type,
                                    "DataElementFragment stream object header has header value 0x06A.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "DataElementFragment stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R138, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             138,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Data Element Fragment, [the Type field is set to]0x06A [and the C-Compound field is set to]0.");

                    break;

                case StreamObjectTypeHeaderStart.FragmentKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(FragmentKnowledge),
                                    type,
                                    "Fragment knowledge stream object header only represents FragmentKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x06B,
                                    (int)header.Type,
                                    "FragmentKnowledge stream object header has header value 0x06B.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "FragmentKnowledge stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R139, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             139,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Fragment Knowledge, [the Type field is set to]0x06B [and the C-Compound field is set to]1.");

                    break;

                case StreamObjectTypeHeaderStart.FragmentKnowledgeEntry:
                    site.Assert.AreEqual<Type>(
                                   typeof(FragmentKnowledgeEntry),
                                   type,
                                   "Fragment knowledge entry stream object header only represents FragmentKnowledgeEntry instance.");
                    site.Assert.AreEqual<int>(
                                    0x06C,
                                    (int)header.Type,
                                    "Fragment knowledge entry stream object header has header value 0x06C.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "Fragment knowledge entry stream object header has compound value 0.");

                    // Directly capture requirement MS-FSSHTTPB_R140, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             140,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Fragment Knowledge entry, [the Type field is set to]0x06C [and the C-Compound field is set to]0.");

                    break;

                case StreamObjectTypeHeaderStart.DataElementPackage:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataElementPackage),
                                    type,
                                    "Data element package stream object header only represents DataElementPackage instance.");
                    site.Assert.AreEqual<int>(
                                    0x15,
                                    (int)header.Type,
                                    "Data element package stream object header has header value 0x15.");
                    site.Assert.AreEqual<int>(
                                    1,
                                    header.Compound,
                                    "Data element package stream object header has compound value 1.");

                    // Directly capture requirement MS-FSSHTTPB_R84, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             84,
                             @"[In 16-bit Stream Object Header Start][If the related Stream Object is type of ] Data Element Package, [the Type field is set to]0x15 [and the B-Compound field is set to]1.");
                    break;
                case StreamObjectTypeHeaderStart.HRESULTError:
                    site.Assert.AreEqual<Type>(
                                   typeof(HRESULTError),
                                   type,
                                   "HRESULTErrorobject header only represents HRESULTError instance.");
                    site.Assert.AreEqual<int>(
                                    0x52,
                                    (int)header.Type,
                                    "Data element package stream object header has header value 0x52.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "Data element package stream object header has compound value 0.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R121
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             121,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Error HRESULT, [the Type field is set to]0x052 [and the C-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.Win32Error:
                    site.Assert.AreEqual<Type>(
                                    typeof(Win32Error),
                                    type,
                                    "Win32Error object header only represents Win32Error instance.");
                    site.Assert.AreEqual<int>(
                                    0x49,
                                    (int)header.Type,
                                    "Data element package stream object header has header value 0x49.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "Data element package stream object header has compound value 0.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R113
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             113,
                             @"[In 32-bit Stream Object Header Start][If the related Stream Object is type of ]Error Win32, [the Type field is set to]0x049 [and the C-Compound field is set to]0.");
                    break;

                case StreamObjectTypeHeaderStart.ProtocolError:
                    break;
                case StreamObjectTypeHeaderStart.ObjectGroupObjectBLOBDataDeclaration:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupObjectBLOBDataDeclaration),
                                    type,
                                    "ObjectGroupObjectBLOBDataDeclaration stream object header only represents ObjectGroupObjectBLOBDataDeclaration instance.");
                    site.Assert.AreEqual<int>(
                                    0x05,
                                    (int)header.Type,
                                    "ObjectGroupObjectBLOBDataDeclaration stream object header has header value 0x49.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "ObjectGroupObjectBLOBDataDeclaration stream object header has compound value 0.");
                    break;

                case StreamObjectTypeHeaderStart.ObjectGroupObjectDataBLOBReference:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupObjectDataBLOBReference),
                                    type,
                                    "ObjectGroupObjectDataBLOBReference stream object header only represents Win32Error instance.");
                    site.Assert.AreEqual<int>(
                                    0x1C,
                                    (int)header.Type,
                                    "ObjectGroupObjectDataBLOBReference stream object header has header value 0x49.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "ObjectGroupObjectDataBLOBReference stream object header has compound value 0.");
                    break;

                case StreamObjectTypeHeaderStart.ErrorStringSupplementalInfo:
                      site.Assert.AreEqual<Type>(
                                    typeof(ErrorStringSupplementalInfo),
                                    type,
                                    "ErrorStringSupplementalInfo stream object header only represents ErrorStringSupplementalInfo instance.");
                    site.Assert.AreEqual<int>(
                                    0x4E,
                                    (int)header.Type,
                                    "ErrorStringSupplementalInfo stream object header has header value 0x4E.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "ErrorStringSupplementalInfo stream object header has compound value 0.");
                    break;

                case StreamObjectTypeHeaderStart.DiagnosticRequestOptionOutput:
                    site.Assert.AreEqual<Type>(
                                  typeof(DiagnosticRequestOptionOutput),
                                  type,
                                  "DiagnosticRequestOptionOutput stream object header only represents DiagnosticRequestOptionOutput instance.");
                    site.Assert.AreEqual<int>(
                                    0x89,
                                    (int)header.Type,
                                    "DiagnosticRequestOptionOutput stream object header has header value 0x89.");
                    site.Assert.AreEqual<int>(
                                    0,
                                    header.Compound,
                                    "DiagnosticRequestOptionOutput stream object header has compound value 0.");
                    break;

                default:
                    site.Assert.Fail("Does not support the stream object type " + header.Type.ToString());
                    break;
            }
        }

        /// <summary>
        /// This method is used to test stream object header end type related requirements.
        /// </summary>
        /// <param name="header">Specify the stream object header end instance.</param>
        /// <param name="type">Specify the stream object header type corresponding type instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void ExpectStreamObjectHeaderEnd(StreamObjectHeaderEnd header, Type type, ITestSite site)
        {
            if(header==null)
            {
                return;
            }
            switch (header.Type)
            {
                case StreamObjectTypeHeaderEnd.RootNodeEnd:
                    site.Assert.AreEqual<Type>(
                                    typeof(IntermediateNodeObject),
                                    type,
                                    "RootNodeEnd stream object header only represents IntermediateNodeObject instance.");
                    site.Assert.AreEqual<int>(
                                    0x20,
                                    (int)header.Type,
                                    "IntermediateNodeObject header has header value 0x20.");
                    break;

                case StreamObjectTypeHeaderEnd.IntermediateNodeEnd:
                    site.Assert.AreEqual<Type>(
                                    typeof(LeafNodeObject),
                                    type,
                                    "IntermediateNodeEnd header only represents LeafNodeObjectData instance.");
                    site.Assert.AreEqual<int>(
                                    0x1F,
                                    (int)header.Type,
                                    "IntermediateNodeEndstream object header has header value 0x1F.");
                    break;

                case StreamObjectTypeHeaderEnd.DataElement:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataElement),
                                    type,
                                    "Data element stream object header only represents DataElement instance.");
                    site.Assert.AreEqual<int>(
                                    0x01,
                                    (int)header.Type,
                                    "Data element stream object header has header value 0x01.");

                    // Directly capture requirement MS-FSSHTTPB_R150, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             150,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Data element, [the Type field is set to]0x01.");

                    break;

                case StreamObjectTypeHeaderEnd.Knowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(Knowledge),
                                    type,
                                    "Knowledge stream object header only represents Knowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x10,
                                    (int)header.Type,
                                    "Knowledge stream object header has header value 0x10.");

                    // Directly capture requirement MS-FSSHTTPB_R151, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             151,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Knowledge, [the Type field is set to]0x10.");

                    break;

                case StreamObjectTypeHeaderEnd.CellKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(CellKnowledge),
                                    type,
                                    "CellKnowledge stream object header only represents CellKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x14,
                                    (int)header.Type,
                                    "CellKnowledge stream object header has header value 0x14.");

                    // Directly capture requirement MS-FSSHTTPB_R152, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             152,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Cell Knowledge, [the Type field is set to]0x14.");

                    break;

                case StreamObjectTypeHeaderEnd.DataElementPackage:
                    site.Assert.AreEqual<Type>(
                                    typeof(DataElementPackage),
                                    type,
                                    "DataElementPackage stream object header only represents DataElementPackage instance.");
                    site.Assert.AreEqual<int>(
                                    0x15,
                                    (int)header.Type,
                                    "DataElementPackage stream object header has header value 0x15.");

                    // Directly capture requirement MS-FSSHTTPB_R153, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             153,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Data Element Package, [the Type field is set to]0x15.");

                    break;

                case StreamObjectTypeHeaderEnd.ObjectGroupDeclarations:
                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupDeclarations),
                                    type,
                                    "ObjectGroupDeclarations stream object header only represents ObjectGroupDeclarations instance.");
                    site.Assert.AreEqual<int>(
                                    0x1D,
                                    (int)header.Type,
                                    "ObjectGroupDeclarations stream object header has header value 0x1D.");

                    // Directly capture requirement MS-FSSHTTPB_R154, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             154,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Object Group declarations, [the Type field is set to]0x1D.");
                    break;

                case StreamObjectTypeHeaderEnd.ObjectGroupData:
                    site.Assert.AreEqual<Type>(
                                   typeof(ObjectGroupData),
                                   type,
                                   "ObjectGroupData stream object header only represents ObjectGroupData instance.");
                    site.Assert.AreEqual<int>(
                                    0x1E,
                                    (int)header.Type,
                                    "ObjectGroupData stream object header has header value 0x1E.");

                    // Directly capture requirement MS-FSSHTTPB_R155, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             155,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Object Group data, [the Type field is set to]0x1E.");
                    break;

                case StreamObjectTypeHeaderEnd.WaterlineKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(WaterlineKnowledge),
                                    type,
                                    "WaterlineKnowledge stream object header only represents WaterlineKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x29,
                                    (int)header.Type,
                                    "WaterlineKnowledge stream object header has header value 0x29.");

                    // Directly capture requirement MS-FSSHTTPB_R156, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             156,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Waterline Knowledge, [the Type field is set to]0x29.");

                    break;

                case StreamObjectTypeHeaderEnd.ContentTagKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(ContentTagKnowledge),
                                    type,
                                    "ContentTagKnowledge stream object header only represents ContentTagKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x2D,
                                    (int)header.Type,
                                    "ContentTagKnowledge stream object header has header value 0x2D.");

                    // Directly capture requirement MS-FSSHTTPB_R157, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             157,
                             @"[In 8-bit Stream Object Header End][If the related Stream Object is type of ]Content Tag Knowledge, [the Type field is set to]0x2D.");

                    break;

                case StreamObjectTypeHeaderEnd.SubResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(FsshttpbSubResponse),
                                    type,
                                    "Sub-response stream object header only represents FSSHTTPBSubResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x041,
                                    (int)header.Type,
                                    "Sub-response stream object header has header value 0x041.");

                    // Directly capture requirement MS-FSSHTTPB_R165, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             165,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Sub-response, [the Type field is set to]0x041.");

                    break;

                case StreamObjectTypeHeaderEnd.ReadAccessResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(ReadAccessResponse),
                                    type,
                                    "ReadAccessResponse stream object header only represents ReadAccessResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x043,
                                    (int)header.Type,
                                    "ReadAccessResponse stream object header has header value 0x043.");

                    // Directly capture requirement MS-FSSHTTPB_R167, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             167,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Read access response, [the Type field is set to]0x043.");

                    break;

                case StreamObjectTypeHeaderEnd.SpecializedKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(SpecializedKnowledge),
                                    type,
                                    "SpecializedKnowledge stream object header only represents SpecializedKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x044,
                                    (int)header.Type,
                                    "SpecializedKnowledge stream object header has header value 0x044.");

                    // Directly capture requirement MS-FSSHTTPB_R168, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             168,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Specialized, [the Type field is set to]0x044.");

                    break;

                case StreamObjectTypeHeaderEnd.WriteAccessResponse:
                    site.Assert.AreEqual<Type>(
                                    typeof(WriteAccessResponse),
                                    type,
                                    "WriteAccessResponse stream object header only represents WriteAccessResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x046,
                                    (int)header.Type,
                                    "WriteAccessResponse stream object header has header value 0x046.");

                    // Directly capture requirement MS-FSSHTTPB_R169, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             169,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Write access response, [the Type field is set to]0x046.");

                    break;

                case StreamObjectTypeHeaderEnd.Error:
                    site.Assert.AreEqual<Type>(
                                    typeof(ResponseError),
                                    type,
                                    "Error stream object header only represents ResponseError instance.");
                    site.Assert.AreEqual<int>(
                                    0x04D,
                                    (int)header.Type,
                                    "ResponseError stream object header has header value 0x04D.");

                    // Directly capture requirement MS-FSSHTTPB_R171, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             171,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Error, [the Type field is set to]0x04D.");

                    break;

                case StreamObjectTypeHeaderEnd.Response:
                    site.Assert.AreEqual<Type>(
                                    typeof(FsshttpbResponse),
                                    type,
                                    "Response stream object header only represents FsshttpbResponse instance.");
                    site.Assert.AreEqual<int>(
                                    0x062,
                                    (int)header.Type,
                                    "Response stream object header has header value 0x062.");

                    // Directly capture requirement MS-FSSHTTPB_R174, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             174,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Response, [the Type field is set to]0x062.");

                    break;

                case StreamObjectTypeHeaderEnd.FragmentKnowledge:
                    site.Assert.AreEqual<Type>(
                                    typeof(FragmentKnowledge),
                                    type,
                                    "FragmentKnowledge stream object header only represents FragmentKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x06B,
                                    (int)header.Type,
                                    "FragmentKnowledge stream object header has header value 0x06B.");

                    // Directly capture requirement MS-FSSHTTPB_R175, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             175,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Fragment Knowledge, [the Type field is set to]0x06B.");

                    break;

                case StreamObjectTypeHeaderEnd.ObjectGroupMetadataDeclarations:

                    site.Assert.AreEqual<Type>(
                                    typeof(ObjectGroupMetadataDeclarations),
                                    type,
                                    "ObjectGroupMetadataDeclarations stream object header only represents FragmentKnowledge instance.");
                    site.Assert.AreEqual<int>(
                                    0x079,
                                    (int)header.Type,
                                    "FragmentKnowledge stream object header has header value 0x079.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R99006
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             99006,
                             @"[In 16-bit Stream Object Header End][If the related Stream Object is type of ]Object Group metadata declarations, [the Type field is set to]0x079.");
                    break;

                default:
                    site.Assert.Fail("Does not support the stream object type " + header.Type.ToString());
                    break;
            }
        }

        /// <summary>
        /// This method is used to test compound object related requirements.
        /// </summary>
        /// <param name="header">Specify the stream object header end instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void ExpectCompoundObject(StreamObjectHeaderStart header, ITestSite site)
        {
            if (header is StreamObjectHeaderStart16bit)
            {
                // When the header is StreamObjectHeaderStart16bit, then capture the following two requirements.
                // Capture requirement MS-FSSHTTPB_R61, if there are no parsing errors and compound value equals to 1.
                site.CaptureRequirementIfAreEqual<int>(
                         1,
                         header.Compound,
                         "MS-FSSHTTPB",
                         61,
                         @"[In 16-bit Stream Object Header Start] B - Compound (1-bit): If set, a bit that specifies a compound parse type is needed.");

                // Capture requirement MS-FSSHTTPB_R62, if there are no parsing errors  and compound value equals to 1. 
                site.CaptureRequirementIfAreEqual<int>(
                         1,
                         header.Compound,
                         "MS-FSSHTTPB",
                         62,
                         @"[In 16-bit Stream Object Header Start] B - Compound (1-bit): MUST end with either an 8-bit Stream Object Header end (section 2.2.1.5.3) or a 16-bit Stream Object Header end (section 2.2.1.5.4).");
            }
            else if (header is StreamObjectHeaderStart32bit)
            {
                // When the header is StreamObjectHeaderStart32bit, then capture the following two requirements.
                // Capture requirement MS-FSSHTTPB_R99, if there are no parsing errors. 
                site.CaptureRequirementIfAreEqual<int>(
                         1,
                         header.Compound,
                         "MS-FSSHTTPB",
                         99,
                         @"[In 32-bit Stream Object Header Start] B - Compound (1-bit): If set, a bit that specifies a compound parse type is needed.");

                // Capture requirement MS-FSSHTTPB_R100, if there are no parsing errors. 
                site.CaptureRequirementIfAreEqual<int>(
                         1,
                         header.Compound,
                         "MS-FSSHTTPB",
                         100,
                         @"[In 32-bit Stream Object Header Start] B - Compound (1-bit): MUST end with either an 8-bit Stream Object Header end (section 2.2.1.5.3) or a 16-bit Stream Object Header end (section 2.2.1.5.4).");
            }
            else
            {
                site.Assert.Fail("Unsupported stream object start type " + header.GetType().Name);
            }
        }

        /// <summary>
        /// This method is used to test single object related requirements.
        /// </summary>
        /// <param name="header">Specify the stream object header end instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void ExpectSingleObject(StreamObjectHeaderStart header, ITestSite site)
        {
            if (header is StreamObjectHeaderStart16bit)
            {
                // When the header is StreamObjectHeaderStart16bit, then capture the following two requirements.
                // Capture requirement MS-FSSHTTPB_R63, if there are no parsing errors  and compound value equals to 0.
                site.CaptureRequirementIfAreEqual<int>(
                         0,
                         header.Compound,
                         "MS-FSSHTTPB",
                         63,
                         @"[In 16-bit Stream Object Header Start] If the bit[B - Compound (1-bit)] is not set, it[B - Compound (1-bit)] specifies a single object.");
            }
            else if (header is StreamObjectHeaderStart32bit)
            {
                // When the header is StreamObjectHeaderStart32bit, then capture the following two requirements.
                // Capture requirement MS-FSSHTTPB_R101, if there are no parsing errors. 
                site.CaptureRequirementIfAreEqual<int>(
                         0,
                         header.Compound,
                         "MS-FSSHTTPB",
                         101,
                         @"[In 32-bit Stream Object Header Start] If the bit[B - Compound (1-bit)] is not set, it[B - Compound (1-bit)] specifies a single object.");
            }
            else
            {
                site.Assert.Fail("Unsupported stream object start type " + header.GetType().Name);
            }
        }

        #endregion
    }
}