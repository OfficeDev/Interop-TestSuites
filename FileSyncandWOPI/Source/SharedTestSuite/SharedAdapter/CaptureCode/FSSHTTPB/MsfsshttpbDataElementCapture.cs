namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This is the partial part of the class MsfsshttpbAdapterCapture for MS-FSSHTTPB data element related parts.
    /// </summary> 
    public partial class MsfsshttpbAdapterCapture
    {
        /// <summary>
        /// This method is used to test Data Element Package related adapter requirements.
        /// </summary> 
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyDataElementPackage(DataElementPackage instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Data Element Package related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type DataElementPackage is null due to parsing error or type casting error.");
            }

            // Directly capture requirement MS-FSSHTTPB_R239, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     239,
                     @"[In Data Element Package] A Data Element Package contains the serialized file data elements made up of Storage Index (section 2.2.1.12.2), Storage Manifest (section 2.2.1.12.3), Cell Manifest (section 2.2.1.12.4), Revision Manifest (section 2.2.1.12.5), and Object Group (section 2.2.1.12.6) or Object Data (section 2.2.1.12.6.4), or both.");

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            bool isVerifyR240 = instance.StreamObjectHeaderStart is StreamObjectHeaderStart16bit;
            site.Assert.IsTrue(
                            isVerifyR240,
                            "Actual stream object header type is {0}, which should be 16-bit stream object header for the requirement MS-FSSHTTPB_R240.",
                            instance.StreamObjectHeaderStart.GetType().Name);

            // Capture requirement MS-FSSHTTPB_R240, if the above assertion was validated.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     240,
                     @"[In Data Element Package] Data Element Package Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a Data Element Package start.");

            // Directly capture requirement MS-FSSHTTPB_R241, if the reserved value equals to 0.
            site.CaptureRequirementIfAreEqual<uint>(
                     0,
                     instance.Reserved,
                     "MS-FSSHTTPB",
                     241,
                     @"[In Data Element Package] Reserved (1 byte): A reserved field that MUST be set to zero.");

            if (instance.DataElements != null && instance.DataElements.Count != 0)
            {
                // Directly capture requirement MS-FSSHTTPB_R243, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         243,
                         @"[In Data Element Package] Data Element (variable): An optional array that contains the serialized file data elements.");
            }

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            bool isVerifyR245 = instance.StreamObjectHeaderStart is StreamObjectHeaderStart16bit;
            site.Assert.IsTrue(
                            isVerifyR245,
                            "Actual stream object header end type is {0}, which should be 8-bit stream object end header for the requirement MS-FSSHTTPB_R245.",
                            instance.StreamObjectHeaderEnd.GetType().Name);

            // Directly capture requirement MS-FSSHTTPB_R245, if the above assertion was validated.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     245,
                     @"[In Data Element Package] Data Element Package End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a Data Element Package end.");
        }

        /// <summary>
        /// This method is used to test Data Element Types related adapter requirements.
        /// </summary> 
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyDataElement(DataElement instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Data Element related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type DataElement is null due to parsing error or type casting error.");
            }

            bool isVerifyR246 = (int)instance.DataElementType == 0x01 ||
                                (int)instance.DataElementType == 0x02 ||
                                (int)instance.DataElementType == 0x03 ||
                                (int)instance.DataElementType == 0x04 ||
                                (int)instance.DataElementType == 0x05 ||
                                (int)instance.DataElementType == 0x06 ||
                                (int)instance.DataElementType == 0x0A;

            site.Assert.IsTrue(
                            isVerifyR246,
                            "For the requirement MS-FSSHTTPB_R246, the data element type value is either 0x01, 0x02, 0x03, 0x04, 0x05, 0x06 or 0x0A");

            // Directly capture requirement MS-FSSHTTPB_R246, if there are no parsing errors. 
            site.CaptureRequirementIfIsTrue(
                     isVerifyR246,
                     "MS-FSSHTTPB",
                     246,
                     @"[In Data Element Types] The following table lists the possible data element types:[Its value must be one of 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x0A].");

            switch ((int)instance.DataElementType)
            {
                case 0x01:
                    site.Assert.AreEqual<Type>(
                            typeof(StorageIndexDataElementData),
                            instance.Data.GetType(),
                            "When the DataElementType value is 0x1, expect the Data type is StorageIndexDataElementData.");

                    // Capture requirement MS-FSSHTTPB_R247, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             247,
                             @"[In Data Element Types][If the related Data Element is type of ] Storage Index (section 2.2.1.12.2), [the Data Element Type field is set to]0x01.");

                    // Verify the storage index data element related requirements.    
                    this.VerifyStorageIndexDataElement(instance, site);
                    break;

                case 0x2:
                    site.Assert.AreEqual<Type>(
                            typeof(StorageManifestDataElementData),
                            instance.Data.GetType(),
                            "When the DataElementType value is 0x2, expect the Data type is StorageManifestDataElementData.");

                    // Capture requirement MS-FSSHTTPB_R248, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             248,
                             @"[In Data Element Types][If the related Data Element is type of ] Storage Manifest (section 2.2.1.12.3), [the Data Element Type field is set to]0x02.");

                    // Verify the storage manifest data element related requirements.    
                    this.VerifyStorageManifestDataElement(instance, site);
                    break;

                case 0x03:
                    site.Assert.AreEqual<Type>(
                            typeof(CellManifestDataElementData),
                            instance.Data.GetType(),
                            "When the DataElementType value is 0x3, expect the Data type is CellManifestDataElementData.");

                    // Directly capture requirement MS-FSSHTTPB_R249, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             249,
                             @"[In Data Element Types][If the related Data Element is type of ] Cell Manifest (section 2.2.1.12.4), [the Data Element Type field is set to]0x03.");

                    // Verify the cell manifest data element related requirements.    
                    this.VerifyCellManifestDataElement(instance, site);
                    break;

                case 0x4:
                    site.Assert.AreEqual<Type>(
                            typeof(RevisionManifestDataElementData),
                            instance.Data.GetType(),
                            "When the DataElementType value is 0x4, expect the Data type is RevisionManifestDataElementData.");

                    // Directly capture requirement MS-FSSHTTPB_R250, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             250,
                             @"[In Data Element Types][If the related Data Element is type of ] Revision Manifest (section 2.2.1.12.5), [the Data Element Type field is set to]0x04.");

                    // Verify the revision manifest data element related requirements.    
                    this.VerifyRevisionManifestDataElement(instance, site);
                    break;

                case 0x05:
                    site.Assert.AreEqual<Type>(
                            typeof(ObjectGroupDataElementData),
                            instance.Data.GetType(),
                            "When the DataElementType value is 0x5, expect the Data type is ObjectGroupDataElementData.");

                    // Directly capture requirement MS-FSSHTTPB_R251, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             251,
                             @"[In Data Element Types][If the related Data Element is type of ] Object Group (section 2.2.1.12.6), [the Data Element Type field is set to]0x05.");

                    // Verify the object group data element related requirements.    
                    this.VerifyObjectGroupDataElement(instance, site);
                    break;

                case 0x6:
                    site.Assert.AreEqual<Type>(
                            typeof(FragmentDataElementData),
                            instance.Data.GetType(),
                            "When the DataElementType value is 0x6, expect the Data type is FragmentDataElementData.");

                    // Directly capture requirement MS-FSSHTTPB_R252, if there are no parsing errors. 
                    site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             252,
                             @"[In Data Element Types][If the related Data Element is type of ] Data Element Fragment (section 2.2.1.12.7), [the Data Element Type field is set to]0x06.");

                    // Verify the object group data element related requirements.    
                    this.VerifyFragmentDataElement(instance, site);
                    break;

                case 0xA:
                    site.Assert.AreEqual<Type>(
                            typeof(ObjectDataBLOBDataElementData),
                            instance.Data.GetType(),
                            "When the DataElementType value is 0xA, expect the Data type is ObjectDataBLOBDataElementData.");

                    // Directly capture requirement MS-FSSHTTPB_R253, if there are no parsing errors. 
                    site.CaptureRequirement(
                        "MS-FSSHTTPB",
                        253,
                        @"[In Data Element Types][If the related Data Element is type of ] Object Data BLOB (section 2.2.1.12.8), [the Data Element Type field is set to]0x0A.");

                    this.VerifyObjectDataBLOBDataElement(instance, site);
                    break;

                default:
                    site.Assert.Fail("Unsupported Data Element Type value " + (int)instance.DataElementType);
                    break;
            }
        }

        /// <summary>
        /// This method is used to test Storage Index Manifest Mapping related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStorageIndexManifestMapping(StorageIndexManifestMapping instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the StorageIndexManifestMapping related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StorageIndexManifestMapping is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R259, if the stream object header is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     259,
                     @"[In Storage Index Data Element] Storage Index Manifest Mapping (2 bytes, optional): Zero or one 16-bit Stream Object Header that specifies the Storage Index Manifest Mappings (with Manifest Mapping Extended GUID and Serial Number).");

            // Directly capture requirement MS-FSSHTTPB_R260, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     260,
                     @"[In Storage Index Data Element] Manifest Mapping Extended GUID (variable, optional): An Extended GUID that specifies the Manifest Mapping.");

            // Directly capture requirement MS-FSSHTTPB_R261, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     261,
                     @"[In Storage Index Data Element] Manifest Mapping Serial Number (variable, optional): A Serial Number that specifies the Manifest Mapping.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Storage Index Cell Mapping related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStorageIndexCellMapping(StorageIndexCellMapping instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the StorageIndexCellMapping related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StorageIndexCellMapping is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R262, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     262,
                     @"[In Storage Index Data Element] Storage Index Cell Mapping (2 bytes, optional): Zero or more 16-bit Stream Object Header that specifies the  Storage Index Cell Mappings (with cell identifier, cell mapping extended GUID, and Cell Mapping Serial Number).");

            // Directly capture requirement MS-FSSHTTPB_R263, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     263,
                     @"[In Storage Index Data Element] Cell ID (variable, optional): A Cell ID (section 2.2.1.10) that specifies the cell identifier.");

            // Directly capture requirement MS-FSSHTTPB_R264, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     264,
                     @"[In Storage Index Data Element] Cell Mapping Extended GUID (variable, optional): An Extended GUID that specifies the Cell Mapping.");

            // Directly capture requirement MS-FSSHTTPB_R265, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     265,
                     @"[In Storage Index Data Element] Cell Mapping Serial Number (variable, optional): A Serial Number that specifies the Cell Mapping.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Storage Index Revision Mapping related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStorageIndexRevisionMapping(StorageIndexRevisionMapping instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the StorageIndexRevisionMapping related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StorageIndexRevisionMapping is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R266, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     266,
                     @"[In Storage Index Data Element] Storage Index Revision Mapping (2 bytes, optional): Zero or more 16-bit Stream Object Headers that specify the Storage Index Revision Mappings (with revision and Revision Mapping Extended GUIDs, and Revision Mapping Serial Number).");

            // Directly capture requirement MS-FSSHTTPB_R267, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     267,
                     @"[In Storage Index Data Element] Revision Extended GUID (variable, optional): An Extended GUID that specifies the revision.");

            // Directly capture requirement MS-FSSHTTPB_R268, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     268,
                     @"[In Storage Index Data Element] Revision Mapping Extended GUID (variable, optional): An Extended GUID that specifies the Revision Mapping.");

            // Directly capture requirement MS-FSSHTTPB_R269, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     269,
                     @"[In Storage Index Data Element] Revision Mapping Serial Number (variable, optional): A Serial Number that specifies the Revision Mapping.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Storage Manifest Schema GUID related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStorageManifestSchemaGUID(StorageManifestSchemaGUID instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the StorageManifestSchemaGUID related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StorageManifestSchemaGUID is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R276, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     276,
                     @"[In Storage Manifest Data Element] Storage Manifest Schema GUID (2 bytes): A 16-bit Stream Object Header that specifies a Storage Manifest schema GUID.");

            // Directly capture requirement MS-FSSHTTPB_R277, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     277,
                     @"[In Storage Manifest Data Element] GUID (16 bytes): A GUID that specifies the schema.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Storage Manifest Root Declare related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyStorageManifestRootDeclare(StorageManifestRootDeclare instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the StorageManifestRootDeclare related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StorageManifestRootDeclare is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R278, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     278,
                     @"[In Storage Manifest Data Element] Storage Manifest Root Declare (2 bytes): A 16-bit Stream Object Header that specifies one or more Storage Manifest root declare(with Root Extended GUID and Cell ID).");

            // Directly capture requirement MS-FSSHTTPB_R279, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     279,
                     @"[In Storage Manifest Data Element] Root Extended GUID (variable): An Extended GUID that specifies the root Storage Manifest.");

            // Directly capture requirement MS-FSSHTTPB_R280, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     280,
                     @"[In Storage Manifest Data Element] Cell ID (variable): A Cell ID (section 2.2.1.10) that specifies the cell identifier.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test cell manifest current revision related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCellManifestCurrentRevision(CellManifestCurrentRevision instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the CellManifestCurrentRevision related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type CellManifestCurrentRevision is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R286, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     286,
                     @"[In Cell Manifest Data Element] Cell Manifest Current Revision (2 bytes): A 16-bit Stream Object Header that specifies a Cell Manifest current revision.");

            // Directly capture requirement MS-FSSHTTPB_R287, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     287,
                     @"[In Cell Manifest Data Element] Cell Manifest Current Revision Extended GUID (variable): An Extended GUID that specifies the revision.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test revision manifest related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyRevisionManifest(RevisionManifest instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the RevisionManifest related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type RevisionManifest is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R293, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     293,
                     @"[In Revision Manifest Data Elements] Revision Manifest (2 bytes): A 16-bit Stream Object Header that specifies a Revision Manifest.");

            // Directly capture requirement MS-FSSHTTPB_R294, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     294,
                     @"[In Revision Manifest Data Elements] Revision ID (variable): An Extended GUID that specifies the revision identifier represented by this data element.");

            // Directly capture requirement MS-FSSHTTPB_R295, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     295,
                     @"[In Revision Manifest Data Elements] Base Revision ID (variable): An Extended GUID that specifies the revision identifier of a base revision that could contain additional information for this revision.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test revision manifest root declaration related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyRevisionManifestRootDeclare(RevisionManifestRootDeclare instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the RevisionManifestRootDeclare related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type RevisionManifestRootDeclare is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R296, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     296,
                     @"[In Revision Manifest Data Elements] Revision Manifest Root Declare (2 bytes, optional): Zero or more 16-bit Stream Object Header that specifies a Revision Manifest root declare, each followed by root and object Extended GUIDs.");

            // Directly capture requirement MS-FSSHTTPB_R297, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     297,
                     @"[In Revision Manifest Data Elements] Root Extended GUID (optional, variable): An Extended GUID that specifies the root revision for each Revision Manifest Root Declare.");

            // Directly capture requirement MS-FSSHTTPB_R298, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     298,
                     @"[In Revision Manifest Data Elements] Object Extended GUID (optional, variable): An Extended GUID that specifies the object for each Revision Manifest Root Declare.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test revision manifest object group references related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyRevisionManifestObjectGroupReferences(RevisionManifestObjectGroupReferences instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the RevisionManifestObjectGroupReferences related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type RevisionManifestObjectGroupReferences is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R299, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     299,
                     @"[In Revision Manifest Data Elements] Revision Manifest Object Group References (2 bytes, optional): Zero or more 16-bit Stream Object Header that specifies a Revision Manifest Object Group references, each followed by Object Group Extended GUIDs.");

            // Directly capture requirement MS-FSSHTTPB_R300, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     300,
                     @"[In Revision Manifest Data Elements] Object Group Extended GUID (variable, optional): An Extended GUID that specifies the Object Group for each Revision Manifest Object Group Reference.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test object group declarations related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupDeclarations(ObjectGroupDeclarations instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the ObjectGroupDeclarations related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ObjectGroupDeclarations is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R308, if the stream object header is StreamObjectHeaderStart16bit or StreamObjectHeaderStart32bit.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     308,
                     @"[In Object Group Data Elements] Object Group Declarations Start (variable): A 16-bit (section 2.2.1.5.1) or 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an Object Group declaration start.");

            // Directly capture requirement MS-FSSHTTPB_R309, if there are no parsing errors. 
            if (instance.ObjectDeclarationList != null && instance.ObjectDeclarationList.Count != 0)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R309
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         309,
                         @"[In Object Group Data Elements] Object Declaration (variable): An optional array of Object Declarations (section 2.2.1.12.6.1) that specifies the object.");
            }
            
            if (instance.ObjectGroupObjectBLOBDataDeclarationList != null && instance.ObjectGroupObjectBLOBDataDeclarationList.Count != 0)
            {
                // Directly capture requirement MS-FSSHTTPB_R3091, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         3091,
                         @"[In Object Group Data Elements] Object Data BLOB Declaration (variable): An optional array of Object Data BLOB declarations (section 2.2.1.12.6.2) that specifies the object.");
            }
            
            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R310, if the stream object end is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     310,
                     @"[In Object Group Data Elements] Object Group Declarations End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies an Object Group declaration end.");
        }

        /// <summary>
        /// This method is used to test object group data related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupData(ObjectGroupData instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the ObjectGroupData related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ObjectGroupData is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R311, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     311,
                     @"[In Object Group Data Elements] Object Group Data Start (variable): A 16-bit or 32-bit Stream Object Header that specifies an Object Group data start.");

            if (instance.ObjectGroupObjectDataList != null && instance.ObjectGroupObjectDataList.Count != 0)
            {
                // Directly capture requirement MS-FSSHTTPB_R312, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         312,
                         @"[In Object Group Data Elements] Object Data (variable): An optional array of Object Data (section 2.2.1.12.6.4) that specifies the Object Data.");
            }
            
            if (instance.ObjectGroupObjectDataBLOBReferenceList != null && instance.ObjectGroupObjectDataBLOBReferenceList.Count != 0)
            {
                // Directly capture requirement MS-FSSHTTPB_R3121, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         3121,
                         @"[In Object Group Data Elements] Object Data BLOB Reference (variable): An optional array of Object Data BLOB references (section 2.2.1.12.6.5) that specifies the Object Data's references.");
            }

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R313, if the stream object end is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     313,
                     @"[In Object Group Data Elements] Object Group Data End (1 byte): An 8-bit Stream Object Header that specifies an Object Group data end.");
        }

        /// <summary>
        /// This method is used to test Object Declaration related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupObjectDeclare(ObjectGroupObjectDeclare instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the ObjectGroupObjectDeclare related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ObjectGroupObjectDeclare is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R315, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     315,
                     @"[In Object Declaration] Object Group Object Declaration (variable): A 16-bit (section 2.2.1.5.1) or 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an Object Group object declaration.");

            // Directly capture requirement MS-FSSHTTPB_R316, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     316,
                     @"[In Object Declaration] Object Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the object.");

            // Directly capture requirement MS-FSSHTTPB_R317, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     317,
                     @"[In Object Declaration] Object Partition ID (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the object partition of the object.");

            // Directly capture requirement MS-FSSHTTPB_R318, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     318,
                     @"[In Object Declaration] Object Data Size (variable): A compact unsigned 64-bit integer that specifies the size in bytes of the binary data opaque to this protocol[MS-FSSHTTPB] for the declared object.");

            // Directly capture requirement MS-FSSHTTPB_R319, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     319,
                     @"[In Object Declaration] Object References Count (variable): A compact unsigned 64-bit integer that specifies the number of object references.");

            // Directly capture requirement MS-FSSHTTPB_R320, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     320,
                     @"[In Object Declaration] Cell References Count (variable): A compact unsigned 64-bit integer that specifies the number of cell references.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Object Data BLOB Declaration related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupObjectBLOBDataDeclaration(ObjectGroupObjectBLOBDataDeclaration instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the DataElementFragment related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type DataElementFragment is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R321, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     321,
                     @"[In Object Data BLOB Declaration] Object Group Object Data BLOB Declaration (variable): A 16-bit (section 2.2.1.5.1) or 32-bit Stream Object Header section 2.2.1.5.2) that specifies an Object Group Object Data BLOB declaration.");

            // Directly capture requirement MS-FSSHTTPB_R322, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     322,
                     @"[In Object Data BLOB Declaration] Object Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the object.");

            // Directly capture requirement MS-FSSHTTPB_R323, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     323,
                     @"[In Object Data BLOB Declaration] Object Data BLOB EXGUID (variable): An Extended GUID that specifies the Object Data BLOB.");

            // Directly capture requirement MS-FSSHTTPB_R324, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     324,
                     @"[In Object Data BLOB Declaration] Object Partition ID (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the object partition of the object.");

            // Directly capture requirement MS-FSSHTTPB_R326, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     326,
                     @"[In Object Data BLOB Declaration] Object References Count (variable): A compact unsigned 64-bit integer that specifies the number of object references.");

            // Directly capture requirement MS-FSSHTTPB_R327, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     327,
                     @"[In Object Data BLOB Declaration] Cell References Count (variable): A compact unsigned 64-bit integer that specifies the number of cell references.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Object Metadata Declaration related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupMetadataDeclarations(ObjectGroupMetadataDeclarations instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Object Metadata Declaration related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ObjectGroupMetadataDeclarations is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R2108, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2108,
                     @"[In Object Metadata Declaration] Object Group Metadata Declarations (variable): 32-bit Stream Object Header section 2.2.1.5.2) that specifies an Object Group metadata declarations.");

            // Directly capture requirement MS-FSSHTTPB_R2109, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2109,
                     @"[In Object Metadata Declaration] Object Metadata (variable): An array of Object Metadata (section 2.2.1.12.6.3.1) that specifies the object metadata.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R2110, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2110,
                     @"[In Object Metadata Declaration] Object Group Metadata Declarations End (2 byte): An 16-bit Stream Object Header (section 2.2.1.5.4) that specifies the end of Object Group metadata declarations.");

            if (Common.Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 21072, site))
            {
                // Directly capture requirement MS-FSSHTTPB_R2110, if the code runs here.
                site.CaptureRequirement(
                    "MS-FSSHTTPB",
                    21072,
                    @"[In Appendix B: Product Behavior] Implementation does support an Object Metadata Declaration.(SharePoint Server 2013 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to test Object Metadata related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupMetadata(ObjectGroupMetadata instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Object Metadata related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ObjectGroupMetadata is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R2112, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2112,
                     @"[In Object Metadata] Object Group Metadata (variable): 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an Object Group metadata.");

            // Directly capture requirement MS-FSSHTTPB_R2113, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     2113,
                     @"[In Object Metadata] Object Change Frequency (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the expected change frequency of the object.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Object Data related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupObjectData(ObjectGroupObjectData instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Object Data related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ObjectGroupObjectData is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture the requirement MS-FSSHTTPB_R328, if there are no parsing errors.
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     328,
                     @"[In Object Data] Object Group Object Data (variable): A 16-bit (section 2.2.1.5.1) or 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an Object Group Object Data.");

            // Directly capture requirement MS-FSSHTTPB_R329, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     329,
                     @"[In Object Data] Object Extended GUID Array (variable): An Extended GUID array (section 2.2.1.8) that specifies the Object Group.");

            // Directly capture requirement MS-FSSHTTPB_R330, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     330,
                     @"[In Object Data] Cell ID Array (variable): A Cell ID Array (section 2.2.1.11) that specifies the Object Group.");

            // Directly capture requirement MS-FSSHTTPB_R331, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     331,
                     @"[In Object Data] Data (variable): A Binary Item (section 2.2.1.3) that specifies the binary data that is opaque to this protocol[MS-FSSHTTPB] in the case of an Object Group Object Data.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Object Data BLOB Reference related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectGroupObjectDataBLOBReference(ObjectGroupObjectDataBLOBReference instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the DataElementFragment related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type DataElementFragment is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R332, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     332,
                     @"[In Object Data BLOB Reference] Object Group Object Data BLOB Reference (variable): A 16-bit (section 2.2.1.5.1) or 32-bit Stream Object Header (section 2.2.1.5.2) that specifies an Object Group Object Data BLOB reference (section 2.2.1.12.6.5).");

            // Directly capture requirement MS-FSSHTTPB_R333, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     333,
                     @"[In Object Data BLOB Reference] Object Extended GUID Array (variable): An Extended GUID Array (section 2.2.1.8) that specifies the object references.");

            // Directly capture requirement MS-FSSHTTPB_R334, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     334,
                     @"[In Object Data BLOB Reference] Cell ID Array (variable): A Cell ID Array (section 2.2.1.11) that specifies the cell references.");

            // Directly capture requirement MS-FSSHTTPB_R335, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     335,
                     @"[In Object Data BLOB Reference] BLOB Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the Object Data BLOB.");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Data Element Fragment related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyDataElementFragment(DataElementFragment instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the DataElementFragment related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type DataElementFragment is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R340, if there are no parsing errors. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     340,
                     @"[In Data Element Fragment Data Elements] Data Element Fragment (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a Data Element Fragment.");

            // Directly capture requirement MS-FSSHTTPB_R341, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     341,
                     @"[In Data Element Fragment Data Elements] Fragment Extended GUID (variable): An Extended GUID that specifies the Data Element Fragment.");

            // Directly capture requirement MS-FSSHTTPB_R342, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     342,
                     @"[In Data Element Fragment Data Elements] Fragment Data Element Size (variable): A compact unsigned 64-bit integer that specifies the size in bytes of the fragmented data element.");

            // Directly capture requirement MS-FSSHTTPB_R343, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     343,
                     @"[In Data Element Fragment Data Elements] Fragment File Chunk Reference (variable): A File Chunk Reference (section 2.2.1.2) that specifies the Data Element Fragment.");

            // Directly capture requirement MS-FSSHTTPB_R344, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     344,
                     @"[In Data Element Fragment Data Elements] Fragment Data (variable): A byte stream that specifies the binary data opaque to this protocol[MS-FSSHTTPB].");

            // Verify the stream object header related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method aims to test object data blob related adapter requirements, but No source code is needed for this method and the method is needed for reflection.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyObjectDataBLOB(ObjectDataBLOB instance, ITestSite site)
        {
            // No source code is needed for this method, but the method is needed for reflection
        }

        #region Private method

        /// <summary>
        /// This method is used to test Storage Index Data Element related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyStorageIndexDataElement(DataElement instance, ITestSite site)
        {
            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R255, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     255,
                     @"[In Storage Index Data Element] Data Element Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a data element start.");

            // Directly capture requirement MS-FSSHTTPB_R256, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     256,
                     @"[In Storage Index Data Element] Data Element Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R257, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     257,
                     @"[In Storage Index Data Element] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R258, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     258,
                     @"[In Storage Index Data Element] Data Element Type (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the value of the Storage Index data element type.");

            // Directly capture requirement MS-FSSHTTPB_R99007, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     99007,
                     @"[In Storage Index Data Element] When serializing a Storage Index data element, there is no sequence for Storage Index Manifest Mapping, Storage Index Cell Mapping and Storage Index Revision Mapping.");

            // Directly capture requirement MS-FSSHTTPB_R4011, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     4011,
                     @"[In Storage Index Data Element] Additionally, the Storage Index contains a set of mappings, and each mapping is assigned a serial number that is unique to that mapping.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R270, if the stream object end header is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     270,
                     @"[In Storage Index Data Element] Data Element End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a data element end.");
        }

        /// <summary>
        /// This method is used to test Storage Manifest Data Element related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyStorageManifestDataElement(DataElement instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Storage Manifest Data Element related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type StorageManifestDataElementData is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R272, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     272,
                     @"[In Storage Manifest Data Element] Data Element Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a data element start.");

            // Directly capture requirement MS-FSSHTTPB_R273, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     273,
                     @"[In Storage Manifest Data Element] Data Element Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R274, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     274,
                     @"[In Storage Manifest Data Element] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R275, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     275,
                     @"[In Storage Manifest Data Element] Data Element Type (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the value of the Storage Manifest data element type.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R281, if the stream object end is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     281,
                     @"[In Storage Manifest Data Element] Data Element End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a data element end.");
        }

        /// <summary>
        /// This method is used to test Cell Manifest Data Element related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyCellManifestDataElement(DataElement instance, ITestSite site)
        {
            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R282, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     282,
                     @"[In Cell Manifest Data Element] Data Element Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a data element start.");

            // Directly capture requirement MS-FSSHTTPB_R283, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     283,
                     @"[In Cell Manifest Data Element] Data Element Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R284, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     284,
                     @"[In Cell Manifest Data Element] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R285, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     285,
                     @"[In Cell Manifest Data Element] Data Element Type (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the value of the Cell Manifest data element type");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R288, if there are no parsing errors. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     288,
                     @"[In Cell Manifest Data Element] Data Element End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a data element end.");
        }

        /// <summary>
        /// This method is used to test Revision Manifest Data Elements related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyRevisionManifestDataElement(DataElement instance, ITestSite site)
        {
            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R289, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     289,
                     @"[In Revision Manifest Data Elements] Data Element Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a data element start.");

            // Directly capture requirement MS-FSSHTTPB_R290, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     290,
                     @"[In Revision Manifest Data Elements] Data Element Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R291, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     291,
                     @"[In Revision Manifest Data Elements] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R292, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     292,
                     @"[In Revision Manifest Data Elements] Data Element Type (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the value of the Revision Manifest data element type.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R301, if the stream object end is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     301,
                     @"[In Revision Manifest Data Elements] Data Element End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a data element end.");
        }

        /// <summary>
        /// This method is used to test Object Group Data Elements related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyObjectGroupDataElement(DataElement instance, ITestSite site)
        {
            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R302, if the stream object header is StreamObjectHeaderStart16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     302,
                     @"[In Object Group Data Elements] Data Element Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a data element start.");

            // Directly capture requirement MS-FSSHTTPB_R303, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     303,
                     @"[In Object Group Data Elements] Data Element Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R304, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     304,
                     @"[In Object Group Data Elements] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R305, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     305,
                     @"[In Object Group Data Elements] Data Element Type (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the value of the Object Group data element type.");

            if (instance.GetData<ObjectGroupDataElementData>().ObjectMetadataDeclaration != null)
            {
                // Directly capture requirement MS-FSSHTTPB_R2103, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         2103,
                         @"[In Object Group Data Elements] Object Metadata Declaration (variable): If Object Metadata (section 2.2.1.12.6.3.1) exists, this field MUST specify an Object Metadata Declaration (section 2.2.1.12.6.3).");
            }

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R314, if the stream object end is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     314,
                     @"[In Object Group Data Elements] Data Element End (1 byte): An 8-bit Stream Object Header that specifies a data element end.");
        }

        /// <summary>
        /// This method is used to test Object Data BLOB Data Elements related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyObjectDataBLOBDataElement(DataElement instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the DataElementFragment related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type DataElementFragment is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R346, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     346,
                     @"[In Object Data BLOB Data Elements] Data Element Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a data element start.");

            // Directly capture requirement MS-FSSHTTPB_R347, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     347,
                     @"[In Object Data BLOB Data Elements] Data Element Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R348, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     348,
                     @"[In Object Data BLOB Data Elements] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R349, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     349,
                     @"[In Object Data BLOB Data Elements] Data Element Type (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the value of the Object Data BLOB data element type.");

            // Directly capture requirement MS-FSSHTTPB_R350, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     350,
                     @"[In Object Data BLOB Data Elements] Object Data BLOB (variable): A 16-bit or 32-bit Stream Object Header that specifies an Object Data BLOB.");

            // Directly capture requirement MS-FSSHTTPB_R351, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     351,
                     @"[In Object Data BLOB Data Elements] Data (variable): A byte stream that specifies the binary data opaque to this protocol[MS-FSSHTTPB].");

            // Directly capture requirement MS-FSSHTTPB_R352, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     352,
                     @"[In Object Data BLOB Data Elements] Data Element End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a data element end.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Data Element Fragment Data Elements related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param> 
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyFragmentDataElement(DataElement instance, ITestSite site)
        {
            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Directly capture requirement MS-FSSHTTPB_R336, if the stream object header is StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     336,
                     @"[In Data Element Fragment Data Elements] Data Element Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a data element start.");

            // Directly capture requirement MS-FSSHTTPB_R337, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     337,
                     @"[In Data Element Fragment Data Elements] Data Element Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R338, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     338,
                     @"[In Data Element Fragment Data Elements] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R339, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     339,
                     @"[In Data Element Fragment Data Elements] Data Element Type (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies the value of the Object Data BLOB data element type (section 2.2.1.12.8).");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R345, if there are no parsing errors. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     345,
                     @"[In Data Element Fragment Data Elements] Data Element End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a data element end.");
        }
        #endregion
    }
}