namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Globalization;
    using Microsoft.Protocols.TestTools;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This is the partial part of the class MsfsshttpbAdapterCapture for MS-FSSHTTPB knowledge part.
    /// </summary>
    public partial class MsfsshttpbAdapterCapture
    {
        /// <summary>
        /// This method is used to verify knowledge related requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyKnowledge(Knowledge instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Knowledge related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type Knowledge is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R359, if stream object start type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     359,
                     @"[In Knowledge] Knowledge Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a Knowledge (section 2.2.1.13) start.");

            // Directly capture requirement MS-FSSHTTPB_R360, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     360,
                     @"[In Knowledge] Specialized Knowledge (variable): Zero or more Specialized Knowledge structures (section 2.2.1.13.1).");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Directly capture requirement MS-FSSHTTPB_R361, if stream object end type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     361,
                     @"[In Knowledge] Knowledge End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies a Knowledge end.");
        }

        /// <summary>
        /// This method is used to test Specialized Knowledge related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifySpecializedKnowledge(SpecializedKnowledge instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Specialized Knowledge related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type SpecializedKnowledge is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R362, if stream object start type is StreamObjectHeaderStart32bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     362,
                     @"[In Specialized Knowledge] Specialized Knowledge Start (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies A Specialized Knowledge start.");

            // Directly capture requirement MS-FSSHTTPB_R363, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     363,
                     @"[In Specialized Knowledge] GUID (16 bytes): A GUID that specifies the type of Specialized Knowledge.");

            bool isVerifyR364 = instance.GUID == SpecializedKnowledge.CellKnowledgeGuid ||
                                instance.GUID == SpecializedKnowledge.ContentTagKnowledgeGuid ||
                                instance.GUID == SpecializedKnowledge.WaterlineKnowledgeGuid ||
                                instance.GUID == SpecializedKnowledge.FragmentKnowledgeGuid;

            site.Log.Add(
                         LogEntryKind.Debug,
                        "Actual GUID value {0}, expect the value either 327A35F6-0761-4414-9686-51E900667A4D, 3A76E90E-8032-4D0C-B9DD-F3C65029433E, 0ABE4F35-01DF-4134-A24A-7C79F0859844 or 10091F13-C882-40FB-9886-6533F934C21D or ,BF12E2C1-E64F-4959-8282-73B9A24A7C44 for MS-FSSHTTPB_R364.",
                         instance.GUID.ToString());

            // Capture requirement MS-FSSHTTPB_R364, if the GUID equals the mentioned four values {327A35F6-0761-4414-9686-51E900667A4D}, {3A76E90E-8032-4D0C-B9DD-F3C65029433E}, {0ABE4F35-01DF-4134-A24A-7C79F0859844}, {10091F13-C882-40FB-9886-6533F934C21D}.
            site.CaptureRequirementIfIsTrue(
                     isVerifyR364,
                     "MS-FSSHTTPB",
                     364,
                     @"[In Specialized Knowledge] The following GUIDs detail the type of Knowledge contained: [Its value must be one of] {327A35F6-0761-4414-9686-51E900667A4D}, {3A76E90E-8032-4D0C-B9DD-F3C65029433E}, {0ABE4F35-01DF-4134-A24A-7C79F0859844}, {10091F13-C882-40FB-9886-6533F934C21D},{BF12E2C1-E64F-4959-8282-73B9A24A7C44}].");

            switch (instance.GUID.ToString("D").ToUpper(CultureInfo.CurrentCulture))
            {
                case "327A35F6-0761-4414-9686-51E900667A4D":

                    // Capture requirement MS-FSSHTTPB_R365, if the knowledge data type is CellKnowledge.
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(CellKnowledge),
                             instance.SpecializedKnowledgeData.GetType(),
                             "MS-FSSHTTPB",
                             365,
                             @"[In Specialized Knowledge][If the GUID field is set to ] {327A35F6-0761-4414-9686-51E900667A4D}, [it indicates the type of the Specialized Knowledge is]Cell Knowledge (section 2.2.1.13.2).");
                    break;

                case "3A76E90E-8032-4D0C-B9DD-F3C65029433E":

                    // Capture requirement MS-FSSHTTPB_R366, if the knowledge data type is WaterlineKnowledge.
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(WaterlineKnowledge),
                             instance.SpecializedKnowledgeData.GetType(),
                             "MS-FSSHTTPB",
                             366,
                             @"[In Specialized Knowledge][If the GUID field is set to ] {3A76E90E-8032-4D0C-B9DD-F3C65029433E}, [it indicates the type of the specialized knowledge is]Waterline Knowledge (section 2.2.1.13.4).");

                    break;

                case "0ABE4F35-01DF-4134-A24A-7C79F0859844":

                    // Capture requirement MS-FSSHTTPB_R367, if the knowledge data type is FragmentKnowledge.
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(FragmentKnowledge),
                             instance.SpecializedKnowledgeData.GetType(),
                             "MS-FSSHTTPB",
                             367,
                             @"[In Specialized Knowledge][If the GUID field is set to ] {0ABE4F35-01DF-4134-A24A-7C79F0859844}, [it indicates the type of the specialized knowledge is]Fragment Knowledge (section 2.2.1.13.3).");

                    break;

                case "10091F13-C882-40FB-9886-6533F934C21D":

                    // Capture requirement MS-FSSHTTPB_R368, if the knowledge data type is ContentTagKnowledge.
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(ContentTagKnowledge),
                             instance.SpecializedKnowledgeData.GetType(),
                             "MS-FSSHTTPB",
                             368,
                             @"[In Specialized Knowledge][If the GUID field is set to ] {10091F13-C882-40FB-9886-6533F934C21D}, [it indicates the type of the specialized knowledge is]Content Tag Knowledge (section 2.2.1.13.5).");
                    break;

                case "BF12E2C1-E64F-4959-8282-73B9A24A7C44":

                    // Capture requirement MS-FSSHTTPB_R1212, if the knowledge data type is VersionTokenKnowledge.
                    site.CaptureRequirementIfAreEqual<Type>(
                             typeof(VersionTokenKnowledge),
                             instance.SpecializedKnowledgeData.GetType(),
                             "MS-FSSHTTPB",
                             1212,
                             @"[In Specialized Knowledge][If the GUID field is set to ] {BF12E2C1-E64F-4959-8282-73B9A24A7C44}, [it indicates the type of the specialized knowledge is]Version Token Knowledge (section 2.2.1.13.6).");
                    break;

                default:
                    site.Assert.Fail("Unsupported specialized knowledge value " + instance.GUID.ToString());
                    break;
            }

            // Directly capture requirement MS-FSSHTTPB_R369, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     369,
                     @"[In Specialized Knowledge] Specialized Knowledge Data (variable): The data for the specific Knowledge type.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R370, if the stream object end type is StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd16bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     370,
                     @"[In Specialized Knowledge] Specialized Knowledge End (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.4) that specifies Specialized Knowledge end.");
        }

        /// <summary>
        /// This method is used to test Cell Knowledge related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCellKnowledge(CellKnowledge instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Cell Knowledge related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type CellKnowledge is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R372, if stream object start type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     372,
                     @"[In Cell Knowledge] Cell Knowledge Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a Cell Knowledge start.");

            if (instance.CellKnowledgeEntryList != null && instance.CellKnowledgeEntryList.Count != 0)
            {
                // Directly capture requirement MS-FSSHTTPB_R373, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         373,
                         @"[In Cell Knowledge] Cell Knowledge Data (variable): An array of Knowledge Entry (section 2.2.1.13.2.2) that specifies one data element Knowledge reference.");
            }
            else if (instance.CellKnowledgeRangeList != null && instance.CellKnowledgeRangeList.Count != 0)
            {
                // Directly capture requirement MS-FSSHTTPB_R3731, if there are no parsing errors. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         3731,
                         @"[In Cell Knowledge] Cell Knowledge Data (variable): An array of Knowledge Range (section 2.2.1.13.2.1) that specifies one or more data element Knowledge references.");
            }
            else
            {
                site.Log.Add(LogEntryKind.Debug, "The CellKnowledgeEntryList and CellKnowledgeRangeList are same null or empty list in the same time.");
            }

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);

            // Capture requirement MS-FSSHTTPB_R374, if the stream object end header is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     374,
                     @"[In Cell Knowledge] Cell Knowledge End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies the Cell Knowledge end.");
        }

        /// <summary>
        /// This method is used to test Cell Knowledge Range related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCellKnowledgeRange(CellKnowledgeRange instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Cell Knowledge Range related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type CellKnowledgeRange is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R376, if stream object start type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     376,
                     @"[In Cell Knowledge Range] Cell Knowledge Range (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies the start of a Cell Knowledge Range.");

            // Directly capture requirement MS-FSSHTTPB_R377, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     377,
                     @"[In Cell Knowledge Range] GUID (16 bytes): A GUID that specifies the data element.");

            // Directly capture requirement MS-FSSHTTPB_R378, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     378,
                     @"[In Cell Knowledge Range] From (variable): A compact unsigned 64-bit integer section 2.2.1.1() that specifies the starting Sequence Number.");

            // Directly capture requirement MS-FSSHTTPB_R558, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     558,
                     @"[In Cell Knowledge Range] To (variable): A compact unsigned 64-bit integer that specifies the ending sequence number.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Cell Knowledge Entry related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyCellKnowledgeEntry(CellKnowledgeEntry instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Cell Knowledge Entry related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type CellKnowledgeEntry is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R560, if stream object start type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     560,
                     @"[In Cell Knowledge Entry] Cell Knowledge Entry (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1), that specifies a Cell Knowledge Entry.");

            // Directly capture requirement MS-FSSHTTPB_R561, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     561,
                     @"[In Cell Knowledge Entry] Serial Number (variable): A Serial Number (section 2.2.1.9) that specifies the cell.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Fragment Knowledge related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyFragmentKnowledge(FragmentKnowledge instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Fragment Knowledge related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type FragmentKnowledge is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R563, if stream object start type is StreamObjectHeaderStart32bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     563,
                     @"[In Fragment Knowledge] Fragment Knowledge Start (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a Fragment Knowledge start.");

            // Directly capture requirement MS-FSSHTTPB_R564, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     564,
                     @"[In Fragment Knowledge] Fragment Knowledge Entries (variable): An optional array of Fragment Knowledge Entry (section 2.2.1.13.3.1) structures specifying the fragments which have been uploaded.");

            // Directly capture requirement MS-FSSHTTPB_R565, if the stream object end is StreamObjectHeaderEnd16bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd16bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     565,
                     @"[In Fragment Knowledge] Fragment Knowledge End (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.4) that specifies a Fragment Knowledge end.");

            // Verify the stream object header end related requirements.
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
        }

        /// <summary>
        /// This method is used to test Fragment Knowledge Entry related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyFragmentKnowledgeEntry(FragmentKnowledgeEntry instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Fragment Knowledge Entry related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type FragmentKnowledgeEntry is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R566, if stream object start type is StreamObjectHeaderStart32bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart32bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     566,
                     @"[In Fragment Knowledge Entry] Fragment Descriptor (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a Fragment Knowledge Entry.");

            // Directly capture requirement MS-FSSHTTPB_R567, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     567,
                     @"[In Fragment Knowledge Entry] Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the data element this Fragment Knowledge Entry contains knowledge about.");

            // Directly capture requirement MS-FSSHTTPB_R568, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     568,
                     @"[In Fragment Knowledge Entry] Data Element Size (variable): A compact unsigned 64-bit integer (section 2.2.1.1) specifying the size in bytes of the data element specified by the preceding Extended GUID.");

            // Directly capture requirement MS-FSSHTTPB_R569, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     569,
                     @"[In Fragment Knowledge Entry] Data Element Chunk Reference (variable): A file chunk reference (section 2.2.1.2) specifying which part of the data element with the preceding GUID this Fragment Knowledge Entry contains knowledge about.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Waterline Knowledge related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyWaterlineKnowledge(WaterlineKnowledge instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Waterline Knowledge related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type WaterlineKnowledge is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R572, if stream object start type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     572,
                     @"[In Waterline Knowledge] Waterline Knowledge Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a Waterline Knowledge start.");

            // Directly capture requirement MS-FSSHTTPB_R573, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     573,
                     @"[In Waterline Knowledge] Waterline Knowledge Data (variable): One or more Waterline Knowledge Entries (section 2.2.1.13.4.1) that specify what the server has already delivered to the client or what the client has already received from the server.");

            // Directly capture requirement MS-FSSHTTPB_R574, if the stream object end type is StreamObjectHeaderEnd8bit.
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderEnd8bit),
                     instance.StreamObjectHeaderEnd.GetType(),
                     "MS-FSSHTTPB",
                     574,
                     @"[In Waterline Knowledge] Waterline Knowledge End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies the Waterline Knowledge end.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Waterline Knowledge Entry related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyWaterlineKnowledgeEntry(WaterlineKnowledgeEntry instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Waterline Knowledge Entry related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type WaterlineKnowledgeEntry is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R575, if stream object start type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     575,
                     @"[In Waterline Knowledge Entry] Waterline Knowledge Entry (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies a Waterline Knowledge Entry.");

            // Directly capture requirement MS-FSSHTTPB_R576, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     576,
                     @"[In Waterline Knowledge Entry] Cell Storage Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the cell storage this entry specifies the waterline for.");

            // Directly capture requirement MS-FSSHTTPB_R577, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     577,
                     @"[In Waterline Knowledge Entry] Waterline (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies a sequential Serial Number (section 2.2.1.9).");

            // Directly capture requirement MS-FSSHTTPB_R379, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     379,
                     @"[In Waterline Knowledge Entry] Reserved (variable): A compact unsigned 64-bit integer that specifies a reserved field that MUST have value of zero.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Content Tag Knowledge related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyContentTagKnowledge(ContentTagKnowledge instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Content Tag Knowledge related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ContentTagKnowledge is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

            // Capture requirement MS-FSSHTTPB_R381, if stream object start type is StreamObjectHeaderStart16bit. 
            site.CaptureRequirementIfAreEqual<Type>(
                     typeof(StreamObjectHeaderStart16bit),
                     instance.StreamObjectHeaderStart.GetType(),
                     "MS-FSSHTTPB",
                     381,
                     @"[In Content Tag Knowledge] Content Tag Start (2 bytes): A 16-bit Stream Object Header (section 2.2.1.5.1) that specifies the Content Tag Knowledge start.");

            // Directly capture requirement MS-FSSHTTPB_R382, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     382,
                     @"[In Content Tag Knowledge] Content Tag Entry Array (variable): An array of Content Tag Knowledge Entry structures (section 2.2.1.13.5.1), each of which specifies changes for a BLOB.");

            // Directly capture requirement MS-FSSHTTPB_R383, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     383,
                     @"[In Content Tag Knowledge] Content Tag Knowledge End (1 byte): An 8-bit Stream Object Header (section 2.2.1.5.3) that specifies the Content Tag Knowledge end.");

            // Verify the stream object header end related requirements.
            this.ExpectStreamObjectHeaderEnd(instance.StreamObjectHeaderEnd, instance.GetType(), site);
            this.ExpectCompoundObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Content Tag Knowledge Entry related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyContentTagKnowledgeEntry(ContentTagKnowledgeEntry instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Content Tag Knowledge Entry related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type ContentTagKnowledgeEntry is null due to parsing error or type casting error.");
            }

            // Verify the stream object header related requirements.
            this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);

           if(instance.StreamObjectHeaderStart.GetType() == typeof(StreamObjectHeaderStart16bit) || instance.StreamObjectHeaderStart.GetType() == typeof(StreamObjectHeaderStart32bit))
            {
                // Capture requirement MS-FSSHTTPB_R385, if stream object start type is StreamObjectHeaderStart16bit or StreamObjectHeaderStart32bit. 
                site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         385,
                         @"[In Content Tag Knowledge Entry] Content Tag Knowledge Entry Start (variable): A 16-bit (section 2.2.1.5.1) or 32-bit Stream Object Header (section 2.2.1.5.2) that specifies the start of a Content Tag Knowledge Entry.");
            }

            // Directly capture requirement MS-FSSHTTPB_R386, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     386,
                     @"[In Content Tag Knowledge Entry] BLOB Heap Extended GUID (variable): An Extended GUID (section 2.2.1.7) that specifies the BLOB this content tag is for.");

            // Directly capture requirement MS-FSSHTTPB_R387, if there are no parsing errors. 
            site.CaptureRequirement(
                     "MS-FSSHTTPB",
                     387,
                     @"[In Content Tag Knowledge Entry] Clock Data (variable): A binary item (section 2.2.1.3) that specifies changes for a BLOB on the server.");

            // Verify the stream object header end related requirements.
            this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
        }

        /// <summary>
        /// This method is used to test Version Token Knowledge related adapter requirements.
        /// </summary>
        /// <param name="instance">Specify the instance which need to be verified.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public void VerifyVersionTokenKnowledge(VersionTokenKnowledge instance, ITestSite site)
        {
            // If the instance is not null and there are no parsing errors, then the Version Token Knowledge related adapter requirements can be directly captured.
            if (null == instance)
            {
                site.Assert.Fail("The instance of type VersionTokenKnowledge is null due to parsing error or type casting error.");
            }

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R1345
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1345, site))
            {
                // Capture requirement MS-FSSHTTPB_R1345, if instance type is VersionTokenKnowledge. 
                site.CaptureRequirementIfAreEqual<Type>(
                        typeof(VersionTokenKnowledge),
                        instance.GetType(),
                        "MS-FSSHTTPB",
                        1345,
                        @"[In Appendix A: Product Behavior]Implementation does support the Version Token Knowledge(SharePoint Server 2016 and above follow this behavior.)");

                // Capture requirement MS-FSSHTTPB_R1214, if stream object start type is StreamObjectHeaderStart32bit. 
                site.CaptureRequirementIfAreEqual<Type>(
                         typeof(StreamObjectHeaderStart32bit),
                         instance.StreamObjectHeaderStart.GetType(),
                         "MS-FSSHTTPB",
                         1214,
                         @"[In Version Token Knowledge]Version Token Knowledge (4 bytes): A 32-bit Stream Object Header (section 2.2.1.5.2) that specifies a Version Token Knowledge.");

                // Capture requirement MS-FSSHTTPB_R1215, if TokenData type is BinaryItem. 
                site.CaptureRequirementIfAreEqual<Type>(
                         typeof(BinaryItem),
                         instance.TokenData.GetType(),
                         "MS-FSSHTTPB",
                         1215,
                         @"[In Version Token Knowledge]Token Data (variable): A byte stream that specifies the version token opaque to this protocol.");

                // Verify the stream object header related requirements.
                this.ExpectStreamObjectHeaderStart(instance.StreamObjectHeaderStart, instance.GetType(), site);
                this.ExpectSingleObject(instance.StreamObjectHeaderStart, site);
            }
        }
    }
}