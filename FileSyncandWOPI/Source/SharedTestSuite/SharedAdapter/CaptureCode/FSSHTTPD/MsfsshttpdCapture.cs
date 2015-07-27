//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains methods which capture requirements related with MS-FSSHTTPD.
    /// </summary>
    public class MsfsshttpdCapture
    {
        /// <summary>
        /// Prevents a default instance of the MsfsshttpdCapture class from being created
        /// </summary>
        private MsfsshttpdCapture()
        { 
        }

        /// <summary>
        /// This method is used to verify the requirements related with the simple chunk method.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifySimpleChunk(ITestSite site)
        {
            // If run here, all the requirements related to simple chunk are carefully examined, so all the requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     129,
                     @"[In Simple Chunking Method] Simple Small File Hash: A 20-byte sequence that specifies the SHA-1 hash code of the file bytes represented by the chunk.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R134
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     134,
                     @"[In Simple Chunking Method] The Signature Data of the Intermediate Node Object MUST be the chunk’s signature.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R133
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     133,
                     @"[In Simple Chunking Method] The Data Size of the Intermediate Node Object MUST be the total number of bytes represented by the chunk.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R128
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     128,
                     @"[In Simple Chunking Method] Files are split into chunks that are each 1 megabyte in size.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R138
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     138,
                     @"[In Simple Chunking Method] The Object Reference Array and the Cell Reference Array of the Data Node Object MUST be empty.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R136
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     136,
                     @"[In Simple Chunking Method] For all Intermediate Node Objects, a Data Node Object MUST be created.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R139
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     139,
                     @"[In Simple Chunking Method] The Object References Array of the Intermediate Node Object associated with this Data Node Object MUST have a single entry, which MUST be the Object ID of the Data Node Object.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R132
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     132,
                     @"[In Simple Chunking Method] For each chunk in the chunk list, an Intermediate Node Object, as specified in section 2.2.3, is created.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R135
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     135,
                     @"[In Simple Chunking Method] The Intermediate Node is referenced by its parent node.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8201
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8201,
                     @"[In Simple Chunking Method] The Object Data of the Data Node Object MUST be the byte sequence from the file tracked by the chunk.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with the RDC analysis chunk method.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyRdcAnalysisChunk(ITestSite site)
        {
            // If runs here successfully, then indicating all the following requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8063,
                     @"[In Generating Chunks] Files are split into chunks by using the RDC FilterMax algorithm, as specified in [MS-RDC] section 3.1.5.1, using a hash window of 48 and a horizon of 16,384.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with zip chunk when the size of zip header and zip content is less than 4096 bytes. 
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyZipFileLessThan4096Bytes(ITestSite site)
        {
            // If runs here successfully, then indicating all the following requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8033,
                     @"[In Zip Files][Data file chunk structure] Uncompressed Size (8 bytes): If the combined size, in bytes, of the local file header chunk and the data file chunk is less than 4,096, a single chunk is produced with a signature that is the local file header chunk signature followed by the data file chunk signature.");

            if (SharedContext.Current.CellStorageVersionType.MinorVersion >= 2)
            {
                // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8034
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8034,
                         @"[In Zip Files][Data file chunk structure] Uncompressed Size (8 bytes): For protocol servers with VersionNumberType, as specified in [MS-FSSHTTP] section 2.2.5.13, greater or equal to  2 and MinorVersionNumberType, as specified in [MS-FSSHTTP] section 2.2.5.10, greater or equal to 2, the signature for the single chunk is a bitwise exclusive OR of the signature bytes of the local file header chunk and the data file chunk.");

                // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8036
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8036,
                         @"[In Zip Files][Data file chunk structure] Uncompressed Size (8 bytes): If the signatures are not of equal length, the extra bytes of the longer signature are appended to the end of the exclusive ORed bytes. ");
            }
            else
            {
                // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8078
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8078,
                         @"[In Zip Files][Data file chunk structure] Uncompressed Size (8 bytes): If the combined size, in bytes, of the local file header chunk and the data file chunk is equal to 4,096, a single chunk is produced with a signature that is the local file header chunk signature followed by the data file chunk signature.");
            }
        }

        /// <summary>
        /// This method is used to verify the requirements related with zip chunk when the size of zip header and zip content is larger than 4096 bytes but less than 1MB.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyZipFileHeaderAndContentSignature(ITestSite site)
        {
            // If runs here successfully, then indicating all the following requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8026,
                     @"[In Zip Files] Local File Header Hash: A 20-byte sequence that specifies the SHA-1 hash code of the file bytes represented by the local file header chunk.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8028
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8028,
                     @"[In Zip Files][Data file chunk structure] CRC (4 bytes): An unsigned 32-bit integer that specifies the value of the local file header crc-32 field, as specified in [PKWARE-Zip].");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8029
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8029,
                     @"[In Zip Files][Data file chunk structure] Compressed Size (8 bytes): An unsigned 64-bit integer that specifies the size, in bytes, of the data file chunk. ");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8030
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8030,
                     @"[In Zip Files][Data file chunk structure] Compressed Size (8 bytes): It MUST be the value of the local file header compressed size field, as specified in [PKWARE-Zip], if the local file header extra field does not include a Zip64 Extended Information Extra Field, as specified in [PKWARE-Zip].");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8031
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8031,
                     @"[In Zip Files][Data file chunk structure] Uncompressed Size (8 bytes): An unsigned 64-bit integer that specifies the size, in bytes, of the uncompressed data represented by the bytes of the data file chunk.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8032
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8032,
                     @"[In Zip Files][Data file chunk structure] Uncompressed Size (8 bytes): It MUST be the value of the local file header uncompressed size field, as specified in [PKWARE-Zip], if the local file header extra field does not include a Zip64 Extended Information Extra Field, as specified in [PKWARE-Zip].");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8046
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8046,
                     @"[In Zip Files] Large Final Chunk Signature: The Signature Data of the Intermediate Node Object MUST be the chunk’s signature.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with zip chunk when the final chunk is less than 1MB. 
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifySmallFinalChunk(ITestSite site)
        {
            // If runs here successfully, then indicating all the following requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8040,
                     @"[In Zip Files][Data file chunk structure] After the analysis of local file headers terminates, the remaining bytes in the file are represented by a final chunk.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8087
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8087,
                     @"[In Zip Files] If the total size, in bytes, of the final chunk is less than or equal to 1 megabyte, the signature for the final chunk has the structure that is shown in the following diagram.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8041
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8041,
                     @"[In Zip Files] Small Final Chunk Signature: A 20-byte sequence that specifies the SHA-1 hash code of the file bytes represented by the final chunk.");
        }

        /// <summary>
        /// This method is used to verify the requirements for intermediate node object when the chunk method is zip chunk.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyIntermediateNodeForZipFileChunk(ITestSite site)
        {
            // If runs here successfully, then indicating all the following requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8045,
                     @"[In Zip Files] Large Final Chunk Signature: The Data Size of the Intermediate Node Object MUST be the total number of bytes represented by the chunk.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8047
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8047,
                     @"[In Zip Files] Large Final Chunk Signature: The Intermediate Node is referenced by its parent node.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8058
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8058,
                     @"[In Zip Files] Sub Chunk Signature: The Object Data of the Data Node Object MUST be the byte sequence from the .ZIP file tracked by the chunk. ");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8059
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8059,
                     @"[In Zip Files] Sub Chunk Signature: The Object Reference Array and the Cell Reference Array of the Data Node Object MUST be empty. ");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8060
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8060,
                     @"[In Zip Files] Sub Chunk Signature: The Object References Array of the Intermediate Node Object associated with this Data Node Object MUST have a single entry, which MUST be the Object ID of the Data Node Object.");
        }

        /// <summary>
        /// This method is used to verify the requirements for the pre-defined GUID values.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyDefinedGUID(ITestSite site)
        {
            // If runs here successfully, then indicating all the following requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     140,
                     @"[In Cell Properties] The storage manifest data element, as specified in [MS-FSSHTTPB] section 2.2.1.12.3, MUST have the Storage Manifest Schema GUID field set to 0EB93394-571D-41E9-AAD3-880D92D31955.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R141 and MS-FSSHTTPD_R8067
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8067,
                     @"[In Cell Properties] The storage manifest data element MUST have the Cell ID field set to[GUIDs] as listed in the following table: 84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073 and 6F2A4665-42C8-46C7-BAB4-E28FDCE1E32B of Type 4 and Value 1.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R141
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     141,
                     @"[In Cell Properties] The storage manifest data element, as specified in [MS-FSSHTTPB] section 2.2.1.12.3, MUST have the Cell ID field set to extended GUID 5-bit Uint values, as specified in [MS-FSSHTTPB] section 2.2.1.7.2.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R142 and MS-FSSHTTPD_R8068
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     142,
                     @"[In Cell Properties] The storage manifest data element, as specified in [MS-FSSHTTPB] section 2.2.1.12.3, MUST have the Root Extended GUID field set to an extended GUID 5-bit Uint value, as specified in [MS-FSSHTTPB] section 2.2.1.7.2.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8068
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8068,
                     @"[In Cell Properties] The storage manifest data element MUST have the Root Extended GUID field set to[GUIDs] as shown in the following table: 84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073 of Type 4 and Value 2.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R73 and  MS-FSSHTTPD_R8089
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     73,
                     @"[In Cell Properties] The revision manifest data element, as specified in [MS-FSSHTTPB] section 2.2.1.12.5, MUST have a Root Extended GUID field set to an extended GUID 5-bit Uint value that represents the primary content stream, as shown in the following table: 84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073 of Type 4 and Value 2.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8069
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8069,
                     @"[In Cell Properties] The revision manifest data element MUST have a Root Extended GUID field set to[GUIDs] as shown in the following table: 84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073 of Type 4 and Value 2.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with the IntermediateNodeObject type.
        /// </summary>
        /// <param name="interNode">Specify the IntermediateNodeObject instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyIntermediateNodeObject(IntermediateNodeObject interNode, ITestSite site)
        {
            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R69
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     69,
                     @"[In Data Node Object Data] Binary data than specifies the contents of the chunk of the file that is represented by this Data Node.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R62
            site.CaptureRequirementIfAreEqual<ulong>(
                     (ulong)interNode.GetContent().Count,
                     interNode.DataSize.DataSize,
                     "MS-FSSHTTPD",
                     62,
                     @"[In Intermediate Node Object Data] Data Size (8 bytes): An unsigned 64-bit integer that specifies the size of the file data represented by this Intermediate Node Object.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R66
            site.CaptureRequirementIfAreEqual<ulong>(
                     (ulong)interNode.GetContent().Count,
                     interNode.DataSize.DataSize,
                     "MS-FSSHTTPD",
                     66,
                     @"[In Intermediate Node Object References] The size of the Data Node Object or the sum of the Data Size values from all of the Intermediate Node Objects MUST equal the Data Size specified in the Object Data of this Intermediate Node Object.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8015
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8015,
                     @"[In Intermediate Node Object References] The ordered set of Object Extended GUIDs MUST contain the Object Extended GUID of a single Data Node Object or an ordered list of Extended GUIDs for the Intermediate Node Object.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8017
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8017,
                     @"[In Intermediate Node Object References] Object Extended GUID Array entries MUST be ordered based on the sequential file bytes represented by each Node Object. ");
        }

        /// <summary>
        /// This method is used to verify the requirements related with the object count for the root node or intermediate node.
        /// </summary>
        /// <param name="data">Specify the object data.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyObjectCount(ObjectGroupObjectData data, ITestSite site)
        {
             RootNodeObject rootNode = null;
             int index = 0;

             if (data.ObjectExGUIDArray.Count.DecodedValue > 1)
             {
                 bool isRootNode = StreamObject.TryGetCurrent<RootNodeObject>(data.Data.Content.ToArray(), ref index, out rootNode);
                 site.Log.Add(
                            TestTools.LogEntryKind.Debug,
                            "If there are more than one objects in the file, the server will respond the Root Node object for SharePoint Server 2013");

                 site.CaptureRequirementIfIsTrue(
                        isRootNode,
                        "MS-FSSHTTPD",
                        8202,
                        @"[In Appendix A: Product Behavior] If there are more than one objects in the file,the implementation does return the Root Node Object. (Microsoft SharePoint Workspace 2010, Microsoft Office 2010 suites/Microsoft SharePoint Server 2010 and above follow this behavior.)");
             }
             else
             {
                 if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 8204, SharedContext.Current.Site))
                 {
                     bool isRootNode = StreamObject.TryGetCurrent<RootNodeObject>(data.Data.Content.ToArray(), ref index, out rootNode);
                     site.Log.Add(
                            TestTools.LogEntryKind.Debug,
                            "If there is an only object in the file, the server will respond the Root Node object for SharePoint Server 2013");

                    site.CaptureRequirementIfIsTrue(
                        isRootNode,
                        "MS-FSSHTTPD",
                        8204,
                        @"[In Appendix A: Product Behavior] If there is only one object in the file,the implementation does return the Root Node Object. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft SharePoint Workspace 2010 follow this behavior.)");          
                 }
             }
        }

        /// <summary>
        /// This method is used to verify the requirements related with the RootNodeObject type.
        /// </summary>
        /// <param name="rootNode">Specify the RootNodeObject instance.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyRootNodeObject(RootNodeObject rootNode, ITestSite site)
        {
            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R44
            site.CaptureRequirementIfAreEqual<ulong>(
                     (ulong)rootNode.GetContent().Count,
                     rootNode.DataSize.DataSize,
                     "MS-FSSHTTPD",
                     44,
                     @"[In Root Node Object Data] Data Size (8 bytes): An unsigned 64-bit integer that specifies the size of the file data represented by this Root Node Object.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R49
            site.CaptureRequirementIfAreEqual<ulong>(
                     (ulong)rootNode.GetContent().Count,
                     rootNode.DataSize.DataSize,
                     "MS-FSSHTTPD",
                     49,
                     @"[In Root Node Object References] The sum of the Data Size values from all of the Intermediate Node Objects MUST equal the Data Size specified in the Object Data of this Root Node Object.");

            // When after build the Root node object successfully, the following requirements can be directly captured.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     48,
                     @"[In Root Node Object References] Each Object Extended GUID MUST specify an Intermediate Node Object.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8011
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8011,
                     @"[In Root Node Object References] The Object Extended GUID Array, as specified in [MS-FSSHTTPB] section 2.2.1.12.6.4, of the Root Node Object MUST specify an ordered set of Object Extended GUIDs. ");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8012
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8012,
                     @"[In Root Node Object References] Object Extended GUID Array entries MUST be ordered based on the sequential file bytes represented by each Node Object.");
        }

        /// <summary>
        /// Verify ObjectGroupObjectData for the Root node object group related requirements.
        /// </summary>
        /// <param name="objectGroupObjectData">Specify the objectGroupObjectData instance.</param>
        /// <param name="rootDeclare">Specify the root declare instance.</param>
        /// <param name="objectGroupList">Specify all the ObjectGroupDataElementData list.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyObjectGroupObjectDataForRootNode(ObjectGroupObjectData objectGroupObjectData, ObjectGroupObjectDeclare rootDeclare, List<ObjectGroupDataElementData> objectGroupList, ITestSite site)
        {
            #region Verify the Object Group Object Data

            // Object Extended GUID Array : Specifies an ordered list of the Object Extended GUIDs for each child of the Root Node.
            ExGUIDArray childObjectExGuidArray = objectGroupObjectData.ObjectExGUIDArray;

            if (childObjectExGuidArray != null && childObjectExGuidArray.Count.DecodedValue != 0)
            {
                // Run here successfully, then capture the requirement MS-FSSHTTPD_R8009, MS-FSSHTTPD_R8005 and MS-FSSHTTPD_R8006.
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8005,
                         @"[In Common Node Object Properties][Root of Object Extended GUID Array field] Specifies an ordered list of the Object Extended GUIDs for each child of the Root Node. ");

                // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8006
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8006,
                         @"[In Common Node Object Properties][Root of Object Extended GUID Array field] Object Extended GUID Array entries MUST be ordered based on the sequential file bytes represented by each Node Object.");

                // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8009
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8009,
                         @"[In Common Node Object Properties][Root of Object Extended GUID field] An extended GUID, as specified in [MS-FSSHTTPB] section 2.2.1.7.");

                foreach (ExGuid guid in childObjectExGuidArray.Content)
                {
                    bool isUnique = IsGuidUnique(guid, objectGroupList);
                    site.Log.Add(LogEntryKind.Debug, "Whether the Object Extended GUID is unique:{0}", isUnique);

                    // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8101
                    site.CaptureRequirementIfIsTrue(
                             isUnique,
                             "MS-FSSHTTPD",
                             8101,
                             @"[In Common Node Object Properties][Root of Object Extended GUID field] This GUID[Object Extended GUID] MUST be different within this file in once response.");
                }
            }

            // Cell ID Array : Specifies an empty list of Cell IDs.
            CellIDArray cellIDArray = objectGroupObjectData.CellIDArray;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The count of Cell ID Array is:{0}", cellIDArray.Count);

            // If the Cell ID Array is an empty list, indicates that the count of the array is 0.
            // So capture these requirements.
            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R4
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     cellIDArray.Count,
                     "MS-FSSHTTPD",
                     4,
                     @"[In Common Node Object Properties][Root of Cell ID Array field] Specifies an empty list of Cell IDs.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R50
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     cellIDArray.Count,
                     "MS-FSSHTTPD",
                     50,
                     @"[In Root Node Object Cell References] The Cell ID Array, as specified in [MS-FSSHTTPB] section 2.2.1.12.6.4, of the Root Node Object MUST specify an empty array.");

            #endregion 

            #region Verify the Object Group Object Declaration

            // Object Extended GUID : An extended GUID which specifies an identifier for this object. This GUID MUST be unique within this file.
            ExGuid currentObjectExGuid = rootDeclare.ObjectExtendedGUID;

            // Check whether Object Extended GUID is unique.
            bool isGuidUnique = IsGuidUnique(currentObjectExGuid, objectGroupList);

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "Whether the Object Extended GUID is unique:{0}", isGuidUnique);

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8103
            site.CaptureRequirementIfIsTrue(
                     isGuidUnique,
                     "MS-FSSHTTPD",
                     8103,
                     @"[In Common Node Object Properties][Data of Object Extended GUID field] This GUID[Object Extended GUID] MUST be different within this file in once response.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R75
            site.CaptureRequirementIfIsTrue(
                     isGuidUnique,
                     "MS-FSSHTTPD",
                     75,
                     @"[In Cell Properties] For each stream, a single Root Node MUST be specified by using a unique root identifier.");

            // Object Partition ID : A compact unsigned 64-bit integer which MUST be 1.
            Compact64bitInt objectPartitionID = rootDeclare.ObjectPartitionID;

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R17
            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectPartitionID.GetType()), "The type of objectPartitionID should be a compact unsigned 64-bit integer.");
            site.Assert.IsTrue(objectPartitionID.DecodedValue == 1, "The actual value of objectPartitionID should be 1.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R17
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     17,
                     @"[In Common Node Object Properties][Root of Object Partition ID field] A compact unsigned 64-bit integer that MUST be ""1"".");

            // Object Data Size :A compact unsigned 64-bit integer which MUST be the size of the Object Data field.
            Compact64bitInt objectDataSize = rootDeclare.ObjectDataSize;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Object Data Size is:{0}", objectDataSize.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectDataSize.GetType()), "The type of objectPartitionID should be a compact unsigned 64-bit integer.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R20
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     20,
                     @"[In Common Node Object Properties][Root of Object Data Size field] A compact unsigned 64-bit integer that MUST be the size of the Object Data field.");

            // Object References Count : A compact unsigned 64-bit integer that specifies the number of object references.
            Compact64bitInt objectReferencesCount = rootDeclare.ObjectReferencesCount;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The count of Object References is:{0}", objectReferencesCount.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectReferencesCount.GetType()), "The type of objectReferencesCount should be a compact unsigned 64-bit integer.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R23
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     23,
                     @"[In Common Node Object Properties][Root  of Object References Count field] A compact unsigned 64-bit integer that specifies the number of object references.");

            // Cell References Count : A compact unsigned 64-bit integer which MUST be 0.
            Compact64bitInt cellReferencesCount = rootDeclare.CellReferencesCount;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Cell References is:{0}", cellReferencesCount.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(cellReferencesCount.GetType()), "The type of cellReferencesCount should be a compact unsigned 64-bit integer.");
            site.Assert.IsTrue(cellReferencesCount.DecodedValue == 0, "The value of cellReferencesCount should be 0.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R26
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     26,
                     @"[In Common Node Object Properties][Root of Cell References Count field] A compact unsigned 64-bit integer that MUST be zero.");
            #endregion

            // Run here successfully, then capture the requirement MS-FSSHTTPD_R8002 and MS-FSSHTTPD_R8004.
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8002,
                     @"[In Common Node Object Properties] A Node Object is contained within an object group data element.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8004
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     8004,
                     @"[In Common Node Object Properties] The Object Group Object Data field MUST be set as shown in the following table, as specified in [MS-FSSHTTPB] section 2.2.1.12.6.4");
        }

        /// <summary>
        /// Verify ObjectGroupObjectData for the Intermediate node object group related requirements.
        /// </summary>
        /// <param name="objectGroupObjectData">Specify the objectGroupObjectData instance.</param>
        /// <param name="intermediateDeclare">Specify the intermediate declare instance.</param>
        /// <param name="objectGroupList">Specify all the ObjectGroupDataElementData list.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyObjectGroupObjectDataForIntermediateNode(ObjectGroupObjectData objectGroupObjectData, ObjectGroupObjectDeclare intermediateDeclare, List<ObjectGroupDataElementData> objectGroupList, ITestSite site)
        {
            #region Verify the Object Group Object Data

            ExGUIDArray childObjectExGuidArray = objectGroupObjectData.ObjectExGUIDArray;

            if (childObjectExGuidArray != null && childObjectExGuidArray.Count.DecodedValue != 0)
            {
                // If the intermediate node can be build then verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R65
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         65,
                         @"[In Intermediate Node Object References] The Object Extended GUID Array, as specified in [MS-FSSHTTPB] section 2.2.1.12.6.4, of the Intermediate Node Object MUST specify an ordered set of Object Extended GUIDs.");

                // If the intermediate node can be build then verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8007 and MS-FSSHTTPD_R8008
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8007,
                         @"[In Common Node Object Properties][Intermediate of Object Extended GUID Array field] Specifies an ordered list of the Object Extended GUIDs for each child of this Node.");

                // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8008
                site.CaptureRequirement(
                         "MS-FSSHTTPD",
                         8008,
                         @"[In Common Node Object Properties][Intermediate of Object Extended GUID Array field] Object Extended GUID Array entries MUST be ordered based on the sequential file bytes represented by each Node Object.");
            }

            // Cell ID Array : Specifies an empty list of Cell IDs.
            CellIDArray cellIDArray = objectGroupObjectData.CellIDArray;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The count of Cell ID Array is:{0}", cellIDArray.Count);

            // If the Cell ID Array is an empty list, indicates that the count of the array is 0.
            // So capture these requirements.
            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R5
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     cellIDArray.Count,
                     "MS-FSSHTTPD",
                     5,
                     @"[In Common Node Object Properties][Intermediate of Cell ID Array field] Specifies an empty list of Cell IDs.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R67
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     cellIDArray.Count,
                     "MS-FSSHTTPD",
                     67,
                     @"[In Intermediate Node Object Cell References] The Cell Reference Array of the Object MUST specify an empty array.");

            #endregion 

            #region Verify the Object Group Object Declaration

            // Object Extended GUID : An extended GUID which specifies an identifier for this object. This GUID MUST be unique within this file.
            ExGuid currentObjectExGuid = intermediateDeclare.ObjectExtendedGUID;

            // Check whether Object Extended GUID is unique.
            bool isVerify8102 = IsGuidUnique(currentObjectExGuid, objectGroupList);

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "Whether the Object Extended GUID is unique:{0}", isVerify8102);

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8102
            site.CaptureRequirementIfIsTrue(
                     isVerify8102,
                     "MS-FSSHTTPD",
                     8102,
                     @"[In Common Node Object Properties][Intermediate of Object Extended GUID Field] This GUID[Object Extended GUID] MUST  be different within this file in once response.");

            // Object Partition ID : A compact unsigned 64-bit integer which MUST be 1.
            Compact64bitInt objectPartitionID = intermediateDeclare.ObjectPartitionID;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Object Partition ID is:{0}", objectPartitionID.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectPartitionID.GetType()), "The type of objectPartitionID should be a compact unsigned 64-bit integer.");
            site.Assert.IsTrue(objectPartitionID.DecodedValue == 1, "The actual value of objectPartitionID should be 1.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R18            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     18,
                     @"[In Common Node Object Properties][Intermediate of Object Partition ID field] A compact unsigned 64-bit integer that MUST be ""1"".");

            // Object Data Size :A compact unsigned 64-bit integer which MUST be the size of the Object Data field.
            Compact64bitInt objectDataSize = intermediateDeclare.ObjectDataSize;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Object Data Size is:{0}", objectDataSize.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectDataSize.GetType()), "The type of objectPartitionID should be a compact unsigned 64-bit integer.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R21            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     21,
                     @"[In Common Node Object Properties][Intermediate of Object Data Size field] A compact unsigned 64-bit integer that MUST be the size of the Object Data field.");

            // Object References Count : A compact unsigned 64-bit integer that specifies the number of object references.
            Compact64bitInt objectReferencesCount = intermediateDeclare.ObjectReferencesCount;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The count of Object References is:{0}", objectReferencesCount.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectReferencesCount.GetType()), "The type of objectReferencesCount should be a compact unsigned 64-bit integer.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R24            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     24,
                     @"[In Common Node Object Properties][Intermediate of Object References Count field] A compact unsigned 64-bit integer that specifies the number of object references.");

            // Cell References Count : A compact unsigned 64-bit integer which MUST be 0.
            Compact64bitInt cellReferencesCount = intermediateDeclare.CellReferencesCount;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Cell References is:{0}", cellReferencesCount.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(cellReferencesCount.GetType()), "The type of cellReferencesCount should be a compact unsigned 64-bit integer.");
            site.Assert.IsTrue(cellReferencesCount.DecodedValue == 0, "The value of cellReferencesCount should be 0.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R27            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     27,
                     @"[In Common Node Object Properties][Intermediate of Cell References Count field] A compact unsigned 64-bit integer that MUST be zero.");

            #endregion 
        }

        /// <summary>
        /// Verify ObjectGroupObjectData for the DataNodeObject related requirements.
        /// </summary>
        /// <param name="objectGroupObjectData">Specify the objectGroupObjectData instance.</param>
        /// <param name="dataNodeDeclare">Specify the data node object declare instance.</param>
        /// <param name="objectGroupList">Specify all the ObjectGroupDataElementData list.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public static void VerifyObjectGroupObjectDataForDataNodeObject(ObjectGroupObjectData objectGroupObjectData, ObjectGroupObjectDeclare dataNodeDeclare, List<ObjectGroupDataElementData> objectGroupList, ITestSite site)
        {
            #region Verify the Object Group Object Data

            // Object Extended GUID Array : Specifies an ordered list of the Object Extended GUIDs for each child of the Root Node.
            ExGUIDArray childObjectExGuidArray = objectGroupObjectData.ObjectExGUIDArray;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The count of Cell ID Array is:{0}", childObjectExGuidArray.Count.DecodedValue);

            // If the Object Extended GUID Array is an empty list, indicates that the count of the array is 0.
            // So capture these requirements.
            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R3
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     childObjectExGuidArray.Count.DecodedValue,
                     "MS-FSSHTTPD",
                     3,
                     @"[In Common Node Object Properties][Data of Object Extended GUID Array field] Specifies an empty list of Object Extended GUIDs.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R71
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     childObjectExGuidArray.Count.DecodedValue,
                     "MS-FSSHTTPD",
                     71,
                     @"[In Data Node Object References] The Object Extended GUID Array, as specified in [MS-FSSHTTPB] section 2.2.1.12.6.4, of the Data Node Object MUST specify an empty array.");

            // Cell ID Array : Specifies an empty list of Cell IDs.
            CellIDArray cellIDArray = objectGroupObjectData.CellIDArray;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The count of Cell ID Array is:{0}", cellIDArray.Count);

            // If the Object Extended GUID Array is an empty list, indicates that the count of the array is 0.
            // So capture these requirements.
            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R6
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     cellIDArray.Count,
                     "MS-FSSHTTPD",
                     6,
                     @"[In Common Node Object Properties][Data of Cell ID Array field] Specifies an empty list of Cell IDs.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R72
            site.CaptureRequirementIfAreEqual<ulong>(
                     0,
                     cellIDArray.Count,
                     "MS-FSSHTTPD",
                     72,
                     @"[In Data Node Object Cell References] The Object Extended GUID Array, as specified in [MS-FSSHTTPB] section 2.2.1.12.6.4, of the Data Node Object MUST specify an empty array.");

            #endregion 

            #region Verify the Object Group Object Declaration

            // Object Extended GUID : An extended GUID which specifies an identifier for this object. This GUID MUST be unique within this file.
            ExGuid currentObjectExGuid = dataNodeDeclare.ObjectExtendedGUID;

            // Check whether Object Extended GUID is unique.
            bool isUnique = IsGuidUnique(currentObjectExGuid, objectGroupList);

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "Whether the Object Extended GUID is unique:{0}", isUnique);

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8103
            site.CaptureRequirementIfIsTrue(
                     isUnique,
                     "MS-FSSHTTPD",
                     8103,
                     @"[In Common Node Object Properties][Data of Object Extended GUID field] This GUID[Object Extended GUID] MUST be different within this file in once response.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R8103
            site.CaptureRequirementIfIsTrue(
                     isUnique,
                     "MS-FSSHTTPD",
                     8103,
                     @"[In Common Node Object Properties][Data of Object Extended GUID field] This GUID[Object Extended GUID] MUST be different within this file in once response.");

            // Object Partition ID : A compact unsigned 64-bit integer which MUST be 1.
            Compact64bitInt objectPartitionID = dataNodeDeclare.ObjectPartitionID;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Object Partition ID is:{0}", objectPartitionID.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectPartitionID.GetType()), "The type of objectPartitionID should be a compact unsigned 64-bit integer.");
            site.Assert.IsTrue(objectPartitionID.DecodedValue == 1, "The actual value of objectPartitionID should be 1.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R19            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     19,
                     @"[In Common Node Object Properties][Data of Object Partition ID field] A compact unsigned 64-bit integer that MUST be ""1"".");

            // Object Data Size : A compact unsigned 64-bit integer which MUST be the size of the Object Data field.
            Compact64bitInt objectDataSize = dataNodeDeclare.ObjectDataSize;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Object Data Size is:{0}", objectDataSize.DecodedValue);

            // Get the size of Object Data.
            ulong sizeInObjectData = (ulong)objectGroupObjectData.Data.Length.DecodedValue;

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectDataSize.GetType()), "The type of objectPartitionID should be a compact unsigned 64-bit integer.");
            site.Assert.IsTrue(objectDataSize.DecodedValue == sizeInObjectData, "The actual value of objectDataSize should be same as the Object Data field");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R22            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     22,
                     @"[In Common Node Object Properties][Data of Object Data Size field] A compact unsigned 64-bit integer that MUST be the size of the Object Data field.");

            // Object References Count : A compact unsigned 64-bit integer that specifies the number of object references.
            Compact64bitInt objectReferencesCount = dataNodeDeclare.ObjectReferencesCount;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The count of Object References is:{0}", objectReferencesCount.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(objectReferencesCount.GetType()), "The type of objectReferencesCount should be a compact unsigned 64-bit integer.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R25            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     25,
                     @"[In Common Node Object Properties][Data of Object References Count field] A compact unsigned 64-bit integer that specifies the number of object references.");

            // Cell References Count : A compact unsigned 64-bit integer which MUST be 0.
            Compact64bitInt cellReferencesCount = dataNodeDeclare.CellReferencesCount;

            // Add the log information.
            site.Log.Add(LogEntryKind.Debug, "The value of Cell References is:{0}", cellReferencesCount.DecodedValue);

            site.Assert.IsTrue(typeof(Compact64bitInt).Equals(cellReferencesCount.GetType()), "The type of cellReferencesCount should be a compact unsigned 64-bit integer.");
            site.Assert.IsTrue(cellReferencesCount.DecodedValue == 0, "The value of cellReferencesCount should be 0.");

            // Verify MS-FSSHTTPD requirement: MS-FSSHTTPD_R28            
            site.CaptureRequirement(
                     "MS-FSSHTTPD",
                     28,
                     @"[In Common Node Object Properties][Data of Cell References Count field] A compact unsigned 64-bit integer that MUST be zero.");

            #endregion 
        }

        /// <summary>
        /// This method is used to check whether the extended GUID is unique in the object group data element data list.
        /// </summary>
        /// <param name="currentObjectExGuid">Specify the object extended GUID.</param>
        /// <param name="objectGroupList">Specify the object group data element data list. </param>
        /// <returns>Return true if the GUID is unique, otherwise return false.</returns>
        private static bool IsGuidUnique(ExGuid currentObjectExGuid, List<ObjectGroupDataElementData> objectGroupList)
        {
            return objectGroupList.SelectMany(dataEle => dataEle.ObjectGroupDeclarations.ObjectDeclarationList).Where(declare => declare.ObjectExtendedGUID.Equals(currentObjectExGuid)).Count() == 1;
        }
    }
}