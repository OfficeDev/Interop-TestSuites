namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.MS_ONESTORE;
    using Microsoft.Protocols.TestTools;
    using System;
    using System.Collections.Generic;

    /// <summary>
    ///  class contains methods which capture requirements related with MS-ONESTORE.
    /// </summary>
    public class MsonestoreCapture
    {
        public void Validate(MSOneStorePackage instance, ITestSite site)
        {
            this.VerifyStorageManifest(instance.StorageManifest, site);
            for(int i=0;i< instance.RevisionManifests.Count;i++)
            {
                RevisionManifestDataElementData revisionManifest = instance.RevisionManifests[i];
                this.VerifyRevisions(revisionManifest, site);
            }

            foreach(RevisionStoreObject revisionStoreObject in instance.RevisionStoreObjects)
            {
                this.VerifyRevisionStoreObject(revisionStoreObject, site);
            }
            this.VerifyRoot(instance.Root,site);
            this.VerifyHeaderCell(instance.HeaderCells, site);    
        }
        /// <summary>
        /// This method is used to verify the requirements related with the Storage Manifest.
        /// </summary>
        /// <param name="instance">The instance of the Storage Manifest.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyStorageManifest(StorageManifestDataElementData instance, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R886
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.StorageManifestSchemaGUID.GUID,
                    typeof(Guid),
                    886,
                    @"[In  Storage Manifest] § GUID (storage manifest schema GUID): A GUID, as specified in [MS-DTYP], that specifies the file type.");

            if (SharedContext.Current.FileUrl.ToLowerInvariant().EndsWith(".one"))
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R888
                site.CaptureRequirementIfAreEqual<String>(
                             "1F937CB4-B26F-445F-B9F8-17E20160E461",
                             instance.StorageManifestSchemaGUID.GUID.ToString().ToUpper(),
                             "MS-ONESTROE",
                             888,
                             @"[In Response Error] Error Type GUID field is set to {8454C8F2-E401-405A-A198-A10B6991B56E}[ specifies the error type is ]HRESULT Error (section 2.2.3.2.4).");
            }
            
            if (SharedContext.Current.FileUrl.ToLowerInvariant().EndsWith(".onetoc2"))
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R889
                site.CaptureRequirementIfAreEqual<String>(
                             "E4DBFD38-E5C7-408B-A8A1-0E7B421E1F5F",
                             instance.StorageManifestSchemaGUID.GUID.ToString().ToUpper(),
                             "MS-ONESTROE",
                             889,
                             @"[In Response Error] Error Type GUID field is set to {8454C8F2-E401-405A-A198-A10B6991B56E}[ specifies the error type is ]HRESULT Error (section 2.2.3.2.4).");
            }

            // If R888 or R889 are verified, then R887 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R887
            site.CaptureRequirement(
                            "MS-ONESTROE",
                             887,
                             @"[In  Storage Manifest] [GUID] MUST be one of the following values[{1F937CB4-B26F-445F-B9F8-17E20160E461},{E4DBFD38-E5C7-408B-A8A1-0E7B421E1F5F}], depending on the file type.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R892
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.StorageManifestRootDeclareList[0].RootExtendedGUID,
                    typeof(ExGuid),
                    "MS-ONESTORE",
                    892,
                    @"[In  Storage Manifest] § Root Extended GUID: An ExtendedGUID structure used to contain the root identifier of the header cell (section 2.7.2).");


            // Verify MS-ONESTORE requirement: MS-ONESTORE_R893
            site.CaptureRequirementIfAreEqual<ExGuid>(
                new ExGuid(1,Guid.Parse("{1A5A319C-C26b-41AA-B9C5-9BD8C44E07D4}")),
                instance.StorageManifestRootDeclareList[0].RootExtendedGUID,
                "MS-ONESTORE",
                893,
                @"[In  Storage Manifest] [Root Extended GUID] MUST be ""{{ 1A5A319C-C26b-41AA-B9C5-9BD8C44E07D4 } , 1}"".");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R894
            if (instance.StorageManifestRootDeclareList[0].RootExtendedGUID.Equals(new ExGuid(1, Guid.Parse("{1A5A319C-C26b-41AA-B9C5-9BD8C44E07D4}"))))
            {
                site.CaptureRequirementIfIsTrue(
                    instance.StorageManifestRootDeclareList[0].CellID.ExtendGUID1.Equals(new ExGuid(1, Guid.Parse("{84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073}"))) &&
                    instance.StorageManifestRootDeclareList[0].CellID.ExtendGUID2.Equals(new ExGuid(1, Guid.Parse("{111E4CF3-7FEF-4087-AF6A-B9544ACD334D}"))),
                    "MS-ONESTORE",
                    894,
                    @"[In  Storage Manifest] § Cell ID: A Cell ID structure (as specified in [MS-FSSHTTPB] section 2.2.1.10) where the EXGUID1 field MUST be equal to ""{{ 84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073} , 1}"" and the EXGUID2 field MUST be equal to ""{{ 111E4CF3-7FEF-4087-AF6A-B9544ACD334D } , 1}"".");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R896
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.StorageManifestRootDeclareList[1].RootExtendedGUID,
                    typeof(ExGuid),
                    "MS-ONESTORE",
                    896,
                    @"[In  Storage Manifest] § Root Extended GUID: An ExtendedGUID structure used to contain the root identifier of the root object space (section 2.1.4).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R897
            site.CaptureRequirementIfAreEqual<ExGuid>(
                new ExGuid(2,Guid.Parse("{84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073}")),
                instance.StorageManifestRootDeclareList[1].RootExtendedGUID,
                "MS-ONESTORE",
                 897,
                @"[In  Storage Manifest] [Root Extended GUID] MUST be the default root ""{{ 84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073} , 2}"".");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R922
            site.CaptureRequirementIfAreEqual<ExGuid>(
                new ExGuid(1,Guid.Parse("{84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073}")),
                instance.StorageManifestRootDeclareList[0].CellID.ExtendGUID1,
                "MS-ONESTORE",
                992,
                @"[In Cells] This value is converted to ""{ { 84DEFAB9 - AAA3 - 4A0D - A3A8 - 520C77AC7073} , 1 }"" when transmitted using the File Synchronization via SOAP over HTTP Protocol.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with the Revisions.
        /// </summary>
        /// <param name="instance">The instance of Revisions.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyRevisions(RevisionManifestDataElementData instance, ITestSite site)
        {
            if(instance.RevisionManifestRootDeclareList.Count>0)
            {
                if (instance.RevisionManifestRootDeclareList[0].RootExtendedGUID.Equals(new ExGuid(1, Guid.Parse("{4A3717F8-1C14-49E7-9526-81D942DE1741}"))) &&
                   instance.RevisionManifestRootDeclareList[0].ObjectExtendedGUID.Equals(new ExGuid(1, Guid.Parse("{B4760B1A-FBDF-4AE3-9D08-53219D8A8D21}"))))
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R905
                    site.CaptureRequirement(
                        "MS-ONESTORE",
                        905,
                        @"[In Header Cell] § Root Extended GUID: MUST be ""{{ 4A3717F8- 1C14-49E7-9526-81D942DE1741 },  1}"".");

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R906
                    site.CaptureRequirement(
                         "MS-ONESTORE",
                         906,
                         @"[In Header Cell] § Object Extended GUID: MUST be ""{{ B4760B1A- FBDF- 4AE3-9D08-53219D8A8D21 }, 1}"".");
                }
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R928,MS-ONESTORE_R929,MS-ONESTORE_R930,MS-ONESTORE_R931
                foreach (RevisionManifestRootDeclare revision in instance.RevisionManifestRootDeclareList)
                {
                    if (revision.RootExtendedGUID.Equals(new ExGuid(1, Guid.Parse("{4A3717F8-1C14-49E7-9526-81D942DE1741}"))))
                    {
                        site.CaptureRequirement(
                            "MS-ONESTORE",
                            928,
                            @"[In Revisions] 0x00000001 means Default content root, specifies Root extended GUID: { { 4A3717F8- 1C14-49E7-9526-81D942DE1741 },  1}");
                    }
                    if (revision.RootExtendedGUID.Equals(new ExGuid(2, Guid.Parse("{4A3717F8-1C14-49E7-9526-81D942DE1741}"))))
                    {
                        site.CaptureRequirement(
                            "MS-ONESTORE",
                            929,
                            @"[In Revisions] 0x00000002 means Metadata root, specifies Root extended GUID: { { 4A3717F8- 1C14-49E7-9526-81D942DE1741 },  2}.");
                    }
                    if (revision.RootExtendedGUID.Equals(new ExGuid(3, Guid.Parse("{4A3717F8-1C14-49E7-9526-81D942DE1741}"))))
                    {
                        site.CaptureRequirement(
                            "MS-ONESTORE",
                            930,
                            @"[In Revisions] 0x00000003 means Encryption Key root, specifies Root extended GUID: { 4A3717F8- 1C14-49E7-9526-81D942DE1741 },  3}");
                    }
                    if (revision.RootExtendedGUID.Equals(new ExGuid(4, Guid.Parse("{4A3717F8-1C14-49E7-9526-81D942DE1741}"))))
                    {
                        site.CaptureRequirement(
                            "MS-ONESTORE",
                            931,
                            @"[In Revisions] 0x00000004 means Version metadata root, specifies Root extended GUID: { { 4A3717F8- 1C14-49E7-9526-81D942DE1741 },  4}");
                    }
                }
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with the Object.
        /// </summary>
        /// <param name="instance">The instance of Object.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyRevisionStoreObject(RevisionStoreObject instance, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R943
            if ( instance.JCID.JCID.IsFileData == 0 && Convert.ToInt32(instance.JCID.ObjectDeclaration.ObjectPartitionID.DecodedValue) == 4)
            {
                site.CaptureRequirementIfAreEqual<Type>(
                    typeof(JCID),
                    instance.JCID.JCID.GetType(),
                    943,
                    @"[In Objects] Object Declaration.PartitionID: 4 (Static Object MetaData), and Object Data.Data:An unsigned integer that specifies the JCID (section 2.6.14) of the object.");

            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R944
            if (instance.JCID.JCID.IsFileData == 0 && Convert.ToInt32(instance.JCID.ObjectDeclaration.ObjectPartitionID.DecodedValue) == 1)
            {
                site.CaptureRequirementIfIsInstanceOfType(
                    instance.PropertySet,
                    typeof(ObjectSpaceObjectPropSet),
                    944,
                    @"[In Objects] Object Declaration.PartitionID: 1 (Object Data), Object Data.Data: MUST be an ObjectSpaceObjectPropSet structure (section 2.6.1), Object Data.Object Extended GUID array: Identifiers of the referenced objects in the revision store, and Object Data.Cell ID array: Identifiers of the referenced object spaces in the revision store.");

            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R946
            if (instance.JCID.JCID.IsFileData == 1 && Convert.ToInt32(instance.JCID.ObjectDeclaration.ObjectPartitionID.DecodedValue) == 4)
            {
                site.CaptureRequirementIfIsInstanceOfType(
                    instance.JCID.JCID,
                    typeof(JCID),
                    946,
                    @"[In Objects] Object Declaration.PartitionID: 4 (Static Object MetaData) and Object Data.Data: An unsigned integer that specifies the JCID (section 2.6.14) of the object.");

            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R947
            if (instance.JCID.JCID.IsFileData == 1 && Convert.ToInt32(instance.JCID.ObjectDeclaration.ObjectPartitionID.DecodedValue) == 1)
            {
                site.CaptureRequirementIfIsInstanceOfType(
                    instance.PropertySet,
                    typeof(ObjectSpaceObjectPropSet),
                    947,
                    @"[In Objects] Object Declaration.PartitionID: 1 (Object Data) and Object Data.Data: MUST be an ObjectSpaceObjectPropSet structure (section 2.6.1) with properties specified later in this section.");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R858
            site.CaptureRequirementIfAreEqual<Type>(
                typeof(Int32),
                instance.JCID.JCID.Index.GetType(),
                858,
                @"[In JCID] index (2 bytes): An unsigned integer that specifies the type of object.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R859
            if (instance.JCID.JCID.IsBinary == 1 || instance.JCID.JCID.IsBinary == 0)
            {
                site.CaptureRequirement(
                    "MS-ONESTORE",
                    859,
                    @"[In JCID] A - IsBinary (1 bit): Specifies whether the object contains encryption data transmitted over the File Synchronization via SOAP over HTTP Protocol, as specified in [MS-FSSHTTP].");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R860
            if (instance.JCID.JCID.IsPropertySet == 1 || instance.JCID.JCID.IsPropertySet == 0)
            {
                site.CaptureRequirement(
                    "MS-ONESTORE",
                    860,
                    @"[In JCID] B - IsPropertySet (1 bit): Specifies whether the object contains a property set (section 2.1.1).");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R862
            if (instance.JCID.JCID.IsFileData == 1 || instance.JCID.JCID.IsFileData == 0)
            {
                site.CaptureRequirement(
                    "MS-ONESTORE",
                    862,
                    @"[In JCID] D - IsFileData (1 bit): Specifies whether the object is a file data object. ");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R864
            if (instance.JCID.JCID.IsReadOnly == 1 || instance.JCID.JCID.IsReadOnly == 0)
            {
                site.CaptureRequirement(
                    "MS-ONESTORE",
                    864,
                    @"[In JCID] E - IsReadOnly (1 bit): Specifies whether the object's data MUST NOT be changed when the object is revised.");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R863
            if (instance.JCID.JCID.IsFileData == 1)
            {
                Boolean R863 = false;
                if (instance.JCID.JCID.IsBinary == 0 && instance.JCID.JCID.IsGraphNode == 0 && instance.JCID.JCID.IsPropertySet == 0 && instance.JCID.JCID.IsReadOnly == 0)
                {
                    R863 = true;
                }
                site.CaptureRequirementIfIsTrue(
                    R863,
                    863,
                    @"[In JCID] If the value of IsFileData is ""true"", then the values of the IsBinary, IsPropertySet, IsGraphNode, and IsReadOnly fields MUST all be false.");
            }


        }
        /// <summary>
        /// This method is used to verify the requirements related with the Object.
        /// </summary>
        /// <param name="instance">The instance of Object.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyRoot(RootObject instance, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R908
            if (instance.ObjectData != null && instance.ObjectDeclaration != null)
            {
                site.CaptureRequirement(
                    "MS-ONESTORE",
                    908,
                    @"[In Header Cell] The root object specified by the Object Extended GUID field (described earlier) MUST be transmitted as a pair of Object Declaration and Object Data structures:");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R909
            if (instance.ObjectDeclaration.ObjectPartitionID.DecodedValue == 1)
            {
                site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ObjectSpaceObjectPropSet),
                    instance.ObjectData.GetType(),
                    "MS-ONENOTE",
                    909,
                    @"[In Header Cell] Object Declaration.PartitionID ""1(Object Data)"": MUST be an ObjectSpaceObjectPropSet structure (section 2.6.1) with properties specified later in this section.");
            }

            //Verify ObjectSpaceObjectPropSet structure in section 2.6.1
            if (instance.ObjectData != null)
            {
                this.VerifyObjectSpaceObjectPropSet(instance.ObjectData, site);
            }

            //verify ObjectSpaceObjectStreamOfOSIDs structure in section 2.6.3
            if (instance.ObjectData.OSIDs != null)
            {
                this.VerifyObjectSpaceObjectStreamOfOSIDs(instance.ObjectData.OSIDs,instance.ObjectData.ContextIDs, site);
            }

            //verify ObjectSpaceObjectStreamOfContextIDs structure2.6.4
            if (instance.ObjectData.ContextIDs != null)
            {
                this.VerifyObjectSpaceObjectStreamOfContextIDs(instance.ObjectData.ContextIDs, site);
            }

            //verify ObjectSpaceObjectStreamOfContextIDs structure in section 2.6.7
            if (instance.ObjectData.Body != null)
            {
                this.VerifyPropertySet(instance.ObjectData.Body, site);

                for (int i = 0; i < instance.ObjectData.Body.RgPrids.Length; i++)
                {
                    PropertyID propId = instance.ObjectData.Body.RgPrids[i];
                    if (propId.Type == 0x1)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R781
                        site.CaptureRequirementIfIsInstanceOfType(
                          instance.ObjectData.Body.RgData[i],
                          typeof(NoData),
                          "MS-ONESTORE",
                          781,
                          @"[In PropertyID] value ""0x1"", name ""NoData"": The property contains no data.");
                    }
                    if (propId.Type == 0x2)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R782
                        site.CaptureRequirementIfIsTrue(
                            propId.BoolValue == 0 || propId.BoolValue == 1,
                            "MS-ONESTORE",
                            782,
                            @"[In PropertyID] value ""0x2"", name ""Bool"": The property is a Boolean value specified by boolValue.");

                        // If R782 is verified, then R799 will be verified.
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R799
                        site.CaptureRequirement(
                            799,
                            @"[In PropertyID] A - boolValue (1 bit): A bit that specifies the value of a Boolean property.");
                    }
                    else
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R79901
                        site.CaptureRequirementIfIsFalse(
                            Convert.ToBoolean(instance.ObjectData.Body.RgPrids[i].BoolValue),
                            79901,
                            @"[In PropertyID] A - boolValue (1 bit):  MUST be false if the value of the type field is not equal to 0x2.");
                    }

                    if (propId.Type == 0x3)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R783
                        site.CaptureRequirementIfIsInstanceOfType(
                            instance.ObjectData.Body.RgData[i],
                            typeof(OneByteOfData),
                            "MS-ONESTORE",
                            783,
                            @"[In PropertyID] value ""0x3"", name ""OneByteOfData"": The property contains 1 byte of data in the PropertySet.rgData stream field.");
                    }
                    if (propId.Type == 0x4)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R784
                        site.CaptureRequirementIfIsInstanceOfType(
                            instance.ObjectData.Body.RgData[i],
                            typeof(TwoBytesOfData),
                            "MS-ONESTORE",
                            784,
                            @"[In PropertyID] value ""0x4"", name ""TwoBytesOfData"": The property contains 2 bytes of data in the PropertySet.rgData stream field.");
                    }
                    if (propId.Type == 0x5)
                    {
                        //Verfiy MS-ONESTORE requirement: MS-ONESTORE_R785
                        site.CaptureRequirementIfIsInstanceOfType(
                            instance.ObjectData.Body.RgData[i],
                            typeof(FourBytesOfData),
                            "MS-ONESTORE",
                            785,
                            @"[In PropertyID] value ""0x5"", name ""FourBytesOfData"": The property contains 4 bytes of data in the PropertySet.rgData stream field.");
                    }
                    if (propId.Type == 0x6)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R786
                        site.CaptureRequirementIfIsInstanceOfType(
                            instance.ObjectData.Body.RgData[i],
                            typeof(EightBytesOfData),
                            "MS-ONESTORE",
                            786,
                            @"[In PropertyID] value ""0x6"", name ""EightBytesOfData"": The property contains 8 bytes of data in the PropertySet.rgData stream field.");
                    }
                    if (propId.Type == 0x7)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R787
                        site.CaptureRequirementIfIsInstanceOfType(
                            instance.ObjectData.Body.RgData[i],
                            typeof(PrtFourBytesOfLengthFollowedByData),
                            "MS-ONESTORE",
                            787,
                            @"[In PropertyID] value ""0x7"", name ""FourBytesOfLengthFollowedByData"": The property contains a prtFourBytesOfLengthFollowedByData (section 2.6.8) in the PropertySet.rgData stream field.");

                        this.VerifyPrtFourBytesOfLengthFollowedByData((PrtFourBytesOfLengthFollowedByData)instance.ObjectData.Body.RgData[i], site);
                    }
                    if (propId.Type == 0x8)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R788
                        site.CaptureRequirementIfIsTrue(
                            instance.ObjectData.OIDs.Body.Length==1,
                            "MS-ONESTORE",
                            788,
                            @"[In PropertyID] value ""0x8"", name ""ObjectID"": The property contains one CompactID (section 2.2.2) in the ObjectSpaceObjectPropSet.OIDs.body stream field.");
                    }
                    if (propId.Type == 0x9)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R789
                        site.CaptureRequirementIfIsTrue(                   
                            instance.ObjectData.OIDs.Body.Length>1,
                            "MS-ONESTORE",
                            789,
                            @"[In PropertyID] value ""0x9"", name ""ArrayOfObjectIDs"": The property contains an array of CompactID structures in the ObjectSpaceObjectPropSet.OIDs.body stream field.");

                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R790
                        site.CaptureRequirementIfAreEqual<uint>(
                            (uint)instance.ObjectData.OIDs.Body.Length,
                            ((ArrayNumber)instance.ObjectData.Body.RgData[i]).Number,
                            "MS-ONESTORE",
                            790,
                            @"[In PropertyID] value ""0x9"", name ""ArrayOfObjectIDs"": The property contains an unsigned integer of size 4 bytes in the PropertySet.rgData stream field that specifies the number of CompactID structures this property contains.");
                    }
                    if (propId.Type == 0xA)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R791
                        site.CaptureRequirementIfIsTrue(
                            instance.ObjectData.OSIDs.Body.Length == 1,
                            "MS-ONESTORE",
                            791,
                            @"[In PropertyID] value ""0xA"", name ""ObjectSpaceID"": The property contains one CompactID structure in the ObjectSpaceObjectPropSet.OSIDs.body stream field.");
                    }
                    if (propId.Type == 0xB)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R792
                        site.CaptureRequirementIfIsTrue(
                            instance.ObjectData.OSIDs.Body.Length > 1,
                            "MS-ONESTORE",
                            792,
                            @"[In PropertyID] value ""0xB"", name ""ArrayOfObjectSpaceIDs"": The property contains an array of CompactID structures in the ObjectSpaceObjectPropSet.OSIDs.body stream field. ");

                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R793
                        site.CaptureRequirementIfAreEqual<uint>(
                            (uint)instance.ObjectData.OSIDs.Body.Length,
                            ((ArrayNumber)instance.ObjectData.Body.RgData[i]).Number,
                            "MS-ONESTORE",
                            793,
                            @"[In PropertyID] value ""0xB"", name ""ArrayOfObjectSpaceIDs"": The property contains an unsigned integer of size 4 bytes in the PropertySet.rgData stream field that specifies the number of CompactID structures this property contains.");
                    }
                    if (propId.Type == 0xC)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R794
                        site.CaptureRequirementIfIsTrue(
                            instance.ObjectData.ContextIDs.Body.Length == 1,
                            "MS-ONESTORE",
                            794,
                            @"[In PropertyID] value ""0xC"", name ""ContextID"": The property contains one CompactID in the ObjectSpaceObjectPropSet.ContextIDs.body stream field.");
                    }
                    if (propId.Type == 0xD)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R795
                        site.CaptureRequirementIfIsTrue(
                            instance.ObjectData.ContextIDs.Body.Length > 1,
                            "MS-ONESTORE",
                            795,
                            @"[In PropertyID] value ""0xD"", name ""ArrayOfContextIDs"": The property contains an array of CompactID structures in the ObjectSpaceObjectPropSet.ContextIDs.body stream field.");

                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R796
                        site.CaptureRequirementIfAreEqual<uint>(
                            (uint)instance.ObjectData.ContextIDs.Body.Length,
                            ((ArrayNumber)instance.ObjectData.Body.RgData[i]).Number,
                            "MS-ONESTORE",
                            796,
                            @"[In PropertyID] value ""0xD"", name ""ArrayOfContextIDs"": The property contains an unsigned integer of size 4 bytes in the PropertySet.rgData stream field that specifies the number of CompactID structures this property contains.");

                    }
                    if (propId.Type == 0x10)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R797
                        site.CaptureRequirementIfIsInstanceOfType(
                            instance.ObjectData.Body.RgData[i],
                            typeof(PrtArrayOfPropertyValues),
                            "MS-ONESTORE",
                            797,
                            @"[In PropertyID] value ""0x10"", name ""ArrayOfPropertyValues"": The property contains a prtArrayOfPropertyValues (section 2.6.9) structure in the PropertySet.rgData stream field.");

                        this.VerifyPrtArrayOfPropertyValues((PrtArrayOfPropertyValues)instance.ObjectData.Body.RgData[i], site);
                    }
                    if (propId.Type == 0x11)
                    {
                        // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R798
                        site.CaptureRequirementIfIsInstanceOfType(                           
                            instance.ObjectData.Body.RgData[i],
                            typeof(PropertySet),
                            "MS-ONESTORE",
                            798,
                            @"[In PropertyID] value ""0x11"", name ""PropertySet"": The property contains a child PropertySet (section 2.6.7) structure in the PropertySet.rgData stream field of the parent PropertySet.");
                    }
                }
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with the ObjectSpaceObjectPropSet structure.
        /// </summary>
        /// <param name="instance">The instance of ObjectSpaceObjectPropSet structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyObjectSpaceObjectPropSet(ObjectSpaceObjectPropSet instance, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R729
            site.CaptureRequirementIfIsInstanceOfType(
                instance.OIDs,
                typeof(ObjectSpaceObjectStreamOfOIDs),
                "MS-ONENOTE",
                729,
                @"[In ObjectSpaceObjectPropSet] OIDs (variable): An ObjectSpaceObjectStreamOfOIDs (section 2.6.2) that specifies the count and list of objects that are referenced by this ObjectSpaceObjectPropSet. ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R731
            site.CaptureRequirementIfAreEqual<uint>(
                    instance.OIDs.Header.Count,
                    (uint)instance.OIDs.Body.Length,
                    "MS-ONENOTE",
                    731,
                    @"[In ObjectSpaceObjectPropSet] [OIDs] This count MUST be equal to the value of OIDs.header.Count field. ");

            this.VerifyObjectSpaceObjectStreamOfOIDs(instance, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R734, MS-ONESTORE_R735
            if (instance.OIDs.Header.OsidStreamNotPresent == 0)
            {
                site.CaptureRequirementIfIsNotNull(
                    instance.OSIDs,
                    "MS-ONENOTE",
                    734,
                    @"[In ObjectSpaceObjectPropSet] [OSIDs] MUST be present if the value of the OIDs.header.OsidStreamNotPresent field is false;");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R733
                site.CaptureRequirementIfIsInstanceOfType(
                    instance.OSIDs,
                    typeof(ObjectSpaceObjectStreamOfOSIDs),
                    "MS-ONENOTE",
                    733,
                    @"[In ObjectSpaceObjectPropSet] OSIDs (variable): An optional ObjectSpaceObjectStreamOfOSIDs structure (section 2.6.3) that specifies the count and list of object spaces referenced by this ObjectSpaceObjectPropSet structure.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R737
                site.CaptureRequirementIfAreEqual<uint>(
                        instance.OSIDs.Header.Count,
                        (uint)instance.OSIDs.Body.Length,
                        "MS-ONENOTE",
                        737,
                        @"[In ObjectSpaceObjectPropSet] [OSIDs] This count MUST be equal to the value of OSIDs.header.Count field.");
            }
            else
            {
                site.CaptureRequirementIfIsNull(
                    instance.OSIDs,
                    "MS-ONENOTE",
                    735,
                    @"[In ObjectSpaceObjectPropSet] [OSIDs] otherwise[if the value of the OIDs.header.OsidStreamNotPresent field is true], the OSIDs field MUST NOT be present. ");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R740, MS-ONESTORE_R741
            if (instance.OSIDs != null && instance.OSIDs.Header.ExtendedStreamsPresent == 1)
            {
                site.CaptureRequirementIfIsNotNull(
                    instance.ContextIDs,
                    "MS-ONENOTE",
                    740,
                    @"[In ObjectSpaceObjectPropSet] [ContextIDs] MUST be present if OSIDs is present and the value of the OSIDs.header.ExtendedStreamsPresent field is true; ");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R739
                site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ObjectSpaceObjectStreamOfContextIDs),
                    instance.ContextIDs.GetType(),
                    "MS-ONENOTE",
                    739,
                    @"[In ObjectSpaceObjectPropSet] ContextIDs (variable): An optional ObjectSpaceObjectStreamOfContextIDs (section 2.6.4) that specifies the count and list of contexts referenced by this ObjectSpaceObjectPropSet structure.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R743
                site.CaptureRequirementIfAreEqual<uint>(
                        instance.ContextIDs.Header.Count,
                        (uint)instance.ContextIDs.Body.Length,
                        "MS-ONENOTE",
                        743,
                        @"[In ObjectSpaceObjectPropSet] [ContextIDs] This count MUST be equal to the value of ContextIDs.header.Count field.");
            }
            else
            {
                site.CaptureRequirementIfIsNull(
                    instance.ContextIDs,
                    "MS-ONENOTE",
                    741,
                    @"[In ObjectSpaceObjectPropSet] [ContextIDs] otherwise[if OSIDs is not present or the value of the OSIDs.header.ExtendedStreamsPresent field is false ], the ContextIDs field MUST NOT be present. ");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R745
            site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertySet),
                instance.Body.GetType(),
                "MS-ONENOTE",
                745,
                @"[In ObjectSpaceObjectPropSet] body (variable): A PropertySet structure (section 2.6.7) that specifies properties that modify this object, and how other objects relate to this object. ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R748
            site.CaptureRequirementIfIsTrue(
                 instance.Padding.Length < 7,
                 "MS-ONENOTE",
                 748,
                 @"[In ObjectSpaceObjectPropSet] The size of the padding field MUST NOT exceed 7 bytes. ");
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the ObjectSpaceObjectStreamOfOIDs structure.
        /// </summary>
        /// <param name="instance">The instance of ObjectSpaceObjectStreamOfOIDs structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyObjectSpaceObjectStreamOfOIDs(ObjectSpaceObjectPropSet instance,ITestSite site)
        {
            //Verify MS-ONESTORE requirement: MS-ONESTORE_R751
            site.CaptureRequirementIfIsInstanceOfType(
                instance.OIDs.Header,
                typeof(ObjectSpaceObjectStreamHeader),
                "MS-ONENOTE",
                751,
                @"[In ObjectSpaceObjectStreamOfOIDs] header (4 bytes): An ObjectSpaceObjectStreamHeader structure (section 2.6.5) that specifies the number of elements in the body field and whether the ObjectSpaceObjectPropSet structure contains an OSIDs field and ContextIDs field.");

            this.VerifyObjectSpaceObjectStreamHeader(instance.OIDs.Header, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R752, MS-ONESTORE_R753
            if (instance.OSIDs != null)
            {
                site.CaptureRequirementIfAreEqual<int>(
                    0,
                    instance.OIDs.Header.OsidStreamNotPresent,
                    "MS-ONENOTE",
                    752,
                    @"[In ObjectSpaceObjectStreamOfOIDs] [header] If the OSIDs field is present, the value of the header.OsidStreamNotPresent field MUST be false;");
            }
            else
            {
                site.CaptureRequirementIfAreEqual<int>(
                    1,
                    instance.OIDs.Header.OsidStreamNotPresent,
                    "MS-ONENOTE",
                    753,
                    @"[In ObjectSpaceObjectStreamOfOIDs] [header] otherwise [the OSIDs field is not present], it [the value of the header.OsidStreamNotPresent field] MUST be true.");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R754, MS-ONESTORE_R755
            if (instance.ContextIDs != null)
            {
                site.CaptureRequirementIfAreEqual<int>(
                    1,
                    instance.OIDs.Header.ExtendedStreamsPresent,
                    "MS-ONENOTE",
                    754,
                    @"[In ObjectSpaceObjectStreamOfOIDs] [header] If the ContextIDs field is present, the value of the header.ExtendedStreamsPresent field MUST be true; ");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R757
                site.CaptureRequirementIfAreEqual<uint>(
                        instance.ContextIDs.Header.Count,
                        (uint)instance.ContextIDs.Body.Length,
                        "MS-ONENOTE",
                        757,
                        @"[In ObjectSpaceObjectStreamOfContextIDs] [body] The number of elements is equal to the value of the header.Count field.");
            }
            else
            {
                site.CaptureRequirementIfAreEqual<int>(
                    0,
                    instance.OIDs.Header.ExtendedStreamsPresent,
                    "MS-ONENOTE",
                    755,
                    @"[In ObjectSpaceObjectStreamOfOIDs]  [header] otherwise[If the ContextIDs field is not present], it [the value of the header.ExtendedStreamsPresent field] MUST be false.");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R756
            site.CaptureRequirementIfIsInstanceOfType(
                instance.OIDs.Body,
                typeof(CompactID[]),
                "MS-ONESTORE",
                756,
                @"[In ObjectSpaceObjectStreamOfOIDs] body (variable): An array of CompactID structures (section 2.2.2) where each element in the array specifies the identity of an object.");
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the ObjectSpaceObjectStreamHeader structure.
        /// </summary>
        /// <param name="instance">The instance of ObjectSpaceObjectStreamHeader structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyObjectSpaceObjectStreamHeader(ObjectSpaceObjectStreamHeader instance,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R771
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Count,
                typeof(UInt32),
                "MS-ONESTORE",
                771,
                @"[In ObjectSpaceObjectStreamHeader] Count (24 bits): An unsigned integer that specifies the number of CompactID structures (section 2.2.2) in the stream that contains this ObjectSpaceObjectStreamHeader structure.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R772
            site.CaptureRequirementIfAreEqual<int>(
                0,
                instance.Reserved,
                "MS-ONESTORE",
                772,
                @"[In ObjectSpaceObjectStreamHeader] Reserved (6 bits): MUST be zero, and MUST be ignored.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R773
            site.CaptureRequirementIfIsTrue(
                instance.ExtendedStreamsPresent == 0 || instance.ExtendedStreamsPresent == 1,
                 "MS-ONESTORE",
                773,
                @"[In ObjectSpaceObjectStreamHeader] A - ExtendedStreamsPresent (1 bit): A bit that specifies whether the ObjectSpaceObjectPropSet structure (section 2.6.1) contains  any additional streams of data following this stream of data.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R774
            site.CaptureRequirementIfIsTrue(
                instance.OsidStreamNotPresent == 0 || instance.OsidStreamNotPresent == 1,
                 "MS-ONESTORE",
                774,
                @"[In ObjectSpaceObjectStreamHeader] B - OsidStreamNotPresent (1 bit): A bit that specifies whether the ObjectSpaceObjectPropSet structure does not contain OSIDs or ContextIDs fields.");
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the ObjectSpaceObjectStreamOfOSIDs structure.
        /// </summary>
        /// <param name="instance">The instance of ObjectSpaceObjectStreamOfOSIDs structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyObjectSpaceObjectStreamOfOSIDs(ObjectSpaceObjectStreamOfOSIDs instance, ObjectSpaceObjectStreamOfContextIDs contextIDs, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R759
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Header,
                typeof(ObjectSpaceObjectStreamHeader),
                "MS-ONESTORE",
                759,
                @"[In ObjectSpaceObjectStreamOfOSIDs] header (4 bytes): An ObjectSpaceObjectStreamHeader structure (section 2.6.5) that specifies the number of elements in the body field and whether the ObjectSpaceObjectPropSet structure contains ContextIDs field. ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R760
            site.CaptureRequirementIfAreEqual<Int32>(
                0,
                instance.Header.OsidStreamNotPresent,
                "MS-ONESTORE",
                760,
                @"[In ObjectSpaceObjectStreamOfOSIDs] The value of the header.OsidStreamNotPresent field MUST be ""false"". ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R764
            site.CaptureRequirementIfAreEqual<uint>(
                    instance.Header.Count,
                    (uint)instance.Body.Length,
                    "MS-ONESTORE",
                    764,
                    @"[In ObjectSpaceObjectStreamOfOSIDs] The number of elements is equal to the value of the header.Count field.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R761,MS-ONESTORE_R762
            if (contextIDs != null)
            {
                site.CaptureRequirementIfAreEqual<Int32>(
                    1,
                    instance.Header.ExtendedStreamsPresent,
                    "MS-ONESTORE",
                    761,
                    @"[In ObjectSpaceObjectStreamOfOSIDs] If the ContextIDs field is present, the value of the header.ExtendedStreamsPresent field MUST be true;");
            }
            else
            {
                site.CaptureRequirementIfAreEqual<Int32>(
                    0,
                    instance.Header.ExtendedStreamsPresent,
                    "MS-ONESTORE",
                    762,
                    @"[In ObjectSpaceObjectStreamOfOSIDs] otherwise[If the ContextIDs field is not present], it[the value of the header.ExtendedStreamsPresent field] MUST be false.");

            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R763
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Body,
                typeof(CompactID[]),
                "MS-ONESTORE",
                763,
                @"[In ObjectSpaceObjectStreamOfOSIDs] body (variable): An array of CompactID structures (section 2.2.2) where each element in the array specifies the identity of an object space.");
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the ObjectSpaceObjectStreamOfContextIDs structure.
        /// </summary>
        /// <param name="instance">The instance of ObjectSpaceObjectStreamOfContextIDs structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyObjectSpaceObjectStreamOfContextIDs(ObjectSpaceObjectStreamOfContextIDs instance,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R766
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Header,
                typeof(ObjectSpaceObjectStreamHeader),
                "MS-ONESTORE",
                766,
                @"[In ObjectSpaceObjectStreamOfContextIDs] header (4 bytes): An ObjectSpaceObjectStreamHeader structure (section 2.6.5) that specifies the number of elements in the body field.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R767
            site.CaptureRequirementIfIsTrue(
                instance.Header.OsidStreamNotPresent==0,
                "MS-ONESTORE",
                767,
                @"[In ObjectSpaceObjectStreamOfContextIDs] The value of header.OsidStreamNotPresent field and header.ExtendedStreamsPresent field MUST be false.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R768
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Body,
                typeof(CompactID[]),
                "MS-ONESTORE",
                768,
                @"[In ObjectSpaceObjectStreamOfContextIDs] body (variable): An array of CompactID structures (section 2.2.2) where each element in the array specifies a context (section 2.1.11).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R769
            site.CaptureRequirementIfAreEqual<uint>(
                 instance.Header.Count,
                 (uint)instance.Body.Length,
                 "MS-ONESTORE",
                 769,
                 @"[In ObjectSpaceObjectStreamOfContextIDs] The number of elements is equal to the value of the header.Count field.");
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the PropertySet structure.
        /// </summary>
        /// <param name="instance">The instance of PropertySet structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyPropertySet(PropertySet instance,ITestSite site)
        {
            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R801
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.CProperties,
                    typeof(UInt16),
                    "MS-ONESTORE",
                    801,
                    @"[In PropertySet] cProperties (2 bytes): An unsigned integer that specifies the number of properties in this PropertySet structure.");

            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R802
            site.CaptureRequirementIfIsInstanceOfType(
                instance.RgPrids,
                typeof(PropertyID[]),
                "MS-ONESTORE",
                802,
                @"[In PropertySet] rgPrids (variable): An array of PropertyID structures (section 2.6.6).");

            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R80201
            site.CaptureRequirementIfAreEqual<int>(
                instance.CProperties,
                instance.RgPrids.Length,
                80201,
                @"[In PropertySet] rgPrids (variable): The number of elements in the array is equal to the value of the cProperties field.");

            for (int i = 0; i < instance.RgPrids.Length; i++)
            {
                this.VerifyPropertyID(instance.RgPrids[i], site);
            }

            // If the rgData is parse successfully,then R803 and R804 will be verified.
            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R803
            site.CaptureRequirementIfIsNotNull(
                instance.RgData,
                "MS-ONESTORE",
                803,
                @"[In PropertySet] rgData (variable): A stream of bytes that specifies the data for each property specified by a rgPrids array.");

            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R804
            site.CaptureRequirement(
                "MS-ONESTORE",
                804,
                @"[In PropertySet] [rgData] The total size, in bytes, of the rgData field is the sum of the sizes specified by the PropertyID.type field for each property in a rgPrids array.");

            if (instance.RgPrids.Length == 0)
            {
                // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R805
                site.CaptureRequirementIfIsTrue(
                    instance.RgData.Count==0,
                    "MS-ONESTORE",
                    805,
                    @"[In PropertySet] [rgData] The total size of rgData MUST be zero if no property in a rgPrids array specifies that it contains data in the rgData field.");
            }
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the PropertyID structure.
        /// </summary>
        /// <param name="instance">The instance of PropertyID structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyPropertyID(PropertyID instance,ITestSite site)
        {
            // Verfiy MS-ONESTORE requirement: MS - ONESTORE_R777,MS - ONESTORE_R778
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Id,
                typeof(UInt32),
                "MS-ONESTORE",
                777,
                @"[In PropertyID] id (26 bits): An unsigned integer that specifies the identity of this property.");

            //Verfiy MS-ONESTORE requirement: MS-ONESTORE_R780
            site.CaptureRequirementIfIsTrue(
                    instance.Type == 0x1 || instance.Type == 0x2 || instance.Type == 0x3 || instance.Type == 0x4 || instance.Type == 0x5 ||
                    instance.Type == 0x6 || instance.Type == 0x7 || instance.Type == 0x8 || instance.Type == 0x9 || instance.Type == 0xA ||
                    instance.Type == 0xB || instance.Type == 0xC || instance.Type == 0xD || instance.Type == 0x10 || instance.Type == 0x11,
                    "MS-ONESTORE",
                    780,
                    @"[In PropertyID] [type] MUST be one of the following values: [0x1,0x2,0x3,0x4,0x5,0x6,0x7,0x8,0x9,0xA,0xB,0xC,0xD,0x10,0x11]");

            //Verfiy MS-ONESTORE requirement: MS-ONESTORE_R779
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Type,
                typeof(UInt32),
                "MS-ONESTORE",
                779,
                @"[In PropertyID] type (5 bits): An unsigned integer that specifies the property type and the size and location of the data for this property.");
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the prtFourBytesOfLengthFollowedByData structure.
        /// </summary>
        /// <param name="instance">The instance of prtFourBytesOfLengthFollowedByData structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyPrtFourBytesOfLengthFollowedByData(PrtFourBytesOfLengthFollowedByData instance, ITestSite site)
        {
            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R807
            site.CaptureRequirementIfAreEqual<uint>(
                (uint)instance.SerializeToByteList().Count,
                instance.CB + 4,
                "MS-ONESTORE",
                807,
                @"[In prtFourBytesOfLengthFollowedByData] The total size, in bytes, of prtFourBytesOfLengthFollowedByData is equal to cb + 4.");

            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R808
            site.CaptureRequirementIfIsTrue(
                instance.CB< 0x40000000,
                "MS-ONESTORE",
                808,
                @"[In prtFourBytesOfLengthFollowedByData] cb (4 bytes): An unsigned integer that specifies the size, in bytes, of the Data field. MUST be less than 0x40000000.");

            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R809
            site.CaptureRequirementIfIsInstanceOfType(
                instance.Data,
                typeof(byte[]),
                "MS-ONESTORE",
                809,
                @"[In prtFourBytesOfLengthFollowedByData] Data (variable): A stream of bytes that specifies the data for the property.");
        }
        /// <summary>
        ///  This method is used to verify the requirements related with the prtArrayOfPropertyValues structure.
        /// </summary>
        /// <param name="instance">The instance of prtArrayOfPropertyValues structure.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        private void VerifyPrtArrayOfPropertyValues(PrtArrayOfPropertyValues instance,ITestSite site)
        {
            // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R811
            site.CaptureRequirementIfIsInstanceOfType(
                instance.CProperties,
                typeof(uint),
                "MS-ONESTORE",
                811,
                @"[In prtArrayOfPropertyValues] cProperties (4 bytes): An unsigned integer that specifies the number of properties in Data.");

            if (instance.CProperties !=0)
            {
                // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R812
                site.CaptureRequirementIfIsInstanceOfType(
                    instance.Prid,
                    typeof(PropertyID),
                    "MS-ONESTORE",
                    812,
                    @"[In prtArrayOfPropertyValues] prid (4 bytes): An optional PropertyID structure (section 2.6.6) that specifies the type of each property in the array.");

                // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R813
                site.CaptureRequirementIfAreEqual<uint>(
                    0x11,
                    instance.Prid.Type,
                    "MS-ONESTORE",
                    813,
                    @"[In prtArrayOfPropertyValues] PropertyID.type MUST be 0x11 (""PropertySet"").");

                // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R816
                site.CaptureRequirement(
                    "MS-ONESTORE",
                    816,
                    @"[In prtArrayOfPropertyValues] otherwise[if cProperties is not zero],[prid] MUST be present.");

                site.CaptureRequirement(
                    "MS-ONESTORE",
                    817,
                    @"[In prtArrayOfPropertyValues] [Data] The total size, in bytes, of the Data field is the sum of the sizes specified by the prid.type field for each property in the array, if prid is present.");

                site.CaptureRequirement(
                    "MS-ONESTORE",
                    818,
                    @"[In prtArrayOfPropertyValues] [Data] The total size, in bytes, of the Data field is the sum of the sizes specified by the prid.type field for each property in the array, if prid is present.");
            }
            else
            {
                // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R815
                site.CaptureRequirementIfIsNull(
                    instance.Prid,
                    "MS-ONESTORE",
                    815,
                    @"[In prtArrayOfPropertyValues] [prid] MUST NOT be present if cProperties is zero;");

                // Verfiy MS-ONESTORE requirement: MS-ONESTORE_R819
                site.CaptureRequirementIfAreEqual<int>(
                    0,
                    instance.Data.Length,
                    "MS-ONESTORE",
                    819,
                    @"[In prtArrayOfPropertyValues] Otherwise[if prid is not present], the total size of Data is zero if cProperties is zero.");
            }
        }
    }
}
