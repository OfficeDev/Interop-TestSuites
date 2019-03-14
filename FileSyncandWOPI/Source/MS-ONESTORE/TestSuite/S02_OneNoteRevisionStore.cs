namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// This scenario is designed to test the requirements related with .one file.
    /// </summary>
    [TestClass]
    public class S02_OneNoteRevisionStore : TestSuiteBase
    {
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #region Test cases
        /// <summary>
        /// The test case is validate that the requirements related with .one file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S02_TC01_LoadOneNoteFileWithFileData()
        {
            string fileName = Common.GetConfigurationPropertyValue("OneFileWithFileData", Site);

            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(fileName);

            for (int i = 0; i <= file.RootFileNodeList.FileNodeListFragments.Count - 1; i++)
            {
                bool isDifferent = file.RootFileNodeList.FileNodeListFragments[i].Header.FileNodeListID != file.RootFileNodeList.FileNodeListFragments[i].Header.nFragmentSequence;

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R3911
                Site.CaptureRequirementIfIsTrue(
                    isDifferent,
                    3911,
                    @"[In FileNodeListHeader] The pair of FileNodeListID and nFragmentSequence field in FileNodeListFragment structures in the file, is different.");
            }

            List<RevisionManifest> revisManifestList = new List<RevisionManifest>();

            foreach (ObjectSpaceManifestList objSpaceManifestList in file.RootFileNodeList.ObjectSpaceManifestList)
            {
                foreach (RevisionManifestList revManifestList in objSpaceManifestList.RevisionManifestList)
                {
                    revisManifestList.AddRange(revManifestList.RevisionManifests);
                }
            }

            List<ExtendedGUID> ridRevisionManifestStart6FND = new List<ExtendedGUID>();
            List<ExtendedGUID> ridDependentRevisionManifestStart6FND = new List<ExtendedGUID>();
            List<ExtendedGUID> ridRevisionManifestStart7FND = new List<ExtendedGUID>();
            List<ExtendedGUID> ridDependentRevisionManifestStart7FND = new List<ExtendedGUID>();
            List<uint> odcsDefault = new List<uint>();

            for (int i = 0; i < revisManifestList.Count; i++)
            {
                for (int j = 0; j < revisManifestList[i].FileNodeSequence.Count; j++)
                {
                    FileNode revision = revisManifestList[i].FileNodeSequence[j];

                    if (revision.FileNodeID == FileNodeIDValues.RevisionManifestStart6FND)
                    {
                        RevisionManifestStart6FND fnd = revision.fnd as RevisionManifestStart6FND;
                        ridRevisionManifestStart6FND.Add(((RevisionManifestStart6FND)revision.fnd).rid);
                        ridDependentRevisionManifestStart6FND.Add(((RevisionManifestStart6FND)revision.fnd).ridDependent);
                        odcsDefault.Add(((RevisionManifestStart6FND)revision.fnd).odcsDefault);
                    }

                    if (revision.FileNodeID == FileNodeIDValues.RevisionManifestStart7FND)
                    {
                       
                        RevisionManifestStart7FND fnd = revision.fnd as RevisionManifestStart7FND;
                        ridRevisionManifestStart7FND.Add(((RevisionManifestStart7FND)revision.fnd).Base.rid);
                        ridDependentRevisionManifestStart7FND.Add(((RevisionManifestStart7FND)revision.fnd).Base.ridDependent);
                    }
                }
            }

            for (int i=0; i<odcsDefault.Count-1; i++)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R535
                Site.CaptureRequirementIfAreEqual<uint>(
                                odcsDefault[i],
                                odcsDefault[i + 1],
                                535,
                                @"[In RevisionManifestStart6FND] [odcsDefault] MUST specify the same type of data encoding as used in the dependency revision (section 2.1.8), if one was specified in the ridDependent field.");
            }

            ExtendedGUID zeroExtendGuid = new ExtendedGUID();
            zeroExtendGuid.Guid = Guid.Empty;
            zeroExtendGuid.N = 0;

            for (int i = 0; i < ridRevisionManifestStart6FND.Count - 1; i++)
            {
                for (int j = i + 1; j < ridRevisionManifestStart6FND.Count; j++)
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R52501
                    Site.CaptureRequirementIfAreNotEqual<ExtendedGUID>(
                                    ridRevisionManifestStart6FND[i],
                                    ridRevisionManifestStart6FND[j],
                                    52501,
                                    @"[In RevisionManifestStart6FND] The rid of two RevisionManifestStart6FND in revision manifest list is different.");
                }

                if (ridDependentRevisionManifestStart6FND[i].Equals(zeroExtendGuid))
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R527
                    Site.CaptureRequirementIfAreNotEqual<ExtendedGUID>(
                        zeroExtendGuid,
                        ridRevisionManifestStart6FND[i],
                                    527,
                                    @"[In RevisionManifestStart6FND] [ridDependent] If the value is ""{ { 00000000 - 0000 - 0000 - 0000 - 000000000000}, 0}
                "", then this revision manifest has no dependency revision. ");
                }
                else
                {
                    if (!ridDependentRevisionManifestStart6FND[i + 1].Equals(zeroExtendGuid))
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R528
                        Site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                                        ridRevisionManifestStart6FND[i],
                                        ridDependentRevisionManifestStart6FND[i + 1],
                                        528,
                                        @"[In RevisionManifestStart6FND] [ridDependent] Otherwise[If the value is not ""{ { 00000000 - 0000 - 0000 - 0000 - 000000000000}, 0}
                    ""], this value MUST be equal to the RevisionManifestStart6FND.rid field [or the RevisionManifestStart7FND.base.rid field] of a previous revision manifest within this revision manifest list.");
                    }
                }
            }

            for (int i = 0; i < ridRevisionManifestStart7FND.Count - 1; i++)
            {
                for (int j = i + 1; j < ridRevisionManifestStart7FND.Count; j++)
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R52502
                    Site.CaptureRequirementIfAreNotEqual<ExtendedGUID>(
                                    ridRevisionManifestStart7FND[i],
                                    ridRevisionManifestStart7FND[i + 1],
                                    52502,
                                    @"[In RevisionManifestStart6FND] The rid of two RevisionManifestStart7FND in revision manifest list is different.");
                }

                if (!ridDependentRevisionManifestStart7FND[i].Equals(zeroExtendGuid) && !ridDependentRevisionManifestStart7FND[i+1].Equals(zeroExtendGuid))
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R52801
                    Site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                                    ridRevisionManifestStart7FND[i],
                                    ridDependentRevisionManifestStart7FND[i+1],
                                    52801,
                                    @"[In RevisionManifestStart6FND] [ridDependent] Otherwise[If the value is not ""{ { 00000000 - 0000 - 0000 - 0000 - 000000000000}, 0}
                    ""], this value MUST be equal to [the RevisionManifestStart6FND.rid field or] the RevisionManifestStart7FND.base.rid field of a previous revision manifest within this revision manifest list.");
                }
            }
            foreach (ObjectSpaceManifestList objectSpaceManifestList in file.RootFileNodeList.ObjectSpaceManifestList)
            {
                foreach (RevisionManifestList revisionManifestList in objectSpaceManifestList.RevisionManifestList)
                {
                    if (revisionManifestList.FileNodeListFragments.Count > 1)
                    {
                        for (int i = 0; i < revisionManifestList.FileNodeListFragments.Count - 1; i++)
                        {
                            for (int j = i + 1; j < revisionManifestList.FileNodeListFragments.Count; j++)
                            {
                                // Verify MS-ONESTORE requirement: MS-ONESTORE_R366
                                Site.CaptureRequirementIfAreEqual<uint>(
                                    revisionManifestList.FileNodeListFragments[i].Header.FileNodeListID,
                                    revisionManifestList.FileNodeListFragments[j].Header.FileNodeListID,
                                    366,
                                    @"[In FileNodeListFragment] All fragments in the same file node list MUST have the same FileNodeListFragment.header.FileNodeListID field.");
                            }
                        }
                    }

                    foreach (FileNodeListFragment fileNodeListFragment in revisionManifestList.FileNodeListFragments)
                    {
                        FileChunkReference64x32 nextFragment = fileNodeListFragment.nextFragment;
                        if ((uint)fileNodeListFragment.rgFileNodes[fileNodeListFragment.rgFileNodes.Count - 1].FileNodeID == 0xFF)
                        {
                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R373
                            Site.CaptureRequirementIfIsInstanceOfType(
                                nextFragment,
                                typeof(FileChunkReference64x32),
                                373,
                                @"[In FileNodeListFragment] [rgFileNodes If a ChunkTerminatorFND structure is present, the value of the nextFragment field MUST be a valid FileChunkReference64x32 structure (section 2.2.4.4) to the next FileNodeListFragment structure.");
                        }

                        if (nextFragment == revisionManifestList.FileNodeListFragments[revisionManifestList.FileNodeListFragments.Count - 1].nextFragment)
                        {
                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R379
                            Site.CaptureRequirementIfIsTrue(
                                nextFragment.IsfcrNil(),
                                379,
                                @"[In FileNodeListFragment] If this is the last fragment, the value of the nextFragment field MUST be ""fcrNil"" (see section 2.2.4). ");
                        }
                        else
                        {
                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R380
                            Site.CaptureRequirementIfIsFalse(
                                nextFragment.IsfcrZero() || nextFragment.IsfcrNil(),
                                380,
                                @"[In FileNodeListFragment] Otherwise [If this is not the last fragment] the value of the nextFragment.stp field MUST specify the location of a valid FileNodeListFragment structure, and the value of the nextFragment.cb field MUST be equal to the size of the referenced fragment including the FileNodeListFragment.header field and the FileNodeListFragment.footer field.");
                        }
                    }
                }
            }

            bool isFileNode = false;

            foreach (FileNode fileDataStoreList in file.RootFileNodeList.FileDataStoreListReference)
            {
                FileDataStoreListReferenceFND fnd = fileDataStoreList.fnd as FileDataStoreListReferenceFND;

                for(int i=0; i<fnd.fileNodeListFragment.rgFileNodes.Count-1; i++)
                {
                    for (int j = i + 1; j < fnd.fileNodeListFragment.rgFileNodes.Count; j++)
                    {
                        FileDataStoreObjectReferenceFND objfnd1 = fnd.fileNodeListFragment.rgFileNodes[i].fnd as FileDataStoreObjectReferenceFND;
                        FileDataStoreObjectReferenceFND objfnd2 = fnd.fileNodeListFragment.rgFileNodes[j].fnd as FileDataStoreObjectReferenceFND;

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R63701
                        Site.CaptureRequirementIfAreNotEqual<Guid>(
                                objfnd1.guidReference,
                                objfnd2.guidReference,
                                63701,
                                @"[In FileDataStoreObjectReferenceFND] The guidReference is different for two FileDataStoreObjectReferenceFND structures.");
                    }
                }

                isFileNode = true;

                for (int k = 0; k < fnd.fileNodeListFragment.rgFileNodes.Count; k++)
                {
                    FileNode fileNode = fnd.fileNodeListFragment.rgFileNodes[k];
                    if ((uint)fileNode.FileNodeID != 0x094)
                    {
                        isFileNode = false;
                        break;
                    }

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R628
                    Site.CaptureRequirementIfIsTrue(
                            isFileNode,
                            628,
                            @"[In FileDataStoreListReferenceFND] The referenced file node list MUST contain only FileNode structures with a FileNodeID field value equal to 00x094 (FileDataStoreObjectReferenceFND structure). ");

                    //If MS-ONESTORE_R628 is verified successfully and according to the definition of FileDataStoreObjectReferenceFND, MS-ONESTORE_R33 can be verified directlly.
                    //Verify MS-ONESTORE requirement: MS-ONESTORE_R633
                    Site.CaptureRequirement(
                            633,
                            @"[[In FileDataStoreObjectReferenceFND] All such FileNode structures MUST be contained in the file node list (section 2.4) specified by a FileDataStoreListReferenceFND structure (section 2.5.21).");
                }
            }

            int objectSpaceCount = file.RootFileNodeList.ObjectSpaceManifestList.Count;

                for (int i = 0; i < file.RootFileNodeList.ObjectSpaceManifestList.Count; i++)
                {
                    ObjectSpaceManifestList objectSpace = file.RootFileNodeList.ObjectSpaceManifestList[i];

                    for (int j = 0; j < objectSpace.RevisionManifestList[0].ObjectGroupList.Count; j++)
                    {
                        ObjectGroupList objectGroupList = objectSpace.RevisionManifestList[0].ObjectGroupList[j];
                        FileNode[] fileNodes = objectGroupList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ReadOnlyObjectDeclaration2RefCountFND).ToArray();

                        foreach (FileNode node in fileNodes)
                        {
                            ReadOnlyObjectDeclaration2RefCountFND fnd = node.fnd as ReadOnlyObjectDeclaration2RefCountFND;

                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R68
                            Site.CaptureRequirementIfIsTrue(
                                    (uint)node.FileNodeID == 0x0C4 && fnd.Base.body.jcid.IsReadOnly == 1,
                                    68,
                                    @"[In Object Space Object] If the value of the JCID.IsReadOnly field is ""true"" then the value of the FileNode.FileNodeID field MUST be 0x0C4 (ReadOnlyObjectDeclaration2RefCountFND structure, section 2.5.29) [or 0x0C5 (ReadOnlyObjectDeclaration2LargeRefCountFND structure, section 2.5.30).]");

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R703
                        Site.CaptureRequirementIfIsNotNull(
                                fnd.Base,
                                703,
                                @"[In ReadOnlyObjectDeclaration2RefCountFND] If this object is revised, all declarations of this object MUST specify identical data. ");
                    }

                    fileNodes = objectGroupList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3RefCountFND).ToArray();

                        foreach (FileNode node in fileNodes)
                        {
                            ObjectDeclarationFileData3RefCountFND fnd = node.fnd as ObjectDeclarationFileData3RefCountFND;

                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R71
                            Site.CaptureRequirementIfIsTrue(
                                    (uint)node.FileNodeID == 0x072 && fnd.jcid.IsFileData == 1,
                                    71,
                                    @"[In Object Space Object]If the value of the JCID.IsFileData field is ""true"" then the value of the FileNode.FileNodeID field MUST be 0x072 (ObjectDeclarationFileData3RefCountFND structure, section 2.5.27) [or 0x073 (ObjectDeclarationFileData3LargeRefCountFND structure, section 2.5.28). ]");
                    }

                    fileNodes = objectGroupList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ReadOnlyObjectDeclaration2LargeRefCountFND).ToArray();

                    foreach (FileNode node in fileNodes)
                    {
                        ReadOnlyObjectDeclaration2LargeRefCountFND fnd = node.fnd as ReadOnlyObjectDeclaration2LargeRefCountFND;

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R710
                        Site.CaptureRequirementIfIsNotNull(
                                fnd.Base,
                                710,
                                @"[In ReadOnlyObjectDeclaration2LargeRefCountFND] If this object is revised, all declarations of this object MUST specify identical data. ");
                    }
                    fileNodes = objectGroupList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.DataSignatureGroupDefinitionFND).ToArray();

                    List<ExtendedGUID> dataSignatureGroupdefinitionFND = new List<ExtendedGUID>();

                    for (int k = 0; k < fileNodes.Length - 1; k++)
                    {
                        DataSignatureGroupDefinitionFND fnd = fileNodes[k].fnd as DataSignatureGroupDefinitionFND;
                        if (!fnd.DataSignatureGroup.Equals(zeroExtendGuid))
                        {
                            dataSignatureGroupdefinitionFND.Add(fnd.DataSignatureGroup);
                        }
                    }

                    for (int k = 0; k < dataSignatureGroupdefinitionFND.Count-1; k++)
                    {
                        for (int h = k + 1; h < dataSignatureGroupdefinitionFND.Count; h++)
                        {
                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R72701
                            Site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                                dataSignatureGroupdefinitionFND[k],
                                    dataSignatureGroupdefinitionFND[h],
                                    72701,
                                    @"[In DataSignatureGroupDefinitionFND] DataSignatureGroup (20 bytes):All declarations of an object (section 2.1.5) with the same identity and the same DataSignatureGroup field not equal to {{00000000-0000-0000-0000-000000000000}, 0} MUST have the same data.");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// The test case is validate that the requirements related with .onetoc2 file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S02_TC02_LoadOnetocFile()
        {
            string fileName = Common.GetConfigurationPropertyValue("OnetocFileLocal", Site);

            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(fileName);

            List<RevisionManifest> revisionManifestList = new List<RevisionManifest>();

            foreach (ObjectSpaceManifestList objSpaceManifestList in file.RootFileNodeList.ObjectSpaceManifestList)
            {
                foreach (RevisionManifestList revManifestList in objSpaceManifestList.RevisionManifestList)
                {
                    revisionManifestList.AddRange(revManifestList.RevisionManifests);
                }
            }

            List<ExtendedGUID> ridRevisionManifestStart4FND = new List<ExtendedGUID>();
            List<ExtendedGUID> ridDependentRevisionManifestStart4FND = new List<ExtendedGUID>();

            for (int i = 0; i < revisionManifestList.Count; i++)
            {
                List<uint> ridGlobalIdTableEntryFNDX = new List<uint>();
                List<uint> ridGlobalIdTableEntry2FNDX = new List<uint>();
                List<uint> ridGlobalIdTableEntry3FNDX = new List<uint>();

                for (int j = 0; j < revisionManifestList[i].FileNodeSequence.Count; j++)
                {
                    FileNode revision = revisionManifestList[i].FileNodeSequence[j];

                    if (revision.FileNodeID == FileNodeIDValues.RevisionManifestStart4FND)
                    {
                        RevisionManifestStart4FND fnd = revision.fnd as RevisionManifestStart4FND;

                        ridRevisionManifestStart4FND.Add(((RevisionManifestStart4FND)revision.fnd).rid);
                        ridDependentRevisionManifestStart4FND.Add(((RevisionManifestStart4FND)revision.fnd).ridDependent);
                    }

                    if (revision.FileNodeID == FileNodeIDValues.GlobalIdTableEntryFNDX)
                    {
                        GlobalIdTableEntryFNDX fnd = revision.fnd as GlobalIdTableEntryFNDX;
                        ridGlobalIdTableEntryFNDX.Add(((GlobalIdTableEntryFNDX)revision.fnd).index);
                    }

                    if (revision.FileNodeID == FileNodeIDValues.GlobalIdTableEntry2FNDX)
                    {
                        GlobalIdTableEntry2FNDX fnd = revision.fnd as GlobalIdTableEntry2FNDX;
                        ridGlobalIdTableEntry2FNDX.Add(((GlobalIdTableEntry2FNDX)revision.fnd).iIndexMapTo);
                    }

                    if (revision.FileNodeID == FileNodeIDValues.GlobalIdTableEntry3FNDX)
                    {
                        GlobalIdTableEntry3FNDX fnd = revision.fnd as GlobalIdTableEntry3FNDX;
                        ridGlobalIdTableEntry3FNDX.Add(((GlobalIdTableEntry3FNDX)revision.fnd).iIndexCopyToStart);
                    }

                    if (revision.FileNodeID == FileNodeIDValues.ObjectInfoDependencyOverridesFND)
                    {
                        ObjectInfoDependencyOverridesFND fnd = revision.fnd as ObjectInfoDependencyOverridesFND;

                        if (fnd.Ref.IsfcrNil())
                        {
                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R620
                            Site.CaptureRequirementIfIsNotNull(
                                            fnd.data,
                                            620,
                                            @"[In ObjectInfoDependencyOverridesFND] otherwise [if the value of the ref field is ""fcrNil""], the override data is specified by the data field. ");
                        }

                    }
                }

                for (int j = 0; j < ridGlobalIdTableEntryFNDX.Count - 1; j++)
                {
                    for (int k = j + 1; k < ridGlobalIdTableEntryFNDX.Count; k++)
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R54901
                        Site.CaptureRequirementIfAreNotEqual<uint>(
                                        ridGlobalIdTableEntryFNDX[j],
                                        ridGlobalIdTableEntryFNDX[k],
                                        54901,
                                        @"[In GlobalIdTableEntryFNDX]  The indexes in two global identification table specified by FileNode structures with the values of the FileNode.FileNodeID fields equal to 0x024 (GlobalIdTableEntryFNDX structure), 0x25 (GlobalIdTableEntry2FNDX structure), and 0x26 (GlobalIdTableEntry3FNDX structure) are different.");
                    }
                }

                for (int j = 0; j < ridGlobalIdTableEntry2FNDX.Count - 1; j++)
                {
                    for (int k = j + 1; k < ridGlobalIdTableEntry2FNDX.Count; k++)
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R55901
                        Site.CaptureRequirementIfAreNotEqual<uint>(
                                        ridGlobalIdTableEntry2FNDX[j],
                                        ridGlobalIdTableEntry2FNDX[k],
                                        55901,
                                        @"[In GlobalIdTableEntry2FNDX] The iIndexMapTo is defferent in two global identification table specified by FileNode structures with the value of the FileNode.FileNodeID field equal to 0x024 (GlobalIdTableEntryFNDX structure), 0x25 (GlobalIdTableEntry2FNDX structure), and 0x26 (GlobalIdTableEntry3FNDX structure).");
                    }
                }

                for(int j=0; j<ridGlobalIdTableEntry3FNDX.Count-1; j++)
                {
                    for (int k = j + 1; k < ridGlobalIdTableEntry3FNDX.Count; k++)
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R57001
                        Site.CaptureRequirementIfAreNotEqual<uint>(
                                        ridGlobalIdTableEntry3FNDX[j],
                                        ridGlobalIdTableEntry3FNDX[k],
                                        57001,
                                        @"[In GlobalIdTableEntry3FNDX] The indices from the value of iIndexCopyToStart to the value of (iIndexCopyToStart + cEntriesToCopy – 1) are different in two global identification table specified by FileNode structures with the values of the FileNode.FileNodeID field equal to 0x024 (GlobalIdTableEntryFNDX structure), 0x025 (GlobalIdTableEntry2FNDX structure), and 0x026 (GlobalIdTableEntry3FNDX structure).");
                    }
                }
            }

            for (int i = 0; i < ridRevisionManifestStart4FND.Count-1; i++)
            {
                for (int j = i + 1; j < ridRevisionManifestStart4FND.Count; j++)
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R51401
                    Site.CaptureRequirementIfAreNotEqual<ExtendedGUID>(
                                    ridRevisionManifestStart4FND[i],
                                    ridRevisionManifestStart4FND[j],
                                    51401,
                                    @"[In RevisionManifestStart4FND] The rid of two RevisionManifestStart4FND in revision manifest list is different");
                }

                ExtendedGUID zeroExtendGuid = new ExtendedGUID();
                zeroExtendGuid.Guid = Guid.Empty;
                zeroExtendGuid.N = 0;

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R516
                Site.CaptureRequirementIfAreNotEqual<ExtendedGUID>(
                    zeroExtendGuid,           
                    ridRevisionManifestStart4FND[i],
                                516,
                                @"[In RevisionManifestStart4FND] [ridDependent] If the value is ""{ { 00000000 - 0000 - 0000 - 0000 - 000000000000}, 0}
                    "", then this revision manifest has no dependency revision. ");

                if (!ridDependentRevisionManifestStart4FND[i].Equals(zeroExtendGuid))
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R517
                    Site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                                    ridRevisionManifestStart4FND[i-1],
                                    ridDependentRevisionManifestStart4FND[i],
                                    517,
                                    @"[In RevisionManifestStart4FND]  [ridDependent] Otherwise [If the value is not ""{ { 00000000 - 0000 - 0000 - 0000 - 000000000000}, 0}
                ""], this value MUST be equal to the RevisionManifestStart4FND.rid field of a previous revision manifest within this (section 2.1.9) revision manifest list.");
                }
            }
        }

        /// <summary>
        /// The test case is validate that the requirements related with .onetoc2 file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S02_TC03_VerifyguildFile()
        {
            string fileName1 = Common.GetConfigurationPropertyValue("OnetocFileLocal", Site);

            OneNoteRevisionStoreFile file1 = this.Adapter.LoadOneNoteFile(fileName1);

            string fileName2 = Common.GetConfigurationPropertyValue("OneFileWithFileData", Site);

            OneNoteRevisionStoreFile file2 = this.Adapter.LoadOneNoteFile(fileName2);
            Site.CaptureRequirementIfIsTrue(
                file1.Header.guidFile != file2.Header.guidFile,
                1421,
                @"[In Header] [guildFile]: The guidFile in two files is different.");


            string fileName3 = Common.GetConfigurationPropertyValue("NoSectionFile", Site);

            OneNoteRevisionStoreFile file3 = this.Adapter.LoadOneNoteFile(fileName3);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R263
            Site.CaptureRequirementIfAreEqual<Guid>(
                    System.Guid.Parse("{00000000-0000-0000-0000-000000000000}"),
                    file3.Header.guidAncestor,
                    263,
                    @"[In Header] If the GUID is ""{00000000-0000-0000-0000-000000000000}"", this field does not reference a table of contents file.");


        }

        /// <summary>
        /// The test case is validate that the requirements related with encryption .one file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S02_TC04_LoadEncryptionFile()
        {
            string fileName = Common.GetConfigurationPropertyValue("OneFileEncryption", Site);

            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(fileName);

            foreach(ObjectSpaceManifestList objectSpaceManifestList in file.RootFileNodeList.ObjectSpaceManifestList)
            {
                foreach(RevisionManifestList revisionManifestList in objectSpaceManifestList.RevisionManifestList)
                {
                    foreach(RevisionManifest revisionManifest in revisionManifestList.RevisionManifests)
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R104
                        this.Site.CaptureRequirementIfIsTrue(
                            revisionManifest.FileNodeSequence[1].FileNodeID == FileNodeIDValues.ObjectDataEncryptionKeyV2FNDX,
                            104,
                            @"[In Revision Manifest] If the object space is encrypted, then the second FileNode in the sequence MUST be a FileNode structure with a FileNodeID equal to 0x07C (ObjectDataEncryptionKeyV2FNDX structure, section 2.5.19).");

                        // If R104 is verified,then R612 will be verified.
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R612
                        this.Site.CaptureRequirement(
                            612,
                            @"[In ObjectDataEncryptionKeyV2FNDX] If any revision manifest (section 2.1.9) for an object space contains this FileNode structure, all other revision manifests for this object space MUST contain this FileNode structure, and these FileNode structures MUST point to structures with identical encryption data.");
                    }
                }
            }
        }
        /// <summary>
        /// The test case is validate that the requirements related with the .one file that have many large references.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void S02_TC05_LoadOneNoteWithLargeReferences()
        {
            string fileName = Common.GetConfigurationPropertyValue("OneWithLarge", Site);

            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(fileName);

            int objectSpaceCount = file.RootFileNodeList.ObjectSpaceManifestList.Count;

            for (int i = 0; i < file.RootFileNodeList.ObjectSpaceManifestList.Count; i++)
            {
                ObjectSpaceManifestList objectSpace = file.RootFileNodeList.ObjectSpaceManifestList[i];

                for (int j = 0; j < objectSpace.RevisionManifestList[0].ObjectGroupList.Count; j++)
                {
                    ObjectGroupList objectGroupList = objectSpace.RevisionManifestList[0].ObjectGroupList[j];
                    FileNode[] fileNodes = objectGroupList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3LargeRefCountFND).ToArray();

                    foreach (FileNode node in fileNodes)
                    {
                        ObjectDeclarationFileData3LargeRefCountFND fnd = node.fnd as ObjectDeclarationFileData3LargeRefCountFND;

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R72
                        Site.CaptureRequirementIfIsTrue(
                                (uint)node.FileNodeID == 0x073 && fnd.jcid.IsFileData == 1,
                                72,
                                @"[In Object Space Object]If the value of the JCID.IsFileData field is ""true"" then the value of the FileNode.FileNodeID field MUST be [0x072 (ObjectDeclarationFileData3RefCountFND structure, section 2.5.27) or] 0x073 (ObjectDeclarationFileData3LargeRefCountFND structure, section 2.5.28).");
                    }
                }
            }
        }
        #endregion
    }
}
