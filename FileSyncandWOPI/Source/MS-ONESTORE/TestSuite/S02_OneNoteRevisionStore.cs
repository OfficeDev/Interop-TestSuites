namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
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

            FileChunkReference64x32 nextFragmentRef = file.RootFileNodeList.FileNodeListFragments[file.RootFileNodeList.FileNodeListFragments.Count - 1].nextFragment;

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R379
            Site.CaptureRequirementIfIsTrue(
                nextFragmentRef.IsfcrNil(),
                379,
                @"[In FileNodeListFragment] If this is the last fragment, the value of the nextFragment field MUST be ""fcrNil"" (see section 2.2.4). ");

            foreach (ObjectSpaceManifestList objectSpaceManifestList in file.RootFileNodeList.ObjectSpaceManifestList)
            {
                foreach (RevisionManifestList revisionManifestList in objectSpaceManifestList.RevisionManifestList)
                {
                    for (int i = 0; i <= revisionManifestList.FileNodeListFragments.Count - 2; i++)
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R366
                        Site.CaptureRequirementIfAreEqual<uint>(
                            revisionManifestList.FileNodeListFragments[i].Header.FileNodeListID,
                            revisionManifestList.FileNodeListFragments[i + 1].Header.FileNodeListID,
                            366,
                            @"[In FileNodeListFragment] All fragments in the same file node list MUST have the same FileNodeListFragment.header.FileNodeListID field.");
                    }

                    foreach (FileNodeListFragment fileNodeListFragment in revisionManifestList.FileNodeListFragments)
                    {
                        FileChunkReference64x32 nextFragment = fileNodeListFragment.nextFragment;
                        if ((uint)fileNodeListFragment.rgFileNodes[fileNodeListFragment.rgFileNodes.Count - 1].FileNodeID == 0xFF)
                        {
                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R373
                            Site.CaptureRequirementIfIsFalse(
                                nextFragment.IsfcrNil(),
                                373,
                                @"[In FileNodeListFragment] [rgFileNodes If a ChunkTerminatorFND structure is present, the value of the nextFragment field MUST be a valid FileChunkReference64x32 structure (section 2.2.4.4) to the next FileNodeListFragment structure.");
                        }
                    }

                    //Verify MS-ONESTORE requirement: MS-ONESTORE_R372
                    Site.CaptureRequirementIfAreEqual<uint>(
                    (uint)revisionManifestList.FileNodeListFragments[0].rgFileNodes[revisionManifestList.FileNodeListFragments[0].rgFileNodes.Count - 1].FileNodeID,
                    0x0FF,
                    372,
                    @"[In FileNodeListFragment]  [rgFileNodes] [The stream is terminated when any of the following conditions is met:] A FileNode structure with a FileNodeID field value equal to 0x0FF (ChunkTerminatorFND structure, section 2.4.3) is read.");

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
                    }
                    fileNodes = objectGroupList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3RefCountFND).ToArray();
                    foreach(FileNode node in fileNodes)
                    {
                        ObjectDeclarationFileData3RefCountFND fnd = node.fnd as ObjectDeclarationFileData3RefCountFND;

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R71
                        Site.CaptureRequirementIfIsTrue(
                                (uint)node.FileNodeID== 0x072 && fnd.jcid.IsFileData==1,
                                71,
                                @"[In Object Space Object]If the value of the JCID.IsFileData field is ""true"" then the value of the FileNode.FileNodeID field MUST be 0x072 (ObjectDeclarationFileData3RefCountFND structure, section 2.5.27) [or 0x073 (ObjectDeclarationFileData3LargeRefCountFND structure, section 2.5.28). ]");
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

        }
        #endregion
    }
}
