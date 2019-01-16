namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using System.Collections.Generic;
    using System;
    using System.Linq;

    /// <summary>
    /// This scenario is designed to test the requirements related with transmission by Using the File Synchronization via SOAP Over HTTP Protocol.
    /// </summary>
    [TestClass]
    public class S01_TransmissionByFSSHTTP : TestSuiteBase
    {
        #region Test Case Initialization
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

        #endregion

        #region Test Cases
        /// <summary>
        /// The test case is validate that call QueryChange to get the specific OneNote file that contains file data.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S01_TC01_QueryOneFileContainsFileData()
        {
            // Get the resource url that contains the file data.
            string resourceName = Common.GetConfigurationPropertyValue("OneFileWithFileData", Site);
            string url = this.GetResourceUrl(resourceName);
            this.InitializeContext(url, this.UserName, this.Password, this.Domain);

            // Call QueryChange to get the data that is uploaded by above step.
            CellSubRequestType cellSubRequest = this.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentSerialNumber());
            CellStorageResponse cellStorageResponse = this.SharedAdapter.CellStorageRequest(url, new SubRequestType[] { cellSubRequest });
            MSOneStorePackage package = this.ConvertOneStorePackage(cellStorageResponse);

            // Call adapter to load same file in local.
            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(resourceName);

            #region Capture code for Revisions
            ExGuid rootObjectId = package.DataRoot[0].ObjectGroupID;
            RevisionManifestDataElementData rootRevision = package.RevisionManifests[0];

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R933 
            Site.CaptureRequirementIfAreEqual<ExGuid>(
                    rootObjectId,
                    rootRevision.RevisionManifestObjectGroupReferencesList[0].ObjectGroupExtendedGUID,
                    933,
                    @"[In Revisions] The Object Extended GUID field in the Revision Manifest Data Element structure MUST be equal to the identity of the corresponding root object in the revision in the revision store.");
            #endregion

            #region Capture code for Objects
            List<RevisionStoreObject> objectsWithFileData = new List<RevisionStoreObject>();
            foreach (RevisionStoreObjectGroup objGroup in package.DataRoot)
            {
                objectsWithFileData.AddRange(objGroup.Objects.Where(o => o.FileDataObject != null).ToArray());
            }
            foreach (RevisionStoreObjectGroup objGroup in package.OtherFileNodeList)
            {
                objectsWithFileData.AddRange(objGroup.Objects.Where(o => o.FileDataObject != null).ToArray());
            }
            string subResponseBase64 = cellStorageResponse.ResponseCollection.Response[0].SubResponse[0].SubResponseData.Text[0];
            byte[] subResponseBinary = Convert.FromBase64String(subResponseBase64);
            FsshttpbResponse fsshttpbResponse = FsshttpbResponse.DeserializeResponseFromByteArray(subResponseBinary, 0);
            DataElement[] objectBlOBElements = fsshttpbResponse.DataElementPackage.DataElements.Where(d => d.DataElementType == DataElementType.ObjectDataBLOBDataElementData).ToArray();

            foreach (RevisionStoreObject obj in objectsWithFileData)
            {
                Guid fileDataObjectGuid = this.GetFileDataObjectGUID(obj);
                string extension = this.GetFileDataObjectExtension(obj);
                bool isFoundBLOB = 
                    objectBlOBElements.Where(b => b.DataElementExtendedGUID.Equals(obj.FileDataObject.ObjectDataBLOBReference.BLOBExtendedGUID)).ToArray().Length > 0;

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R948
                Site.CaptureRequirementIfIsTrue(
                        isFoundBLOB && obj.FileDataObject.ObjectDataBLOBDeclaration.ObjectPartitionID.DecodedValue==2,
                        948,
                        @"[In Objects] Object Data BLOB Declaration.PartitionID: 2 (File Data) and Object Data BLOB Reference. BLOB Extended GUID: MUST have a reference to an Object Data BLOB Data Element structure, as specified in [MS-FSSHTTPB] section 2.2.1.12.8, used to transmit the data of the file data object.");

                ObjectGroupList objectGropListByLocal = this.FindObjectGroup(file, obj.ObjectGroupID);

                if (objectGropListByLocal != null)
                {
                    FileNode[] fileData3Refs = objectGropListByLocal.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3RefCountFND).ToArray();

                    foreach(FileNode fn in fileData3Refs)
                    {
                        ObjectDeclarationFileData3RefCountFND fnd = fn.fnd as ObjectDeclarationFileData3RefCountFND;
                        if (fnd.FileDataReference.StringData.Contains(fileDataObjectGuid.ToString()))
                        {
                            // Verify MS-ONESTORE requirement: MS-ONESTORE_R951
                            Site.CaptureRequirementIfIsTrue(
                                    fnd.FileDataReference.StringData.StartsWith("<invfdo>"),
                                    951,
                                    @"[In Objects] This property MUST be set only if the prefix specified by the ObjectDeclarationFileData3RefCountFND.FileDataReference field (section 2.5.27) [or ObjectDeclarationFileData3LargeRefCountFND.FileDataReference field (section 2.5.28)] is not <invfdo>.");


                            Site.CaptureRequirementIfAreEqual<string>(
                                    fnd.Extension.StringData,
                                    extension,
                                    958,
                                    @"[In Objects] MUST be the value specified by the ObjectDeclarationFileData3RefCountFND.Extension field [or the ObjectDeclarationFileData3LargeRefCountFND.Extension] field.");
                            break;
                        }
                    }
                }
            }
            #endregion
        }

        /// <summary>
        /// The test case is validate that call QueryChange to get the specific encryption OneNote file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S01_TC02_QueryEncyptionOneFile()
        {
            // Get the resource url that contains the file data.
            string resourceName = Common.GetConfigurationPropertyValue("OneFileEncryption", Site);
            string url = this.GetResourceUrl(resourceName);
            this.InitializeContext(url, this.UserName, this.Password, this.Domain);

            // Call QueryChange to get the data that is uploaded by above step.
            CellSubRequestType cellSubRequest = this.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentSerialNumber());
            CellStorageResponse cellStorageResponse = this.SharedAdapter.CellStorageRequest(url, new SubRequestType[] { cellSubRequest });
            MSOneStorePackage package = this.ConvertOneStorePackage(cellStorageResponse);

            #region Capture code
            RevisionManifestDataElementData rootRevisionManifest = package.RevisionManifests[0];
            bool isFoundEncryptionKeyRoot = false;

            isFoundEncryptionKeyRoot = 
                rootRevisionManifest.RevisionManifestRootDeclareList.Where(r => r.RootExtendedGUID.Value == 0x00000003).ToArray().Length == 1;

            Site.CaptureRequirementIfIsTrue(
                    isFoundEncryptionKeyRoot,
                    932,
                    @"[In Revisions] The root object with RootObjectReference3FND.rootRole value set to 0x00000003 MUST be present only when the file is encrypted. (see section 2.7.7).");
            #endregion
        }

        /// <summary>
        /// The test case is validate that call QueryChange to get the specific encryption OneNote file.
        /// </summary>
        [TestCategory("MSONESTORE"), TestMethod]
        public void MSONESTORE_S01_TC03_QueryOneFileWithoutFileData()
        {
            // Get the resource url that contains the file data.
            string resourceName = Common.GetConfigurationPropertyValue("OneFileWithoutFileData", Site);
            string url = this.GetResourceUrl(resourceName);
            this.InitializeContext(url, this.UserName, this.Password, this.Domain);

            // Call QueryChange to get the data that is uploaded by above step.
            CellSubRequestType cellSubRequest = this.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentSerialNumber());
            CellStorageResponse cellStorageResponse = this.SharedAdapter.CellStorageRequest(url, new SubRequestType[] { cellSubRequest });
            MSOneStorePackage package = this.ConvertOneStorePackage(cellStorageResponse);

            // Call adapter to load same file in local.
            OneNoteRevisionStoreFile file = this.Adapter.LoadOneNoteFile(resourceName);
            #region Capture code for Header Cell
            for (int i = 0; i < package.HeaderCell.ObjectData.Body.RgPrids.Length; i++)
            {
                PropertyID propId = package.HeaderCell.ObjectData.Body.RgPrids[i];
                if (propId.Value == 0x14001D93)
                {
                    FourBytesOfData crcData = package.HeaderCell.ObjectData.Body.RgData[i] as FourBytesOfData;

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R917
                    Site.CaptureRequirementIfAreEqual<uint>(
                            file.Header.crcName,
                            BitConverter.ToUInt32(crcData.Data, 0),
                            917,
                            @"[In Header Cell] FileNameCRC's PropertyID 0x14001D93 with value: MUST be the Header.crcName field.");
                }
                else if (propId.Value == 0x1C001D94)
                {
                    PrtFourBytesOfLengthFollowedByData guidFileData = package.HeaderCell.ObjectData.Body.RgData[i] as PrtFourBytesOfLengthFollowedByData;

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R912
                    Site.CaptureRequirementIfAreEqual<Guid>(
                            file.Header.guidFile,
                            new Guid(guidFileData.Data),
                            912,
                            @"[In Header Cell] [FileIdentityGuid] MUST be the value specified by the Header.guidFile field.");
                }
                else if (propId.Value == 0x1C001D95)
                {
                    PrtFourBytesOfLengthFollowedByData guidAncestorData = package.HeaderCell.ObjectData.Body.RgData[i] as PrtFourBytesOfLengthFollowedByData;

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R914
                    Site.CaptureRequirementIfAreEqual<Guid>(
                            file.Header.guidAncestor,
                            new Guid(guidAncestorData.Data),
                            914,
                            @"[In Header Cell] [FileAncestorIdentityGuid] MUST be the value specified by the Header.guidAncestor field.");
                }
            }
            #endregion

            #region Capture code for Revision
            List<RevisionManifest> revisionManifestList = new List<RevisionManifest>();

            foreach (ObjectSpaceManifestList objSpaceManifestList in file.RootFileNodeList.ObjectSpaceManifestList)
            {
                foreach (RevisionManifestList revManifestList in objSpaceManifestList.RevisionManifestList)
                {
                    revisionManifestList.AddRange(revManifestList.RevisionManifests);
                }
            }

            foreach (RevisionManifestDataElementData revisionManifestData in package.RevisionManifests)
            {
                ExGuid revisionId = revisionManifestData.RevisionManifest.RevisionID;
                ExGuid baseRevisionId = revisionManifestData.RevisionManifest.BaseRevisionID;

                ExtendedGUID rid = null;
                ExtendedGUID ridDependent = null;
                bool isFound = false;
                for (int i = 0; i < revisionManifestList.Count; i++)
                {
                    FileNode revisionStart = revisionManifestList[i].FileNodeSequence[0];
                    if (revisionStart.FileNodeID == FileNodeIDValues.RevisionManifestStart6FND)
                    {
                        rid = ((RevisionManifestStart6FND)revisionStart.fnd).rid;
                        ridDependent = ((RevisionManifestStart6FND)revisionStart.fnd).ridDependent;
                    }
                    else if (revisionStart.FileNodeID == FileNodeIDValues.RevisionManifestStart7FND)
                    {
                        rid = ((RevisionManifestStart7FND)revisionStart.fnd).Base.rid;
                        ridDependent = ((RevisionManifestStart7FND)revisionStart.fnd).Base.ridDependent;
                    }

                    if (rid.Guid == revisionId.GUID && rid.N == revisionId.Value)
                    {
                        isFound = true;
                        break;
                    }
                }

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R925
                Site.CaptureRequirementIfIsTrue(
                        isFound,
                        925,
                        @"[In Revisions] § Revision ID: MUST be equal to the revision store file revision identifier (section 2.1.8).");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R926
                Site.CaptureRequirementIfIsTrue(
                        baseRevisionId.GUID == ridDependent.Guid && baseRevisionId.Value == ridDependent.N,
                        926,
                        @"[In Revisions] § Base Revision ID: MUST be equal to the revision store file dependency revision identifier (section 2.1.9).");

                ExGuid objectGroupId = revisionManifestData.RevisionManifestObjectGroupReferencesList[0].ObjectGroupExtendedGUID;
                bool isFoundObjectGroup = false;

                isFoundObjectGroup = package.DataRoot.Where(o => o.ObjectGroupID == objectGroupId).ToArray().Length > 0 ||
                    package.OtherFileNodeList.Where(o => o.ObjectGroupID == objectGroupId).ToArray().Length > 0;

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R935
                Site.CaptureRequirementIfIsTrue(
                         isFoundObjectGroup,
                         935,
                         @"[In Object Groups] The Revision Manifest Data Element structure, as specified in [MS-FSSHTTPB] section 2.2.1.12.5, that references an object group MUST specify the object group extended GUID to be equal to the revision store object group identifier.");
            }
            #endregion
        }
        #endregion Test Cases

        #region Private methods
        /// <summary>
        /// initialize the shared context based on the specified request file URL, user name, password and domain.
        /// </summary>
        /// <param name="requestFileUrl">Specify the request file URL.</param>
        /// <param name="userName">Specify the user name.</param>
        /// <param name="password">Specify the password.</param>
        /// <param name="domain">Specify the domain.</param>
        private void InitializeContext(string requestFileUrl, string userName, string password, string domain)
        {
            SharedContext context = SharedContext.Current;

            if (string.Equals("HTTP", Common.GetConfigurationPropertyValue("TransportType", this.Site), System.StringComparison.OrdinalIgnoreCase))
            {
                context.TargetUrl = Common.GetConfigurationPropertyValue("HttpTargetServiceUrl", this.Site);
                context.EndpointConfigurationName = Common.GetConfigurationPropertyValue("HttpEndPointName", this.Site);
            }
            else
            {
                context.TargetUrl = Common.GetConfigurationPropertyValue("HttpsTargetServiceUrl", this.Site);
                context.EndpointConfigurationName = Common.GetConfigurationPropertyValue("HttpsEndPointName", this.Site);
            }
            context.Site = this.Site;
            context.OperationType = OperationType.FSSHTTPCellStorageRequest;
            context.UserName = userName;
            context.Password = password;
            context.Domain = domain;
        }

        /// <summary>
        /// A method used to create a CellRequest object and initialize it.
        /// </summary>
        /// <returns>A return value represents the CellRequest object.</returns>
        private FsshttpbCellRequest CreateFsshttpbCellRequest()
        {
            FsshttpbCellRequest cellRequest = new FsshttpbCellRequest();

            // MUST be great or equal to OxFA12994 
            cellRequest.Version = 0xFA12994;

            // MUST be 12 
            cellRequest.ProtocolVersion = 12;

            // MUST be 11 
            cellRequest.MinimumVersion = 11;

            // MUST be 0x9B069439F329CF9C 
            cellRequest.Signature = 0x9B069439F329CF9C;

            // Set the user agent GUID. 
            cellRequest.GUID = FsshttpbCellRequest.UserAgentGuid;

            // Set the value which MUST be 1. 
            cellRequest.RequestHashingSchema = new Compact64bitInt(1u);
            return cellRequest;
        }

        /// <summary>
        /// A method used to create a QueryChanges CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents subRequest id.</param>
        /// <param name="reserved">A parameter that must be set to zero.</param>
        /// <param name="isAllowFragments">A parameter represents that if to allow fragments.</param>
        /// <param name="isExcludeObjectData">A parameter represents if to exclude object data.</param>
        /// <param name="isIncludeFilteredOutDataElementsInKnowledge">A parameter represents if to include the serial numbers of filtered out data elements in the response knowledge.</param>
        /// <param name="reserved1">A parameter represents a 4-bit reserved field that must be set to zero.</param>
        /// <param name="isStorageManifestIncluded">A parameter represents if to include the storage manifest.</param>
        /// <param name="isCellChangesIncluded">A parameter represents if to include the cell changes.</param>
        /// <param name="reserved2">A parameter represents a 6-bit reserved field that must be set to zero.</param>
        /// <param name="cellId">A parameter represents if the Query Changes are scoped to a specific cell. If the Cell ID is 0x0000, no scoping restriction is specified.</param>
        /// <param name="maxDataElements">A parameter represents the maximum data elements to return.</param>
        /// <param name="queryChangesFilterList">A parameter represents how the results of the query will be filtered before it is returned to the client.</param>
        /// <param name="knowledge">A parameter represents what the client knows about a state of a file.</param>
        /// <returns>A return value represents QueryChanges CellSubRequest object.</returns>
        private QueryChangesCellSubRequest BuildFsshttpbQueryChangesSubRequest(
                                ulong subRequestId,
                                int reserved = 0,
                                bool isAllowFragments = false,
                                bool isExcludeObjectData = false,
                                bool isIncludeFilteredOutDataElementsInKnowledge = true,
                                int reserved1 = 0,
                                bool isStorageManifestIncluded = true,
                                bool isCellChangesIncluded = true,
                                int reserved2 = 0,
                                CellID cellId = null,
                                ulong? maxDataElements = null,
                                List<Filter> queryChangesFilterList = null,
                                Knowledge knowledge = null)
        {
            QueryChangesCellSubRequest queryChange = new QueryChangesCellSubRequest(subRequestId);

            queryChange.Reserved = reserved;
            queryChange.AllowFragments = Convert.ToInt32(isAllowFragments);
            queryChange.ExcludeObjectData = Convert.ToInt32(isExcludeObjectData);
            queryChange.IncludeFilteredOutDataElementsInKnowledge = Convert.ToInt32(isIncludeFilteredOutDataElementsInKnowledge);
            queryChange.Reserved1 = reserved1;

            queryChange.IncludeStorageManifest = Convert.ToInt32(isStorageManifestIncluded);
            queryChange.IncludeCellChanges = Convert.ToInt32(isCellChangesIncluded);
            queryChange.Reserved2 = reserved2;

            if (cellId == null)
            {
                cellId = new CellID(new ExGuid(0, Guid.Empty), new ExGuid(0, Guid.Empty));
            }

            queryChange.CellId = cellId;

            if (maxDataElements != null)
            {
                queryChange.MaxDataElements = new Compact64bitInt(maxDataElements.Value);
            }

            queryChange.QueryChangeFilters = queryChangesFilterList;
            queryChange.Knowledge = knowledge;

            return queryChange;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="requestToken">A parameter represents Request token.</param>
        /// <param name="base64Content">A parameter represents serialized subRequest.</param>
        /// <returns>A return value represents CellSubRequest object.</returns>
        private CellSubRequestType CreateCellSubRequest(ulong requestToken, string base64Content)
        {
            return this.CreateCellSubRequest(requestToken, base64Content, Convert.FromBase64String(base64Content).Length);
        }

        /// <summary>
        /// A method used to create a CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="requestToken">A parameter represents Request token.</param>
        /// <param name="base64Content">A parameter represents serialized subRequest.</param>
        /// <param name="binaryDataSize">A parameter represents the number of bytes of data in the SubRequestData element of a cell sub-request.</param>
        /// <returns>A return value represents CellSubRequest object.</returns>
        private CellSubRequestType CreateCellSubRequest(ulong requestToken, string base64Content, long binaryDataSize)
        {
            CellSubRequestType cellRequestType = new CellSubRequestType();
            cellRequestType.SubRequestToken = requestToken.ToString();
            CellSubRequestDataType subRequestData = new CellSubRequestDataType();
            subRequestData.BinaryDataSize = binaryDataSize;
            subRequestData.Text = new string[1];
            subRequestData.Text[0] = base64Content;

            cellRequestType.SubRequestData = subRequestData;

            return cellRequestType;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object for QueryChanges and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents the subRequest identifier.</param>
        /// <returns>A return value represents the CellRequest object for QueryChanges.</returns>
        private CellSubRequestType CreateCellSubRequestEmbeddedQueryChanges(ulong subRequestId)
        {
            FsshttpbCellRequest cellRequest = this.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = this.BuildFsshttpbQueryChangesSubRequest(subRequestId);
            cellRequest.AddSubRequest(queryChange, null);

            CellSubRequestType cellSubRequest = this.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            return cellSubRequest;
        }
        /// <summary>
        /// Parse the structure of revision store file.
        /// </summary>
        /// <param name="cellStorageResponse">the CellStorageResponse message received from the server.</param>
        /// <returns>Returns the revision store file from the server.</returns>
        private MSOneStorePackage ConvertOneStorePackage(CellStorageResponse cellStorageResponse)
        {
            MSOneStorePackage package = null;
            string subResponseBase64 = cellStorageResponse.ResponseCollection.Response[0].SubResponse[0].SubResponseData.Text[0];
            byte[] subResponseBinary = Convert.FromBase64String(subResponseBase64);
            FsshttpbResponse fsshttpbResponse = FsshttpbResponse.DeserializeResponseFromByteArray(subResponseBinary, 0);
            if (fsshttpbResponse.DataElementPackage != null && fsshttpbResponse.DataElementPackage.DataElements != null)
            {
                MSONESTOREParser onenoteParser = new MSONESTOREParser();
                package = onenoteParser.Parse(fsshttpbResponse.DataElementPackage);
            }

            return package;
        }
        /// <summary>
        /// Get the url of resource.
        /// </summary>
        /// <returns>Returns the url of resource.</returns>
        private string GetResourceUrl(string resourceName)
        {
            string transportType = Common.GetConfigurationPropertyValue("TransportType", Site);
            string sut = Common.GetConfigurationPropertyValue("SUTComputerName", Site);
            string site = Common.GetConfigurationPropertyValue("SiteCollectionName", Site);
            string documentLibray = Common.GetConfigurationPropertyValue("MSONESTORELibraryName", Site);

            return string.Format("{0}://{1}/sites/{2}/{3}/{4}", transportType, sut, site, documentLibray, resourceName);
        }
        /// <summary>
        /// Find the specific object group by object group ID.
        /// </summary>
        /// <param name="file">The instance of OneNoteRevisionStoreFile.</param>
        /// <param name="ObjectGroupId">The object group ID.</param>
        /// <returns>Returns the specify object group.</returns>
        private ObjectGroupList FindObjectGroup(OneNoteRevisionStoreFile file, ExGuid ObjectGroupId)
        {
            foreach (ObjectSpaceManifestList objSpaceManifest in file.RootFileNodeList.ObjectSpaceManifestList)
            {
                foreach(RevisionManifestList revManifestList in objSpaceManifest.RevisionManifestList)
                {
                    for (int i = 0; i < revManifestList.RevisionManifests.Count; i++)
                    {
                        RevisionManifest revManifest = revManifestList.RevisionManifests[i];
                        ObjectGroupListReferenceFND objGroupListRef = revManifest.FileNodeSequence[1].fnd as ObjectGroupListReferenceFND;
                        if (objGroupListRef.ObjectGroupID.Guid == ObjectGroupId.GUID && objGroupListRef.ObjectGroupID.N == ObjectGroupId.Value)
                        {
                            return revManifestList.ObjectGroupList[i];
                        }
                    }
                }
            }

            return null;
        }
        /// <summary>
        /// Get the value of FileDataObject_GUID property.
        /// </summary>
        /// <param name="objectData">The instance of object.</param>
        /// <returns>Return the value of FileDataObject_GUID property.</returns>
        private Guid GetFileDataObjectGUID(RevisionStoreObject objectData)
        {
            for (int i = 0; i < objectData.PropertySet.ObjectSpaceObjectPropSet.Body.RgPrids.Length; i++)
            {
                PropertyID propId = objectData.PropertySet.ObjectSpaceObjectPropSet.Body.RgPrids[i];
                if(propId.Value== 0x1C00343E)
                {
                    PrtFourBytesOfLengthFollowedByData data = objectData.PropertySet.ObjectSpaceObjectPropSet.Body.RgData[i] as PrtFourBytesOfLengthFollowedByData;
                    Guid fileDataObjectGUID = new Guid(data.Data);
                  
                    return fileDataObjectGUID;
                }
            }

            return Guid.Empty;
        }

        /// <summary>
        /// Get the value of FileDataObject_Extension property.
        /// </summary>
        /// <param name="objectData">The instance of object.</param>
        /// <returns>Return the value of FileDataObject_Extension property.</returns>
        private string GetFileDataObjectExtension(RevisionStoreObject objectData)
        {
            for (int i = 0; i < objectData.PropertySet.ObjectSpaceObjectPropSet.Body.RgPrids.Length; i++)
            {
                PropertyID propId = objectData.PropertySet.ObjectSpaceObjectPropSet.Body.RgPrids[i];
                if (propId.Value == 0x1C003424)
                {
                    PrtFourBytesOfLengthFollowedByData data = objectData.PropertySet.ObjectSpaceObjectPropSet.Body.RgData[i] as PrtFourBytesOfLengthFollowedByData;
                    string extension = System.Text.Encoding.Unicode.GetString(data.Data);

                    return extension;
                }
            }

            return string.Empty;
        }

        #endregion
    }
}
