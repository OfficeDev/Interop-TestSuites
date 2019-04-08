namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestSuites.MS_ONESTORE;
    using System.Linq;
    using System.Collections.Generic;
    using System;
    using System.Collections;

    /// <summary>
    /// This class is used to parse the revision-based file transport by FSSHTTP.
    /// </summary>
    public class MSONESTOREParser
    {
        // The DataElements of Storage Index
        private DataElement[] storageIndexDataElements;
        // The DataElements of Storage Manifest
        private DataElement[] storageManifestDataElements;
        // The DataElements of Cell Manifest
        private DataElement[] cellManifestDataElements;
        // The DataElements of Revision Manifest
        private DataElement[] revisionManifestDataElements;
        // The DataElements of Object Group Data
        private DataElement[] objectGroupDataElements;
        // The DataElements of Object BLOB
        private DataElement[] objectBlOBElements;

        private HashSet<CellID> storageIndexHashTab = new HashSet<CellID>();
        public MSOneStorePackage Parse(DataElementPackage dataElementPackage)
        {
            MSOneStorePackage package = new MSOneStorePackage();

            storageIndexDataElements = dataElementPackage.DataElements.Where(d => d.DataElementType == DataElementType.StorageIndexDataElementData).ToArray();
            storageManifestDataElements = dataElementPackage.DataElements.Where(d => d.DataElementType == DataElementType.StorageManifestDataElementData).ToArray();
            cellManifestDataElements = dataElementPackage.DataElements.Where(d => d.DataElementType == DataElementType.CellManifestDataElementData).ToArray();
            revisionManifestDataElements = dataElementPackage.DataElements.Where(d => d.DataElementType == DataElementType.RevisionManifestDataElementData).ToArray();
            objectGroupDataElements = dataElementPackage.DataElements.Where(d => d.DataElementType == DataElementType.ObjectGroupDataElementData).ToArray();
            objectBlOBElements = dataElementPackage.DataElements.Where(d => d.DataElementType == DataElementType.ObjectDataBLOBDataElementData).ToArray();

            package.StorageIndex = storageIndexDataElements[0].Data as StorageIndexDataElementData;
            package.StorageManifest = storageManifestDataElements[0].Data as StorageManifestDataElementData;

            // Parse Header Cell
            CellID headerCellID= package.StorageManifest.StorageManifestRootDeclareList[0].CellID;
            StorageIndexCellMapping headerCellStorageIndexCellMapping = package.FindStorageIndexCellMapping(headerCellID);
            storageIndexHashTab.Add(headerCellID);
            package.HeaderCellCellManifest = this.FindCellManifest(headerCellStorageIndexCellMapping.CellMappingExtendedGUID);
            StorageIndexRevisionMapping headerCellRevisionManifestMapping =
                package.FindStorageIndexRevisionMapping(package.HeaderCellCellManifest.CellManifestCurrentRevision.CellManifestCurrentRevisionExtendedGUID);
            package.HeaderCellRevisionManifest = this.FindRevisionManifestDataElement(headerCellRevisionManifestMapping.RevisionMappingExtendedGUID);
            package.HeaderCell = this.ParseHeaderCell(package.HeaderCellRevisionManifest);

            // Parse Data root
            CellID dataRootCellID = package.StorageManifest.StorageManifestRootDeclareList[1].CellID;
            storageIndexHashTab.Add(dataRootCellID);
            package.DataRoot = this.ParseObjectGroup(dataRootCellID, package);
            // Parse other data
            foreach(StorageIndexCellMapping storageIndexCellMapping in package.StorageIndex.StorageIndexCellMappingList)
            {
                if (storageIndexHashTab.Contains(storageIndexCellMapping.CellID) == false)
                {
                    package.OtherFileNodeList.AddRange(this.ParseObjectGroup(storageIndexCellMapping.CellID,package));
                    storageIndexHashTab.Add(storageIndexCellMapping.CellID);
                }
            }

            return package;
        }

        /// <summary>
        /// Find the CellManifestDataElementData
        /// </summary>
        /// <param name="cellMappingExtendedGUID">The ExGuid of Cell Mapping Extended GUID.</param>
        /// <returns>Return the CellManifestDataElementData instance.</returns>
        private CellManifestDataElementData FindCellManifest(ExGuid cellMappingExtendedGUID)
        {
            return (CellManifestDataElementData)this.cellManifestDataElements
                .Where(d => d.DataElementExtendedGUID.Equals(cellMappingExtendedGUID)).SingleOrDefault().Data;  
        }
        /// <summary>
        /// Find the Revision Manifest from Data Elements.
        /// </summary>
        /// <param name="revisionMappingExtendedGUID">The Revision Mapping Extended GUID.</param>
        /// <returns>Returns the instance of RevisionManifestDataElementData</returns>
        private RevisionManifestDataElementData FindRevisionManifestDataElement(ExGuid revisionMappingExtendedGUID)
        {
            return (RevisionManifestDataElementData)revisionManifestDataElements
                .Where(d => d.DataElementExtendedGUID.Equals(revisionMappingExtendedGUID)).SingleOrDefault().Data;
        }

        private HeaderCell ParseHeaderCell(RevisionManifestDataElementData headerCellRevisionManifest)
        {
            ExGuid rootObjectId = headerCellRevisionManifest.RevisionManifestObjectGroupReferencesList[0].ObjectGroupExtendedGUID;
            DataElement element = this.objectGroupDataElements
                .Where(d => d.DataElementExtendedGUID.Equals(rootObjectId)).SingleOrDefault();

            return HeaderCell.CreateInstance((ObjectGroupDataElementData)element.Data);
        }

        private List<RevisionStoreObjectGroup> ParseObjectGroup(CellID objectGroupCellID, MSOneStorePackage package)
        {
            StorageIndexCellMapping storageIndexCellMapping = package.FindStorageIndexCellMapping(objectGroupCellID);
            CellManifestDataElementData cellManifest = this.FindCellManifest(storageIndexCellMapping.CellMappingExtendedGUID);
            List<RevisionStoreObjectGroup> objectGroups = new List<RevisionStoreObjectGroup>();
            package.CellManifests.Add(cellManifest);
            StorageIndexRevisionMapping revisionMapping =
                package.FindStorageIndexRevisionMapping(cellManifest.CellManifestCurrentRevision.CellManifestCurrentRevisionExtendedGUID);
            RevisionManifestDataElementData revisionManifest =
                this.FindRevisionManifestDataElement(revisionMapping.RevisionMappingExtendedGUID);
            package.RevisionManifests.Add(revisionManifest);
            RevisionManifestRootDeclare encryptionKeyRoot = revisionManifest.RevisionManifestRootDeclareList.Where(r => r.RootExtendedGUID.Equals(new ExGuid(3, Guid.Parse("4A3717F8-1C14-49E7-9526-81D942DE1741")))).SingleOrDefault();
            bool isEncryption = encryptionKeyRoot != null;
            foreach (RevisionManifestObjectGroupReferences objRef in revisionManifest.RevisionManifestObjectGroupReferencesList)
            {
                ObjectGroupDataElementData dataObject = objectGroupDataElements.Where(d => d.DataElementExtendedGUID.Equals(
                                         objRef.ObjectGroupExtendedGUID)).Single().Data as ObjectGroupDataElementData;

                RevisionStoreObjectGroup objectGroup = RevisionStoreObjectGroup.CreateInstance(objRef.ObjectGroupExtendedGUID, dataObject, isEncryption);
                objectGroups.Add(objectGroup);
            }

            return objectGroups;
        }
    }
}
