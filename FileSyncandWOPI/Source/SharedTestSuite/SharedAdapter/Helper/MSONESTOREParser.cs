namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestSuites.MS_ONESTORE;
    using System.Linq;
    using System.Collections.Generic;
    using System;

    public static class MSONESTOREParser
    {
        public static MSOneStorePackage Parse(DataElementPackage dataElementPackage)
        {
            MSOneStorePackage package = new MSOneStorePackage();

            foreach (DataElement storageIndex in dataElementPackage.DataElements.Where(dataElement => dataElement.DataElementType == DataElementType.StorageIndexDataElementData))
            {
                package.StorageIndex = storageIndex;
            }

            foreach (DataElement storageManifest in dataElementPackage.DataElements.Where(dataElement => dataElement.DataElementType == DataElementType.StorageManifestDataElementData))
            {
                package.StorageManifest = storageManifest.Data as StorageManifestDataElementData;
            }
            DataElement[] objectElements = dataElementPackage.DataElements.Where(dataElement => dataElement.DataElementType == DataElementType.ObjectGroupDataElementData).ToArray();
            List<DataElement> objectBlOBElements = dataElementPackage.DataElements.Where(dataElement => dataElement.DataElementType == DataElementType.ObjectDataBLOBDataElementData).ToList();
            package.HeaderCells = new List<CellManifestDataElementData>();
            foreach (DataElement cellManifestData in dataElementPackage.DataElements.Where(dataElement => dataElement.DataElementType == DataElementType.CellManifestDataElementData))
            {
                package.HeaderCells.Add((CellManifestDataElementData)cellManifestData.Data);
            }
            package.RevisionManifests = new List<RevisionManifestDataElementData>();
            package.RevisionStoreObjects = new List<RevisionStoreObject>();
 
            foreach (DataElement revisionManifestElement in dataElementPackage.DataElements.Where(dataElement => dataElement.DataElementType == DataElementType.RevisionManifestDataElementData))
            {
                RevisionManifestDataElementData revisionManifest = revisionManifestElement.Data as RevisionManifestDataElementData;
                DataElement instance = objectElements.Where(dataElement => dataElement.DataElementExtendedGUID.Equals(
                                         revisionManifest.RevisionManifestObjectGroupReferencesList[0].ObjectGroupExtendedGUID)).Single();
                if (revisionManifest.RevisionManifestRootDeclareList.Count > 0)
                {
                    if (revisionManifest.RevisionManifestRootDeclareList[0].RootExtendedGUID.Equals(new ExGuid(1, Guid.Parse("{4A3717F8-1C14-49E7-9526-81D942DE1741}"))) &&
                       revisionManifest.RevisionManifestRootDeclareList[0].ObjectExtendedGUID.Equals(new ExGuid(1, Guid.Parse("{B4760B1A-FBDF-4AE3-9D08-53219D8A8D21}"))))
                    {
                        package.Root = ParseRootObject((ObjectGroupDataElementData)instance.Data);
                    }
                }
                else
                {
                    package.RevisionStoreObjects.AddRange(ParseObject((ObjectGroupDataElementData)instance.Data, objectBlOBElements));
                }
                package.RevisionManifests.Add(revisionManifest);
            }

            return package;
        }

        private static RootObject ParseRootObject(ObjectGroupDataElementData rootDataElement)
        {
            RootObject root = new RootObject();
            root.ObjectDeclaration = rootDataElement.ObjectGroupDeclarations.ObjectDeclarationList[0];
            ObjectGroupObjectData objectData= rootDataElement.ObjectGroupData.ObjectGroupObjectDataList[0];
            root.ObjectData = new ObjectSpaceObjectPropSet();
            root.ObjectData.DoDeserializeFromByteArray(objectData.Data.Content.ToArray(), 0);

            return root;
        }

        private static RevisionStoreObject[] ParseObject(ObjectGroupDataElementData instance, List<DataElement> objectBlOBElements)
        {
            Dictionary<ExGuid, RevisionStoreObject> revisionStoreObjectDict = new Dictionary<ExGuid, RevisionStoreObject>();
            RevisionStoreObject revisionObject = null;
            for (int i = 0; i < instance.ObjectGroupDeclarations.ObjectDeclarationList.Count; i++)
            {
                ObjectGroupObjectDeclare objectDeclaration = instance.ObjectGroupDeclarations.ObjectDeclarationList[i];
                ObjectGroupObjectData objectData = instance.ObjectGroupData.ObjectGroupObjectDataList[i];
                
                if (!revisionStoreObjectDict.ContainsKey(objectDeclaration.ObjectExtendedGUID))
                {
                    revisionObject = new RevisionStoreObject();
                    revisionStoreObjectDict.Add(objectDeclaration.ObjectExtendedGUID, revisionObject);
                }
                else
                {
                    revisionObject = revisionStoreObjectDict[objectDeclaration.ObjectExtendedGUID];
                }

                if(objectDeclaration.ObjectPartitionID.DecodedValue==4)
                {
                    revisionObject.JCID = new JCIDObject();
                    revisionObject.JCID.ObjectDeclaration = objectDeclaration;
                    revisionObject.JCID.JCID = new JCID();
                    revisionObject.JCID.JCID.DoDeserializeFromByteArray(objectData.Data.Content.ToArray(), 0);
                }
                else if(objectDeclaration.ObjectPartitionID.DecodedValue == 1)
                {
                    revisionObject.PropertySet = new PropertySetObject();
                    revisionObject.PropertySet.ObjectDeclaration = objectDeclaration;
                    revisionObject.PropertySet.ObjectSpaceObjectPropSet = new ObjectSpaceObjectPropSet();
                    revisionObject.PropertySet.ObjectSpaceObjectPropSet.DoDeserializeFromByteArray(objectData.Data.Content.ToArray(), 0);
                }
            }

            for (int i = 0; i < instance.ObjectGroupDeclarations.ObjectGroupObjectBLOBDataDeclarationList.Count; i++)
            {
                ObjectGroupObjectBLOBDataDeclaration objectGroupObjectBLOBDataDeclaration = instance.ObjectGroupDeclarations.ObjectGroupObjectBLOBDataDeclarationList[i];
                ObjectGroupObjectDataBLOBReference objectGroupObjectDataBLOBReference = instance.ObjectGroupData.ObjectGroupObjectDataBLOBReferenceList[i];
                if (!revisionStoreObjectDict.ContainsKey(objectGroupObjectBLOBDataDeclaration.ObjectExGUID))
                {
                    revisionObject = new RevisionStoreObject();
                    revisionStoreObjectDict.Add(objectGroupObjectBLOBDataDeclaration.ObjectExGUID, revisionObject);
                }
                else
                {
                    revisionObject = revisionStoreObjectDict[objectGroupObjectBLOBDataDeclaration.ObjectExGUID];
                }
                if (objectGroupObjectBLOBDataDeclaration.ObjectPartitionID.DecodedValue == 2)
                {
                    revisionObject.FileDataObject = new FileDataObject();
                    revisionObject.FileDataObject.ObjectDataBLOBDeclaration = objectGroupObjectBLOBDataDeclaration;
                    revisionObject.FileDataObject.ObjectDataBLOBReference = objectGroupObjectDataBLOBReference;
                    revisionObject.FileDataObject.ObjectDataBLOBDataElement =
                        objectBlOBElements.Where(element => element.DataElementExtendedGUID.Equals(objectGroupObjectDataBLOBReference.BLOBExtendedGUID)).Single();
                }
            }
            return revisionStoreObjectDict.Values.ToArray();
        }
    }
}
