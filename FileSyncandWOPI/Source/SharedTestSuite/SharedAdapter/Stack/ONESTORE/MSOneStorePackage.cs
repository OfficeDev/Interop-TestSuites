namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using MS_ONESTORE;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// The class s used to represent a package of a revision-based file.
    /// </summary>
    public class MSOneStorePackage
    {
        public MSOneStorePackage()
        {
            this.RevisionManifests = new List<RevisionManifestDataElementData>();
            this.CellManifests = new List<CellManifestDataElementData>();
            this.OtherFileNodeList = new List<RevisionStoreObjectGroup>();
        }
        /// <summary>
        /// Gets or sets the Storage Index.
        /// </summary>
        public StorageIndexDataElementData StorageIndex { get; set; }
        /// <summary>
        /// Gets or sets the Storage Manifest.
        /// </summary>
        public StorageManifestDataElementData StorageManifest { get; set; }
        /// <summary>
        /// Gets or sets the Cell Manifest of Header Cell.
        /// </summary>
        public CellManifestDataElementData HeaderCellCellManifest { get; set; }
        /// <summary>
        /// Gets or sets the Revision Manifest of Header Cell.
        /// </summary>
        public RevisionManifestDataElementData HeaderCellRevisionManifest { get; set; }
        /// <summary>
        /// Gets or sets the Revision Manifests.
        /// </summary>
        public List<RevisionManifestDataElementData> RevisionManifests { get; set; }
        /// <summary>
        /// Gets or sets the Cell Manifests.
        /// </summary>
        public List<CellManifestDataElementData> CellManifests { get; set; }
        /// <summary>
        /// Gets or sets the Header Cell.
        /// </summary>
        public HeaderCell HeaderCell { get; set; }
        /// <summary>
        /// Gets or sets the root objects of the revision store file.
        /// </summary>
        public List<RevisionStoreObjectGroup> DataRoot { get; set; }
        /// <summary>
        /// Gets or sets the other objects of the revision store file.
        /// </summary>
        public List<RevisionStoreObjectGroup> OtherFileNodeList { get; set; }

        /// <summary>
        /// This method is used to find the Storage Index Cell Mapping matches the Cell ID.
        /// </summary>
        /// <param name="cellID">Specify the Cell ID.</param>
        /// <returns>Return the specific Storage Index Cell Mapping.</returns>
        public StorageIndexCellMapping FindStorageIndexCellMapping(CellID cellID)
        {
            StorageIndexCellMapping storageIndexCellMapping = null;
            if (this.StorageIndex != null)
            {
                storageIndexCellMapping = this.StorageIndex.StorageIndexCellMappingList.Where(s => s.CellID.Equals(cellID)).SingleOrDefault();
            }

            return storageIndexCellMapping;
        }
        /// <summary>
        /// This method is used to find the Storage Index Revision Mapping that matches the Revision Mapping Extended GUID.
        /// </summary>
        /// <param name="revisionExtendedGUID">Specify the Revision Mapping Extended GUID.</param>
        /// <returns>Return the instance of Storage Index Revision Mapping.</returns>
        public StorageIndexRevisionMapping FindStorageIndexRevisionMapping(ExGuid revisionExtendedGUID)
        {
            StorageIndexRevisionMapping instance = null;
            if(this.StorageIndex!=null)
            {
                instance = this.StorageIndex.StorageIndexRevisionMappingList.Where(r => r.RevisionExtendedGUID.Equals(revisionExtendedGUID)).SingleOrDefault();
            }

            return instance;
        }
    }
    /// <summary>
    /// The Class is used to represent the root object
    /// </summary>
    public class RootObject
    {
        /// <summary>
        /// Gets or sets the Object Declaration of root object 
        /// </summary>
        public ObjectGroupObjectDeclare ObjectDeclaration { get; set; }
        /// <summary>
        /// Gets or sets the data of Object Data in root object.
        /// </summary>
        public ObjectSpaceObjectPropSet ObjectData { get; set; }
    }
    /// <summary>
    ///  The Class is used to represent the Header Cell.
    /// </summary>
    public class HeaderCell
    {
        /// <summary>
        /// Gets or sets the Object Declaration of root object 
        /// </summary>
        public ObjectGroupObjectDeclare ObjectDeclaration { get; set; }
        /// <summary>
        /// Gets or sets the data of Object Data in root object.
        /// </summary>
        public ObjectSpaceObjectPropSet ObjectData { get; set; }
        /// <summary>
        /// Create the instacne of Header Cell.
        /// </summary>
        /// <param name="objectElement">The instance of ObjectGroupDataElementData.</param>
        /// <returns>Returns the instacne of HeaderCell.</returns>
        public static HeaderCell CreateInstance(ObjectGroupDataElementData objectElement)
        {
            HeaderCell instance = new HeaderCell();

            for (int i = 0; i < objectElement.ObjectGroupDeclarations.ObjectDeclarationList.Count; i++)
            {
                if (objectElement.ObjectGroupDeclarations.ObjectDeclarationList[i].ObjectPartitionID != null && objectElement.ObjectGroupDeclarations.ObjectDeclarationList[i].ObjectPartitionID.DecodedValue == 1)
                {
                    instance.ObjectDeclaration = objectElement.ObjectGroupDeclarations.ObjectDeclarationList[0];
                    ObjectGroupObjectData objectData = objectElement.ObjectGroupData.ObjectGroupObjectDataList[0];
                    instance.ObjectData = new ObjectSpaceObjectPropSet();
                    instance.ObjectData.DoDeserializeFromByteArray(objectData.Data.Content.ToArray(), 0);
                    break;
                }
            }

            return instance;
        }
    }

    /// <summary>
    /// This class is used to represent the JCID object.
    /// </summary>
    public class JCIDObject
    {
        /// <summary>
        /// Construct the JCIDObject instance.
        /// </summary>
        /// <param name="objectDeclaration">The Object Declaration structure.</param>
        /// <param name="objectData">The Object Data structure.</param>
        public JCIDObject(ObjectGroupObjectDeclare objectDeclaration, ObjectGroupObjectData objectData)
        {
            this.ObjectDeclaration = objectDeclaration;
            this.JCID = new JCID();
            this.JCID.DoDeserializeFromByteArray(objectData.Data.Content.ToArray(), 0);
        }
        /// <summary>
        /// Gets or sets the value of Object Declaration.
        /// </summary>
        public ObjectGroupObjectDeclare ObjectDeclaration { get; set; }
        /// <summary>
        /// Gets or sets the data of object data.
        /// </summary>
        public JCID JCID { get; set; }
    }
    /// <summary>
    /// This class is used to represent the property set.
    /// </summary>
    public class PropertySetObject
    {
        /// <summary>
        /// Construct the PropertySetObject instance.
        /// </summary>
        /// <param name="objectDeclaration">The Object Declaration structure.</param>
        /// <param name="objectData">The Object Data structure.</param>
        public PropertySetObject(ObjectGroupObjectDeclare objectDeclaration, ObjectGroupObjectData objectData)
        {
            this.ObjectDeclaration = objectDeclaration;
            this.ObjectSpaceObjectPropSet = new ObjectSpaceObjectPropSet();
            this.ObjectSpaceObjectPropSet.DoDeserializeFromByteArray(objectData.Data.Content.ToArray(), 0);
        }
        /// <summary>
        /// Gets or sets the value of Object Declaration.
        /// </summary>
        public ObjectGroupObjectDeclare ObjectDeclaration { get; set; }
        /// <summary>
        /// Gets or sets the data of object data.
        /// </summary>
        public ObjectSpaceObjectPropSet ObjectSpaceObjectPropSet { get; set; }
    }

    /// <summary>
    /// This class is used to represent the file data.
    /// </summary>
    public class FileDataObject
    {
        /// <summary>
        /// Gets or sets the value of Object Data BLOB Declaration.
        /// </summary>
        public ObjectGroupObjectBLOBDataDeclaration ObjectDataBLOBDeclaration { get; set; }
        /// <summary>
        /// Gets or sets the value of Object Data BLOB Reference.
        /// </summary>
        public ObjectGroupObjectDataBLOBReference ObjectDataBLOBReference { get; set; }

        /// <summary>
        /// Gets or sets the data of file data object.
        /// </summary>
        public DataElement ObjectDataBLOBDataElement { get; set; }
    }
    /// <summary>
    /// This class is used to represent the Object Group
    /// </summary>
    public class RevisionStoreObjectGroup
    {
        public RevisionStoreObjectGroup(ExGuid objectGroupId)
        {
            this.Objects = new List<RevisionStoreObject>();
            this.EncryptionObjects = new List<EncryptionObject>();
            this.ObjectGroupID = objectGroupId;
        }
        /// <summary>
        /// Gets or sets the revision store object group identifier.
        /// </summary>
        public ExGuid ObjectGroupID { get; set; }
        /// <summary>
        /// Gets or sets the Objects in object group.
        /// </summary>
        public List<RevisionStoreObject> Objects { get; set; }
        /// <summary>
        /// Gets or sets the encryption objects.
        /// </summary>
        public List<EncryptionObject> EncryptionObjects { get; set; }

        public static RevisionStoreObjectGroup CreateInstance(ExGuid objectGroupId, ObjectGroupDataElementData dataObject, bool isEncryption)
        {
            RevisionStoreObjectGroup objectGroup = new RevisionStoreObjectGroup(objectGroupId);
            Dictionary<ExGuid, RevisionStoreObject> objectDict = new Dictionary<ExGuid, RevisionStoreObject>();
            if (isEncryption == false)
            {
                RevisionStoreObject revisionObject = null;
                for (int i = 0; i < dataObject.ObjectGroupDeclarations.ObjectDeclarationList.Count; i++)
                {
                    ObjectGroupObjectDeclare objectDeclaration = dataObject.ObjectGroupDeclarations.ObjectDeclarationList[i];
                    ObjectGroupObjectData objectData = dataObject.ObjectGroupData.ObjectGroupObjectDataList[i];

                    if (!objectDict.ContainsKey(objectDeclaration.ObjectExtendedGUID))
                    {
                        revisionObject = new RevisionStoreObject();
                        revisionObject.ObjectGroupID = objectGroupId;
                        revisionObject.ObjectID = objectDeclaration.ObjectExtendedGUID;
                        objectDict.Add(objectDeclaration.ObjectExtendedGUID, revisionObject);
                    }
                    else
                    {
                        revisionObject = objectDict[objectDeclaration.ObjectExtendedGUID];
                    }
                    if (objectDeclaration.ObjectPartitionID.DecodedValue == 4)
                    {
                        revisionObject.JCID = new JCIDObject(objectDeclaration, objectData);
                    }
                    else if (objectDeclaration.ObjectPartitionID.DecodedValue == 1)
                    {
                        revisionObject.PropertySet = new PropertySetObject(objectDeclaration, objectData);
                        if (Convert.ToBoolean(revisionObject.JCID.JCID.IsFileData) == false)
                        {
                            revisionObject.ReferencedObjectID = objectData.ObjectExGUIDArray;
                            revisionObject.ReferencedObjectSpacesID = objectData.CellIDArray;
                        }
                    }
                }

                for (int i = 0; i < dataObject.ObjectGroupDeclarations.ObjectGroupObjectBLOBDataDeclarationList.Count; i++)
                {
                    ObjectGroupObjectBLOBDataDeclaration objectGroupObjectBLOBDataDeclaration = dataObject.ObjectGroupDeclarations.ObjectGroupObjectBLOBDataDeclarationList[i];
                    ObjectGroupObjectDataBLOBReference objectGroupObjectDataBLOBReference = dataObject.ObjectGroupData.ObjectGroupObjectDataBLOBReferenceList[i];
                    if (!objectDict.ContainsKey(objectGroupObjectBLOBDataDeclaration.ObjectExGUID))
                    {
                        revisionObject = new RevisionStoreObject();
                        objectDict.Add(objectGroupObjectBLOBDataDeclaration.ObjectExGUID, revisionObject);
                    }
                    else
                    {
                        revisionObject = objectDict[objectGroupObjectBLOBDataDeclaration.ObjectExGUID];
                    }
                    if (objectGroupObjectBLOBDataDeclaration.ObjectPartitionID.DecodedValue == 2)
                    {
                        revisionObject.FileDataObject = new FileDataObject();
                        revisionObject.FileDataObject.ObjectDataBLOBDeclaration = objectGroupObjectBLOBDataDeclaration;
                        revisionObject.FileDataObject.ObjectDataBLOBReference = objectGroupObjectDataBLOBReference;
                    }
                }
                objectGroup.Objects.AddRange(objectDict.Values.ToArray());
            }
            else
            {
                for (int i = 0; i < dataObject.ObjectGroupDeclarations.ObjectDeclarationList.Count; i++)
                {
                    ObjectGroupObjectDeclare objectDeclaration = dataObject.ObjectGroupDeclarations.ObjectDeclarationList[i];
                    ObjectGroupObjectData objectData = dataObject.ObjectGroupData.ObjectGroupObjectDataList[i];

                    if(objectDeclaration.ObjectPartitionID.DecodedValue==1)
                    {
                        EncryptionObject encrypObject = new EncryptionObject();
                        encrypObject.ObjectDeclaration = objectDeclaration;
                        encrypObject.ObjectData = objectData.Data.Content.ToArray();
                        objectGroup.EncryptionObjects.Add(encrypObject);
                    }
                }
            }

            return objectGroup;
        }
    }
    /// <summary>
    /// The class is used to represent the revision store object.
    /// </summary>
    public class RevisionStoreObject
    {
        /// <summary>
        ///  Initialize the class.
        /// </summary>
        public RevisionStoreObject()
        {
           
        }
        /// <summary>
        /// Gets or sets the object identifier.
        /// </summary>
        public ExGuid ObjectID { get; set; }
        /// <summary>
        /// Gets or sets the object group identifier.
        /// </summary>
        public ExGuid ObjectGroupID { get; set; }
        /// <summary>
        /// Gets or sets the value of Object Declaration.
        /// </summary>
        public JCIDObject JCID { get; set; }
        /// <summary>
        /// Gets or sets the value of PropertySet.
        /// </summary>
        public PropertySetObject PropertySet { get; set; }
        /// <summary>
        /// Gets or sets the value of FileDataObject.
        /// </summary>
        public FileDataObject FileDataObject { get; set; }
        /// <summary>
        /// Gets or sets the identifiers of the referenced objects in the revision store.
        /// </summary>
        public ExGUIDArray ReferencedObjectID { get; set; }
        /// <summary>
        /// Gets or sets the identifiers of the referenced object spaces in the revision store.
        /// </summary>
        public CellIDArray ReferencedObjectSpacesID { get; set; }
    }
    /// <summary>
    /// The class is used to represent the encryption revision store object.
    /// </summary>
    public class EncryptionObject
    {
        /// <summary>
        /// Gets or sets the value of Object Declaration.
        /// </summary>
        public ObjectGroupObjectDeclare ObjectDeclaration { get; set; }
        /// <summary>
        /// Gets or sets the data of object.
        /// </summary>
        public byte[] ObjectData { get; set; }
    }
}
