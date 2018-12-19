namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using MS_ONESTORE;
    using System.Collections.Generic;

    /// <summary>
    /// The class s used to represent a package of a revision-based file.
    /// </summary>
    public class MSOneStorePackage
    {
        public DataElement StorageIndex { get; set; }

        public StorageManifestDataElementData StorageManifest { get; set; }

        public List<RevisionManifestDataElementData> RevisionManifests { get; set; }

        public List<CellManifestDataElementData> HeaderCells { get; set; }
        /// <summary>
        /// Gets or sets the root object of the revision store file.
        /// </summary>
        public RootObject Root { get; set; }
        /// <summary>
        /// Gets or sets the revision store objects.
        /// </summary>
        public List<RevisionStoreObject> RevisionStoreObjects { get; set; }
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
    /// This class is used to represent the JCID object.
    /// </summary>
    public class JCIDObject
    {
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
    /// The class is used to represent the revision store object.
    /// </summary>
    public class RevisionStoreObject
    {
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
    }
}
