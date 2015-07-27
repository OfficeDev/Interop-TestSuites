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

    /// <summary>
    /// Object Group data element 
    /// </summary>
    public class ObjectGroupDataElementData : DataElementData
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupDataElementData class.
        /// </summary>
        public ObjectGroupDataElementData()
        {
            this.ObjectGroupDeclarations = new ObjectGroupDeclarations();

            // The ObjectMetadataDeclaration is only present for MOSS2013, so leave null for default value.
            this.ObjectMetadataDeclaration = null;

            // The DataElementHash is only present for MOSS2013, so leave null for default value.
            this.DataElementHash = null;
            this.ObjectGroupData = new ObjectGroupData();
        }

        /// <summary>
        ///  Gets or sets an optional data element hash for the object data group.
        /// </summary>
        public DataElementHash DataElementHash { get; set; }

        /// <summary>
        /// Gets or sets an optional array of object declarations that specifies the object.
        /// </summary>
        public ObjectGroupDeclarations ObjectGroupDeclarations { get; set; }

        /// <summary>
        /// Gets or sets an object metadata declaration. If no object metadata exists, this field must be omitted.
        /// </summary>
        public ObjectGroupMetadataDeclarations ObjectMetadataDeclaration { get; set; }

        /// <summary>
        /// Gets or sets an object group data.
        /// </summary>
        public ObjectGroupData ObjectGroupData { get; set; }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <returns>A Byte list</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> result = new List<byte>();

            if (this.DataElementHash != null)
            {
                result.AddRange(this.DataElementHash.SerializeToByteList());
            }

            result.AddRange(this.ObjectGroupDeclarations.SerializeToByteList());
            if (this.ObjectMetadataDeclaration != null)
            {
                result.AddRange(this.ObjectMetadataDeclaration.SerializeToByteList());
            }

            result.AddRange(this.ObjectGroupData.SerializeToByteList());
            return result;
        }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="startIndex">Start position</param>
        /// <returns>The length of the element</returns>
        public override int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;

            DataElementHash dataElementHash;
            if (StreamObject.TryGetCurrent<DataElementHash>(byteArray, ref index, out dataElementHash))
            {
                this.DataElementHash = dataElementHash;
            }

            this.ObjectGroupDeclarations = StreamObject.GetCurrent<ObjectGroupDeclarations>(byteArray, ref index);

            ObjectGroupMetadataDeclarations objectMetadataDeclaration = new ObjectGroupMetadataDeclarations();
            if (StreamObject.TryGetCurrent<ObjectGroupMetadataDeclarations>(byteArray, ref index, out objectMetadataDeclaration))
            {
                this.ObjectMetadataDeclaration = objectMetadataDeclaration;
            }

            this.ObjectGroupData = StreamObject.GetCurrent<ObjectGroupData>(byteArray, ref index);

            return index - startIndex;
        }

        /// <summary>
        /// The internal class for build a list of DataElement from an node object.
        /// </summary>
        public class Builder
        {
            /// <summary>
            /// This method is used to build  a list of DataElement from an node object.
            /// </summary>
            /// <param name="node">Specify the node object.</param>
            /// <returns>Return the list of data elements build from the specified node object.</returns>
            public List<DataElement> Build(NodeObject node)
            {
                List<DataElement> dataElements = new List<DataElement>();
                this.TravelNodeObject(node, ref dataElements);
                return dataElements;
            }

            /// <summary>
            /// This method is used to travel the node tree and build the ObjectGroupDataElementData and the extra data element list.
            /// </summary>
            /// <param name="node">Specify the object node.</param>
            /// <param name="dataElements">Specify the list of data elements.</param>
            private void TravelNodeObject(NodeObject node, ref List<DataElement> dataElements)
            {
                if (node is RootNodeObject)
                {
                    ObjectGroupDataElementData data = new ObjectGroupDataElementData();
                    data.ObjectGroupDeclarations.ObjectDeclarationList.Add(this.CreateObjectDeclare(node));
                    data.ObjectGroupData.ObjectGroupObjectDataList.Add(this.CreateObjectData(node as RootNodeObject));

                    dataElements.Add(new DataElement(DataElementType.ObjectGroupDataElementData, data));

                    foreach (IntermediateNodeObject child in (node as RootNodeObject).IntermediateNodeObjectList)
                    {
                        this.TravelNodeObject(child, ref dataElements);
                    }
                }
                else if (node is IntermediateNodeObject)
                {
                    IntermediateNodeObject intermediateNode = node as IntermediateNodeObject;

                    ObjectGroupDataElementData data = new ObjectGroupDataElementData();
                    data.ObjectGroupDeclarations.ObjectDeclarationList.Add(this.CreateObjectDeclare(node));
                    data.ObjectGroupData.ObjectGroupObjectDataList.Add(this.CreateObjectData(intermediateNode));

                    if (intermediateNode.DataNodeObjectData != null)
                    {
                        data.ObjectGroupDeclarations.ObjectDeclarationList.Add(this.CreateObjectDeclare(intermediateNode.DataNodeObjectData));
                        data.ObjectGroupData.ObjectGroupObjectDataList.Add(this.CreateObjectData(intermediateNode.DataNodeObjectData));
                        dataElements.Add(new DataElement(DataElementType.ObjectGroupDataElementData, data));
                        return;
                    }

                    if (intermediateNode.DataNodeObjectData == null && intermediateNode.IntermediateNodeObjectList != null)
                    {
                        dataElements.Add(new DataElement(DataElementType.ObjectGroupDataElementData, data));

                        foreach (IntermediateNodeObject child in intermediateNode.IntermediateNodeObjectList)
                        {
                            this.TravelNodeObject(child, ref dataElements);
                        }

                        return;
                    }
                   
                    throw new System.InvalidOperationException("The DataNodeObjectData and IntermediateNodeObjectList properties in IntermediateNodeObject type cannot be null in the same time.");
                }
            }

            /// <summary>
            /// This method is used to create ObjectGroupObjectDeclare instance from a node object.
            /// </summary>
            /// <param name="node">Specify the node object.</param>
            /// <returns>Return the ObjectGroupObjectDeclare instance.</returns>
            private ObjectGroupObjectDeclare CreateObjectDeclare(NodeObject node)
            {
                ObjectGroupObjectDeclare objectGroupObjectDeclare = new ObjectGroupObjectDeclare();

                objectGroupObjectDeclare.ObjectExtendedGUID = node.ExGuid;
                objectGroupObjectDeclare.ObjectPartitionID = new Compact64bitInt(1u);
                objectGroupObjectDeclare.CellReferencesCount = new Compact64bitInt(0u);
                objectGroupObjectDeclare.ObjectReferencesCount = new Compact64bitInt(0u);
                objectGroupObjectDeclare.ObjectDataSize = new Compact64bitInt((ulong)node.GetContent().Count);

                return objectGroupObjectDeclare;
            }

            /// <summary>
            /// This method is used to create ObjectGroupObjectDeclare instance from a data node object.
            /// </summary>
            /// <param name="node">Specify the node object.</param>
            /// <returns>Return the ObjectGroupObjectDeclare instance.</returns>
            private ObjectGroupObjectDeclare CreateObjectDeclare(DataNodeObjectData node)
            {
                ObjectGroupObjectDeclare objectGroupObjectDeclare = new ObjectGroupObjectDeclare();

                objectGroupObjectDeclare.ObjectExtendedGUID = node.ExGuid;
                objectGroupObjectDeclare.ObjectPartitionID = new Compact64bitInt(1u);
                objectGroupObjectDeclare.CellReferencesCount = new Compact64bitInt(0u);
                objectGroupObjectDeclare.ObjectReferencesCount = new Compact64bitInt(1u);
                objectGroupObjectDeclare.ObjectDataSize = new Compact64bitInt((ulong)node.ObjectData.LongLength);

                return objectGroupObjectDeclare;
            }

            /// <summary>
            /// This method is used to create ObjectGroupObjectData instance from a root node object.
            /// </summary>
            /// <param name="node">Specify the node object.</param>
            /// <returns>Return the ObjectGroupObjectData instance.</returns>
            private ObjectGroupObjectData CreateObjectData(RootNodeObject node)
            {
                ObjectGroupObjectData objectData = new ObjectGroupObjectData();

                objectData.CellIDArray = new CellIDArray(0u, null);

                List<ExGuid> extendedGuidList = new List<ExGuid>();
                foreach (IntermediateNodeObject child in node.IntermediateNodeObjectList)
                {
                    extendedGuidList.Add(child.ExGuid);
                }

                objectData.ObjectExGUIDArray = new ExGUIDArray(extendedGuidList);
                objectData.Data = new BinaryItem(node.SerializeToByteList());

                return objectData;
            }

            /// <summary>
            /// This method is used to create ObjectGroupObjectData instance from a intermediate node object.
            /// </summary>
            /// <param name="node">Specify the node object.</param>
            /// <returns>Return the ObjectGroupObjectData instance.</returns>
            private ObjectGroupObjectData CreateObjectData(IntermediateNodeObject node)
            {
                ObjectGroupObjectData objectData = new ObjectGroupObjectData();

                objectData.CellIDArray = new CellIDArray(0u, null);
                List<ExGuid> extendedGuidList = new List<ExGuid>();

                if (node.DataNodeObjectData != null)
                {
                    extendedGuidList.Add(node.DataNodeObjectData.ExGuid);
                }
                else if (node.IntermediateNodeObjectList != null)
                {
                    foreach (IntermediateNodeObject child in node.IntermediateNodeObjectList)
                    {
                        extendedGuidList.Add(child.ExGuid);
                    }
                }

                objectData.ObjectExGUIDArray = new ExGUIDArray(extendedGuidList);
                objectData.Data = new BinaryItem(node.SerializeToByteList());

                return objectData;
            }

            /// <summary>
            /// This method is used to create ObjectGroupObjectData instance from a data node object.
            /// </summary>
            /// <param name="node">Specify the node object.</param>
            /// <returns>Return the ObjectGroupObjectData instance.</returns>
            private ObjectGroupObjectData CreateObjectData(DataNodeObjectData node)
            {
                ObjectGroupObjectData objectData = new ObjectGroupObjectData();
                objectData.CellIDArray = new CellIDArray(0u, null);
                objectData.ObjectExGUIDArray = new ExGUIDArray(new List<ExGuid>());
                objectData.Data = new BinaryItem(node.ObjectData);
                return objectData;
            }
        }
    }

    /// <summary>
    /// object declaration 
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class ObjectGroupObjectDeclare : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupObjectDeclare class.
        /// </summary>
        public ObjectGroupObjectDeclare()
            : base(StreamObjectTypeHeaderStart.ObjectGroupObjectDeclare)
        {
            this.ObjectExtendedGUID = new ExGuid();
            this.ObjectPartitionID = new Compact64bitInt();
            this.ObjectDataSize = new Compact64bitInt();
            this.ObjectReferencesCount = new Compact64bitInt();
            this.CellReferencesCount = new Compact64bitInt();

            this.ObjectPartitionID.DecodedValue = 1;
            this.ObjectReferencesCount.DecodedValue = 1;
            this.CellReferencesCount.DecodedValue = 0;
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the data element hash.
        /// </summary>
        public ExGuid ObjectExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the partition.
        /// </summary>
        public Compact64bitInt ObjectPartitionID { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the size in bytes of the object.binary data opaque 
        /// to this protocol for the declared object.
        /// This MUST match the size of the binary item in the corresponding object data for this object.
        /// </summary>
        public Compact64bitInt ObjectDataSize { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the number of object references.
        /// </summary>
        public Compact64bitInt ObjectReferencesCount { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the number of cell references.
        /// </summary>
        public Compact64bitInt CellReferencesCount { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            this.ObjectExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.ObjectPartitionID = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.ObjectDataSize = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.ObjectReferencesCount = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.CellReferencesCount = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupObjectDeclare", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of the element</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.ObjectExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.ObjectPartitionID.SerializeToByteList());
            byteList.AddRange(this.ObjectDataSize.SerializeToByteList());
            byteList.AddRange(this.ObjectReferencesCount.SerializeToByteList());
            byteList.AddRange(this.CellReferencesCount.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// object data BLOB declaration 
    /// </summary>
    public class ObjectGroupObjectBLOBDataDeclaration : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupObjectBLOBDataDeclaration class.
        /// </summary>
        public ObjectGroupObjectBLOBDataDeclaration()
            : base(StreamObjectTypeHeaderStart.ObjectGroupObjectBLOBDataDeclaration)
        {
            this.ObjectExGUID = new ExGuid();
            this.ObjectDataBLOBExGUID = new ExGuid();
            this.ObjectPartitionID = new Compact64bitInt();
            this.ObjectDataSize = new Compact64bitInt();
            this.ObjectReferencesCount = new Compact64bitInt();
            this.CellReferencesCount = new Compact64bitInt();
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the object.
        /// </summary>
        public ExGuid ObjectExGUID { get; set; }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the object data BLOB.
        /// </summary>
        public ExGuid ObjectDataBLOBExGUID { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the partition.
        /// </summary>
        public Compact64bitInt ObjectPartitionID { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the size in bytes of the object.opaque binary data  for the declared object. 
        /// This MUST match the size of the binary item in the corresponding object data BLOB referenced by the Object Data BLOB reference for this object.
        /// </summary>
        public Compact64bitInt ObjectDataSize { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the number of object references.
        /// </summary>
        public Compact64bitInt ObjectReferencesCount { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the number of cell references.
        /// </summary>
        public Compact64bitInt CellReferencesCount { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            this.ObjectExGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.ObjectDataBLOBExGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.ObjectPartitionID = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.ObjectReferencesCount = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.CellReferencesCount = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupObjectBLOBDataDeclaration", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of the element</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.ObjectExGUID.SerializeToByteList());
            byteList.AddRange(this.ObjectDataBLOBExGUID.SerializeToByteList());
            byteList.AddRange(this.ObjectPartitionID.SerializeToByteList());
            byteList.AddRange(this.ObjectDataSize.SerializeToByteList());
            byteList.AddRange(this.ObjectReferencesCount.SerializeToByteList());
            byteList.AddRange(this.CellReferencesCount.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// object data 
    /// </summary>
    public partial class ObjectGroupObjectData : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupObjectData class.
        /// </summary>
        public ObjectGroupObjectData()
            : base(StreamObjectTypeHeaderStart.ObjectGroupObjectData)
        {
            this.ObjectExGUIDArray = new ExGUIDArray();
            this.CellIDArray = new CellIDArray();
            this.Data = new BinaryItem();
        }

        /// <summary>
        /// Gets or sets an extended GUID array that specifies the object group.
        /// </summary>
        public ExGUIDArray ObjectExGUIDArray { get; set; }

        /// <summary>
        /// Gets or sets a cell ID array that specifies the object group.
        /// </summary>
        public CellIDArray CellIDArray { get; set; }

        /// <summary>
        /// Gets or sets a byte stream that specifies the binary data which is opaque to this protocol.
        /// </summary>
        public BinaryItem Data { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.ObjectExGUIDArray = BasicObject.Parse<ExGUIDArray>(byteArray, ref index);
            this.CellIDArray = BasicObject.Parse<CellIDArray>(byteArray, ref index);
            this.Data = BasicObject.Parse<BinaryItem>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupObjectData", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of the element</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.ObjectExGUIDArray.SerializeToByteList());
            byteList.AddRange(this.CellIDArray.SerializeToByteList());
            byteList.AddRange(this.Data.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// object data BLOB reference 
    /// </summary>
    public class ObjectGroupObjectDataBLOBReference : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupObjectDataBLOBReference class.
        /// </summary>
        public ObjectGroupObjectDataBLOBReference()
            : base(StreamObjectTypeHeaderStart.ObjectGroupObjectDataBLOBReference)
        {
            this.ObjectExtendedGUIDArray = new ExGUIDArray();
            this.CellIDArray = new CellIDArray();
            this.BLOBExtendedGUID = new ExGuid();
        }

        /// <summary>
        /// Gets or sets an extended GUID array that specifies the object references.
        /// </summary>
        public ExGUIDArray ObjectExtendedGUIDArray { get; set; }

        /// <summary>
        /// Gets or sets a cell ID array that specifies the cell references.
        /// </summary>
        public CellIDArray CellIDArray { get; set; }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the object data BLOB.
        /// </summary>
        public ExGuid BLOBExtendedGUID { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.ObjectExtendedGUIDArray = BasicObject.Parse<ExGUIDArray>(byteArray, ref index);
            this.CellIDArray = BasicObject.Parse<CellIDArray>(byteArray, ref index);
            this.BLOBExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupObjectDataBLOBReference", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of the elements</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.ObjectExtendedGUIDArray.SerializeToByteList());
            byteList.AddRange(CellIDArray.SerializeToByteList());
            byteList.AddRange(this.BLOBExtendedGUID.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// Object Group Declarations
    /// </summary>
    public class ObjectGroupDeclarations : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupDeclarations class.
        /// </summary>
        public ObjectGroupDeclarations()
            : base(StreamObjectTypeHeaderStart.ObjectGroupDeclarations)
        {
            this.ObjectDeclarationList = new List<ObjectGroupObjectDeclare>();
            this.ObjectGroupObjectBLOBDataDeclarationList = new List<ObjectGroupObjectBLOBDataDeclaration>();
        }

        /// <summary>
        /// Gets or sets a list of declarations that specifies the object.
        /// </summary>
        public List<ObjectGroupObjectDeclare> ObjectDeclarationList { get; set; }

        /// <summary>
        /// Gets or sets a list of object data BLOB declarations that specifies the object.
        /// </summary>
        public List<ObjectGroupObjectBLOBDataDeclaration> ObjectGroupObjectBLOBDataDeclarationList { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupDeclarations", "Stream object over-parse error", null);
            }

            int index = currentIndex;
            int headerLength = 0;
            StreamObjectHeaderStart header;
            this.ObjectDeclarationList = new List<ObjectGroupObjectDeclare>();
            this.ObjectGroupObjectBLOBDataDeclarationList = new List<ObjectGroupObjectBLOBDataDeclaration>();
            while ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) != 0)
            {
                if (header.Type == StreamObjectTypeHeaderStart.ObjectGroupObjectDeclare)
                {
                    index += headerLength;
                    this.ObjectDeclarationList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as ObjectGroupObjectDeclare);
                }
                else if (header.Type == StreamObjectTypeHeaderStart.ObjectGroupObjectBLOBDataDeclaration)
                {
                    index += headerLength;
                    this.ObjectGroupObjectBLOBDataDeclarationList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as ObjectGroupObjectBLOBDataDeclaration);
                }
                else
                {
                    throw new StreamObjectParseErrorException(index, "ObjectGroupDeclarations", "Failed to parse ObjectGroupDeclarations, expect the inner object type either ObjectGroupObjectDeclare or ObjectGroupObjectBLOBDataDeclaration, but actual type value is " + header.Type, null);
                }
            }

            currentIndex = index;
        }
        
        /// <summary>
        /// Used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">The Byte list</param>
        /// <returns>A constant value 0</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.ObjectDeclarationList != null)
            {
                foreach (ObjectGroupObjectDeclare objectGroupObjectDeclare in this.ObjectDeclarationList)
                {
                    byteList.AddRange(objectGroupObjectDeclare.SerializeToByteList());
                }
            }

            if (this.ObjectGroupObjectBLOBDataDeclarationList != null)
            {
                foreach (ObjectGroupObjectBLOBDataDeclaration objectGroupObjectBLOBDataDeclaration in this.ObjectGroupObjectBLOBDataDeclarationList)
                {
                    byteList.AddRange(objectGroupObjectBLOBDataDeclaration.SerializeToByteList());
                }
            }

            return 0;
        }
    }

    /// <summary>
    /// Object Data
    /// </summary>
    public class ObjectGroupData : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupData class.
        /// </summary>
        public ObjectGroupData()
            : base(StreamObjectTypeHeaderStart.ObjectGroupData)
        {
            this.ObjectGroupObjectDataList = new List<ObjectGroupObjectData>();
            this.ObjectGroupObjectDataBLOBReferenceList = new List<ObjectGroupObjectDataBLOBReference>();
        }

        /// <summary>
        /// Gets or sets a list of Object Data.
        /// </summary>
        public List<ObjectGroupObjectData> ObjectGroupObjectDataList { get; set; }

        /// <summary>
        /// Gets or sets a list of object data BLOB references that specifies the object.
        /// </summary>
        public List<ObjectGroupObjectDataBLOBReference> ObjectGroupObjectDataBLOBReferenceList { get; set; }

        /// <summary>
        /// Used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>A constant value 0</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.ObjectGroupObjectDataList != null)
            {
                foreach (ObjectGroupObjectData objectGroupObjectData in this.ObjectGroupObjectDataList)
                {
                    byteList.AddRange(objectGroupObjectData.SerializeToByteList());
                }
            }

            if (this.ObjectGroupObjectDataBLOBReferenceList != null)
            {
                foreach (ObjectGroupObjectDataBLOBReference objectGroupObjectDataBLOBReference in this.ObjectGroupObjectDataBLOBReferenceList)
                {
                    byteList.AddRange(objectGroupObjectDataBLOBReference.SerializeToByteList());
                }
            }

            return 0;
        }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupDeclarations", "Stream object over-parse error", null);
            }

            int index = currentIndex;
            int headerLength = 0;
            StreamObjectHeaderStart header;

            this.ObjectGroupObjectDataList = new List<ObjectGroupObjectData>();
            this.ObjectGroupObjectDataBLOBReferenceList = new List<ObjectGroupObjectDataBLOBReference>();

            while ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) != 0)
            {
                if (header.Type == StreamObjectTypeHeaderStart.ObjectGroupObjectData)
                {
                    index += headerLength;
                    this.ObjectGroupObjectDataList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as ObjectGroupObjectData);
                }
                else if (header.Type == StreamObjectTypeHeaderStart.ObjectGroupObjectDataBLOBReference)
                {
                    index += headerLength;
                    this.ObjectGroupObjectDataBLOBReferenceList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as ObjectGroupObjectDataBLOBReference);
                }
                else
                {
                    throw new StreamObjectParseErrorException(index, "ObjectGroupDeclarations", "Failed to parse ObjectGroupData, expect the inner object type either ObjectGroupObjectData or ObjectGroupObjectDataBLOBReference, but actual type value is " + header.Type, null);
                }
            }

            currentIndex = index;
        }
    }

    /// <summary>
    /// Object Metadata Declaration
    /// </summary>
    public class ObjectGroupMetadataDeclarations : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupMetadataDeclarations class.
        /// </summary>
        public ObjectGroupMetadataDeclarations()
            : base(StreamObjectTypeHeaderStart.ObjectGroupMetadataDeclarations)
        {
            this.ObjectGroupMetadataList = new List<ObjectGroupMetadata>();
        }

        /// <summary>
        /// Gets or sets a list of Object Metadata.
        /// </summary>
        public List<ObjectGroupMetadata> ObjectGroupMetadataList { get; set; }

        /// <summary>
        /// Used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>A constant value 0</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.ObjectGroupMetadataList != null)
            {
                foreach (ObjectGroupMetadata objectGroupMetadata in this.ObjectGroupMetadataList)
                {
                    byteList.AddRange(objectGroupMetadata.SerializeToByteList());
                }
            }

            return 0;
        }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupMetadataDeclarations", "Stream object over-parse error", null);
            }

            int index = currentIndex;
            int headerLength = 0;
            StreamObjectHeaderStart header;
            this.ObjectGroupMetadataList = new List<ObjectGroupMetadata>();

            while ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) != 0)
            {
                index += headerLength;
                if (header.Type == StreamObjectTypeHeaderStart.ObjectGroupMetadata)
                {
                    this.ObjectGroupMetadataList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as ObjectGroupMetadata);
                }
                else
                {
                    throw new StreamObjectParseErrorException(index, "ObjectGroupDeclarations", "Failed to parse ObjectGroupMetadataDeclarations, expect the inner object type ObjectGroupMetadata, but actual type value is " + header.Type, null);
                }
            }

            currentIndex = index;
        }
    }

    /// <summary>
    /// Specifies an object group metadata.
    /// </summary>
    public class ObjectGroupMetadata : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ObjectGroupMetadata class.
        /// </summary>
        public ObjectGroupMetadata()
            : base(StreamObjectTypeHeaderStart.ObjectGroupMetadata)
        {
        }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the expected change frequency of the object.
        /// This value MUST be:
        /// 0, if the change frequency is not known.
        /// 1, if the object is known to change frequently.
        /// 2, if the object is known to change infrequently.
        /// 3, if the object is known to change independently of any other objects.
        /// </summary>
        public Compact64bitInt ObjectChangeFrequency { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.ObjectChangeFrequency = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ObjectGroupMetadata", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of elements actually contained in the list</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            List<byte> tmpList = this.ObjectChangeFrequency.SerializeToByteList();
            byteList.AddRange(tmpList);
            return tmpList.Count;
        }
    }

    /// <summary>
    /// Specifies an data element hash stream object.
    /// </summary>
    public class DataElementHash : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the DataElementHash class.
        /// </summary>
        public DataElementHash()
            : base(StreamObjectTypeHeaderStart.DataElementHash)
        {
        }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the hash schema. This value MUST be 1, indicating Content Information Data Structure Version 1.0.
        /// </summary>
        public Compact64bitInt DataElementHashScheme { get; set; }

        /// <summary>
        /// Gets or sets the data element hash data.
        /// </summary>
        public BinaryItem DataElementHashData { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.DataElementHashScheme = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.DataElementHashData = BasicObject.Parse<BinaryItem>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "DataElementHash", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List 
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of elements actually contained in the list</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int startPoint = byteList.Count;
            byteList.AddRange(this.DataElementHashScheme.SerializeToByteList());
            byteList.AddRange(this.DataElementHashData.SerializeToByteList());

            return byteList.Count - startPoint;
        }
    }
}