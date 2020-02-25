namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This class is used to represent a node object.
    /// </summary>
    public abstract class NodeObject : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the NodeObject class.
        /// </summary>
        /// <param name="headerType">Specify the node object header type.</param>
        protected NodeObject(StreamObjectTypeHeaderStart headerType)
            : base(headerType)
        {
        }

        /// <summary>
        /// Gets or sets the extended GUID of this node object.
        /// </summary>
        public ExGuid ExGuid { get; set; }

        /// <summary>
        /// Gets or sets the intermediate node object list.
        /// </summary>
        public List<LeafNodeObject> IntermediateNodeObjectList { get; set; }

        /// <summary>
        /// Gets or sets the signature.
        /// </summary>
        public SignatureObject Signature { get; set; }

        /// <summary>
        /// Gets or sets the data size.
        /// </summary>
        public DataSizeObject DataSize { get; set; }

        /// <summary>
        /// Get all the content which is represented by the node object.
        /// </summary>
        /// <returns>Return the byte list of node object content.</returns>
        public abstract List<byte> GetContent();
    }

    /// <summary>
    /// The data of Root Node Object.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class IntermediateNodeObject : NodeObject
    {
        /// <summary>
        /// Initializes a new instance of the IntermediateNodeObject class.
        /// </summary>
        public IntermediateNodeObject()
            : base(StreamObjectTypeHeaderStart.IntermediateNodeObject)
        {
            this.IntermediateNodeObjectList = new List<LeafNodeObject>();
        }

        /// <summary>
        /// Get all the content which is represented by the root node object.
        /// </summary>
        /// <returns>Return the byte list of root node object content.</returns>
        public override List<byte> GetContent()
        {
            List<byte> content = new List<byte>();

            foreach (LeafNodeObject intermediateNode in this.IntermediateNodeObjectList)
            {
                content.AddRange(intermediateNode.GetContent());
            }

            return content;
        }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            if (lengthOfItems != 0)
            {
                throw new StreamObjectParseErrorException(currentIndex, "IntermediateNodeObject", "Stream Object over-parse error", null);
            }

            this.Signature = StreamObject.GetCurrent<SignatureObject>(byteArray, ref index);
            this.DataSize = StreamObject.GetCurrent<DataSizeObject>(byteArray, ref index);

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The Byte list</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            byteList.AddRange(this.Signature.SerializeToByteList());
            byteList.AddRange(this.DataSize.SerializeToByteList());
            return 0;
        }

        /// <summary>
        /// The class is used to build a root node object.
        /// </summary>
        public class RootNodeObjectBuilder
        {
            /// <summary>
            /// This method is used to build a root node object from an data element list with the specified storage index extended GUID.
            /// </summary>
            /// <param name="dataElements">Specify the data element list.</param>
            /// <param name="storageIndexExGuid">Specify the storage index extended GUID.</param>
            /// <returns>Return a root node object build from the data element list.</returns>
            public IntermediateNodeObject Build(List<DataElement> dataElements, ExGuid storageIndexExGuid)
            {
                if (DataElementUtils.TryAnalyzeWhetherFullDataElementList(dataElements, storageIndexExGuid)
                    && DataElementUtils.TryAnalyzeWhetherConfirmSchema(dataElements, storageIndexExGuid))
                {
                    ExGuid rootObjectExGUID;
                    List<ObjectGroupDataElementData> objectGroupList = DataElementUtils.GetDataObjectDataElementData(dataElements, storageIndexExGuid, out rootObjectExGUID);

                    // If the root object extend GUID can be found, then the root node can be build.
                    if (rootObjectExGUID != null)
                    {
                        // If can analyze for here, then can directly capture all the GUID values related requirements
                        if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                        {
                            MsfsshttpdCapture.VerifyDefinedGUID(SharedContext.Current.Site);
                        }

                        return this.Build(objectGroupList, rootObjectExGUID);
                    }
                    else
                    {
                        throw new InvalidOperationException(string.Format("There is no root extended GUID value {0}", DataElementUtils.RootExGuid.ToString()));
                    }
                }

                return null;
            }

            /// <summary>
            /// This method is used to build a root node object from a byte array.
            /// </summary>
            /// <param name="fileContent">Specify the byte array.</param>
            /// <returns>Return a root node object build from the byte array.</returns>
            public IntermediateNodeObject Build(byte[] fileContent)
            {
                IntermediateNodeObject rootNode = new IntermediateNodeObject();
                rootNode.Signature = new SignatureObject();
                rootNode.DataSize = new DataSizeObject();
                rootNode.DataSize.DataSize = (ulong)fileContent.Length;
                rootNode.ExGuid = new ExGuid(SequenceNumberGenerator.GetCurrentSerialNumber(), Guid.NewGuid());
                rootNode.IntermediateNodeObjectList = ChunkingFactory.CreateChunkingInstance(fileContent).Chunking();
                return rootNode;
            }

            /// <summary>
            /// This method is used to build a root node object from an object group data element list with the specified root extended GUID.
            /// </summary>
            /// <param name="objectGroupList">Specify the object group data element list.</param>
            /// <param name="rootExGuid">Specify the root extended GUID.</param>
            /// <returns>Return a root node object build from the object group data element list.</returns>
            private IntermediateNodeObject Build(List<ObjectGroupDataElementData> objectGroupList, ExGuid rootExGuid)
            {
                ObjectGroupObjectDeclare rootDeclare;
                ObjectGroupObjectData root = this.FindByExGuid(objectGroupList, rootExGuid, out rootDeclare);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    MsfsshttpdCapture.VerifyObjectCount(root, SharedContext.Current.Site);
                }

                int index = 0;
                IntermediateNodeObject rootNode = null;

                if (StreamObject.TryGetCurrent<IntermediateNodeObject>(root.Data.Content.ToArray(), ref index, out rootNode))
                {
                    rootNode.ExGuid = rootExGuid;

                    foreach (ExGuid extGuid in root.ObjectExGUIDArray.Content)
                    {
                        ObjectGroupObjectDeclare intermediateDeclare;
                        ObjectGroupObjectData intermediateData = this.FindByExGuid(objectGroupList, extGuid, out intermediateDeclare);
                        rootNode.IntermediateNodeObjectList.Add(new LeafNodeObject.IntermediateNodeObjectBuilder().Build(objectGroupList, intermediateData, extGuid));

                        // Capture the intermediate related requirements
                        if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                        {
                            MsfsshttpdCapture.VerifyObjectGroupObjectDataForIntermediateNode(intermediateData, intermediateDeclare, objectGroupList, SharedContext.Current.Site);
                            MsfsshttpdCapture.VerifyLeafNodeObject(rootNode.IntermediateNodeObjectList.Last(), SharedContext.Current.Site);
                        }
                    }

                    if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                    {
                        // Capture the root node related requirements. 
                        MsfsshttpdCapture.VerifyObjectGroupObjectDataForRootNode(root, rootDeclare, objectGroupList, SharedContext.Current.Site);
                        MsfsshttpdCapture.VerifyIntermediateNodeObject(rootNode, SharedContext.Current.Site);
                    }
                }
                else
                {
                    // If there is only one object in the file, SharePoint Server 2010 does not return the Root Node Object, but an Intermediate Node Object at the beginning.
                    // At this case, we will add the root node object for the further parsing.
                    rootNode = new IntermediateNodeObject();
                    rootNode.ExGuid = rootExGuid;
                    
                    rootNode.IntermediateNodeObjectList.Add(new LeafNodeObject.IntermediateNodeObjectBuilder().Build(objectGroupList, root, rootExGuid));
                    rootNode.DataSize = new DataSizeObject();
                    rootNode.DataSize.DataSize = (ulong)rootNode.IntermediateNodeObjectList.Sum(o => (float)o.DataSize.DataSize);
                }

                // Capture all the signature related requirements.
                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    AbstractChunking chunking = ChunkingFactory.CreateChunkingInstance(rootNode);

                    if (chunking != null)
                    {
                        chunking.AnalyzeChunking(rootNode, SharedContext.Current.Site);
                    }
                }
                
                return rootNode;
            }

            /// <summary>
            /// This method is used to find the object group data element using the specified extended GUID.
            /// </summary>
            /// <param name="objectGroupList">Specify the object group data element list.</param>
            /// <param name="extendedGuid">Specify the extended GUID.</param>
            /// <param name="declare">Specify the output of ObjectGroupObjectDeclare.</param>
            /// <returns>Return the object group data element if found.</returns>
            /// <exception cref="InvalidOperationException">If not found, throw the InvalidOperationException exception.</exception>
            private ObjectGroupObjectData FindByExGuid(List<ObjectGroupDataElementData> objectGroupList, ExGuid extendedGuid, out ObjectGroupObjectDeclare declare)
            {
                foreach (ObjectGroupDataElementData objectGroup in objectGroupList)
                {
                    int findIndex = objectGroup.ObjectGroupDeclarations.ObjectDeclarationList.FindIndex(objDeclare => objDeclare.ObjectExtendedGUID.Equals(extendedGuid));

                    if (findIndex == -1)
                    {
                        continue;
                    }

                    declare = objectGroup.ObjectGroupDeclarations.ObjectDeclarationList[findIndex];
                    return objectGroup.ObjectGroupData.ObjectGroupObjectDataList[findIndex];
                }

                throw new InvalidOperationException("Cannot find the " + extendedGuid.GUID.ToString());
            }
        }
    }

    /// <summary>
    /// The data of Intermediate Node Object.
    /// </summary>
    public class LeafNodeObject : NodeObject
    {
        /// <summary>
        /// Initializes a new instance of the LeafNodeObjectData class.
        /// </summary>
        public LeafNodeObject()
            : base(StreamObjectTypeHeaderStart.LeafNodeObject)
        {
        }

        /// <summary>
        /// Gets or sets the data node object.
        /// </summary>
        public DataNodeObjectData DataNodeObjectData { get; set; }

        /// <summary>
        /// Gets or sets the data size.
        /// </summary>
        public DataHashObject DataHash { get; set; }

        /// <summary>
        /// Get all the content which is represented by the intermediate node object.
        /// </summary>
        /// <returns>Return the byte list of intermediate node object content.</returns>
        public override List<byte> GetContent()
        {
            List<byte> content = new List<byte>();

            if (this.DataNodeObjectData != null)
            {
                content.AddRange(this.DataNodeObjectData.ObjectData);
            }
            else if (this.IntermediateNodeObjectList != null)
            {
                foreach (LeafNodeObject intermediateNode in this.IntermediateNodeObjectList)
                {
                    content.AddRange(intermediateNode.GetContent());
                }
            }
            else
            {
                throw new InvalidOperationException("The DataNodeObjectData and IntermediateNodeObjectList properties in LeafNodeObjectData cannot be null at the same time.");
            }
            
            return content;
        }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            if (lengthOfItems != 0)
            {
                throw new StreamObjectParseErrorException(currentIndex, "LeafNodeObjectData", "Stream Object over-parse error", null);
            }

            this.Signature = StreamObject.GetCurrent<SignatureObject>(byteArray, ref index);
            this.DataSize = StreamObject.GetCurrent<DataSizeObject>(byteArray, ref index);

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>A constant value</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            byteList.AddRange(this.Signature.SerializeToByteList());
            byteList.AddRange(this.DataSize.SerializeToByteList());
            return 0;
        }

        /// <summary>
        /// The class is used to build a intermediate node object.
        /// </summary>
        public class IntermediateNodeObjectBuilder
        {
            /// <summary>
            /// This method is used to build intermediate node object from an list of object group data element.
            /// </summary>
            /// <param name="objectGroupList">Specify the list of object group data elements.</param>
            /// <param name="dataObj">Specify the object group object.</param>
            /// <param name="intermediateGuid">Specify the intermediate extended GUID.</param>
            /// <returns>Return the intermediate node object.</returns>
            public LeafNodeObject Build(List<ObjectGroupDataElementData> objectGroupList, ObjectGroupObjectData dataObj, ExGuid intermediateGuid)
            {
                LeafNodeObject node = null;
                IntermediateNodeObject rootNode = null;

                int index = 0;
                if (StreamObject.TryGetCurrent<LeafNodeObject>(dataObj.Data.Content.ToArray(), ref index, out node))
                {
                    if (dataObj.ObjectExGUIDArray == null)
                    {
                        throw new InvalidOperationException("Failed to build intermediate node because the object extend GUID array does not exist.");
                    }

                    node.ExGuid = intermediateGuid;

                    // Contain a single Data Node Object.
                    if (dataObj.ObjectExGUIDArray.Count.DecodedValue == 1u)
                    {
                        ObjectGroupObjectDeclare dataNodeDeclare;
                        ObjectGroupObjectData dataNodeData = this.FindByExGuid(objectGroupList, dataObj.ObjectExGUIDArray.Content[0], out dataNodeDeclare);
                        BinaryItem data = dataNodeData.Data;
                        
                        node.DataNodeObjectData = new DataNodeObjectData(data.Content.ToArray(), 0, (int)data.Length.DecodedValue);
                        node.DataNodeObjectData.ExGuid = dataObj.ObjectExGUIDArray.Content[0];
                        node.IntermediateNodeObjectList = null;

                        if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                        {
                            MsfsshttpdCapture.VerifyObjectGroupObjectDataForDataNodeObject(dataNodeData, dataNodeDeclare, objectGroupList, SharedContext.Current.Site);
                        }
                    }
                    else
                    {
                        // Contain a list of LeafNodeObjectData
                        node.IntermediateNodeObjectList = new List<LeafNodeObject>();
                        node.DataNodeObjectData = null;
                        foreach (ExGuid extGuid in dataObj.ObjectExGUIDArray.Content)
                        {
                            ObjectGroupObjectDeclare intermediateDeclare;
                            ObjectGroupObjectData intermediateData = this.FindByExGuid(objectGroupList, extGuid, out intermediateDeclare);
                            node.IntermediateNodeObjectList.Add(new IntermediateNodeObjectBuilder().Build(objectGroupList, intermediateData, extGuid));

                            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                            {
                                MsfsshttpdCapture.VerifyObjectGroupObjectDataForIntermediateNode(intermediateData, intermediateDeclare, objectGroupList, SharedContext.Current.Site);
                            }
                        }
                    }
                }
                else if (StreamObject.TryGetCurrent<IntermediateNodeObject>(dataObj.Data.Content.ToArray(), ref index, out rootNode))
                {
                    // In Sub chunking for larger than 1MB zip file, MOSS2010 could return IntermediateNodeObject.
                    // For easy further process, the rootNode will be replaced by intermediate node instead.
                    node = new LeafNodeObject();
                    node.IntermediateNodeObjectList = new List<LeafNodeObject>();
                    node.DataSize = rootNode.DataSize;
                    node.ExGuid = rootNode.ExGuid;
                    node.Signature = rootNode.Signature;
                    node.DataNodeObjectData = null;
                    foreach (ExGuid extGuid in dataObj.ObjectExGUIDArray.Content)
                    {
                        ObjectGroupObjectDeclare intermediateDeclare;
                        ObjectGroupObjectData intermediateData = this.FindByExGuid(objectGroupList, extGuid, out intermediateDeclare);
                        node.IntermediateNodeObjectList.Add(new IntermediateNodeObjectBuilder().Build(objectGroupList, intermediateData, extGuid));

                        if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                        {
                            MsfsshttpdCapture.VerifyObjectGroupObjectDataForIntermediateNode(intermediateData, intermediateDeclare, objectGroupList, SharedContext.Current.Site);
                        }
                    }
                }
                else
                {
                    throw new InvalidOperationException("In the ObjectGroupDataElement cannot only contain the IntermediateNodeObject or IntermediateNodeOBject.");
                }

                return node;
            }

            /// <summary>
            /// This method is used to build intermediate node object from a byte array with a signature.
            /// </summary>
            /// <param name="array">Specify the byte array.</param>
            /// <param name="signature">Specify the signature.</param>
            /// <returns>Return the intermediate node object.</returns>
            public LeafNodeObject Build(byte[] array, SignatureObject signature)
            {
                LeafNodeObject nodeObject = new LeafNodeObject();
                nodeObject.DataSize = new DataSizeObject();
                nodeObject.DataSize.DataSize = (ulong)array.Length;

                nodeObject.Signature = signature;
                nodeObject.ExGuid = new ExGuid(SequenceNumberGenerator.GetCurrentSerialNumber(), Guid.NewGuid());

                nodeObject.DataNodeObjectData = new DataNodeObjectData(array, 0, array.Length);
                nodeObject.IntermediateNodeObjectList = null;

                // Now in the current implementation, one intermediate node only contain one single data object node.
                return nodeObject;
            }

            /// <summary>
            /// This method is used to find the object group data element using the specified extended GUID.
            /// </summary>
            /// <param name="objectGroupList">Specify the object group data element list.</param>
            /// <param name="extendedGuid">Specify the extended GUID.</param>
            /// <param name="declare">Specify the output of ObjectGroupObjectDeclare.</param>
            /// <returns>Return the object group data element if found.</returns>
            /// <exception cref="InvalidOperationException">If not found, throw the InvalidOperationException exception.</exception>
            private ObjectGroupObjectData FindByExGuid(List<ObjectGroupDataElementData> objectGroupList, ExGuid extendedGuid, out ObjectGroupObjectDeclare declare)
            {
                foreach (ObjectGroupDataElementData objectGroup in objectGroupList)
                {
                    int findIndex = objectGroup.ObjectGroupDeclarations.ObjectDeclarationList.FindIndex(objDeclare => objDeclare.ObjectExtendedGUID.Equals(extendedGuid));

                    if (findIndex == -1)
                    {
                        continue;
                    }

                    declare = objectGroup.ObjectGroupDeclarations.ObjectDeclarationList[findIndex];
                    return objectGroup.ObjectGroupData.ObjectGroupObjectDataList[findIndex];
                }

                throw new InvalidOperationException("Cannot find the " + extendedGuid.GUID.ToString());
            }
        }
    }

    /// <summary>
    /// Data Node Object data.
    /// </summary>
    public class DataNodeObjectData
    {
        /// <summary>
        /// Initializes a new instance of the DataNodeObjectData class.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="startIndex">Start position</param>
        /// <param name="length">The element length</param>
        public DataNodeObjectData(byte[] byteArray, int startIndex, int length) : this()
        {
            this.ObjectData = new byte[length];
            Array.Copy(byteArray, startIndex, this.ObjectData, 0, length);
        }

        /// <summary>
        /// Initializes a new instance of the DataNodeObjectData class.
        /// </summary>
        internal DataNodeObjectData()
        {
            this.ExGuid = new ExGuid(SequenceNumberGenerator.GetCurrentSerialNumber(), Guid.NewGuid());
        }

        /// <summary>
        /// Gets or sets the extended GUID of the data node object.
        /// </summary>
        public ExGuid ExGuid { get; set; }

        /// <summary>
        /// Gets or sets the Data field for the Intermediate Node Object.
        /// </summary>
        public byte[] ObjectData { get; set; }
    }

    /// <summary>
    /// Signature Object
    /// </summary>
    public class SignatureObject : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the SignatureObject class.
        /// </summary>
        public SignatureObject()
            : base(StreamObjectTypeHeaderStart.SignatureObject)
        {
            this.SignatureData = new BinaryItem();
        }

        /// <summary>
        ///  Gets or sets a binary item as specified in [MS-FSSHTTPB] section 2.2.1.3 that specifies a value that is unique to the file data represented by this root node object. 
        ///  The value of this item depends on the file chunking algorithm used, as specified in section 2.4. 
        /// </summary>
        public BinaryItem SignatureData { get; set; }

        /// <summary>
        /// Override the equals method.
        /// </summary>
        /// <param name="obj">Specify the compared instance.</param>
        /// <returns>If equals return true, otherwise return false.</returns>
        public override bool Equals(object obj)
        {
            SignatureObject so = obj as SignatureObject;

            if (so == null)
            {
                return false;
            }

            if (so.SignatureData != null && this.SignatureData != null)
            {
                return so.SignatureData.Equals(this.SignatureData);
            }
            else if (so.SignatureData == null && this.SignatureData == null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Override the GetHashCode method.
        /// </summary>
        /// <returns>Return the hash code value.</returns>
        public override int GetHashCode()
        {
            if (this.SignatureData != null)
            {
                return this.SignatureData.GetHashCode();
            }
            else
            {
                return base.GetHashCode();
            }
        }

        /// <summary>
        /// Override the ToString method.
        /// </summary>
        /// <returns>Return the string represent the instance.</returns>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (byte value in this.SignatureData.Content)
            {
                sb.Append(value);
                sb.Append(",");
            }

            sb.Remove(sb.Length - 1, 1);

            return sb.ToString();
        }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            this.SignatureData = BasicObject.Parse<BinaryItem>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "Signature", "Stream Object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of elements</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int length = byteList.Count;
            byteList.AddRange(this.SignatureData.SerializeToByteList());
            return byteList.Count - length;
        }
    }

    /// <summary>
    /// Data Size Object
    /// </summary>
    public class DataSizeObject : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the DataSizeObject class.
        /// </summary>
        public DataSizeObject()
            : base(StreamObjectTypeHeaderStart.DataSizeObject)
        {
        }

        /// <summary>
        /// Gets or sets an unsigned 64-bit integer that specifies the size of the file data represented by this root node object.
        /// </summary>
        public ulong DataSize { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 8)
            {
                throw new StreamObjectParseErrorException(currentIndex, "DataSize", "Stream Object over-parse error", null);
            }

            this.DataSize = LittleEndianBitConverter.ToUInt64(byteArray, currentIndex);
            currentIndex += 8;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>A constant value 8</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            byteList.AddRange(LittleEndianBitConverter.GetBytes(this.DataSize));
            return 8;
        }
    }

    /// <summary>
    /// Data Hash Object
    /// </summary>
    public class DataHashObject : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the DataHashObject class.
        /// </summary>
        public DataHashObject()
            : base(StreamObjectTypeHeaderStart.DataHashObject)
        {
        }
    }
}