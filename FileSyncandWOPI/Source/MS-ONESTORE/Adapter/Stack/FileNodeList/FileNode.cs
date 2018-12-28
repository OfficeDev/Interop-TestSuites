namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Common;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to the FileNode structure.
    /// </summary>
    public class FileNode
    {
        /// <summary>
        /// Gets or sets the value of FileNodeID field.
        /// </summary>
        public FileNodeIDValues FileNodeID { get; set; }

        /// <summary>
        /// Gets or sets the value of Size field.
        /// </summary>
        public uint Size { get; set; }

        /// <summary>
        /// Gets or sets the value of StpFormat field,
        /// </summary>
        public uint StpFormat { get; set; }

        /// <summary>
        /// Gets or sets the value of CbFormat field.
        /// </summary>
        public uint CbFormat { get; set; }

        /// <summary>
        /// Gets or sets the value of BaseType field.
        /// </summary>
        public uint BaseType { get; set; }

        /// <summary>
        /// Gets or sets the value of Reserved field.
        /// </summary>
        public uint Reserved { get; set; }

        /// <summary>
        /// Gets or sets the value of fnd field.
        /// </summary>
        public FileNodeBase fnd { get; set; }

        /// <summary>
        /// This method is used to deserialize the FileNodeListHeader object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileNodeListHeader object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            using (BitReader bitReader = new BitReader(byteArray, index))
            {
                this.FileNodeID = (FileNodeIDValues)bitReader.ReadUInt32(10);
                this.Size = bitReader.ReadUInt32(13);
                this.StpFormat = bitReader.ReadUInt32(2);
                this.CbFormat = bitReader.ReadUInt32(2);
                this.BaseType = bitReader.ReadUInt32(4);
                this.Reserved = bitReader.ReadUInt32(1);
            }
            index += 4;
            switch (this.FileNodeID)
            {
                case FileNodeIDValues.ObjectSpaceManifestRootFND:
                    this.fnd = new ObjectSpaceManifestRootFND();
                    break;
                case FileNodeIDValues.ObjectSpaceManifestListReferenceFND:
                    this.fnd = new ObjectSpaceManifestListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectSpaceManifestListStartFND:
                    this.fnd = new ObjectSpaceManifestListStartFND();
                    break;
                case FileNodeIDValues.RevisionManifestListReferenceFND:
                    this.fnd = new RevisionManifestListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.RevisionManifestListStartFND:
                    this.fnd = new RevisionManifestListStartFND();
                    break;
                case FileNodeIDValues.RevisionManifestStart4FND:
                    this.fnd = new RevisionManifestStart4FND();
                    break;
                case FileNodeIDValues.RevisionManifestStart6FND:
                    this.fnd = new RevisionManifestStart6FND();
                    break;
                case FileNodeIDValues.RevisionManifestStart7FND:
                    this.fnd = new RevisionManifestStart7FND();
                    break;
                case FileNodeIDValues.GlobalIdTableStartFNDX:
                    this.fnd = new GlobalIdTableStartFNDX();
                    break;
                case FileNodeIDValues.GlobalIdTableEntryFNDX:
                    this.fnd = new GlobalIdTableEntryFNDX();
                    break;
                case FileNodeIDValues.GlobalIdTableEntry2FNDX:
                    this.fnd = new GlobalIdTableEntry2FNDX();
                    break;
                case FileNodeIDValues.GlobalIdTableEntry3FNDX:
                    this.fnd = new GlobalIdTableEntry3FNDX();
                    break;
                case FileNodeIDValues.ObjectDeclarationWithRefCountFNDX:
                    this.fnd = new ObjectDeclarationWithRefCountFNDX(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectDeclarationWithRefCount2FNDX:
                    this.fnd = new ObjectDeclarationWithRefCount2FNDX(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectRevisionWithRefCountFNDX:
                    this.fnd = new ObjectRevisionWithRefCountFNDX(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectRevisionWithRefCount2FNDX:
                    this.fnd = new ObjectRevisionWithRefCount2FNDX(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.RootObjectReference2FNDX:
                    this.fnd = new RootObjectReference2FNDX();
                    break;
                case FileNodeIDValues.RootObjectReference3FND:
                    this.fnd = new RootObjectReference3FND();
                    break;
                case FileNodeIDValues.RevisionRoleDeclarationFND:
                    this.fnd = new RevisionRoleDeclarationFND();
                    break;
                case FileNodeIDValues.RevisionRoleAndContextDeclarationFND:
                    this.fnd = new RevisionRoleAndContextDeclarationFND();
                    break;
                case FileNodeIDValues.ObjectDeclarationFileData3RefCountFND:
                    this.fnd = new ObjectDeclarationFileData3RefCountFND();
                    break;
                case FileNodeIDValues.ObjectDeclarationFileData3LargeRefCountFND:
                    this.fnd = new ObjectDeclarationFileData3LargeRefCountFND();
                    break;
                case FileNodeIDValues.ObjectDataEncryptionKeyV2FNDX:
                    this.fnd = new ObjectDataEncryptionKeyV2FNDX(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectInfoDependencyOverridesFND:
                    this.fnd = new ObjectInfoDependencyOverridesFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.DataSignatureGroupDefinitionFND:
                    this.fnd = new DataSignatureGroupDefinitionFND();
                    break;
                case FileNodeIDValues.FileDataStoreListReferenceFND:
                    this.fnd = new FileDataStoreListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.FileDataStoreObjectReferenceFND:
                    this.fnd = new FileDataStoreObjectReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectDeclaration2RefCountFND:
                    this.fnd = new ObjectDeclaration2RefCountFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectDeclaration2LargeRefCountFND:
                    this.fnd = new ObjectDeclaration2LargeRefCountFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectGroupListReferenceFND:
                    this.fnd = new ObjectGroupListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ObjectGroupStartFND:
                    this.fnd = new ObjectGroupStartFND();
                    break;
                case FileNodeIDValues.HashedChunkDescriptor2FND:
                    this.fnd = new HashedChunkDescriptor2FND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ReadOnlyObjectDeclaration2RefCountFND:
                    this.fnd = new ReadOnlyObjectDeclaration2RefCountFND(this.StpFormat, this.CbFormat);
                    break;
                case FileNodeIDValues.ReadOnlyObjectDeclaration2LargeRefCountFND:
                    this.fnd = new ReadOnlyObjectDeclaration2LargeRefCountFND(this.StpFormat, this.CbFormat);
                    break;
                default:
                    this.fnd = null;
                    break;

            }
            if (this.fnd != null)
            {
                int len = this.fnd.DoDeserializeFromByteArray(byteArray, index);
                index += len;
            }

            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of FileNodeListHeader object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileNodeListHeader.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            BitWriter bitWriter = new BitWriter(4);
            bitWriter.AppendUInit32((uint)this.FileNodeID, 10);
            bitWriter.AppendUInit32(this.Size, 13);
            bitWriter.AppendUInit32(this.StpFormat, 2);
            bitWriter.AppendUInit32(this.CbFormat, 2);
            bitWriter.AppendUInit32(this.BaseType, 4);
            bitWriter.AppendUInit32(this.Reserved, 1);
            byteList.AddRange(bitWriter.Bytes);
            if(this.fnd!=null)
            {
                byteList.AddRange(this.fnd.SerializeToByteList());
            }

            return byteList;
        }
    }
}
