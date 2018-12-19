namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Common;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to the FileNode structure.
    /// </summary>
    public class FileNode
    {
        /// <summary>
        /// Gets or sets the value of FileNodeID field.
        /// </summary>
        public uint FileNodeID { get; set; }

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
                this.FileNodeID = bitReader.ReadUInt32(10);
                this.Size = bitReader.ReadUInt32(13);
                this.StpFormat = bitReader.ReadUInt32(2);
                this.CbFormat = bitReader.ReadUInt32(2);
                this.BaseType = bitReader.ReadUInt32(4);
                this.Reserved = bitReader.ReadUInt32(1);
            }
            index += 4;
            switch (this.FileNodeID)
            {
                case 0x004:
                    this.fnd = new ObjectSpaceManifestRootFND();
                    break;
                case 0x008:
                    this.fnd = new ObjectSpaceManifestListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x00C:
                    this.fnd = new ObjectSpaceManifestListStartFND();
                    break;
                case 0x010:
                    this.fnd = new RevisionManifestListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x014:
                    this.fnd = new RevisionManifestListStartFND();
                    break;
                case 0x01B:
                    this.fnd = new RevisionManifestStart4FND();
                    break;
                case 0x01E:
                    this.fnd = new RevisionManifestStart6FND();
                    break;
                case 0x01F:
                    this.fnd = new RevisionManifestStart7FND();
                    break;
                case 0x021:
                    this.fnd = new GlobalIdTableStartFNDX();
                    break;
                case 0x024:
                    this.fnd = new GlobalIdTableEntryFNDX();
                    break;
                case 0x025:
                    this.fnd = new GlobalIdTableEntry2FNDX();
                    break;
                case 0x026:
                    this.fnd = new GlobalIdTableEntry3FNDX();
                    break;
                case 0x02D:
                    this.fnd = new ObjectDeclarationWithRefCountFNDX(this.StpFormat, this.CbFormat);
                    break;
                case 0x02E:
                    this.fnd = new ObjectDeclarationWithRefCount2FNDX(this.StpFormat, this.CbFormat);
                    break;
                case 0x041:
                    this.fnd = new ObjectRevisionWithRefCountFNDX(this.StpFormat, this.CbFormat);
                    break;
                case 0x042:
                    this.fnd = new ObjectRevisionWithRefCount2FNDX(this.StpFormat, this.CbFormat);
                    break;
                case 0x059:
                    this.fnd = new RootObjectReference2FNDX();
                    break;
                case 0x05A:
                    this.fnd = new RootObjectReference3FND();
                    break;
                case 0x05C:
                    this.fnd = new RevisionRoleDeclarationFND();
                    break;
                case 0x05D:
                    this.fnd = new RevisionRoleAndContextDeclarationFND();
                    break;
                case 0x072:
                    this.fnd = new ObjectDeclarationFileData3RefCountFND();
                    break;
                case 0x073:
                    this.fnd = new ObjectDeclarationFileData3LargeRefCountFND();
                    break;
                case 0x07C:
                    this.fnd = new ObjectDataEncryptionKeyV2FNDX(this.StpFormat, this.CbFormat);
                    break;
                case 0x084:
                    this.fnd = new ObjectInfoDependencyOverridesFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x08C:
                    this.fnd = new DataSignatureGroupDefinitionFND();
                    break;
                case 0x090:
                    this.fnd = new FileDataStoreListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x094:
                    this.fnd = new FileDataStoreObjectReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x0A4:
                    this.fnd = new ObjectDeclaration2RefCountFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x0A5:
                    this.fnd = new ObjectDeclaration2LargeRefCountFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x0B0:
                    this.fnd = new ObjectGroupListReferenceFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x0B4:
                    this.fnd = new ObjectGroupStartFND();
                    break;
                case 0x0C2:
                    this.fnd = new HashedChunkDescriptor2FND(this.StpFormat, this.CbFormat);
                    break;
                case 0x0C4:
                    this.fnd = new ReadOnlyObjectDeclaration2RefCountFND(this.StpFormat, this.CbFormat);
                    break;
                case 0x0C5:
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
            bitWriter.AppendUInit32(this.FileNodeID, 10);
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
