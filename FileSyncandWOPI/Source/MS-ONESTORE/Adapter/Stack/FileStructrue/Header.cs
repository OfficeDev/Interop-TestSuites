namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// This class is used to represent the header of a revision store file.
    /// </summary>
    public class Header
    {
        /// <summary>
        /// Gets or sets the value of guidFileType.
        /// </summary>
        public Guid guidFileType { get; set; }
        /// <summary>
        /// Gets or sets the value of guidFile.
        /// </summary>
        public Guid guidFile { get; set; }
        /// <summary>
        /// Gets or sets the value of guidLegacyFileVersion.
        /// </summary>
        public Guid guidLegacyFileVersion { get; set; }
        /// <summary>
        /// Gets or sets the value of guidFileFormat.
        /// </summary>
        public Guid guidFileFormat { get; set; }
        /// <summary>
        /// Gets or sets the value of ffvLastCodeThatWroteToThisFile.
        /// </summary>
        public uint ffvLastCodeThatWroteToThisFile { get; set; }
        /// <summary>
        /// Gets or sets the value of ffvOldestCodeThatHasWrittenToThisFile.
        /// </summary>
        public uint ffvOldestCodeThatHasWrittenToThisFile { get; set; }
        /// <summary>
        /// Gets or sets the value of ffvNewestCodeThatHasWrittenToThisFile.
        /// </summary>
        public uint ffvNewestCodeThatHasWrittenToThisFile { get; set; }
        /// <summary>
        /// Gets or sets the value of ffvOldestCodeThatMayReadThisFile.
        /// </summary>
        public uint ffvOldestCodeThatMayReadThisFile { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrLegacyFreeChunkList field.
        /// </summary>
        public FileChunkReference32 fcrLegacyFreeChunkList { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrLegacyTransactionLog field.
        /// </summary>
        public FileChunkReference32 fcrLegacyTransactionLog { get; set; }
        /// <summary>
        /// Gets or sets the value of cTransactionsInLog field.
        /// </summary>
        public uint cTransactionsInLog { get; set; }
        /// <summary>
        /// Gets or sets the value of cbLegacyExpectedFileLength field.
        /// </summary>
        public uint cbLegacyExpectedFileLength { get; set; }
        /// <summary>
        /// Gets or sets the value of rgbPlaceholder field.
        /// </summary>
        public UInt64 rgbPlaceholder { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrLegacyFileNodeListRoot field.
        /// </summary>
        public FileChunkReference32 fcrLegacyFileNodeListRoot { get; set; }
        /// <summary>
        /// Gets or sets the value of cbLegacyFreeSpaceInFreeChunkList field.
        /// </summary>
        public uint cbLegacyFreeSpaceInFreeChunkList { get; set; }
        /// <summary>
        /// Gets or sets the value of fNeedsDefrag field.
        /// </summary>
        public byte fNeedsDefrag { get; set; }
        /// <summary>
        /// Gets or sets the value of fRepairedFile field.
        /// </summary>
        public byte fRepairedFile { get; set; }
        /// <summary>
        /// Gets or sets the value of fNeedsGarbageCollect field.
        /// </summary>
        public byte fNeedsGarbageCollect { get; set; }
        /// <summary>
        /// Gets or sets the value of fHasNoEmbeddedFileObjects field.
        /// </summary>
        public byte fHasNoEmbeddedFileObjects { get; set; }
        /// <summary>
        /// Gets or sets the value of guidAncestor field.
        /// </summary>
        public Guid guidAncestor { get; set; }
        /// <summary>
        /// Gets or sets the value of crcName field.
        /// </summary>
        public uint crcName { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrHashedChunkList field.
        /// </summary>
        public FileChunkReference64x32 fcrHashedChunkList { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrTransactionLog field.
        /// </summary>
        public FileChunkReference64x32 fcrTransactionLog { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrFileNodeListRoot field.
        /// </summary>
        public FileChunkReference64x32 fcrFileNodeListRoot { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrFreeChunkList field.
        /// </summary>
        public FileChunkReference64x32 fcrFreeChunkList { get; set; }
        /// <summary>
        /// Gets or sets the value of cbExpectedFileLength field.
        /// </summary>
        public UInt64 cbExpectedFileLength { get; set; }
        /// <summary>
        /// Gets or sets the value of cbFreeSpaceInFreeChunkList field.
        /// </summary>
        public UInt64 cbFreeSpaceInFreeChunkList { get; set; }
        /// <summary>
        /// Gets or sets the value of guidFileVersion field.
        /// </summary>
        public Guid guidFileVersion { get; set; }
        /// <summary>
        /// Gets or sets the value of nFileVersionGeneration field.
        /// </summary>
        public UInt64 nFileVersionGeneration { get; set; }
        /// <summary>
        /// Gets or sets the value of guidDenyReadFileVersion field.
        /// </summary>
        public Guid guidDenyReadFileVersion { get; set; }
        /// <summary>
        /// Gets or sets the value of grfDebugLogFlags field.
        /// </summary>
        public uint grfDebugLogFlags { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrDebugLog field.
        /// </summary>
        public FileChunkReference64x32 fcrDebugLog { get; set; }
        /// <summary>
        /// Gets or sets the value of fcrAllocVerificationFreeChunkList field.
        /// </summary>
        public FileChunkReference64x32 fcrAllocVerificationFreeChunkList { get; set; }
        /// <summary>
        /// Gets or sets the value of bnCreated field.
        /// </summary>
        public uint bnCreated { get; set; }
        /// <summary>
        /// Gets or sets the value of bnLastWroteToThisFile field.
        /// </summary>
        public uint bnLastWroteToThisFile { get; set; }
        /// <summary>
        /// Gets or sets the value of bnOldestWritten field.
        /// </summary>
        public uint bnOldestWritten { get; set; }
        /// <summary>
        /// Gets or sets the value of bnNewestWritten field.
        /// </summary>
        public uint bnNewestWritten { get; set; }
        /// <summary>
        /// Gets or sets the value of rgbReserved field.
        /// </summary>
        public byte[] rgbReserved { get; set; }

        /// <summary>
        /// This method is used to convert the element of Header object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of Header</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.guidFileType.ToByteArray());
            byteList.AddRange(this.guidFile.ToByteArray());
            byteList.AddRange(this.guidLegacyFileVersion.ToByteArray());
            byteList.AddRange(BitConverter.GetBytes(this.ffvLastCodeThatWroteToThisFile));
            byteList.AddRange(BitConverter.GetBytes(this.ffvOldestCodeThatHasWrittenToThisFile));
            byteList.AddRange(BitConverter.GetBytes(this.ffvNewestCodeThatHasWrittenToThisFile));
            byteList.AddRange(BitConverter.GetBytes(this.ffvOldestCodeThatMayReadThisFile));
            byteList.AddRange(this.fcrLegacyFreeChunkList.SerializeToByteList());
            byteList.AddRange(this.fcrLegacyTransactionLog.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.cTransactionsInLog));
            byteList.AddRange(BitConverter.GetBytes(this.cbLegacyExpectedFileLength));
            byteList.AddRange(BitConverter.GetBytes(this.rgbPlaceholder));
            byteList.AddRange(this.fcrLegacyFileNodeListRoot.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.cbLegacyFreeSpaceInFreeChunkList));
            byteList.Add(this.fNeedsDefrag);
            byteList.Add(this.fRepairedFile);
            byteList.Add(this.fNeedsGarbageCollect);
            byteList.Add(this.fHasNoEmbeddedFileObjects);
            byteList.AddRange(this.guidAncestor.ToByteArray());
            byteList.AddRange(BitConverter.GetBytes(this.crcName));
            byteList.AddRange(this.fcrHashedChunkList.SerializeToByteList());
            byteList.AddRange(this.fcrTransactionLog.SerializeToByteList());
            byteList.AddRange(this.fcrFileNodeListRoot.SerializeToByteList());
            byteList.AddRange(this.fcrFreeChunkList.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.cbExpectedFileLength));
            byteList.AddRange(BitConverter.GetBytes(this.cbFreeSpaceInFreeChunkList));
            byteList.AddRange(this.guidFileVersion.ToByteArray());
            byteList.AddRange(BitConverter.GetBytes(this.nFileVersionGeneration));
            byteList.AddRange(this.guidDenyReadFileVersion.ToByteArray());
            byteList.AddRange(BitConverter.GetBytes(this.grfDebugLogFlags));
            byteList.AddRange(this.fcrDebugLog.SerializeToByteList());
            byteList.AddRange(this.fcrAllocVerificationFreeChunkList.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.bnCreated));
            byteList.AddRange(BitConverter.GetBytes(this.bnLastWroteToThisFile));
            byteList.AddRange(BitConverter.GetBytes(this.bnOldestWritten));
            byteList.AddRange(BitConverter.GetBytes(this.bnNewestWritten));
            byteList.AddRange(this.rgbReserved);

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the Header object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the Header object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.guidFileType = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            this.guidFile = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            this.guidLegacyFileVersion = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            this.guidFileFormat = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            this.ffvLastCodeThatWroteToThisFile = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.ffvOldestCodeThatHasWrittenToThisFile = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.ffvNewestCodeThatHasWrittenToThisFile = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.ffvOldestCodeThatMayReadThisFile = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.fcrLegacyFreeChunkList = new FileChunkReference32();
            int len = this.fcrLegacyFreeChunkList.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.fcrLegacyTransactionLog = new FileChunkReference32();
            len = this.fcrLegacyTransactionLog.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cTransactionsInLog = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.cbLegacyExpectedFileLength = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.rgbPlaceholder = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.fcrLegacyFileNodeListRoot = new FileChunkReference32();
            len = this.fcrLegacyFileNodeListRoot.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cbLegacyFreeSpaceInFreeChunkList = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.fNeedsDefrag = byteArray[index];
            index += 1;
            this.fRepairedFile= byteArray[index];
            index += 1;
            this.fNeedsGarbageCollect = byteArray[index];
            index += 1;
            this.fHasNoEmbeddedFileObjects = byteArray[index];
            index += 1;
            this.guidAncestor = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            this.crcName = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.fcrHashedChunkList = new FileChunkReference64x32();
            len=this.fcrHashedChunkList.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.fcrTransactionLog = new FileChunkReference64x32();
            len = this.fcrTransactionLog.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.fcrFileNodeListRoot = new FileChunkReference64x32();
            len=this.fcrFileNodeListRoot.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.fcrFreeChunkList = new FileChunkReference64x32();
            len=this.fcrFreeChunkList.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cbExpectedFileLength = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.cbFreeSpaceInFreeChunkList = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.guidFileVersion = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            this.nFileVersionGeneration = BitConverter.ToUInt64(byteArray, index);
            index += 8;
            this.guidDenyReadFileVersion = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            this.grfDebugLogFlags = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.fcrDebugLog = new FileChunkReference64x32();
            len = this.fcrDebugLog.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.fcrAllocVerificationFreeChunkList = new FileChunkReference64x32();
            len = this.fcrAllocVerificationFreeChunkList.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.bnCreated = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.bnLastWroteToThisFile = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.bnOldestWritten = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.bnNewestWritten = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.rgbReserved = new byte[728];
            Array.Copy(byteArray, index, this.rgbReserved, 0, 728);
            index += 728;

            return index - startIndex;
        }
    }
}
