namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the Alternative Packaging structure.
    /// </summary>
    public class AlternativePackaging
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
        /// Gets or sets the value of rgbReserved.
        /// </summary>
        public uint rgbReserved { get; set; }
        /// <summary>
        /// Gets or sets the value of PackagingStart field.
        /// </summary>
        public StreamObjectHeaderStart32bit packagingStart { get; set; }
        /// <summary>
        /// Gets or sets the value of StorageIndexExtendedGUID field.
        /// </summary>
        public ExGuid storageIndexExtendedGUID { get; set; }
        /// <summary>
        /// Gets or sets the value of guidCellSchemaId field.
        /// </summary>
        public Guid guidCellSchemaId { get; set; }
        /// <summary>
        /// Gets or sets the value of dataElementPackage field.
        /// </summary>
        public DataElementPackage dataElementPackage { get; set; }
        /// <summary>
        /// Gets or sets the value of packagingEnd field.
        /// </summary>
        public StreamObjectHeaderEnd packagingEnd { get; set; }

        /// <summary>
        /// This method is used to convert the element of Alternative Packaging object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of Alternative Packaging</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.guidFileType.ToByteArray());
            byteList.AddRange(this.guidFile.ToByteArray());
            byteList.AddRange(this.guidLegacyFileVersion.ToByteArray());
            byteList.AddRange(this.guidFileFormat.ToByteArray());
            byteList.AddRange(BitConverter.GetBytes(this.rgbReserved));
            byteList.AddRange(this.packagingStart.SerializeToByteList());
            byteList.AddRange(this.storageIndexExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.guidCellSchemaId.ToByteArray());
            byteList.AddRange(this.dataElementPackage.SerializeToByteList());
            byteList.AddRange(this.packagingEnd.SerializeToByteList());
            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the Alternative Packaging object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the Alternative Packaging object.</returns>
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
            this.rgbReserved = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.packagingStart = new StreamObjectHeaderStart32bit();
            this.packagingStart.DeserializeFromByteArray(byteArray, index);
            index += 4;
            this.storageIndexExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.guidCellSchemaId = AdapterHelper.ReadGuid(byteArray, index);
            index += 16;
            DataElementPackage package;
            this.dataElementPackage = StreamObject.TryGetCurrent<DataElementPackage>(byteArray, ref index, out package) ? package : null;
            this.packagingEnd = new StreamObjectHeaderEnd16bit();
            this.packagingEnd.DeserializeFromByteArray(byteArray, index);
            index += 2;

            return index - startIndex;
        }
    }
}
