namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Common;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent a PropertyID.
    /// </summary>
    public class PropertyID
    {
        /// <summary>
        /// Gets or sets the value of id field.
        /// </summary>
        public uint Id { get; set; }
        /// <summary>
        /// Gets or sets the value of type field.
        /// </summary>
        public uint Type { get; set; }
        /// <summary>
        /// Gets or sets the value of boolValue field.
        /// </summary>
        public int BoolValue { get; set; }

        /// <summary>
        /// Gets or sets the value of PropertyID.
        /// </summary>
        public int Value { get; set; }

        /// <summary>
        /// This method is used to convert the element of PropertyID object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of PropertyID</returns>
        public List<byte> SerializeToByteList()
        {
            BitWriter bitWriter = new BitWriter(4);
            bitWriter.AppendUInit32(this.Id, 26);
            bitWriter.AppendUInit32(this.Type, 5);
            bitWriter.AppendInit32(this.BoolValue, 1);
           
            return new List<byte>(bitWriter.Bytes);
        }

        /// <summary>
        /// This method is used to deserialize the PropertyID object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the PropertyID object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            using (BitReader bitReader = new BitReader(byteArray, startIndex))
            {
                this.Id = bitReader.ReadUInt32(26);
                this.Type = bitReader.ReadUInt32(5);
                this.BoolValue = bitReader.ReadInt32(1);
                this.Value = BitConverter.ToInt32(byteArray, startIndex);
                return 4;
            }
        }
    }
    /// <summary>
    /// The types of properties.
    /// </summary>
    public enum PropertyType:uint
    {
        /// <summary>
        /// The property contains no data.
        /// </summary>
        NoData = 0x1,
        /// <summary>
        /// The property is a Boolean value specified by boolValue.
        /// </summary>
        Bool = 0x2,
        /// <summary>
        /// The property contains 1 byte of data in the PropertySet.rgData stream field.
        /// </summary>
        OneByteOfData = 0x3,
        /// <summary>
        /// The property contains 2 bytes of data in the PropertySet.rgData stream field.
        /// </summary>
        TwoBytesOfData = 0x4,
        /// <summary>
        /// The property contains 4 bytes of data in the PropertySet.rgData stream field.
        /// </summary>
        FourBytesOfData = 0x5,
        /// <summary>
        /// The property contains 8 bytes of data in the PropertySet.rgData stream field.
        /// </summary>
        EightBytesOfData = 0x6,
        /// <summary>
        /// The property contains a prtFourBytesOfLengthFollowedByData in the PropertySet.rgData stream field.
        /// </summary>
        FourBytesOfLengthFollowedByData = 0x7,
        /// <summary>
        /// The property contains one CompactID in the ObjectSpaceObjectPropSet.OIDs.body stream field.
        /// </summary>
        ObjectID = 0x8,
        /// <summary>
        /// The property contains an array of CompactID structures in the ObjectSpaceObjectPropSet.OSIDs.body stream field. 
        /// </summary>
        ArrayOfObjectIDs = 0x9,
        /// <summary>
        /// The property contains one CompactID structure in the ObjectSpaceObjectPropSet.OSIDs.body stream field.
        /// </summary>
        ObjectSpaceID = 0xA,
        /// <summary>
        /// The property contains an array of CompactID structures in the ObjectSpaceObjectPropSet.OSIDs.body stream field. 
        /// </summary>
        ArrayOfObjectSpaceIDs = 0xB,
        /// <summary>
        /// The property contains one CompactID in the ObjectSpaceObjectPropSet.ContextIDs.body stream field.
        /// </summary>
        ContextID = 0xC,
        /// <summary>
        /// The property contains an array of CompactID structures in the ObjectSpaceObjectPropSet.ContextIDs.body stream field. 
        /// </summary>
        ArrayOfContextIDs = 0xD,
        /// <summary>
        /// The property contains a prtArrayOfPropertyValues structure in the PropertySet.rgData stream field.
        /// </summary>
        ArrayOfPropertyValues = 0x10,
        /// <summary>
        /// The property contains a child PropertySet structure in the PropertySet.rgData stream field of the parent PropertySet.
        /// </summary>
        PropertySet = 0x11
    }
}
