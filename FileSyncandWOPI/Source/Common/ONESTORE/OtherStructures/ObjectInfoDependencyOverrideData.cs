namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// This class is used to represent the ObjectInfoDependencyOverrideData structure. 
    /// </summary>
    public class ObjectInfoDependencyOverrideData
    {
        /// <summary>
        /// Gets or sets the value of c8BitOverrides field.
        /// </summary>
        public uint c8BitOverrides { get; set; }

        /// <summary>
        /// Gets or sets the value of c32BitOverrides field.
        /// </summary>
        public uint c32BitOverrides { get; set; }

        /// <summary>
        /// Gets or sets the value of crc field.
        /// </summary>
        public uint crc { get; set; }
        /// <summary>
        /// Gets or sets the value of Overrides1 field.
        /// </summary>
        public ObjectInfoDependencyOverride8 Overrides1 { get; set; }

        /// <summary>
        /// Gets or sets the value of Overrides2 field.
        /// </summary>
        public ObjectInfoDependencyOverride32 Overrides2 { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectInfoDependencyOverrideData object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectInfoDependencyOverrideData object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.c8BitOverrides = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.c32BitOverrides= BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.crc = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.Overrides1 = new ObjectInfoDependencyOverride8();
            int len = this.Overrides1.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.Overrides2 = new ObjectInfoDependencyOverride32();
            len = this.Overrides2.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectInfoDependencyOverrideData object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectInfoDependencyOverrideData.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.c8BitOverrides));
            byteList.AddRange(BitConverter.GetBytes(this.c32BitOverrides));
            byteList.AddRange(BitConverter.GetBytes(this.crc));
            byteList.AddRange(this.Overrides1.SerializeToByteList());
            byteList.AddRange(this.Overrides2.SerializeToByteList());

            return byteList;
        }
    }
    /// <summary>
    /// This class is used to represent the ObjectInfoDependencyOverride8 structure.
    /// </summary>
    public class ObjectInfoDependencyOverride8
    {
        /// <summary>
        /// Gets or sets the value of oid field.
        /// </summary>
        public CompactID oid { get; set; }

        /// <summary>
        /// Gets or sets the value of cRef field.
        /// </summary>
        public byte cRef { get; set; }
        /// <summary>
        /// This method is used to deserialize the ObjectInfoDependencyOverride8 object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectInfoDependencyOverride8 object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.oid = new CompactID();
            int len = this.oid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cRef = byteArray[index];
            index += 1;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectInfoDependencyOverride8 object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectInfoDependencyOverride8.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.oid.SerializeToByteList());
            byteList.Add(this.cRef);

            return byteList;
        }
    }
    /// <summary>
    /// This class is used to represent the ObjectInfoDependencyOverride32 structure.
    /// </summary>
    public class ObjectInfoDependencyOverride32
    {
        /// <summary>
        /// Gets or sets the value of oid field.
        /// </summary>
        public CompactID oid { get; set; }
        /// <summary>
        /// Gets or sets the value of cRef field.
        /// </summary>
        public uint cRef { get; set; }

        /// <summary>
        /// This method is used to deserialize the ObjectInfoDependencyOverride32 object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectInfoDependencyOverride32 object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.oid = new CompactID();
            int len = this.oid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.cRef = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of ObjectInfoDependencyOverride32 object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectInfoDependencyOverride32.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.oid.SerializeToByteList());
            byteList.AddRange(BitConverter.GetBytes(this.cRef));

            return byteList;
        }
    }
}
