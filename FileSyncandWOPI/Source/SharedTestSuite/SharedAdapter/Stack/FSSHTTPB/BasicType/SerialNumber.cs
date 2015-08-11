namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies variable-width encoding that specifies a combination of a GUID and an unsigned 64-bit integer.  
    /// </summary>
   public class SerialNumber : BasicObject
    {
        /// <summary>
        /// Initializes a new instance of the SerialNumber class with specified values.
        /// </summary>
        /// <param name="identifier">Specify the Guid value of the serialNumber.</param>
        /// <param name="value">Specify the value of the serialNumber.</param>
        public SerialNumber(Guid identifier, ulong value)
        {
            this.GUID = identifier;
            this.Value = value;
        }

        /// <summary>
        /// Initializes a new instance of the SerialNumber class, this is the copy constructor.
        /// </summary>
        /// <param name="sn">Specify the serial number where copy from.</param>
        public SerialNumber(SerialNumber sn)
        {
            this.GUID = sn.GUID;
            this.Value = sn.Value;
        }

        /// <summary>
        /// Initializes a new instance of the SerialNumber class, this is default contractor
        /// </summary>
        public SerialNumber()
        {
        }
 
        /// <summary>
        /// Gets or sets a value which indicate the SerialNumber type.
        /// </summary>
        public uint Type { get; set; }

        /// <summary>
        /// Gets or sets a GUID that specifies the item. 
        /// </summary>
        public Guid GUID { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the value of the serial number.
        /// </summary>
        public ulong Value { get; set; }

        /// <summary>
        /// This method is used to convert the element of SerialNumber basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of SerialNumber.</returns>
        public override List<byte> SerializeToByteList()
        {
            BitWriter bitWriter = null;
            if (this.GUID == Guid.Empty)
            {
                bitWriter = new BitWriter(1);
                bitWriter.AppendUInit32(0, 8);
            }
            else
            {
                bitWriter = new BitWriter(25);
                bitWriter.AppendUInit32(128, 8);
                bitWriter.AppendGUID(this.GUID);
                bitWriter.AppendUInt64(this.Value, 64);
            }

            return new List<byte>(bitWriter.Bytes);
        }

        /// <summary>
        /// This method is used to deserialize the SerialNumber basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the SerialNumber basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex) // return the length consumed
        {
            using (BitReader bitField = new BitReader(byteArray, startIndex))
            {
                int type = bitField.ReadInt32(8);

                if (type == 0)
                {
                    this.GUID = Guid.Empty;
                    this.Type = 0;

                    return 1;
                }
                else if (type == 128)
                {
                    this.GUID = bitField.ReadGuid();
                    this.Value = bitField.ReadUInt64(64);
                    this.Type = 128;
                    return 25;
                }
                else
                {
                    throw new System.IO.IOException("Failed to parse SerialNumber object, Expect the type value is either 0 or 128, but the actual value is " + this.Type);
                }
            }
        }
    }
}