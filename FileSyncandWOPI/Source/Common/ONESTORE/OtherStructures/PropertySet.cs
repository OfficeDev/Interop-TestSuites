namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    /// <summary>
    /// This class is used to represent a PropertySet.
    /// </summary>
    public class PropertySet: IProperty
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of properties in this PropertySet structure.
        /// </summary>
        public ushort CProperties { get; set; }

        /// <summary>
        /// Gets or sets the value of rgPrids.
        /// </summary>
        public PropertyID[] RgPrids { get; set; }
        /// <summary>
        /// Gets or sets the value of rgData field.
        /// </summary>
        public List<IProperty> RgData { get; set; }

        /// <summary>
        /// This method is used to convert the element of PropertySet into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of PropertySet.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.CProperties));

            foreach(PropertyID propertyId in this.RgPrids)
            {
                byteList.AddRange(propertyId.SerializeToByteList());
            }

            foreach(IProperty property in this.RgData)
            {
                byteList.AddRange(property.SerializeToByteList());
            }

            return byteList;
        }
        /// <summary>
        /// This method is used to deserialize the PropertySet from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the PropertySet.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.CProperties = BitConverter.ToUInt16(byteArray, index);
            index += 2;
            this.RgPrids = new PropertyID[this.CProperties];
            for (int i = 0; i < this.CProperties; i++)
            {
                PropertyID propertyID = new PropertyID();
                propertyID.DoDeserializeFromByteArray(byteArray, index);
                this.RgPrids[i] = propertyID;
                index += 4;
            }
            this.RgData = new List<IProperty>();
            foreach(PropertyID propertyID in this.RgPrids)
            {
                IProperty property = null;
                switch ((PropertyType)propertyID.Type)
                {
                    case PropertyType.NoData:
                    case PropertyType.Bool:
                    case PropertyType.ObjectID:
                    case PropertyType.ContextID:
                    case PropertyType.ObjectSpaceID:
                        property = new NoData();
                        break;
                    case PropertyType.ArrayOfObjectIDs:
                    case PropertyType.ArrayOfObjectSpaceIDs:
                    case PropertyType.ArrayOfContextIDs:
                        property = new ArrayNumber();
                        break;
                    case PropertyType.OneByteOfData:
                        property = new OneByteOfData();
                        break;
                    case PropertyType.TwoBytesOfData:
                        property = new TwoBytesOfData();
                        break;
                    case PropertyType.FourBytesOfData:
                        property = new FourBytesOfData();
                        break;
                    case PropertyType.EightBytesOfData:
                        property = new EightBytesOfData();
                        break;
                    case PropertyType.FourBytesOfLengthFollowedByData:
                        property = new PrtFourBytesOfLengthFollowedByData();
                        break;
                    case PropertyType.ArrayOfPropertyValues:
                        property = new PrtArrayOfPropertyValues();
                        break;
                    case PropertyType.PropertySet:
                        property = new PropertySet();
                        break;
                    default:
                        break;
                }
                if (property != null)
                {
                    int len = property.DoDeserializeFromByteArray(byteArray, index);
                    this.RgData.Add(property);
                    index += len;
                }
            }

            return index - startIndex;
        }
    }
    /// <summary>
    /// This class is used to represent the property contains no data.
    /// </summary>
    public class NoData : IProperty
    {
        /// <summary>
        /// This method is used to deserialize the NoData from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the NoData.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            return 0;
        }
        /// <summary>
        /// This method is used to convert the element of NoData into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of NoData.</returns>
        public List<byte> SerializeToByteList()
        {
            return new List<byte>();
        }
    }
    /// <summary>
    /// This class is used to represent the property contains 1 byte of data in the PropertySet.rgData stream field.
    /// </summary>
    public class OneByteOfData : IProperty
    {
        /// <summary>
        /// Gets or sets the data of property.
        /// </summary>
        public byte Data { get; set; }

        /// <summary>
        /// This method is used to deserialize the OneByteOfData from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the OneByteOfData.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.Data = byteArray[startIndex];
            return 1;
        }
        /// <summary>
        /// This method is used to convert the element of OneByteOfData into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of OneByteOfData.</returns>
        public List<byte> SerializeToByteList()
        {
            return new List<byte>(this.Data);
        }
    }
    /// <summary>
    /// This class is used to represent the property contains 2 bytes of data in the PropertySet.rgData stream field.
    /// </summary>
    public class TwoBytesOfData : IProperty
    {
        /// <summary>
        /// Gets or sets the data of property.
        /// </summary>
        public byte[] Data { get; set; }

        /// <summary>
        /// This method is used to deserialize the TwoBytesOfData from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the TwoBytesOfData.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.Data = new byte[2] { byteArray[startIndex], byteArray[startIndex + 1] };

            return 2;
        }
        /// <summary>
        /// This method is used to convert the element of TwoBytesOfData into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of TwoBytesOfData.</returns>
        public List<byte> SerializeToByteList()
        {
            return new List<byte>(this.Data);
        }
    }

    /// <summary>
    /// This class is used to represent the property contains 4 bytes of data in the PropertySet.rgData stream field.
    /// </summary>
    public class FourBytesOfData : IProperty
    {
        /// <summary>
        ///  Gets or sets the data of property.
        /// </summary>
        public byte[] Data { get; set; }
        /// <summary>
        /// This method is used to deserialize the FourBytesOfData from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FourBytesOfData.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.Data = new byte[4];
            Array.Copy(byteArray, startIndex, this.Data, 0, 4);

            return 4;
        }

        /// <summary>
        /// This method is used to convert the element of FourBytesOfData into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FourBytesOfData.</returns>
        public List<byte> SerializeToByteList()
        {
            return new List<byte>(this.Data);
        }
    }

    /// <summary>
    /// This class is used to represent the property contains 8 bytes of data in the PropertySet.rgData stream field.
    /// </summary>
    public class EightBytesOfData : IProperty
    {
        /// <summary>
        /// Gets or sets the data of property.
        /// </summary>
        public byte[] Data { get; set; }

        /// <summary>
        /// This method is used to deserialize the EightBytesOfData from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the EightBytesOfData.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.Data = new byte[8];
            Array.Copy(byteArray, startIndex, this.Data, 0, 8);
            return 8;
        }
        /// <summary>
        /// This method is used to convert the element of EightBytesOfData into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of EightBytesOfData.</returns>
        public List<byte> SerializeToByteList()
        {
            return new List<byte>(this.Data);
        }
    }
    /// <summary>
    /// This class is used to represent the prtFourBytesOfLengthFollowedByData.
    /// </summary>
    public class PrtFourBytesOfLengthFollowedByData : IProperty
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the size, in bytes, of the Data field.
        /// </summary>
        public uint CB { get; set; }

        /// <summary>
        /// Gets or sets the value of Data field.
        /// </summary>
        public byte[] Data { get; set; }
        /// <summary>
        /// This method is used to deserialize the prtFourBytesOfLengthFollowedByData from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the prtFourBytesOfLengthFollowedByData.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.CB = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.Data = new byte[this.CB];
            Array.Copy(byteArray, index, this.Data, 0, this.CB);
            index += (int)this.CB;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of prtFourBytesOfLengthFollowedByData into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of prtFourBytesOfLengthFollowedByData.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.CB));
            byteList.AddRange(this.Data);

            return byteList;
        }
    }
    /// <summary>
    /// The class is used to represent the number of the array.
    /// </summary>
    public class ArrayNumber : IProperty
    {
        /// <summary>
        /// Gets or sets the number of array.
        /// </summary>
        public uint Number { get; set; }
        /// <summary>
        /// This method is used to deserialize the number of array from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the number of array.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            this.Number = BitConverter.ToUInt32(byteArray, startIndex);
            return 4;
        }
        /// <summary>
        /// This method is used to convert the element of the number of array into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of the number of array.</returns>
        public List<byte> SerializeToByteList()
        {
            return new List<byte>(BitConverter.GetBytes(this.Number));
        }
    }
    /// <summary>
    /// The class is used to represent the prtArrayOfPropertyValues . 
    /// </summary>
    public class PrtArrayOfPropertyValues : IProperty
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of properties in Data.
        /// </summary>
        public uint CProperties { get; set; }
        /// <summary>
        /// Gets or sets the value of prid field.
        /// </summary>
        public PropertyID Prid { get; set; }
        /// <summary>
        /// Gets or sets the value of Data field.
        /// </summary>
        public PropertySet[] Data { get; set; }
        /// <summary>
        /// This method is used to deserialize the prtArrayOfPropertyValues from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the prtArrayOfPropertyValues.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.CProperties = BitConverter.ToUInt32(byteArray, index);
            index += 4;
            this.Prid = new PropertyID();
            int len = this.Prid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.Data = new PropertySet[this.CProperties];
            for (int i = 0; i < this.CProperties; i++)
            {
                this.Data[i] = new PropertySet();
                int length = this.Data[i].DoDeserializeFromByteArray(byteArray, index);
                index += length;
            }

            return index - startIndex;
        }
        /// <summary>
        /// This method is used to convert the element of the prtArrayOfPropertyValues into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of the prtArrayOfPropertyValues.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(BitConverter.GetBytes(this.CProperties));
            byteList.AddRange(this.Prid.SerializeToByteList());
            foreach (PropertySet ps in this.Data)
            {
                byteList.AddRange(ps.SerializeToByteList());
            }
            return byteList;
        }
    }
}
