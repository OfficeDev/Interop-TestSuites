namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the ExtendedGUID structrue.
    /// </summary>
    public class ExtendedGUID
    {
        /// <summary>
        /// Gets or sets the value of guid field.
        /// </summary>
        public Guid Guid { get; set; }

        /// <summary>
        /// Gets or sets the value of n field.
        /// </summary>
        public uint N { get; set; }

        /// <summary>
        /// This method is used to convert the element of ExtendedGUID object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ExtendedGUID</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Guid.ToByteArray());
            byteList.AddRange(BitConverter.GetBytes(this.N));

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the ExtendedGUID object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ExtendedGUID object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            byte[] guidBuffer = new byte[16];
            Array.Copy(byteArray, index, guidBuffer, 0, 16);
            index += 16;
            this.Guid = new Guid(guidBuffer);
            this.N = BitConverter.ToUInt32(byteArray, index);
            index += 4;

            return index - startIndex;
        }
        /// <summary>
        /// Determines whether the ExtendedGuid object is equal to the current object.
        /// </summary>
        /// <param name="obj">The object to compare with the current object.</param>
        /// <returns>true if the ExtendedGuid object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if(obj.GetType()!=typeof(ExtendedGUID))
            {
                return false;
            } 
            else if(((ExtendedGUID)obj).Guid!=this.Guid || ((ExtendedGUID)obj).N!=this.N)
            {
                return false;
            }

            return true;
        }
        /// <summary>
        ///  Calculate a hash code for the ExtendedGUID.
        /// </summary>
        /// <returns>Return a hash code for the ExtendedGUID object.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = Guid.GetHashCode();
                hashCode ^= N.GetHashCode();

                return hashCode;
            }
        }
    }
}
