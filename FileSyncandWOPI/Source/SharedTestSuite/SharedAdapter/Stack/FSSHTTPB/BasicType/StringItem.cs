namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// The count and content of an arbitrary wide character string.
    /// </summary>
    public class StringItem : BasicObject
    {
        /// <summary>
        /// Initializes a new instance of the StringItem class.
        /// </summary>
        public StringItem()
        {
            this.Count = new Compact64bitInt();
        }

        /// <summary>
        /// Gets or sets the count of character.
        /// </summary>
        public Compact64bitInt Count { get; set; }

        /// <summary>
        /// Gets or sets an array of UTF-16 characters that specify the string. It MUST NOT be null-terminated.
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// This method is used to convert the element of StringItem basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of StringItem.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Count.SerializeToByteList());

            ushort[] content = Encoding.Unicode.GetBytes(this.Content).ToUInt16Array();
            byte[] binaryString = new byte[content.Length * 2];
            Buffer.BlockCopy(content, 0, binaryString, 0, content.Length * 2);
            byteList.AddRange(binaryString);

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the StringItem basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the StringItem basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Count = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.Content = System.Text.Encoding.Unicode.GetString(byteArray, index, Convert.ToInt32(this.Count.DecodedValue * 2));
            index += (int)this.Count.DecodedValue * 2;
            return index - startIndex;
        }
    }

    /// <summary>
    /// This class is used to specify the array of String Items.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class StringItemArray : BasicObject
    {
        /// <summary>
        /// Initializes a new instance of the StringItemArray class.
        /// </summary>
        /// <param name="count">Specify the number of StringItem in the StringItem array.</param>
        /// <param name="content">Specify the list of StringItem.</param>
        public StringItemArray(ulong count, List<StringItem> content)
        {
            this.Count = count;
            this.Content = content;
        }

        /// <summary>
        /// Initializes a new instance of the StringItemArray class, this is default constructor.
        /// </summary>
        public StringItemArray()
        {
            this.Count = 0;
            this.Content = new List<StringItem>();
        }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the count of string Items in the array. 
        /// </summary>
        public ulong Count { get; set; }

        /// <summary>
        /// Gets or sets a String Items list that specifies a list of cells.
        /// </summary>
        public List<StringItem> Content { get; set; }

        /// <summary>
        /// This method is used to convert the element of StringItemArray basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of StringItemArray.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange((new Compact64bitInt(this.Count)).SerializeToByteList());
            if (this.Content != null)
            {
                foreach (StringItem DoDeserializeFromByteArray in this.Content)
                {
                    byteList.AddRange(DoDeserializeFromByteArray.SerializeToByteList());
                }
            }

            return byteList;
        }

        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            throw new NotImplementedException();
        }
    }

    // Extension method to convert a byte array to an array of UInt16
    public static class ByteArrayExtensions
    {
        public static ushort[] ToUInt16Array(this byte[] bytes)
        {
            if (bytes.Length % 2 != 0)
            {
                throw new ArgumentException("Byte array length must be even");
            }

            ushort[] result = new ushort[bytes.Length / 2];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = BitConverter.ToUInt16(bytes, i * 2);
            }

            return result;
        }
    }

}
