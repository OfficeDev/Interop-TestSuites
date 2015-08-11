namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies arbitrary binary stream of data.
    /// </summary>
    public class BinaryItem : BasicObject
    {
        /// <summary>
        /// A byte stream that specifies the data for the item.
        /// </summary>
        private List<byte> content = null;

        /// <summary>
        /// Initializes a new instance of the BinaryItem class.
        /// </summary>
        public BinaryItem()
        {
            this.Length = new Compact64bitInt();
            this.content = new List<byte>();
        }

        /// <summary>
        /// Initializes a new instance of the BinaryItem class with the specified content.
        /// </summary>
        /// <param name="content">Specify the binary content.</param>
        public BinaryItem(IEnumerable<byte> content)
            : this()
        {
            this.content = new List<byte>();
            this.content.AddRange(content);
            this.Length.DecodedValue = (ulong)this.Content.Count;
        }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the count of bytes of Content of the item. 
        /// </summary>
        public Compact64bitInt Length { get; set; }
       
        /// <summary>
        /// Gets or sets a byte stream that specifies the data for the item.
        /// </summary>
        public List<byte> Content 
        {
            get
            {
                return this.content;
            }

            set
            {
                this.content = value;
            }
        }

        /// <summary>
        /// This method is used to convert the element of BinaryItem basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of BinaryItem.</returns>
        public override List<byte> SerializeToByteList()
        {
            this.Length.DecodedValue = (uint)this.content.Count;

            List<byte> result = new List<byte>();
            result.AddRange(this.Length.SerializeToByteList());
            result.AddRange(this.content);

            return result;
        }

        /// <summary>
        /// Override the GetHashCode method.
        /// </summary>
        /// <returns>Return the hash code value.</returns>
        public override int GetHashCode()
        {
            return this.content.GetHashCode();
        }

        /// <summary>
        /// Override the Equals method.
        /// </summary>
        /// <param name="obj">Specify the object.</param>
        /// <returns>Return true if equals, otherwise return false.</returns>
        public override bool Equals(object obj)
        {
            BinaryItem bi = obj as BinaryItem;

            if (bi == null)
            {
                return false;
            }

            return this.Length.DecodedValue == bi.Length.DecodedValue && AdapterHelper.ByteArrayEquals(this.content.ToArray(), bi.content.ToArray());
        }

        /// <summary>
        /// This method is used to de-serialize the BinaryItem basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the BinaryItem basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex) 
        {
            int index = startIndex;

            this.Length = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            
            this.Content.Clear();
            for (ulong i = 0; i < this.Length.DecodedValue; i++)
            {
                this.Content.Add(byteArray[index++]);
            }

            return index - startIndex;
        }
    }
}