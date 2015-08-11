namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies the starting offset and length of a contiguous portion of a file.
    /// </summary>
    public class FileChunk : BasicObject
    {
        /// <summary>
        /// Initializes a new instance of the FileChunk class.
        /// </summary>
        public FileChunk()
        {
            this.Start = new Compact64bitInt();
            this.Length = new Compact64bitInt();
        }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the byte-offset within the file of the beginning of the file chunk. 
        /// </summary>
        public Compact64bitInt Start { get; set; }
        
        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the count of bytes included in the file chunk.
        /// </summary>
        public Compact64bitInt Length { get; set; }

        /// <summary>
        /// This method is used to convert the element of FileChunk basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileChunk.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> bytelist = new List<byte>();
            bytelist.AddRange(this.Start.SerializeToByteList());
            bytelist.AddRange(this.Length.SerializeToByteList());
            return bytelist;
        }

        /// <summary>
        /// This method is used to deserialize the FileChunk basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileChunk basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex) 
        {
            int index = startIndex;

            this.Start = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.Length = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            return index - startIndex;
        }
    }
}