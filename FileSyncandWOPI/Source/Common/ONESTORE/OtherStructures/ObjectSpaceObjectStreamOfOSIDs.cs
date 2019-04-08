namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;
    /// <summary>
    /// This class is used to represent a ObjectSpaceObjectStreamOfOSIDs.
    /// </summary>
    public class ObjectSpaceObjectStreamOfOSIDs
    {
        /// <summary>
        /// Gets or sets the value of header field.
        /// </summary>
        public ObjectSpaceObjectStreamHeader Header { get; set; }
        /// <summary>
        /// Gets or sets the value of body field.
        /// </summary>
        public CompactID[] Body { get; set; }

        /// <summary>
        /// This method is used to convert the element of ObjectSpaceObjectStreamOfOSIDs object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectSpaceObjectStreamOfOSIDs</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Header.SerializeToByteList());
            foreach (CompactID compactID in this.Body)
            {
                byteList.AddRange(compactID.SerializeToByteList());
            }

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the ObjectSpaceObjectStreamOfOSIDs object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectSpaceObjectStreamOfOSIDs object.</returns>
        public int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Header = new ObjectSpaceObjectStreamHeader();
            int headerCount = this.Header.DoDeserializeFromByteArray(byteArray, index);
            index += headerCount;

            this.Body = new CompactID[this.Header.Count];
            for (int i = 0; i < this.Header.Count; i++)
            {
                CompactID compactID = new CompactID();
                int count = compactID.DoDeserializeFromByteArray(byteArray, index);
                this.Body[i] = compactID;
                index += count;
            }

            return index - startIndex;
        }
    }
}
