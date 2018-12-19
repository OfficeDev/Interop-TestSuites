namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent a ObjectSpaceObjectStreamOfOIDs.
    /// </summary>
    public class ObjectSpaceObjectStreamOfOIDs
    {
        /// <summary>
        /// Gets or sets an ObjectSpaceObjectStreamHeader that specifies the number of elements in the body field and whether the ObjectSpaceObjectPropSet structure contains an OSIDs field and ContextIDs field. 
        /// </summary>
        public ObjectSpaceObjectStreamHeader Header { get; set; }
        /// <summary>
        /// Gets or sets an array of CompactID structures.
        /// </summary>
        public CompactID[] Body { get; set; }
        /// <summary>
        /// This method is used to convert the element of ObjectSpaceObjectStreamOfOIDs object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ObjectSpaceObjectStreamOfOIDs</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Header.SerializeToByteList());
            foreach(CompactID compactID in this.Body)
            {
                byteList.AddRange(compactID.SerializeToByteList());
            }

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the ObjectSpaceObjectStreamOfOIDs object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ObjectSpaceObjectStreamOfOIDs object.</returns>
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
                int count = compactID.DoDeserializeFromByteArray(byteArray, startIndex);
                this.Body[i] = compactID;
                index += count;
            }

            return index - startIndex;
        }
    }
}
