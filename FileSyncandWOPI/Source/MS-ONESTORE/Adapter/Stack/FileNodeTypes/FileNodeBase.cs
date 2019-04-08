namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// The abstract class is used to represent FileNode structure.
    /// </summary>
    public abstract class FileNodeBase
    {
        /// <summary>
        /// This method is used to deserialize the FileNode object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileNode object.</returns>
        public abstract int DoDeserializeFromByteArray(byte[] byteArray, int startIndex);
        /// <summary>
        /// This method is used to convert the element of FileNode object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileNode</returns>
        public abstract List<byte> SerializeToByteList();
    }
}
