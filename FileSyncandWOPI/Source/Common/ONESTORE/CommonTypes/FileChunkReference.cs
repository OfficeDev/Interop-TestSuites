namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the File Chunk Reference.
    /// </summary>
    public abstract class FileChunkReference
    {
        /// <summary>
        /// This method is used to convert the element of FileChunkReference object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of FileChunkReference</returns>
        public abstract List<byte> SerializeToByteList();

        /// <summary>
        /// This method is used to deserialize the FileChunkReference object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the FileChunkReference object.</returns>
        public abstract int DoDeserializeFromByteArray(byte[] byteArray, int startIndex);
        /// <summary>
        /// This method is used the check the instance whether is fcrNil.
        /// </summary>
        /// <returns>return the whether the instance is fcrNil.</returns>
        public abstract bool IsfcrNil();

        /// <summary>
        /// This method is used the check the instance whether is fcrZero.
        /// </summary>
        /// <returns>return the whether the instance is fcrZero.</returns>
        public abstract bool IsfcrZero();
    }
}
