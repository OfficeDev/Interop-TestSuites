namespace Microsoft.Protocols.TestSuites.MS_OXWSITEMID
{
    using System.IO;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSITEMID
    /// </summary>
    public interface IMS_OXWSITEMIDAdapter : IAdapter
    {
        /// <summary>
        /// Parse an ItemId's Id from a base64 string to a ItemIdId object according to the defined format 
        /// </summary>
        /// <param name="itemId">An ItemIdType object</param>
        /// <returns>An ItemIdId object as the result of parsing</returns>
        ItemIdId ParseItemId(ItemIdType itemId);

        /// <summary>
        /// Compresses the passed byte array using a simple RLE compression scheme.
        /// </summary>
        /// <param name="streamIn">input stream to compress</param>
        /// <param name="compressorId">id of the compressor</param>
        /// <returns>compressed bytes</returns>
        byte[] Compress(byte[] streamIn, byte compressorId);

        /// <summary>
        /// Decompresses the passed byte array using RLE scheme.
        /// </summary>
        /// <param name="input">Bytes to decompress</param>
        /// <param name="maxLength">Max allowed length for the byte array</param>
        /// <returns>Decompressed bytes minus the first byte of input</returns>
        MemoryStream Decompress(byte[] input, int maxLength);
    }
}