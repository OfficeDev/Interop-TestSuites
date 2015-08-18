namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// The struct of Global Identifier (GID)
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct GlobalIdentifier : ISerializable, IDeserializable
    {
        /// <summary>
        /// A 128-bit unsigned integer identifying a Store object.
        /// </summary>
        public byte[] ReplGuid;

        /// <summary>
        /// An unsigned 48-bit integer identifying the folder within its Store object.
        /// </summary>
        public byte[] GlobalCounter;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] resultBytes = new byte[this.Size()];
            Array.Copy(this.ReplGuid, 0, resultBytes, index, sizeof(byte) * 16);
            index += sizeof(byte) * 16;
            Array.Copy(this.GlobalCounter, 0, resultBytes, index, sizeof(byte) * 6);
            index += sizeof(byte) * 6;
            Array.Copy(BitConverter.GetBytes(0), 0, resultBytes, index, sizeof(ushort));
            return resultBytes;
        }

        /// <summary>
        /// Return the size of this struct.
        /// </summary>
        /// <returns>The size of this struct.</returns>
        public int Size()
        {
            return sizeof(byte) * 22;
        }

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer struct.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.ReplGuid = new byte[16];
            Array.Copy(ropBytes, index, this.ReplGuid, 0, sizeof(byte) * 16);
            index += sizeof(byte) * 16;
            this.GlobalCounter = new byte[6];
            Array.Copy(ropBytes, index, this.GlobalCounter, 0, sizeof(byte) * 6);
            index += sizeof(byte) * 6;
            return index - startIndex;
        }
    }
}