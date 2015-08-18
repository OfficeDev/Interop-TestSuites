namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Action Data buffer format for ActionType: OP_BOUNCE
    /// </summary>
    public class BounceActionData : IActionData
    {
        /// <summary>
        /// Specifies a bounce code
        /// </summary>
        private BounceCode bounce;

        /// <summary>
        /// Gets or sets a bounce code
        /// </summary>
        public BounceCode Bounce
        {
            get { return this.bounce; }
            set { this.bounce = value; }
        }

        /// <summary>
        /// The total Size of this ActionData buffer
        /// </summary>
        /// <returns>Number of bytes in this ActionData buffer.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this ActionData
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            return BitConverter.GetBytes((uint)this.Bounce);
        }

        /// <summary>
        /// Deserialized byte array to a BounceActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contains data of an ActionData instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            this.Bounce = (BounceCode)bufferReader.ReadUInt32();
            return bufferReader.Position;
        }
    }
}