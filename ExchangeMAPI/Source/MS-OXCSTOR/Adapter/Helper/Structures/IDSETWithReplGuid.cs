namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Command type
    /// </summary>
    public enum CommandType : byte
    {
        /// <summary>
        /// Push Command
        /// </summary>
        PushCommand = 0x06,

        /// <summary>
        /// Pop Command
        /// </summary>
        PopCommand = 0x50,

        /// <summary>
        /// BitMask Command
        /// </summary>
        BitMaskCommand = 0x42,

        /// <summary>
        /// Rang Command
        /// </summary>
        RangCommand = 0x52,

        /// <summary>
        /// End Command
        /// </summary>
        EndCommand = 0x00,

        /// <summary>
        /// Not Supported
        /// </summary>
        NotSupported,
    }

    /// <summary>
    /// RopLogon request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public class IDSETWithReplGuid
    {
        /// <summary>
        /// Const variable ReplGuidLength
        /// </summary>
        private const int ReplGuidLength = 16;

        /// <summary>
        /// A 128-bit unsigned integer identifying a Store object (DatabaseGuid)
        /// </summary>
        private byte[] replGuid;

        /// <summary>
        /// A GLOBSET is a set of GLOBCNT values that are typically reduced to GLOBCNT ranges.
        /// </summary>
        private Globset globset;

        /// <summary>
        /// Gets or sets the ReplGuid
        /// </summary>
        public byte[] ReplGuid
        {
            get
            {
                return this.replGuid;
            }

            set
            {
                this.replGuid = value;
            }
        }

        /// <summary>
        /// Gets or sets the globset
        /// </summary>
        public Globset Globset
        {
            get
            {
                return this.globset;
            }

            set
            {
                this.globset = value;
            }
        }

        /// <summary>
        /// Serialize this structure.
        /// </summary>
        /// <returns>The serialized XID structure specified in [MS-OXCFXICS].</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] resultBytes = new byte[this.Size()];
            Array.Copy(this.ReplGuid, 0, resultBytes, index, ReplGuidLength);
            index += ReplGuidLength;
            Array.Copy(this.Globset.GetBytes(), 0, resultBytes, index, this.Globset.GetBytes().Length);
            return resultBytes;
        }

        /// <summary>
        /// Deserialize the stream. 
        /// </summary>
        /// <param name="streamByte">data array in byte</param>
        /// <param name="startIndex">The startIndex in the data array</param>
        /// <returns>The size of this structure</returns>
        public int Deserialize(byte[] streamByte, int startIndex)
        {
            int index = startIndex;

            this.ReplGuid = new byte[ReplGuidLength];
            Array.Copy(streamByte, index, this.ReplGuid, 0, ReplGuidLength);
            index += this.ReplGuid.Length;
            byte[] globsetBytes = new byte[streamByte.Length - ReplGuidLength];
            Array.Copy(streamByte, index, globsetBytes, 0, globsetBytes.Length);
            this.Globset = new Globset(globsetBytes);
            return this.Size();
        }

        /// <summary>
        /// Get the size of the IDSETWithReplGuid
        /// </summary>
        /// <returns>The size of the IDSETWithReplGuid</returns>
        public int Size()
        {
            return ReplGuidLength + this.Globset.GetBytes().Length;
        }
    }
}