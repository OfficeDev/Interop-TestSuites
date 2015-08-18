namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// BackoffRop response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct BackoffRop : IDeserializable
    {
        /// <summary>
        /// This value identifies the remote operation for which to apply the ROP backoff.
        /// </summary>
        public byte RopIdBackoff;

        /// <summary>
        /// This value specifies the number of milliseconds to apply a ROP backoff.
        /// </summary>
        public uint Duration;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (BackoffRop)Marshal.PtrToStructure(responseBuffer, typeof(BackoffRop));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }

    /// <summary>
    /// RopBackoffResponse response buffer structure.
    /// </summary>
    public struct RopBackoffResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0xF9.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon to which the RopBackoff response applies.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This value specifies the number of milliseconds to apply a logon backoff.
        /// </summary>
        public uint Duration;

        /// <summary>
        /// This value specifies the number of structures in the BackoffRopData field.
        /// </summary>
        public byte BackoffRopCount;

        /// <summary>
        /// Array of BackoffRop structures. 
        /// The format of the BackoffRop structure is specified in section 2.2.14.2.1.1. 
        /// This array specifies the operations to be backed off and also the backoff duration for each.
        /// </summary>
        public BackoffRop[] BackoffRopData;

        /// <summary>
        /// This value specifies the size of the AdditionalData field.
        /// </summary>
        public ushort AdditionalDataSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the AdditionalDataSize. 
        /// This array specifies additional information about the backoff response.
        /// </summary>
        public byte[] AdditionalData;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.LogonId = ropBytes[index++];

            this.Duration = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);
            this.BackoffRopCount = ropBytes[index++];
            if (this.BackoffRopCount > 0)
            {
                this.BackoffRopData = new BackoffRop[this.BackoffRopCount];
                for (int i = 0; i < this.BackoffRopCount; i++)
                {
                    index += this.BackoffRopData[i].Deserialize(ropBytes, index);
                }
            }

            this.AdditionalDataSize = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            if (this.AdditionalDataSize > 0)
            {
                this.AdditionalData = new byte[this.AdditionalDataSize];
                Array.Copy(ropBytes, index, this.AdditionalData, 0, this.AdditionalDataSize);
                index += this.AdditionalDataSize;
            }

            return index - startIndex;
        }
    }
}