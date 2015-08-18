namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    ///  A form of encoding of an internal identifier that makes it globally unique 
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct GID
    {
        /// <summary>
        /// A value that represents a namespace for IDs. 
        /// </summary>
        public Guid DatabaseGuid;

        /// <summary>
        /// An auto-incrementing 6-byte value.
        /// </summary>
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 6)]
        public byte[] GlobalCounter;
    }
}