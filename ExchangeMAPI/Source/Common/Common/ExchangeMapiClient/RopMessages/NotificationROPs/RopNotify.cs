namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopNotify response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopNotifyResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the Type of remote operation. For this operation, this field is set to 0x2A.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This handle specifies the notification server object associated with this notification event.
        /// </summary>
        public uint NotificationHandle;

        /// <summary>
        /// This value specifies the logon associated with this notification event.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Various structures. The notification structures that can be found here are specified in [MS-OXCDATA].
        /// </summary>
        public NotificationData NotificationData;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer struct.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.NotificationHandle = (uint)BitConverter.ToInt32(ropBytes, index);
            index += 4;
            this.LogonId = ropBytes[index++];
            this.NotificationData = new NotificationData();
            index += this.NotificationData.Deserialize(ropBytes, index);

            return index - startIndex;
        }
    }
}