namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System;

    /// <summary>
    /// This structure represent the input and output method and the value type of the message ID
    /// </summary>
    public struct MessageID : IPropertyInterface
    {
        /// <summary>
        /// The value of this property is the message ID
        /// </summary>
        private ulong messageId;

        /// <summary>
        /// Input the value from rawData to structure
        /// </summary>
        /// <param name="rawData">The byte array returned from the GetLists</param>
        /// <param name="count">The count point to the current digit</param>
        /// <returns>A IPropertyInterface structure contains the value</returns>
        public IPropertyInterface InputValue(byte[] rawData, ref int count)
        {
            MessageID value;
            value.messageId = (ulong)BitConverter.ToInt64(rawData, count);
            count += 8;
            return value;
        }

        /// <summary>
        /// Output the value saved in the IPropertyInterface structure
        /// </summary>
        /// <param name="list">The list from the input</param>
        /// <returns>Certain value of each property</returns>
        public object OutputValue(IPropertyInterface list)
        {
            MessageID fid1 = (MessageID)list;
            ulong value = fid1.messageId;
            return value;
        }
    }
}